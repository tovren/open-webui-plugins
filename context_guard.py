"""
title: Context Guard
author: <AUTHOR NAME>
author_url: <URL>
version: 1.0.0
license: MIT
required_open_webui_version: 0.11.4
description: Estimates the prompt size right before each request is sent to the model - after document excerpts and tool results have been added - and shows a warning or stops with a clear, friendly message before the model's context window overflows.

How it works
- Runs on the "request" hook (new in Open WebUI 0.11), i.e. after file/RAG context has been
  injected and again after every tool-call round, so everything the model will see is counted.
- Shows a status hint above the answer once the prompt uses warn_percent of the usable context
  (context size minus the space reserved for the answer).
- Stops the request with an explanation and practical next steps when the prompt would not fit.
  Without this, llama-server and vLLM reject the request with a technical HTTP 400 error
  ("exceeds the available context size"), and Ollama silently drops the oldest part of the
  conversation, so the model quietly forgets earlier messages and documents.
- The token count is an estimate (characters / chars_per_token plus a flat amount per image),
  not a real tokenizer. Lower chars_per_token to be more cautious.

Security & privacy
- Read-only: the request is never modified. No network access, nothing is stored, and no
  message content is logged.
"""

import json
import logging
import math
from collections.abc import Awaitable, Callable
from functools import lru_cache
from typing import Any, Literal

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

EventEmitter = Callable[[dict[str, Any]], Awaitable[Any]]

MESSAGE_OVERHEAD_CHARS = 16  # role and chat-template tokens per message

TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "blocked": (
            "This request is too large for the AI model (about {percent}% of what it can handle). "
            "The chat history, documents and your message no longer fit into its working memory "
            "together.\n\n"
            "How to continue:\n"
            "• Start a new chat - always a good idea for a new topic.\n"
            "• Paste only the relevant section of long texts, or upload them as a file or add them "
            "to a knowledge base instead.\n"
            "• Attach fewer documents at once."
        ),
        "warning": (
            "Note: this chat is getting very long (about {percent}% of the capacity). "
            "Please start a new chat for a new topic."
        ),
    },
    "de": {
        "blocked": (
            "Diese Anfrage ist zu umfangreich für das KI-Modell (ca. {percent} % der möglichen Menge). "
            "Chatverlauf, Dokumente und Ihre Nachricht passen nicht mehr zusammen in sein "
            "Arbeitsgedächtnis.\n\n"
            "So geht es weiter:\n"
            "• Einen neuen Chat beginnen – bei einem neuen Thema am besten immer.\n"
            "• Lange Texte nicht komplett einfügen, sondern nur den wichtigen Abschnitt – oder als "
            "Datei hochladen bzw. in eine Wissensdatenbank legen.\n"
            "• Weniger Dokumente gleichzeitig anhängen."
        ),
        "warning": (
            "Hinweis: Dieser Chat ist schon sehr lang (ca. {percent} % der Kapazität). "
            "Für ein neues Thema bitte einen neuen Chat beginnen."
        ),
    },
}


class ContextLimitError(Exception):
    """Raised to abort a request; Open WebUI shows the message to the user."""


def _json_length(value: Any) -> int:
    return len(json.dumps(value, ensure_ascii=False, default=str))


def _content_size(content: Any) -> tuple[int, int]:
    """Return (characters, images) of a message content string or list of content parts."""
    if isinstance(content, str):
        return len(content), 0
    if not isinstance(content, list):
        return 0, 0
    chars = images = 0
    for part in content:
        if not isinstance(part, dict):
            continue
        if part.get("type") == "text":
            chars += len(part.get("text") or "")
        elif part.get("type") == "image_url":
            images += 1
    return chars, images


def estimate_tokens(body: dict[str, Any], chars_per_token: float, tokens_per_image: int) -> int:
    """Rough token estimate for messages, tool calls and tool definitions of a chat request."""
    chars = images = 0
    for message in body.get("messages") or []:
        if not isinstance(message, dict):
            continue
        message_chars, message_images = _content_size(message.get("content"))
        chars += message_chars + MESSAGE_OVERHEAD_CHARS
        images += message_images
        chars += sum(_json_length(call) for call in message.get("tool_calls") or [])
    if body.get("tools"):
        chars += _json_length(body["tools"])
    return math.ceil(chars / chars_per_token) + images * tokens_per_image


@lru_cache(maxsize=4)
def parse_model_overrides(text: str) -> dict[str, int]:
    """Parse 'model-a=8192, model-b=32768' into a dict; invalid entries are logged and skipped.

    Cached per valve value, so the text is parsed (and warned about) once, not on every request.
    The returned dict is shared and must not be modified."""
    overrides: dict[str, int] = {}
    for entry in filter(None, (item.strip() for item in text.split(","))):
        model, _, tokens = entry.rpartition("=")
        if model.strip() and tokens.strip().isdigit():
            overrides[model.strip()] = int(tokens)
        else:
            log.warning("Context Guard: ignoring invalid model_context_tokens entry %r", entry)
    return overrides


class Filter:
    class Valves(BaseModel):
        priority: int = Field(
            default=100,
            description="Filter order (lower runs first). Keep it high so that content added by other "
            "filters is counted too.",
        )
        language: Literal["en", "de"] = Field(
            default="en", description="Language of the warning and error messages."
        )
        context_tokens: int = Field(
            default=32768,
            ge=1,
            description="Context size of the model in tokens (llama-server --ctx-size, vLLM --max-model-len, "
            "Ollama num_ctx).",
        )
        model_context_tokens: str = Field(
            default="",
            description="Optional per-model context sizes as comma-separated model_id=tokens pairs, "
            "e.g. 'qwen3-32b=40960, llama3.1:8b=8192'. Other models use context_tokens.",
        )
        answer_reserve_tokens: int = Field(
            default=6000,
            ge=0,
            description="Tokens kept free for the answer, including the model's reasoning.",
        )
        chars_per_token: float = Field(
            default=3.5,
            ge=1.0,
            description="Characters per token for the estimate (English about 4, German about 3.5 to 4.3; "
            "lower = more cautious).",
        )
        tokens_per_image: int = Field(default=300, ge=0, description="Estimated tokens per attached image.")
        warn_percent: int = Field(
            default=75,
            ge=1,
            le=100,
            description="Show a hint once the prompt uses this share of the usable context (in percent).",
        )

    def __init__(self) -> None:
        self.valves = self.Valves()

    def _context_tokens(self, model_id: str | None) -> int:
        overrides = parse_model_overrides(self.valves.model_context_tokens)
        return overrides.get(model_id or "", self.valves.context_tokens)

    async def request(
        self, body: dict[str, Any], __event_emitter__: EventEmitter | None = None
    ) -> dict[str, Any]:
        """Check the final payload; warn via status event or abort with ContextLimitError."""
        usable = self._context_tokens(body.get("model")) - self.valves.answer_reserve_tokens
        if usable <= 0:
            log.warning("Context Guard: answer_reserve_tokens >= context size, check is skipped")
            return body
        used = estimate_tokens(body, self.valves.chars_per_token, self.valves.tokens_per_image)
        percent = math.ceil(used * 100 / usable)
        texts = TRANSLATIONS[self.valves.language]
        if used > usable:
            raise ContextLimitError(texts["blocked"].format(percent=percent))
        if used * 100 >= self.valves.warn_percent * usable and __event_emitter__ is not None:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {"description": texts["warning"].format(percent=percent), "done": True},
                }
            )
        return body
