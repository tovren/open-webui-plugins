"""
title: Context Guard
author: tovren
author_url: https://github.com/tovren/open-webui-plugins
version: 1.3.0
license: MIT
required_open_webui_version: 0.11.4
description: Estimates the prompt size right before each request is sent to the model - after document excerpts and tool results have been added - and shows a warning or stops with a clear, friendly message before the model's context window overflows.

How it works
- Runs on the "request" hook (new in Open WebUI 0.11), i.e. after file/RAG context has been
  injected and again after every tool-call round. The model's system prompt, which Open WebUI only
  adds after this hook, and earlier reasoning sent back to the model are counted too.
- Shows a status hint above the answer once the prompt uses warn_percent of the usable context
  (context size minus the space reserved for the answer).
- Stops the request with an explanation and practical next steps when the prompt would not fit.
  If the user's own message makes up most of it (a long pasted text), the message says so. Without
  this, llama-server and vLLM reject the request with a technical HTTP 400 error
  ("exceeds the available context size"), and Ollama silently drops the oldest part of the
  conversation, so the model quietly forgets earlier messages and documents.
- The token count is an estimate (characters / chars_per_token plus a flat amount per image),
  not a real tokenizer. Lower chars_per_token to be more cautious. Open WebUI's own estimate
  (Context Compaction) uses 4 characters per token, so the guard's default of 3.5 counts higher.
- Works together with Open WebUI's Context Compaction: compaction runs first, so the guard checks
  the shortened history. While compaction will summarize the chat before the limit, the hint to
  start a new chat is not shown. The summary request contains the whole chat, so the compaction
  threshold must leave room for it (COMPACTION_RESERVE_TOKENS below the usable context); otherwise a
  warning is logged once and the hint is shown as usual. If compaction could not get a summary from
  the model and fell back to cutting the old messages short, a warning is logged once per chat.

Security & privacy
- Read-only: the request is never modified. No network access, nothing is stored, and no
  message content is logged. Open WebUI's Context Compaction settings are only read.
"""

import json
import logging
import math
import re
from collections.abc import Awaitable, Callable
from functools import lru_cache
from typing import Any, Literal

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

EventEmitter = Callable[[dict[str, Any]], Awaitable[Any]]

MESSAGE_OVERHEAD_CHARS = 16  # role and chat-template tokens per message
REASONING_FIELDS = ("reasoning_content", "reasoning", "thinking")
# Context Compaction: room the summary request needs besides the chat itself (prompt template
# about 300 tokens, the summary up to Open WebUI's default max_tokens of 1000, plus a margin)
COMPACTION_RESERVE_TOKENS = 1500
SUMMARY_MARKER = "[CONVERSATION SUMMARY]\n"
# Open WebUI's fallback when the summary request fails: "- role: first 500 characters" per message
FALLBACK_SUMMARY_LINE = re.compile(r"^- (?:user|assistant|system|tool): ", re.MULTILINE)

# The error box of Open WebUI (0.11, 0.12) shows plain text without line breaks, so the messages are
# written as flowing text.
TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "blocked": (
            "This request is too large for the AI model (about {percent}% of what it can handle): "
            "the chat history, documents and your message no longer fit into its working memory "
            "together. Please start a new chat, which is a good idea for a new topic anyway. Paste only "
            "the relevant section of long texts, or upload them as a file or add them to a knowledge "
            "base, and attach fewer documents at once."
        ),
        "message_too_long": (
            "Your message is too long for the AI model (together with the chat about {percent}% of "
            "what it can handle). Please paste only the relevant section, or upload the text as a "
            "file or add it to a knowledge base instead."
        ),
        "warning": (
            "Note: this chat is getting very long (about {percent}% of the capacity). "
            "Please start a new chat for a new topic."
        ),
    },
    "de": {
        "blocked": (
            "Diese Anfrage ist zu umfangreich für das KI-Modell (ca. {percent} % der möglichen Menge): "
            "Chatverlauf, Dokumente und Ihre Nachricht passen nicht mehr zusammen in sein "
            "Arbeitsgedächtnis. Bitte einen neuen Chat beginnen, bei einem neuen Thema ohnehin am "
            "besten. Lange Texte nur mit dem wichtigen Abschnitt einfügen oder als Datei hochladen bzw. "
            "in eine Wissensdatenbank legen und weniger Dokumente gleichzeitig anhängen."
        ),
        "message_too_long": (
            "Ihre Nachricht ist zu lang für das KI-Modell (zusammen mit dem Chat ca. {percent} % der "
            "möglichen Menge). Bitte nur den wichtigen Abschnitt einfügen oder den Text als Datei "
            "hochladen bzw. in eine Wissensdatenbank legen."
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


def _text(content: Any) -> str:
    if isinstance(content, list):
        return "\n".join(str(part.get("text") or "") for part in content if isinstance(part, dict))
    return content if isinstance(content, str) else ""


def missing_system_prompt_chars(messages: list[Any], metadata: dict[str, Any] | None) -> int:
    """Characters of the model's system prompt that Open WebUI adds only after the request hook.

    Before the hook Open WebUI stores "model prompt\nsystem message" in the metadata (0.11:
    system_prompt, 0.12: base_system_prompt); the router puts the model prompt in front of the
    system message later. In tool-call rounds the body already contains it."""
    metadata = metadata or {}
    prompt = metadata.get("base_system_prompt") or metadata.get("system_prompt")
    if not isinstance(prompt, str) or not prompt:
        return 0
    system = next((m for m in messages if isinstance(m, dict) and m.get("role") == "system"), None)
    current = _text(system.get("content")) if system else ""
    if current.startswith(prompt):
        return 0
    # the system message starts with the part after one of the prompt's line breaks (RAG context
    # added); a part longer than the system message cannot match, so the search starts after them
    index = prompt.find("\n", max(0, len(prompt) - len(current) - 1))
    while index >= 0:
        if current.startswith(prompt[index + 1 :]):
            return index + 1
        index = prompt.find("\n", index + 1)
    return len(prompt) + 1


def estimate_tokens(
    body: dict[str, Any],
    chars_per_token: float,
    tokens_per_image: int,
    metadata: dict[str, Any] | None = None,
) -> int:
    """Rough token estimate for messages, reasoning, tool calls, tool definitions and the model's
    system prompt of a chat request."""
    messages = body.get("messages") or []
    chars, images = missing_system_prompt_chars(messages, metadata), 0
    for message in messages:
        if not isinstance(message, dict):
            continue
        message_chars, message_images = _content_size(message.get("content"))
        chars += message_chars + MESSAGE_OVERHEAD_CHARS
        images += message_images
        # earlier reasoning sent back: llama.cpp reasoning_content, Ollama thinking
        chars += sum(len(value) for key in REASONING_FIELDS if isinstance(value := message.get(key), str))
        chars += sum(_json_length(call) for call in message.get("tool_calls") or [])
    if body.get("tools"):
        chars += _json_length(body["tools"])
    return math.ceil(chars / chars_per_token) + images * tokens_per_image


def typed_message_chars(metadata: dict[str, Any] | None) -> int:
    """Characters of the message the user typed, before Open WebUI added document excerpts to it
    (0.12: base_user_prompt, 0.11: user_message); 0 if unknown."""
    metadata = metadata or {}
    user_message = metadata.get("user_message")
    content = metadata.get("base_user_prompt") or (
        user_message.get("content") if isinstance(user_message, dict) else None
    )
    return _content_size(content)[0]


def has_fallback_summary(messages: list[Any]) -> bool:
    """True if Context Compaction could not get a summary from the model: Open WebUI then stores the
    first 500 characters of each dropped message ("- user: ...") as the summary."""
    for message in messages:
        if isinstance(message, dict) and message.get("role") == "system":
            _, marker, summary = _text(message.get("content")).partition(SUMMARY_MARKER)
            if marker and FALLBACK_SUMMARY_LINE.search(summary):
                return True
    return False


def model_ids(body: dict[str, Any], model: dict[str, Any] | None) -> list[str]:
    """The requested model ID, then the base model of a workspace model (if any)."""
    info = (model or {}).get("info") or {}
    ids = [body.get("model"), (model or {}).get("id"), info.get("base_model_id")]
    return [model_id for model_id in dict.fromkeys(ids) if isinstance(model_id, str) and model_id]


def user_language(language: str, metadata: dict[str, Any] | None) -> str:
    """The language valve; 'auto' takes the user's interface language, English if it has no translation."""
    if language != "auto":
        return language
    variables = (metadata or {}).get("variables") or {}
    locale = str(variables.get("{{USER_LANGUAGE}}") or "") if isinstance(variables, dict) else ""
    code = locale.replace("_", "-").split("-")[0].strip().lower()
    return code if code in TRANSLATIONS else "en"


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


def _positive_int(value: Any) -> int | None:
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return None
    return parsed if parsed > 0 else None


async def compaction_threshold(metadata: dict[str, Any] | None) -> int | None:
    """Token threshold of Open WebUI's Context Compaction for this request, None if it is off.

    Resolved like open_webui.utils.context_compaction (0.11.4): the model's compact_token_threshold
    or the global threshold (default 80000), capped by the token cap."""
    try:
        from open_webui.models.config import Config

        values = await Config.get_many(
            "chat.context_compaction.enable",
            "chat.context_compaction.token_threshold",
            "chat.context_compaction.token_cap",
        )
    except Exception:
        log.debug("Context Guard: Context Compaction settings not available", exc_info=True)
        return None
    if not values.get("chat.context_compaction.enable"):
        return None
    threshold = _positive_int(values.get("chat.context_compaction.token_threshold")) or 80000
    cap = _positive_int(values.get("chat.context_compaction.token_cap")) or threshold
    model_threshold = _positive_int(((metadata or {}).get("params") or {}).get("compact_token_threshold"))
    return min(model_threshold or threshold, cap)


class Filter:
    class Valves(BaseModel):
        priority: int = Field(
            default=100,
            description="Filter order (lower runs first). Only matters among the request hooks of several "
            "filters: keep it high so that changes of other request hooks are counted. Inlet hooks (e.g. "
            "Company Context) always run before and are counted at any priority.",
        )
        language: Literal["auto", "en", "de"] = Field(
            default="auto",
            description="Language of the warning and error messages. auto = the user's interface language "
            "(English if there is no translation).",
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
            "e.g. 'qwen3-32b=40960, llama3.1:8b=8192'. A workspace model uses the size of its base model "
            "unless it has its own entry. Other models use context_tokens.",
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
        self._reported_thresholds: set[tuple[str, int, int]] = set()
        self._reported_fallbacks: set[str] = set()

    def _context_tokens(self, ids: list[str]) -> int:
        overrides = parse_model_overrides(self.valves.model_context_tokens)
        return next(
            (overrides[model_id] for model_id in ids if model_id in overrides), self.valves.context_tokens
        )

    def _report_threshold_too_high(self, model_id: str, threshold: int, usable: int) -> None:
        if (model_id, threshold, usable) in self._reported_thresholds:
            return
        self._reported_thresholds.add((model_id, threshold, usable))
        log.warning(
            "Context Guard: the Context Compaction threshold for model %r (%d tokens) leaves too little "
            "room: the summary request contains the whole chat plus about %d tokens and must fit into "
            "the usable context (%d tokens). Otherwise long chats are stopped, or compacted without a "
            "real summary. Set CONTEXT_COMPACTION_TOKEN_THRESHOLD or the model's compact_token_threshold "
            "to about half the context size.",
            model_id,
            threshold,
            COMPACTION_RESERVE_TOKENS,
            usable,
        )

    def _report_fallback_summary(self, chat_id: str) -> None:
        if chat_id in self._reported_fallbacks:
            return
        if len(self._reported_fallbacks) >= 1000:
            self._reported_fallbacks.clear()
        self._reported_fallbacks.add(chat_id)
        log.warning(
            "Context Guard: Context Compaction could not get a summary from the model for chat %s (the "
            "summary request probably did not fit into the context or failed), so the older messages "
            "were only cut short. Lower the compaction threshold or check the compaction model.",
            chat_id or "(unknown)",
        )

    async def request(
        self,
        body: dict[str, Any],
        __event_emitter__: EventEmitter | None = None,
        __metadata__: dict[str, Any] | None = None,
        __model__: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Check the final payload; warn via status event or abort with ContextLimitError."""
        model_id = body.get("model") or ""
        usable = self._context_tokens(model_ids(body, __model__)) - self.valves.answer_reserve_tokens
        if usable <= 0:
            log.warning("Context Guard: answer_reserve_tokens >= context size, check is skipped")
            return body
        used = estimate_tokens(body, self.valves.chars_per_token, self.valves.tokens_per_image, __metadata__)
        fallback_summary = has_fallback_summary(body.get("messages") or [])
        if fallback_summary:
            self._report_fallback_summary(str((__metadata__ or {}).get("chat_id") or ""))
        if used * 100 < self.valves.warn_percent * usable:
            return body
        threshold = await compaction_threshold(__metadata__)
        compaction_fits = threshold is not None and threshold + COMPACTION_RESERVE_TOKENS <= usable
        if threshold is not None and not compaction_fits:
            self._report_threshold_too_high(model_id, threshold, usable)
        percent = math.ceil(used * 100 / usable)
        texts = TRANSLATIONS[user_language(self.valves.language, __metadata__)]
        if used > usable:
            typed = math.ceil(typed_message_chars(__metadata__) / self.valves.chars_per_token)
            # most of the prompt is the user's own message, e.g. a long pasted text
            key = "message_too_long" if typed * 2 >= used else "blocked"
            raise ContextLimitError(texts[key].format(percent=percent))
        # Normal chat growth: compaction summarizes the chat before it reaches the limit.
        if compaction_fits and not fallback_summary and used < threshold:
            return body
        if __event_emitter__ is not None:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {"description": texts["warning"].format(percent=percent), "done": True},
                }
            )
        return body
