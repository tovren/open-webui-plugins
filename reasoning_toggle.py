"""
title: Reasoning Toggle
author: <AUTHOR NAME>
author_url: <URL>
version: 1.0.0
license: MIT
required_open_webui_version: 0.11.4
description: Adds a lightbulb button to the chat input that lets the model think before it answers - for local reasoning models on vLLM or llama-server (Qwen3, Gemma 4, gpt-oss and others), with an optional effort level per model and a token budget for the reasoning.

How it works
- A toggle filter: users switch it on in the Integrations menu (the button next to "+" in the
  chat input). While it is on, a lightbulb shows in the chat input; clicking it switches it off.
  While it is off, the filter does not run at all and the model behaves as configured, so make
  "no thinking" the model's default where needed (see README).
- Works without the Chat Controls side panel: unlike Open WebUI's Model Controls (0.12+), a
  toggle filter needs neither the "Allow Chat Controls" nor the "Allow Chat Params" permission.
- When the button is on, the request gets chat_template_kwargs.enable_thinking = true. Both
  vLLM and llama-server hand this to the chat template (Qwen3.x, Gemma 4, GLM and others).
- Optional effort level per model, sent as reasoning_effort, e.g. low/medium/xhigh for Qwen3.8
  or low/medium/high for gpt-oss. Use only levels the model's chat template supports: Qwen3.8
  rejects "high", for example, and the request fails.
- Optional thinking budget, a hard cap on reasoning tokens. It is sent as thinking_budget_tokens
  (llama-server) and as thinking_token_budget (vLLM, which then needs --reasoning-parser).
- Meant for local model servers. Cloud APIs reject unknown parameters such as
  chat_template_kwargs, so enable the filter only for local reasoning models.

Security & privacy
- No network access, nothing is stored or logged. Only the request parameters named above are
  set; messages are left untouched.
"""

import base64
import logging
from collections.abc import Iterator
from functools import lru_cache
from typing import Any

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

# Values vLLM accepts for reasoning_effort ("none" is left out: the button means "think").
EFFORT_LEVELS = ("minimal", "low", "medium", "high", "xhigh", "max")

ICON_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#000" '
    'stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M9.5 21h5M9 18h6M12 3a6 6 0 0 0-3.6 10.8c.7.5 1.1 1.3 1.1 2.1V16h5v-.1c0-.8.4-1.6 '
    '1.1-2.1A6 6 0 0 0 12 3z"/></svg>'
)
ICON = "data:image/svg+xml;base64," + base64.b64encode(ICON_SVG.encode()).decode()


def _pairs(text: str) -> Iterator[tuple[str, str, str]]:
    """Yield (entry, model_id, value) for each 'model_id=value' entry of a comma-separated list."""
    for entry in filter(None, (item.strip() for item in text.split(","))):
        model, _, value = entry.rpartition("=")
        yield entry, model.strip(), value.strip()


@lru_cache(maxsize=4)
def parse_efforts(text: str) -> dict[str, str]:
    """Parse 'model-a=low, model-b=xhigh'; invalid entries are logged once and skipped.

    The returned dict is cached per valve value and must not be modified."""
    efforts: dict[str, str] = {}
    for entry, model, value in _pairs(text):
        if model and value.lower() in EFFORT_LEVELS:
            efforts[model] = value.lower()
        else:
            log.warning("Reasoning Toggle: ignoring invalid model_reasoning_effort entry %r", entry)
    return efforts


@lru_cache(maxsize=4)
def parse_budgets(text: str) -> dict[str, int]:
    """Parse 'model-a=2000, model-b=8000'; invalid entries are logged once and skipped.

    The returned dict is cached per valve value and must not be modified."""
    budgets: dict[str, int] = {}
    for entry, model, value in _pairs(text):
        if model and value.isdigit():
            budgets[model] = int(value)
        else:
            log.warning("Reasoning Toggle: ignoring invalid model_thinking_budget entry %r", entry)
    return budgets


def model_ids(body: dict[str, Any], model: dict[str, Any] | None) -> list[str]:
    """The requested model ID, then the base model of a workspace model (if any)."""
    info = (model or {}).get("info") or {}
    ids = [body.get("model"), (model or {}).get("id"), info.get("base_model_id")]
    return [model_id for model_id in dict.fromkeys(ids) if isinstance(model_id, str) and model_id]


class Filter:
    class Valves(BaseModel):
        priority: int = Field(default=0, description="Filter order (lower runs first).")
        model_reasoning_effort: str = Field(
            default="",
            description="Optional effort level per model as comma-separated model_id=level pairs, "
            "e.g. 'qwen3.8-27b=medium, gpt-oss-120b=high'. Use only levels the model supports "
            "(Qwen3.8: low, medium, xhigh). Other models use their chat template's default.",
        )
        thinking_budget: int = Field(
            default=0,
            ge=0,
            description="Maximum reasoning tokens per answer, 0 = no limit. vLLM needs "
            "--reasoning-parser for this, otherwise it rejects the request.",
        )
        model_thinking_budget: str = Field(
            default="",
            description="Optional per-model budgets as comma-separated model_id=tokens pairs, "
            "e.g. 'gemma-4-26b=4000'. Other models use thinking_budget.",
        )

    def __init__(self) -> None:
        self.valves = self.Valves()
        self.toggle = True
        self.icon = ICON

    async def inlet(self, body: dict[str, Any], __model__: dict[str, Any] | None = None) -> dict[str, Any]:
        """Turn thinking on for this request, with the configured effort level and budget."""
        ids = model_ids(body, __model__)
        kwargs = body.get("chat_template_kwargs")
        body["chat_template_kwargs"] = {
            **(kwargs if isinstance(kwargs, dict) else {}),
            "enable_thinking": True,
        }

        efforts = parse_efforts(self.valves.model_reasoning_effort)
        if effort := next((efforts[model_id] for model_id in ids if model_id in efforts), None):
            body["reasoning_effort"] = effort
        elif body.get("reasoning_effort") == "none":
            # A model default of "none" keeps thinking off while the button is off. vLLM would still
            # pass "none" to the chat template, where Qwen3.x rejects it; null means template default.
            body["reasoning_effort"] = None

        budgets = parse_budgets(self.valves.model_thinking_budget)
        budget = next(
            (budgets[model_id] for model_id in ids if model_id in budgets), self.valves.thinking_budget
        )
        if budget:
            body["thinking_budget_tokens"] = budget  # llama-server
            body["thinking_token_budget"] = budget  # vLLM
        return body
