"""
title: Company Context
author: <AUTHOR NAME>
author_url: <URL>
version: 1.0.0
license: MIT
required_open_webui_version: 0.11.4
description: Adds a short system note to every request with your company name, today's date, the response language, the form of address, the user's name and your own house rules - the same for every model and assistant, including newly created ones.

How it works
- Runs on the "inlet" hook. If the request already has a system message, the note is put in
  front of it (the existing system prompt is kept); otherwise a new system message is added.
- The note starts with a marker line, so it is never added twice to the same request.
- The note itself is written in English or German (language valve); the date follows that
  language, e.g. "Monday, 12 October 2026" or "Montag, 12. Oktober 2026".
- The note contains the date but not the time: it opens the prompt, so it has to stay the same
  all day for the model server to reuse its prompt cache (llama-server, vLLM, Ollama) instead
  of processing the whole chat again on every turn.

Security & privacy
- No network access, nothing is stored. Only the system message is changed; user messages
  are left untouched.
- If include_user_name is on, the user's display name (as stored in Open WebUI) is sent to
  the model as part of the system note.
"""

import logging
from datetime import datetime
from typing import Any, Literal
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

MAX_USER_NAME_CHARS = 100

TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "marker": "[Company context]",
        "company": "You are the internal AI assistant of {company}.",
        "date": "Today is {weekday}, {day} {month} {year}.",
        "response_language": "English",
        "language_rule": (
            "Always answer in {language} unless the user writes in another language "
            "or explicitly asks for a different one."
        ),
        "formal": (
            "Address the user formally. Texts you write for them (e-mails, letters) should "
            "also use a formal tone unless they ask otherwise."
        ),
        "informal": (
            "Address the user in a friendly, informal tone. Texts you write for them (e-mails, "
            "letters) should also be informal unless they ask otherwise."
        ),
        "user_name": "The user's name is {name}.",
    },
    "de": {
        "marker": "[Firmen-Kontext]",
        "company": "Du bist der interne KI-Assistent der Firma {company}.",
        "date": "Heute ist {weekday}, {day}. {month} {year}.",
        "response_language": "Deutsch",
        "language_rule": (
            "Antworte immer auf {language}, außer der Nutzer schreibt in einer anderen Sprache "
            "oder bittet ausdrücklich um eine andere Sprache."
        ),
        "formal": (
            "Sprich den Nutzer mit „Sie“ an. Texte, die du für ihn verfasst (E-Mails, Briefe), "
            "schreibst du ebenfalls in der Sie-Form, sofern er nichts anderes wünscht."
        ),
        "informal": (
            "Sprich den Nutzer mit „du“ an. Texte, die du für ihn verfasst (E-Mails, Briefe), "
            "schreibst du ebenfalls in der Du-Form, sofern er nichts anderes wünscht."
        ),
        "user_name": "Der Nutzer heißt {name}.",
    },
}
WEEKDAYS = {
    "en": ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"),
    "de": ("Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"),
}
MONTHS = {
    "en": (
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ),
    "de": (
        "Januar", "Februar", "März", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Dezember",
    ),
}  # fmt: skip
MARKERS = tuple(texts["marker"] for texts in TRANSLATIONS.values())


def has_marker(content: Any) -> bool:
    """True if a system message content (string or list of parts) already contains the note."""
    if isinstance(content, str):
        return any(marker in content for marker in MARKERS)
    if isinstance(content, list):
        return any(isinstance(part, dict) and has_marker(part.get("text")) for part in content)
    return False


def prepend_note(note: str, content: Any) -> str | list[Any]:
    """Put the note in front of an existing system message content."""
    if has_marker(content):
        return content
    if isinstance(content, list):
        return [{"type": "text", "text": note}, *content]
    if isinstance(content, str) and content.strip():
        return f"{note}\n\n{content}"
    return note


class Filter:
    class Valves(BaseModel):
        priority: int = Field(default=0, description="Filter order (lower runs first).")
        language: Literal["en", "de"] = Field(
            default="en", description="Language of the system note and of the date format."
        )
        company_name: str = Field(
            default="", description="Your company name. Empty = no company line in the note."
        )
        response_language: str = Field(
            default="",
            description="Language the assistant should answer in, written in the note's language "
            "(e.g. 'English', 'Deutsch'). Empty = the language of the note.",
        )
        form_of_address: Literal["formal", "informal"] = Field(
            default="formal",
            description="How to address the user and write texts for them (German: Sie / du).",
        )
        timezone: str = Field(
            default="UTC", description="IANA time zone that determines today's date, e.g. 'Europe/Berlin'."
        )
        include_user_name: bool = Field(
            default=True, description="Tell the model the user's name (e.g. for e-mail signatures)."
        )
        extra_rules: str = Field(
            default="",
            description="Additional house rules appended to the note, e.g. 'Always write to customers "
            "in the first person plural.'",
        )

    def __init__(self) -> None:
        self.valves = self.Valves()

    def _now(self) -> datetime:
        try:
            return datetime.now(ZoneInfo(self.valves.timezone))
        except (ZoneInfoNotFoundError, ValueError):
            log.warning("Company Context: unknown time zone %r, using server time", self.valves.timezone)
            return datetime.now()

    def build_note(self, user: dict[str, Any] | None, now: datetime) -> str:
        """Compose the system note for the given user and day."""
        language = self.valves.language
        texts = TRANSLATIONS[language]
        lines = [texts["marker"]]
        if company := self.valves.company_name.strip():
            lines.append(texts["company"].format(company=company))
        lines.append(
            texts["date"].format(
                weekday=WEEKDAYS[language][now.weekday()],
                day=now.day,
                month=MONTHS[language][now.month - 1],
                year=now.year,
            )
        )
        response_language = self.valves.response_language.strip() or texts["response_language"]
        lines.append(texts["language_rule"].format(language=response_language))
        lines.append(texts[self.valves.form_of_address])
        if self.valves.include_user_name and (
            user_name := " ".join(str((user or {}).get("name") or "").split())[:MAX_USER_NAME_CHARS]
        ):
            lines.append(texts["user_name"].format(name=user_name))
        if extra_rules := self.valves.extra_rules.strip():
            lines.append(extra_rules)
        return "\n".join(lines)

    async def inlet(self, body: dict[str, Any], __user__: dict[str, Any] | None = None) -> dict[str, Any]:
        """Prepend the note to the system message (or add one); idempotent per request."""
        messages = body.get("messages") or []
        note = self.build_note(__user__, self._now())
        if messages and isinstance(messages[0], dict) and messages[0].get("role") == "system":
            messages[0]["content"] = prepend_note(note, messages[0].get("content"))
        else:
            messages.insert(0, {"role": "system", "content": note})
        body["messages"] = messages
        return body
