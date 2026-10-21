"""
title: Company Context
author: tovren
author_url: https://github.com/tovren/open-webui-plugins
version: 1.1.0
license: MIT
required_open_webui_version: 0.11.4
description: Adds a short system note to every request with your company name, today's date and calendar week, the response language, the form of address, the user's name and your own house rules - the same for every model and assistant, including newly created ones.

How it works
- Runs on the "inlet" hook. If the request already has a system message (the chat's or the
  user's system prompt), the note is put in front of it; otherwise a new system message is added.
  Open WebUI adds the model's own system prompt later, in front of the note.
- The note starts with a marker line, so it is never added twice to the same request.
- The note itself is written in English or German (language valve, by default the user's interface
  language); the date follows that language, e.g. "Monday, 12 October 2026 (calendar week 42)" or
  "Montag, 12. Oktober 2026 (Kalenderwoche 42)".
- The date is taken in the time zone of the timezone valve, by default in the user's browser time
  zone.
- The note contains the date but not the time: it opens the prompt, so it has to stay the same
  all day for the model server to reuse its prompt cache (llama-server, vLLM, Ollama) instead
  of processing the whole chat again on every turn.

Security & privacy
- No network access, nothing is stored. Only the system message is changed; user messages
  are left untouched.
- If include_user_name is on, the user's display name (as stored in Open WebUI) is sent to
  the model as part of the system note.
"""

from datetime import datetime, tzinfo
from typing import Any, Literal
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from pydantic import BaseModel, Field

MAX_USER_NAME_CHARS = 100

TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "marker": "[Company context]",
        "company": "You are the internal AI assistant of {company}.",
        "date": "Today is {weekday}, {day} {month} {year} (calendar week {week}).",
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
        "date": "Heute ist {weekday}, {day}. {month} {year} (Kalenderwoche {week}).",
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


def zone(name: str) -> tzinfo | None:
    """The IANA time zone, None if the name is empty or unknown."""
    try:
        return ZoneInfo(name.strip()) if name.strip() else None
    except (ZoneInfoNotFoundError, ValueError, OSError):  # OSError: on some systems for a name too long for a file
        return None


def chat_variable(metadata: dict[str, Any] | None, name: str) -> str:
    """A variable the browser sends with every chat request, e.g. {{USER_LANGUAGE}}; '' if missing."""
    variables = (metadata or {}).get("variables")
    return str(variables.get(name) or "") if isinstance(variables, dict) else ""


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
        language: Literal["auto", "en", "de"] = Field(
            default="auto",
            description="Language of the system note and of the date format. auto = the user's interface "
            "language (English if there is no translation).",
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
            default="",
            description="IANA time zone that determines today's date, e.g. 'Europe/Berlin'. Empty = the "
            "time zone of the user's browser, or the server's if the browser sends none. An unknown name "
            "counts as empty (a stored typo must not stop every chat).",
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

    def _now(self, metadata: dict[str, Any] | None = None) -> datetime:
        """Now in the valve's time zone, else the browser's, else server time."""
        timezone = zone(self.valves.timezone) or zone(chat_variable(metadata, "{{CURRENT_TIMEZONE}}"))
        return datetime.now(timezone)

    def _language(self, metadata: dict[str, Any] | None = None) -> str:
        if self.valves.language != "auto":
            return self.valves.language
        locale = chat_variable(metadata, "{{USER_LANGUAGE}}").replace("_", "-").split("-")[0].strip().lower()
        return locale if locale in TRANSLATIONS else "en"

    def build_note(self, user: dict[str, Any] | None, now: datetime, language: str | None = None) -> str:
        """Compose the system note for the given user and day."""
        language = language or self._language()
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
                week=now.isocalendar().week,
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

    async def inlet(
        self,
        body: dict[str, Any],
        __user__: dict[str, Any] | None = None,
        __metadata__: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Prepend the note to the system message (or add one); idempotent per request."""
        messages = body.get("messages") or []
        note = self.build_note(__user__, self._now(__metadata__), self._language(__metadata__))
        if messages and isinstance(messages[0], dict) and messages[0].get("role") == "system":
            messages[0]["content"] = prepend_note(note, messages[0].get("content"))
        else:
            messages.insert(0, {"role": "system", "content": note})
        body["messages"] = messages
        return body
