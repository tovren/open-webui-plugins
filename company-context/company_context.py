"""
title: Company Context
author: tovren
author_url: https://github.com/tovren/open-webui-plugins
version: 1.3.1
license: MIT
required_open_webui_version: 0.12.4
description: Adds a short system note to every request with your company name, today's date and calendar week, the response language, the form of address, the user's name, your own house rules, a short company profile and texts for single user groups (e.g. departments) - the same for every model and assistant, including newly created ones.

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
- The parts that are the same for everybody come first (rules, company profile), then the texts
  of the user's groups, then the date and the user's name. So the model server can share the
  cached start of the prompt between all users, and between the members of a group.
- Each user can pick the form of address per model, through Open WebUI's Model Controls (open to
  every user of the model): an option sets the custom parameter company_context, e.g.
  {"address": "informal"}. The filter takes it out of the request, so it never reaches the model
  server, and writes that line at the end of the note, in front of the date: the start of the prompt
  stays the same for all users.
- group_profiles holds one section per group, headed "## <group name>". A user gets the sections
  of the groups they belong to (also through a subgroup), in the order of the
  sections. Section names that match no group are written to the log once.

Security & privacy
- No network access, nothing is stored. Only the system message is changed; user messages
  are left untouched. With group_profiles set, the user's group names are read from Open WebUI's
  database.
- If include_user_name is on, the user's display name (as stored in Open WebUI) is sent to
  the model as part of the system note.
"""

import logging
import re
import time
from datetime import datetime, tzinfo
from typing import Any, Literal
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

MAX_USER_NAME_CHARS = 100
NOTE_WARN_CHARS = 8000  # about 2000 tokens; every request carries the note
GROUP_CHECK_SECONDS = 600  # how often group_profiles is compared with the existing groups
GROUP_HEADING = re.compile(r"^##[ \t]+(.+?)[ \t#]*$")  # "## Sales"; "###" stays text of the section

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
        "profile": "Company information:",
        "group_profile": "Information for {group}:",
        "groups": "The user belongs to: {groups}.",
        "no_invention": (
            "Use this information where it fits the question. Do not make up company details that "
            "are neither given here nor in the sources of this chat (people, phone numbers, prices, "
            "deadlines, rules); say that you do not know them instead."
        ),
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
        "profile": "Informationen zur Firma:",
        "group_profile": "Informationen für {group}:",
        "groups": "Der Nutzer gehört zu: {groups}.",
        "no_invention": (
            "Nutze diese Angaben, wenn sie zur Frage passen. Erfinde keine Firmendetails, die weder "
            "hier noch in den Quellen dieses Chats stehen (Personen, Telefonnummern, Preise, Fristen, "
            "Regeln); sag stattdessen, dass du sie nicht kennst."
        ),
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
CHOICES_PARAM = "company_context"  # custom parameter set by Model Controls, removed from the request
CHOICES = {"address": ("formal", "informal")}


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


def pop_choices(body: dict[str, Any]) -> Any:
    """Remove the custom parameter from the request: top level, or Ollama's options."""
    value = body.pop(CHOICES_PARAM, None)
    options = body.get("options")
    if isinstance(options, dict) and (in_options := options.pop(CHOICES_PARAM, None)) is not None:
        value = in_options if value is None else value
    return value


def group_key(name: str) -> str:
    """A group name for comparison: case and runs of spaces do not matter."""
    return " ".join(name.split()).casefold()


def parse_group_profiles(text: str) -> tuple[list[tuple[str, str]], bool]:
    """The sections of the group_profiles valve as (group name, text) in their order, and whether
    there is text before the first heading (it belongs to no group). A name used twice gets one
    section with both texts."""
    sections: dict[str, tuple[str, list[str]]] = {}
    current: list[str] | None = None
    text_before = False
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if heading := GROUP_HEADING.match(line.strip()):
            name = " ".join(heading.group(1).split())
            current = sections.setdefault(group_key(name), (name, []))[1]
        elif current is not None:
            current.append(line)
        elif line.strip():
            text_before = True
    return [(name, "\n".join(lines).strip()) for name, lines in sections.values()], text_before


async def member_group_names(user_id: str) -> list[str]:
    """Names of the user's groups, with the parent groups of their groups. Raises if Open WebUI's
    groups are not available."""
    from open_webui.models.groups import Groups

    groups = await Groups.get_groups_by_member_id(user_id, include_inherited=True)
    return [str(getattr(group, "name", "") or "") for group in groups]


async def all_group_names() -> list[str]:
    """Names of all groups. Raises if Open WebUI's groups are not available."""
    from open_webui.models.groups import Groups

    return [str(getattr(group, "name", "") or "") for group in await Groups.get_all_groups()]


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
            description="Additional house rules for the note, e.g. 'Always write to customers "
            "in the first person plural.'",
        )
        company_profile: str = Field(
            default="",
            description="Facts every user may need, e.g. sites, departments, products, internal "
            "abbreviations, who to ask for what. Every request carries it, so keep it short (at most "
            "about 2000 tokens, ~8000 characters); longer texts belong in a knowledge base. The model is "
            "told not to make up company details that are not given.",
        )
        group_profiles: str = Field(
            default="",
            description="Texts for single user groups (e.g. departments), one section per group headed "
            "'## <group name>' as in Admin Panel → Users → Groups. A user gets the sections of their "
            "groups. A heading without text only names the group (include_group_names).",
        )
        include_group_names: bool = Field(
            default=True,
            description="Tell the model which of the groups in group_profiles the user belongs to.",
        )

    def __init__(self) -> None:
        self.valves = self.Valves()
        self._warned: set[Any] = set()  # warnings already written to the log
        self._group_check: tuple[str, float, list[str]] = ("", 0.0, [])  # valve text, time, unknown names

    def _now(self, metadata: dict[str, Any] | None = None) -> datetime:
        """Now in the valve's time zone, else the browser's, else server time."""
        timezone = zone(self.valves.timezone) or zone(chat_variable(metadata, "{{CURRENT_TIMEZONE}}"))
        return datetime.now(timezone)

    def _language(self, metadata: dict[str, Any] | None = None) -> str:
        if self.valves.language != "auto":
            return self.valves.language
        locale = chat_variable(metadata, "{{USER_LANGUAGE}}").replace("_", "-").split("-")[0].strip().lower()
        return locale if locale in TRANSLATIONS else "en"

    def build_note(
        self,
        user: dict[str, Any] | None,
        now: datetime,
        language: str | None = None,
        groups: list[tuple[str, str]] | None = None,
        choices: dict[str, str] | None = None,
    ) -> str:
        """Compose the system note for the given user and day; groups = the sections of
        group_profiles that apply to the user, choices = the user's validated Model Controls picks."""
        language = language or self._language()
        texts = TRANSLATIONS[language]
        head = [texts["marker"]]
        if company := self.valves.company_name.strip():
            head.append(texts["company"].format(company=company))
        response_language = self.valves.response_language.strip() or texts["response_language"]
        head.append(texts["language_rule"].format(language=response_language))
        choices = choices or {}
        if "address" not in choices:
            head.append(texts[self.valves.form_of_address])
        if extra_rules := self.valves.extra_rules.strip():
            head.append(extra_rules)

        blocks = []
        if profile := self.valves.company_profile.strip():
            blocks.append(f"{texts['profile']}\n{profile}")
        blocks += [
            f"{texts['group_profile'].format(group=name)}\n{text}" for name, text in groups or [] if text
        ]
        facts = bool(blocks)
        if groups and self.valves.include_group_names:
            blocks.append(texts["groups"].format(groups=", ".join(name for name, _ in groups)))
        if facts:
            blocks.append(texts["no_invention"])

        tail = []
        if address := choices.get("address"):
            tail.append(texts[address])
        tail.append(
            texts["date"].format(
                weekday=WEEKDAYS[language][now.weekday()],
                day=now.day,
                month=MONTHS[language][now.month - 1],
                year=now.year,
                week=now.isocalendar().week,
            )
        )
        if self.valves.include_user_name and (
            user_name := " ".join(str((user or {}).get("name") or "").split())[:MAX_USER_NAME_CHARS]
        ):
            tail.append(texts["user_name"].format(name=user_name))
        if not blocks:
            return "\n".join(head + tail)
        return "\n\n".join(["\n".join(head), *blocks, "\n".join(tail)])

    def _warn_once(self, key: Any, message: str, *args: Any, exc_info: bool = False) -> None:
        if key not in self._warned:
            self._warned.add(key)
            log.warning(message, *args, exc_info=exc_info)

    async def _check_group_names(self, sections: list[tuple[str, str]]) -> None:
        """Log sections that match no group (a typo, a renamed or deleted group): when the valve
        changes and then every GROUP_CHECK_SECONDS, but only if the list of such names changed."""
        text, checked_at, reported = self._group_check
        if text == self.valves.group_profiles and time.monotonic() - checked_at < GROUP_CHECK_SECONDS:
            return
        try:
            known = {group_key(name) for name in await all_group_names()}
        except Exception:
            return  # reported when the user's groups cannot be read either
        unknown = [name for name, _ in sections if group_key(name) not in known]
        if unknown and unknown != reported:
            log.warning("Company Context: group_profiles names groups that do not exist: %s", ", ".join(unknown))
        self._group_check = (self.valves.group_profiles, time.monotonic(), unknown)

    def _choices(self, value: Any) -> dict[str, str]:
        """The user's picks from the company_context parameter; unknown values are left out (logged once)."""
        if value is None:
            return {}
        if not isinstance(value, dict):
            self._warn_once(
                ("choices", repr(value)[:200]),
                "Company Context: the custom parameter %s must be an object such as "
                '{"address": "informal"}, got %r',
                CHOICES_PARAM,
                value,
            )
            return {}
        choices = {}
        for key, choice in value.items():
            if choice in CHOICES.get(key, ()):
                choices[key] = choice
            else:
                self._warn_once(
                    ("choice", str(key), repr(choice)[:200]),
                    "Company Context: unknown %s.%s = %r is left out (allowed: %s)",
                    CHOICES_PARAM,
                    key,
                    choice,
                    ", ".join(f"{name}: {'/'.join(values)}" for name, values in CHOICES.items()),
                )
        return choices

    async def _user_groups(self, user: dict[str, Any] | None) -> list[tuple[str, str]]:
        """The sections of group_profiles for the user's groups, in the order of the sections."""
        sections, text_before = parse_group_profiles(self.valves.group_profiles)
        if not sections:
            return []
        if text_before:
            self._warn_once(
                ("text before", self.valves.group_profiles),
                "Company Context: group_profiles has text before the first '## <group name>' heading; "
                "it belongs to no group and is left out",
            )
        await self._check_group_names(sections)
        user_id = str((user or {}).get("id") or "")
        if not user_id:
            return []
        try:
            member_of = {group_key(name) for name in await member_group_names(user_id)}
        except Exception:
            self._warn_once(
                "groups unavailable",
                "Company Context: cannot read the user's groups, group_profiles is left out",
                exc_info=True,
            )
            return []
        return [(name, text) for name, text in sections if group_key(name) in member_of]

    async def inlet(
        self,
        body: dict[str, Any],
        __user__: dict[str, Any] | None = None,
        __metadata__: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Prepend the note to the system message (or add one); idempotent per request."""
        choices = self._choices(pop_choices(body))
        messages = body.get("messages") or []
        groups = await self._user_groups(__user__)
        note = self.build_note(__user__, self._now(__metadata__), self._language(__metadata__), groups, choices)
        if len(note) > NOTE_WARN_CHARS:
            self._warn_once(
                ("size", self.valves.company_profile, self.valves.group_profiles, tuple(name for name, _ in groups)),
                "Company Context: the note has %d characters (about %d tokens) and every request carries "
                "it; keep company_profile and group_profiles short and put longer texts into a knowledge base",
                len(note),
                len(note) // 4,
            )
        if messages and isinstance(messages[0], dict) and messages[0].get("role") == "system":
            messages[0]["content"] = prepend_note(note, messages[0].get("content"))
        else:
            messages.insert(0, {"role": "system", "content": note})
        body["messages"] = messages
        return body
