# Company Context

An Open WebUI **filter** that adds a short system note to every request, the same for every model and
assistant, including newly created ones. Example with `company_name = Example Ltd` and
`timezone = Europe/Berlin`:

```text
[Company context]
You are the internal AI assistant of Example Ltd.
Today is Monday, 12 October 2026.
Always answer in English unless the user writes in another language or explicitly asks for a different one.
Address the user formally. Texts you write for them (e-mails, letters) should also use a formal tone unless they ask otherwise.
The user's name is Ada Lovelace.
```

With `language = de` the note is written in German, with German dates
(`Heute ist Montag, 12. Oktober 2026.`) and *Sie*/*du* as the form
of address.

- If the request already has a system message (model or chat system prompt), the note is put in
  front of it and the existing prompt is kept. Otherwise a new system message is added.
- The note starts with a marker line, so it is never added twice to the same request.
- The note contains the date but not the time. It opens the prompt, so it has to stay the same all
  day: then the model server (llama-server, vLLM, Ollama) can reuse its prompt cache instead of
  processing the whole chat again on every turn.

![An answer that uses today's date and the user's name](images/answer.png)

## Install

File: [`company_context.py`](company_context.py). Install it under **Admin Panel → Functions** and
enable it globally, see [Installation](../README.md#installation). Set `company_name` and
`timezone` in its valves.

## Valves

| Valve | Default | Description |
|---|---|---|
| `priority` | `0` | Filter order (lower runs first) |
| `language` | `en` | Language of the note and date format (`en` or `de`) |
| `company_name` | _(empty)_ | Your company name. Empty leaves out the company line |
| `response_language` | _(empty)_ | Language the assistant should answer in, written in the note's language (e.g. `English`, `Deutsch`, `French`). Empty means the note's language |
| `form_of_address` | `formal` | `formal` or `informal` (German: *Sie* / *du*), also for texts written for the user |
| `timezone` | `UTC` | IANA time zone that determines today's date, e.g. `Europe/Berlin`. Invalid values fall back to server time |
| `include_user_name` | `true` | Tell the model the user's display name, e.g. for e-mail signatures |
| `extra_rules` | _(empty)_ | Additional house rules appended to the note |

## Security & privacy

No network access, nothing is stored. Only the system message is changed and user messages are left
untouched. With `include_user_name` on, the user's display name from Open WebUI is sent to the model
(line breaks removed, at most 100 characters).
