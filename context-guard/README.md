# Context Guard

An Open WebUI **filter** that estimates the prompt size right before each request is sent to the model
and warns or stops before the context window overflows.

- Runs on the `request` hook (new in Open WebUI 0.11). It runs **after** document excerpts
  (RAG/files) have been injected and again after every tool-call round, so everything the model
  will see is counted: also the model's system prompt (Workspace → Models), which Open WebUI only
  adds after this hook, and earlier reasoning that is sent back to the model (`reasoning_content`
  for llama.cpp, `thinking` for Ollama).
- Once the prompt uses `warn_percent` of the usable context, a status line appears above the
  answer. The usable context is the context size minus `answer_reserve_tokens`, or minus the
  request's `max_tokens` if that is larger (a model or chat parameter, or the default parameters):
  vLLM rejects a request whose prompt and `max_tokens` together exceed `--max-model-len`.
- If the prompt would not fit, the request is stopped with a friendly explanation and next steps:
  start a new chat, paste less text, attach fewer documents. If the user's own message makes up
  most of the prompt (a long pasted text), the message says so and suggests pasting only the
  relevant section or uploading the text as a file. Document excerpts that Open WebUI adds to the
  message do not count as the user's text.
- The messages are flowing text without line breaks, because Open WebUI shows errors as plain text
  and would join the lines.

Why this helps: **llama-server** and **vLLM** reject oversized requests with a technical HTTP 400
error ("exceeds the available context size"). **Ollama** silently cuts off the oldest part of the
prompt, so the model quietly forgets earlier messages or documents.

The token count is an **estimate** (characters ÷ `chars_per_token`, plus a flat amount per image
and the JSON size of tool definitions and tool calls), not a real tokenizer. Lower
`chars_per_token` to be more cautious. Open WebUI's own estimate for Context Compaction uses 4
characters per token (or the real token count of the last answer), so with the default of 3.5 the
guard counts a little higher than Open WebUI does.

![Warning status line above the answer](images/warning.png)

![A request that is too large is stopped with a message](images/blocked.png)

## Install

File: [`context_guard.py`](context_guard.py). Install it under **Admin Panel → Functions** and
enable it globally or per model, see [Installation](../README.md#installation). Set
`context_tokens` to the context size of your model, unless the guard finds it on its own:

1. an entry for the model in `model_context_tokens`,
2. Ollama: the model's `num_ctx` parameter (Workspace → Models → Advanced Params), which Open WebUI
   sends with the request,
3. vLLM: the `max_model_len` that vLLM reports in its model list (also for workspace models on a
   vLLM base model),
4. otherwise `context_tokens`.

llama-server reports no context size that fits (only the training size), so set `context_tokens` or
`model_context_tokens` for it.

## Valves

| Valve | Default | Description |
|---|---|---|
| `priority` | `100` | Filter order (lower runs first). Only matters among the `request` hooks of several filters; keep it high so their changes are counted. Inlet hooks such as Company Context always run before and are counted at any priority |
| `language` | `auto` | Language of the warning and error messages: `auto` (the user's interface language, English if there is no translation), `en` or `de` |
| `context_tokens` | `32768` | Context size of your model: llama-server `--ctx-size`, vLLM `--max-model-len`, Ollama `num_ctx`. Used when the request has no `num_ctx` and the server reports no `max_model_len` |
| `model_context_tokens` | _(empty)_ | Optional per-model sizes as `model_id=tokens` pairs, e.g. `qwen3-32b=40960, llama3.1:8b=8192`. A workspace model (assistant) uses the size of its base model unless it has its own entry. An entry wins over `num_ctx` and `max_model_len` |
| `answer_reserve_tokens` | `6000` | Tokens kept free for the answer, including reasoning; at most half the context size. A larger `max_tokens` of the request is kept free instead |
| `chars_per_token` | `3.5` | Characters per token for the estimate (English about 4, German about 3.5 to 4.3) |
| `tokens_per_image` | `300` | Estimated tokens per attached image |
| `warn_percent` | `75` | Show the hint from this share of the usable context (percent) |

At most half the context size is kept free for the answer, so a model with a small context (an
Ollama `num_ctx` of 4096, say) still gets a check with the default `answer_reserve_tokens` of 6000.

## Built-in tools of Open WebUI 0.12

From 0.12 on, Open WebUI sends the definitions of its built-in tools (knowledge, memories, notes,
chats, skills, automations, calendar …) with every chat that is started in the browser: 39 tools
with about 27,000 characters of JSON in 0.12.0. How many tokens that is depends on the model and its
chat template. Measured with the same request:

| Model | Tokens of the built-in tools | Guard's estimate (`chars_per_token` 3.5) |
|---|---|---|
| Qwen3.8 27B | 7,109 | 7,793 (+10 %) |
| Gemma 4 26B-A4B | 6,216 | 7,793 (+25 %) |
| gpt-oss-20b | 4,207 | 7,793 (+85 %) |

So about 4,000 to 7,000 tokens of the context are taken before the user has typed anything, and a
model with an 8k context hardly has room for a chat. The guard counts these tokens on purpose: they
take up the model's context like any other text. Its estimate is close for Qwen and on the cautious
side for models whose tokenizer or chat template is more compact.

If a model does not need the built-in tools, switch them off in **Workspace → Models → (model) →
Capabilities → Builtin Tools**, or set the model's function calling to `legacy`. The individual
groups (knowledge, memory, notes …) can also be switched off there.

## With Context Compaction

Open WebUI's built-in [Context Compaction](https://docs.openwebui.com/troubleshooting/context-window/)
(off by default) summarizes older messages once a chat passes a token threshold. It works together
with the guard:

- Compaction runs before the filters, so the guard checks the shortened history including the
  summary.
- As long as compaction will summarize the chat before the limit, the hint to start a new chat is
  not shown. It still appears if the prompt stays above the compaction threshold, for example
  because of large documents.
- The guard stays the safety net. Compaction skips a turn when no earlier message can be cut off,
  and the summary request contains the whole chat, so after a large paste it can overflow itself.
  The guard then stops the request with its message.
- Long tool-call loops: Open WebUI 0.12.0 also compacts between the rounds, before the guard checks
  each round. In 0.11.4 only the guard covers a loop that keeps growing.
- Since 0.12.4, compaction also runs in temporary chats and in short chats that are already over the
  threshold (before, chats with up to three messages were never compacted), and the chat shows a
  "Context compacted" line with the summary. The guard needs no change for this.
- The summary request contains the whole chat plus the summary prompt (about 500 tokens) and room
  for the summary (`max_tokens` of the task model parameters, 1000 if they set none), so it needs
  about 1500 tokens more than the threshold. It goes to the compaction model if one is set
  (`CONTEXT_COMPACTION_MODEL`), otherwise to the chat model. If the threshold plus these tokens is not
  below the usable context of that model, the guard logs a warning once and shows its hint as usual. The default of 80000 tokens is above most local context sizes, so
  compaction never runs before the guard stops a chat.
- If the summary request fails, for example because vLLM rejects a prompt plus `max_tokens` that
  is larger than `--max-model-len`, Open WebUI does not report it: it stores the first 500
  characters of each older message as the "summary" and carries on. The guard recognizes this,
  logs a warning once per chat (without message content) and shows its hint as usual.

Set `CONTEXT_COMPACTION_TOKEN_THRESHOLD`, or a model's `compact_token_threshold`, to about **half
the context size**, e.g. `16000` for a 32k model. Compaction is only checked at the start of a turn,
and the summary request is about as large as the whole chat, so a threshold close to the limit
leaves too little room.

In Open WebUI 0.11.4 and 0.12.0, running `/compact` right after the guard stopped a request does
not help: Open WebUI stores the summary on the empty error message and skips that message when it
builds the next request. That is why the stop message suggests a new chat, not `/compact`.

## Security & privacy

Read-only: the request is never modified. No network access, nothing is stored, and no message
content is logged. Open WebUI's Context Compaction settings are only read.
