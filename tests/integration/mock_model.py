"""Fake OpenAI-compatible model server with a small context, for the Context Guard integration tests.

python -I mock_model.py <port> <log.jsonl> <llama|vllm>

- "Tokenizer": characters / 3.5 over all texts, tool calls and tool definitions.
- llama mode rejects a prompt larger than the context with HTTP 400 like llama-server; vllm mode
  also rejects it when prompt + max_tokens is larger, and reports max_model_len in /v1/models, like vLLM.
- Summary requests of Open WebUI's Context Compaction get a short summary.
- The answer length is set with "[answer:1234]" in the last user message; "[tool:8]" makes the
  model call the first tool 8 times before it answers.
- Every request is written to the log as one JSON line.
"""

import json
import math
import re
import sys
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

PORT, LOG, MODE = int(sys.argv[1]), sys.argv[2], sys.argv[3]
N_CTX = 8192
CHARS_PER_TOKEN = 3.5
MODEL = "mock-8k"
TOOL_ROUNDS: dict[str, int] = {}  # tool rounds per user message (key: start of its text)


def text_of(content):
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        return "".join(p.get("text", "") for p in content if isinstance(p, dict))
    return ""


def count_tokens(body):
    chars = 0
    for m in body.get("messages") or []:
        chars += len(text_of(m.get("content"))) + 16
        chars += sum(len(json.dumps(c, ensure_ascii=False)) for c in m.get("tool_calls") or [])
    if body.get("tools"):
        chars += len(json.dumps(body["tools"], ensure_ascii=False))
    return math.ceil(chars / CHARS_PER_TOKEN)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def _json(self, code, obj):
        data = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _log(self, entry):
        with open(LOG, "a") as f:
            f.write(json.dumps(entry) + "\n")

    def do_GET(self):
        if self.path.rstrip("/").endswith("/models"):
            model = {"id": MODEL, "object": "model", "owned_by": "mock"}
            if MODE == "vllm":
                model["max_model_len"] = N_CTX
            self._json(200, {"object": "list", "data": [model]})
        else:
            self._json(404, {"error": "not found"})

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
        messages = body.get("messages") or []
        all_text = "\n".join(text_of(m.get("content")) for m in messages)
        is_summary = "Summarize the conversation history" in all_text
        is_task = not is_summary and len(messages) == 1 and "### Task:" in all_text
        last_user = next((text_of(m.get("content")) for m in reversed(messages) if m.get("role") == "user"), "")
        entry = {
            "kind": "summary" if is_summary else ("task" if is_task else "chat"),
            "prompt_tokens": count_tokens(body),
            "max_tokens": body.get("max_tokens") or body.get("max_completion_tokens"),
            "has_summary": "[CONVERSATION SUMMARY]" in all_text,
            "system": text_of(messages[0].get("content")) if messages and messages[0].get("role") == "system" else "",
        }
        needed = entry["prompt_tokens"] + (int(entry["max_tokens"] or 0) if MODE == "vllm" else 0)
        if needed > N_CTX:
            entry["rejected"] = True
            self._log(entry)
            self._json(400, {"error": {"code": 400, "type": "exceed_context_size_error",
                                       "message": "the request exceeds the available context size"}})
            return
        tool_call = None
        wanted = re.search(r"\[tool:(\d+)\]", last_user)
        if body.get("tools") and wanted and not is_summary:
            key = last_user[:300]
            TOOL_ROUNDS[key] = TOOL_ROUNDS.get(key, 0) + 1
            entry["tool_round"] = TOOL_ROUNDS[key]
            if TOOL_ROUNDS[key] <= int(wanted.group(1)):
                name = body["tools"][0]["function"]["name"]
                tool_call = {"index": 0, "id": f"call_{TOOL_ROUNDS[key]}", "type": "function",
                             "function": {"name": name, "arguments": json.dumps({"number": TOOL_ROUNDS[key]})}}
        if tool_call:
            reply = ""
        elif is_summary:
            reply = "SUMMARY: " + "Discussed points A, B and C. " * 12
        elif is_task:
            reply = '{"title": "Test", "tags": [], "follow_ups": []}'
        else:
            m = re.search(r"\[answer:(\d+)\]", last_user)
            n = int(m.group(1)) if m else 800
            reply = ("Answer text " * (n // 12 + 1))[:n]
        self._log(entry)
        usage = {"prompt_tokens": entry["prompt_tokens"], "completion_tokens": math.ceil(len(reply) / CHARS_PER_TOKEN)}
        usage["total_tokens"] = usage["prompt_tokens"] + usage["completion_tokens"]
        created = int(time.time())
        if body.get("stream"):
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self.end_headers()
            delta = {"role": "assistant", "tool_calls": [tool_call]} if tool_call else {"role": "assistant", "content": reply}
            chunks = [
                {"choices": [{"index": 0, "delta": delta, "finish_reason": None}]},
                {"choices": [{"index": 0, "delta": {}, "finish_reason": "tool_calls" if tool_call else "stop"}]},
                {"choices": [], "usage": usage},
            ]
            for c in chunks:
                c.update({"id": "mock", "object": "chat.completion.chunk", "created": created, "model": MODEL})
                self.wfile.write(f"data: {json.dumps(c)}\n\n".encode())
            self.wfile.write(b"data: [DONE]\n\n")
            return
        message = {"role": "assistant", "content": reply, **({"tool_calls": [tool_call]} if tool_call else {})}
        self._json(200, {"id": "mock", "object": "chat.completion", "created": created, "model": MODEL,
                         "choices": [{"index": 0, "message": message,
                                      "finish_reason": "tool_calls" if tool_call else "stop"}],
                         "usage": usage})


ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
