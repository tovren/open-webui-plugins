"""Context Guard together with Open WebUI's Context Compaction, in a real Open WebUI.

Each test starts a mock model server (8k context, see mock_model.py) and Open WebUI with its own
data folder, installs the guard and a test tool, and plays a chat. Skipped unless OWUI_BIN points
to the `open-webui` executable of an installation, so different versions can be checked:

    OWUI_BIN=/path/to/venv/bin/open-webui python -I -m pytest -q tests/integration
"""

import json
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pytest

OWUI_BIN = os.environ.get("OWUI_BIN", "")
pytestmark = pytest.mark.skipif(not OWUI_BIN, reason="set OWUI_BIN to an open-webui executable")

HERE = Path(__file__).parent
GUARD = HERE.parent.parent / "context-guard" / "context_guard.py"
MODEL = "mock-8k"
# 8192 tokens context, 1000 reserved: usable 7192, hint from 75 % = 5394 tokens
GUARD_VALVES = {"language": "en", "context_tokens": 8192, "answer_reserve_tokens": 1000, "warn_percent": 75}


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def filler(chars: int, label: str) -> str:
    base = f"{label}: a paragraph of text that fills the chat history. "
    return (base * (chars // len(base) + 1))[:chars]


def turns(n: int) -> list[str]:
    """Turns of about 400 tokens question and 600 tokens answer each."""
    return [filler(1400, f"Turn {k}") + " [answer:2100]" for k in range(1, n + 1)]


@dataclass
class Turn:
    calls: list[dict[str, Any]]
    status: list[str]
    error: str

    def kinds(self, kind: str) -> list[dict[str, Any]]:
        return [call for call in self.calls if call["kind"] == kind]


@dataclass
class OpenWebUI:
    base: str
    log: Path
    mock_log: Path
    token: str = ""
    log_pos: int = 0
    chat_id: str | None = None
    parent: str | None = None
    turns: list[Turn] = field(default_factory=list)

    def api(self, method: str, path: str, body: Any = None) -> tuple[int, Any]:
        data = json.dumps(body).encode() if body is not None else None
        request = urllib.request.Request(self.base + path, method=method, data=data)
        request.add_header("Content-Type", "application/json")
        if self.token:
            request.add_header("Authorization", f"Bearer {self.token}")
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                raw = response.read()
                code = response.status
        except urllib.error.HTTPError as error:
            raw, code = error.read(), error.code
        try:
            return code, json.loads(raw) if raw else None
        except ValueError:
            return code, raw.decode(errors="replace")

    def setup(self) -> None:
        user = {"name": "Test", "email": "test@example.com", "password": "Test-1234-Test"}
        code, res = self.api("POST", "/api/v1/auths/signup", user)
        assert code == 200, res
        self.token = res["token"]
        for kind, ident, path in (("functions", "context_guard", GUARD), ("tools", "sections", HERE / "sections_tool.py")):
            body = {"id": ident, "name": ident, "content": path.read_text(), "meta": {"description": "test"}}
            code, res = self.api("POST", f"/api/v1/{kind}/create", body)
            assert code == 200, res
        fn = self.api("GET", "/api/v1/functions/id/context_guard")[1]
        if not fn.get("is_active"):
            self.api("POST", "/api/v1/functions/id/context_guard/toggle")
        if not fn.get("is_global"):
            self.api("POST", "/api/v1/functions/id/context_guard/toggle/global")
        self.update_guard()

    def update_guard(self, **valves: Any) -> None:
        code, res = self.api("POST", "/api/v1/functions/id/context_guard/valves/update", {**GUARD_VALVES, **valves})
        assert code == 200, res

    def new_calls(self) -> list[dict[str, Any]]:
        lines = self.mock_log.read_text().splitlines() if self.mock_log.exists() else []
        calls, self.log_pos = [json.loads(line) for line in lines[self.log_pos :]], len(lines)
        return calls

    def send(
        self, content: str, tools: bool = False, params: dict[str, Any] | None = None, model: str = MODEL
    ) -> Turn:
        user_id, assistant_id = str(uuid.uuid4()), str(uuid.uuid4())
        user_message = {"id": user_id, "parentId": self.parent, "childrenIds": [], "role": "user",
                        "content": content, "timestamp": int(time.time()), "models": [model]}
        body = {"model": model, "stream": tools, "messages": [{"role": "user", "content": content}],
                "id": assistant_id, "parent_id": self.parent, "user_message": user_message, "params": params or {}}
        if self.chat_id:
            body["chat_id"] = self.chat_id
        if tools:
            body["tool_ids"] = ["sections"]
        code, res = self.api("POST", "/api/chat/completions", body)
        if not self.chat_id:
            self.chat_id = self.api("GET", "/api/v1/chats/")[1][0]["id"]
        self.parent = assistant_id
        message: dict[str, Any] = {}
        for _ in range(240):  # tool loops may still be running
            chat = self.api("GET", f"/api/v1/chats/{self.chat_id}")[1]
            message = chat["chat"]["history"]["messages"].get(assistant_id, {})
            if message.get("done") or message.get("error"):
                break
            time.sleep(0.25)
        error = message.get("error")
        error = error.get("content", "") if isinstance(error, dict) else str(error or "")
        if code != 200 and not error:
            error = str(res)
        status = [s.get("description", "") for s in message.get("statusHistory") or []]
        turn = Turn(self.new_calls(), status, error)
        self.turns.append(turn)
        return turn

    def play(self, contents: list[str]) -> list[Turn]:
        return [self.send(content) for content in contents]

    def version(self) -> tuple[int, ...]:
        return tuple(int(part) for part in self.api("GET", "/api/version")[1]["version"].split(".")[:3])

    def server_log(self) -> str:
        return self.log.read_text(errors="replace")

    @property
    def calls(self) -> list[dict[str, Any]]:
        return [call for turn in self.turns for call in turn.calls]


@pytest.fixture
def start(tmp_path: Path):
    processes: list[subprocess.Popen] = []

    def start(mode: str = "llama", **env: str) -> OpenWebUI:
        mock_port, port = free_port(), free_port()
        mock_log = tmp_path / "mock.jsonl"
        processes.append(subprocess.Popen(
            [sys.executable, "-I", str(HERE / "mock_model.py"), str(mock_port), str(mock_log), mode]
        ))
        log = tmp_path / "open-webui.log"
        settings = {
            "DATA_DIR": str(tmp_path / "data"), "WEBUI_SECRET_KEY": "test", "ENABLE_PERSISTENT_CONFIG": "False",
            "ENABLE_OLLAMA_API": "false", "OPENAI_API_BASE_URL": f"http://127.0.0.1:{mock_port}/v1",
            "OPENAI_API_KEY": "x", "RAG_EMBEDDING_ENGINE": "openai",
            "RAG_OPENAI_API_BASE_URL": f"http://127.0.0.1:{mock_port}/v1", "RAG_OPENAI_API_KEY": "x",
            "OFFLINE_MODE": "true", "HF_HUB_OFFLINE": "1", "ENABLE_VERSION_UPDATE_CHECK": "false",
            "ANONYMIZED_TELEMETRY": "false", "ENABLE_TITLE_GENERATION": "false",
            "ENABLE_TAGS_GENERATION": "false", "ENABLE_FOLLOW_UP_GENERATION": "false", **env,
        }
        with open(log, "w") as out:
            processes.append(subprocess.Popen(
                [OWUI_BIN, "serve", "--host", "127.0.0.1", "--port", str(port)],
                env={**os.environ, **settings}, stdout=out, stderr=subprocess.STDOUT,
            ))
        owui = OpenWebUI(f"http://127.0.0.1:{port}", log, mock_log)
        for _ in range(360):
            try:
                if owui.api("GET", "/health")[0] == 200:
                    break
            except OSError:
                pass
            time.sleep(0.5)
        else:
            pytest.fail("Open WebUI does not start:\n" + log.read_text()[-3000:])
        owui.setup()
        return owui

    yield start
    for process in processes:
        process.terminate()
        process.wait(timeout=30)


COMPACTION = {"ENABLE_CONTEXT_COMPACTION": "true"}
HALF = {**COMPACTION, "CONTEXT_COMPACTION_TOKEN_THRESHOLD": "4000"}


def test_without_compaction_the_guard_warns_then_stops(start) -> None:
    owui = start()
    played = owui.play(turns(9))
    assert not any(call.get("rejected") for call in owui.calls)  # the model never gets too much
    first_hint = next(i for i, turn in enumerate(played) if turn.status)
    first_stop = next(i for i, turn in enumerate(played) if turn.error)
    assert first_hint < first_stop and "very long" in played[first_hint].status[0]
    assert played[first_stop].error.startswith("This request is too large for the AI model")
    assert played[first_stop].calls == []
    assert all(turn.error for turn in played[first_stop:])


def test_compaction_in_time_keeps_the_guard_quiet(start) -> None:
    owui = start(**HALF)
    played = owui.play(turns(12))
    assert len([call for turn in played for call in turn.kinds("summary")]) >= 2
    assert all(not turn.status and not turn.error for turn in played)
    assert not any(call.get("rejected") for call in owui.calls)
    assert played[-1].kinds("chat")[0]["has_summary"]
    assert "Context Guard" not in owui.server_log()


def test_default_threshold_is_reported(start) -> None:
    # 80000 tokens: compaction never runs before the guard stops a chat of an 8k model
    owui = start(**COMPACTION)
    played = owui.play(turns(8))
    assert not any(turn.kinds("summary") for turn in played)
    assert played[-1].error
    assert owui.server_log().count("Context Compaction threshold for model 'mock-8k' (80000 tokens)") == 1


def test_tool_loop_is_compacted_between_rounds(start) -> None:
    owui = start(**HALF)
    owui.play(turns(2))
    turn = owui.send("Read sections 1 to 8 and summarize them. [tool:8]", tools=True)
    assert not any(call.get("rejected") for call in owui.calls)
    if owui.version() < (0, 12, 0):  # 0.11 compacts only at the start of a turn: the guard stops the loop
        assert not turn.kinds("summary") and turn.error.startswith("This request is too large")
        return
    assert turn.kinds("summary") and not turn.error and not turn.status
    assert len(turn.kinds("chat")) == 9


def test_long_paste_names_the_message(start) -> None:
    # the summary request of a 24k-character paste does not fit; the guard names the message
    owui = start(**HALF)
    owui.play(turns(3))
    turn = owui.send(filler(24000, "Pasted") + " [answer:1000]")
    assert turn.error.startswith("Your message is too long for the AI model")
    assert not turn.kinds("chat")


def test_threshold_too_close_for_vllm(start) -> None:
    # 7000 of 8192 tokens: the summary request is chat + template + max_tokens 1000, which vLLM
    # rejects; Open WebUI then keeps only the first 500 characters of each old message
    owui = start("vllm", **COMPACTION, CONTEXT_COMPACTION_TOKEN_THRESHOLD="7000")
    played = owui.play(turns(12))
    assert any(call.get("rejected") for call in owui.calls if call["kind"] == "summary")
    assert not any(call.get("rejected") for call in owui.calls if call["kind"] == "chat")
    log = owui.server_log()
    assert log.count("(7000 tokens) leaves too little room") == 1
    assert log.count(f"could not get a summary from the model for chat {owui.chat_id}") == 1
    assert any(turn.status for turn in played)  # no false sense of safety: the hint is shown


def test_vllm_context_size_and_max_tokens(start) -> None:
    # the valve says 100k, but vLLM reports max_model_len 8192; the chat asks for max_tokens 3000,
    # which vLLM adds to the prompt: the guard stops the chat before vLLM rejects a request
    owui = start("vllm")
    owui.update_guard(context_tokens=100000)
    played = [owui.send(content, params={"max_tokens": 3000}) for content in turns(8)]
    assert all(call["max_tokens"] == 3000 for call in owui.calls)
    assert not any(call.get("rejected") for call in owui.calls)
    assert any(turn.error.startswith("This request is too large for the AI model") for turn in played)
