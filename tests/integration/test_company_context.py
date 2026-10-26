"""Company Context's group_profiles with the groups of a real Open WebUI (0.11: flat groups,
0.12: subgroups count for their parent group). Uses the harness of test_compaction.py:

    OWUI_BIN=/path/to/venv/bin/open-webui python -I -m pytest -q tests/integration
"""

from pathlib import Path

from test_compaction import OWUI_BIN, OpenWebUI, pytestmark, start  # noqa: F401 (fixture, skip mark)

COMPANY_CONTEXT = Path(__file__).parent.parent.parent / "company-context" / "company_context.py"
PROFILES = "## Vertrieb\nPreisliste: P:/Vertrieb\n## Personal\nUrlaub über das Portal\n## Einkauf alt\nx"


def install_company_context(owui: OpenWebUI) -> None:
    body = {"id": "company_context", "name": "company_context", "content": COMPANY_CONTEXT.read_text(),
            "meta": {"description": "test"}}
    code, res = owui.api("POST", "/api/v1/functions/create", body)
    assert code == 200, res
    fn = owui.api("GET", "/api/v1/functions/id/company_context")[1]
    if not fn.get("is_active"):
        owui.api("POST", "/api/v1/functions/id/company_context/toggle")
    if not fn.get("is_global"):
        owui.api("POST", "/api/v1/functions/id/company_context/toggle/global")
    valves = {"language": "de", "company_profile": "Standorte: Stuttgart, Ulm", "group_profiles": PROFILES}
    code, res = owui.api("POST", "/api/v1/functions/id/company_context/valves/update", valves)
    assert code == 200, res


def create_group(owui: OpenWebUI, name: str, user_ids: list[str], parent: str | None = None) -> str:
    body = {"name": name, "description": "", **({"parent_group_id": parent} if parent else {})}
    code, group = owui.api("POST", "/api/v1/groups/create", body)
    assert code == 200, group
    if user_ids:
        code, res = owui.api("POST", f"/api/v1/groups/id/{group['id']}/users/add", {"user_ids": user_ids})
        assert code == 200, res
    return group["id"]


def test_group_profiles_follow_the_users_groups(start) -> None:
    owui = start()
    install_company_context(owui)
    user_id = owui.api("GET", "/api/v1/auths/")[1]["id"]

    system = owui.send("Hallo").kinds("chat")[0]["system"]
    assert "Informationen zur Firma:\nStandorte: Stuttgart, Ulm" in system and "Informationen für" not in system

    sales = create_group(owui, "Vertrieb", [] if owui.version() >= (0, 12) else [user_id])
    if owui.version() >= (0, 12):  # a member of a subgroup counts as a member of its parent group
        create_group(owui, "Vertrieb Süd", [user_id], parent=sales)
    create_group(owui, "personal", [user_id])
    system = owui.send("Noch einmal").kinds("chat")[0]["system"]
    assert "Informationen für Vertrieb:\nPreisliste: P:/Vertrieb\n\nInformationen für Personal:" in system
    assert "Der Nutzer gehört zu: Vertrieb, Personal." in system
    assert system.index("Informationen zur Firma") < system.index("Informationen für") < system.index("Heute ist")
    # before the groups existed all three names were unknown; the next check names the one left
    assert "names groups that do not exist: Vertrieb, Personal, Einkauf alt" in owui.server_log()
    code, res = owui.api("POST", "/api/v1/functions/id/company_context/valves/update",
                         {"language": "de", "group_profiles": PROFILES + "\n"})
    assert code == 200, res
    owui.send("Und noch einmal")
    assert "names groups that do not exist: Einkauf alt\n" in owui.server_log()
    assert "cannot read the user's groups" not in owui.server_log()


CONTROLS = {
    "anrede": {"label": "Anrede", "default": "sie", "options": {
        "sie": {"label": "Sie", "params": {"company_context": {"address": "formal"}}},
        "du": {"label": "Du", "params": {"company_context": {"address": "informal"}}},
    }},
}


def test_model_controls_pick_the_form_of_address(start) -> None:
    owui = start()
    if owui.version() < (0, 12):
        return  # Model Controls came with 0.12.0
    install_company_context(owui)
    owui.api("GET", "/api/models")  # fills the model list the model check needs
    model = {"id": "assistent", "base_model_id": "mock-8k", "name": "Assistent", "meta": {},
             "params": {"model_controls": CONTROLS}}
    code, res = owui.api("POST", "/api/v1/models/create", model)
    assert code == 200, res
    owui.api("GET", "/api/models")  # the chat only finds the new model in the refreshed list
    texts = {"formal": "„Sie“", "informal": "„du“"}

    call = owui.send("Hallo", model="assistent").kinds("chat")[0]  # the default: Sie
    assert "company_context" not in call["keys"]
    assert texts["formal"] in call["system"] and texts["informal"] not in call["system"]

    picks = {"model_controls": {"assistent": {"anrede": "du"}}}
    call = owui.send("Noch einmal", model="assistent", params=picks).kinds("chat")[0]
    assert "company_context" not in call["keys"]
    system = call["system"]
    assert texts["informal"] in system and texts["formal"] not in system
    assert system.index("Informationen zur Firma") < system.index(texts["informal"]) < system.index("Heute ist")

    # without the filter the parameter reaches the model server (vLLM ignores unknown fields)
    owui.api("POST", "/api/v1/functions/id/company_context/toggle")
    call = owui.send("Ohne Filter", model="assistent", params=picks).kinds("chat")[0]
    assert "company_context" in call["keys"] and "[Firmen-Kontext]" not in call["system"]
