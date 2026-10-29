from duty_swap_automation.config import Settings
from duty_swap_automation.mailbox import GraphMailbox
from duty_swap_automation.models import CellChange
from duty_swap_automation.roster import GraphWorkbook, load_roster


class FakeGraph:
    user_base = "https://graph.microsoft.com/v1.0/users/me%40corp.com"

    def __init__(self, responses):
        self.responses = responses
        self.calls = []

    def request(self, method, url, *, json=None, params=None, headers=None, expect_json=True):
        self.calls.append((method, url, json, params))
        for (m, fragment), resp in self.responses.items():
            if m == method and fragment in url:
                return resp
        return {}


def test_graph_workbook_reads_offset_range_and_writes_cells():
    client = FakeGraph({
        ("POST", "createSession"): {"id": "sess"},
        ("GET", "worksheets('Roster')/usedRange"): {
            "address": "Roster!B2:E5",
            "values": [["Duty Roster", "", "", ""], ["Date", "Day", "Duty", "Assigned"],
                       [46330, "Tuesday", "Gate", "Alice"], [46331, "Wednesday", "Gate", "Bob"]],
        },
        ("GET", "worksheets('Staff')/usedRange"): {"address": "Staff!A1:B2", "values": [["Name", "Email"], ["Alice", "alice@corp.com"]]},
    })
    settings = Settings(roster_file_path="Shared/Duty Roster.xlsx", graph_user="me@corp.com")
    wb = GraphWorkbook(client, settings)
    assert wb.base.endswith("/drive/root:/Shared/Duty%20Roster.xlsx:/workbook")

    roster = load_roster(wb, settings)
    assert [(r.sheet_row, r.assignee) for r in roster.rows] == [(4, "Alice"), (5, "Bob")]
    assert roster.name_col == 5  # column E
    assert roster.email_for("alice") == "alice@corp.com"

    wb.write_cells([CellChange(sheet="Roster", row=4, col=5, old_value="Alice", new_value="Bob")])
    method, url, body, _ = client.calls[-1]
    assert method == "PATCH" and url.endswith("worksheets('Roster')/range(address='E4')")
    assert body == {"values": [["Bob"]]}


def test_graph_mailbox_filters_processed_and_keywords():
    msg = lambda i, subject, cats=(): {  # noqa: E731
        "id": i, "subject": subject, "from": {"emailAddress": {"name": "A", "address": "a@corp.com"}},
        "toRecipients": [], "ccRecipients": [], "receivedDateTime": "2026-10-28T15:00:00Z",
        "body": {"content": "body"}, "categories": list(cats),
    }
    client = FakeGraph({("GET", "mailFolders/inbox/messages"): {"value": [
        msg("1", "Duty swap please"), msg("2", "Lunch"), msg("3", "swap again", ["Duty Swap Processed"]),
    ]}})
    settings = Settings(subject_keywords=["swap"])
    emails = GraphMailbox(client, settings).fetch_candidates()
    assert [e.id for e in emails] == ["1"]
    params = client.calls[0][3]
    assert params["$filter"].startswith("receivedDateTime ge ") and params["$orderby"] == "receivedDateTime asc"
