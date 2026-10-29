from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Protocol

from duty_swap_automation.config import Settings
from duty_swap_automation.graph_client import GraphClient
from duty_swap_automation.models import Email


class Mailbox(Protocol):
    def fetch_candidates(self) -> list[Email]: ...

    def reply(self, email: Email, to: list[str], cc: list[str], body: str) -> None: ...

    def mark_processed(self, email: Email, category: str) -> None: ...


def _matches_keywords(email: Email, keywords: list[str]) -> bool:
    if not keywords:
        return True
    haystack = f"{email.subject}\n{email.body[:2000]}".lower()
    return any(k.lower() in haystack for k in keywords)


class GraphMailbox:
    def __init__(self, client: GraphClient, settings: Settings):
        self.client = client
        self.settings = settings

    def fetch_candidates(self) -> list[Email]:
        since = datetime.now(timezone.utc) - timedelta(hours=self.settings.lookback_hours)
        url = f"{self.client.user_base}/mailFolders/inbox/messages"
        params = {
            # Graph requires $orderby properties to lead the $filter expression.
            "$filter": f"receivedDateTime ge {since.strftime('%Y-%m-%dT%H:%M:%SZ')} and isRead eq false",
            "$orderby": "receivedDateTime asc",
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,categories",
            "$top": "50",
        }
        headers = {"Prefer": 'outlook.body-content-type="text"'}
        emails: list[Email] = []
        data = self.client.request("GET", url, params=params, headers=headers)
        while True:
            for msg in data.get("value", []):
                email = Email(
                    id=msg["id"],
                    subject=msg.get("subject") or "",
                    sender_name=(msg.get("from") or {}).get("emailAddress", {}).get("name", ""),
                    sender_email=(msg.get("from") or {}).get("emailAddress", {}).get("address", ""),
                    to=[r["emailAddress"]["address"] for r in msg.get("toRecipients", [])],
                    cc=[r["emailAddress"]["address"] for r in msg.get("ccRecipients", [])],
                    received_at=msg.get("receivedDateTime"),
                    body=(msg.get("body") or {}).get("content", ""),
                    categories=msg.get("categories", []),
                )
                skip = {self.settings.processed_category, self.settings.attention_category}
                if skip.intersection(email.categories):
                    continue
                if _matches_keywords(email, self.settings.subject_keywords):
                    emails.append(email)
            next_link = data.get("@odata.nextLink")
            if not next_link:
                break
            data = self.client.request("GET", next_link, headers=headers)
        return emails

    def reply(self, email: Email, to: list[str], cc: list[str], body: str) -> None:
        payload = {
            "message": {
                "toRecipients": [{"emailAddress": {"address": a}} for a in to],
                "ccRecipients": [{"emailAddress": {"address": a}} for a in cc],
                "body": {"contentType": "Text", "content": body},
            }
        }
        self.client.request(
            "POST", f"{self.client.user_base}/messages/{email.id}/reply", json=payload, expect_json=False
        )

    def mark_processed(self, email: Email, category: str) -> None:
        categories = sorted(set(email.categories) | {category})
        self.client.request(
            "PATCH",
            f"{self.client.user_base}/messages/{email.id}",
            json={"isRead": True, "categories": categories},
        )


class LocalMailbox:
    """Inbox = <dir>/inbox/*.json, replies written to <dir>/outbox/, processed state in <dir>/processed.json."""

    def __init__(self, settings: Settings):
        self.settings = settings
        self.root = Path(settings.local_mail_dir)
        self.state_path = self.root / "processed.json"

    def _state(self) -> dict[str, str]:
        if self.state_path.exists():
            return json.loads(self.state_path.read_text())
        return {}

    def fetch_candidates(self) -> list[Email]:
        processed = self._state()
        emails = []
        for path in sorted((self.root / "inbox").glob("*.json")):
            email = Email.model_validate_json(path.read_text())
            if email.id in processed:
                continue
            if _matches_keywords(email, self.settings.subject_keywords):
                emails.append(email)
        return emails

    def reply(self, email: Email, to: list[str], cc: list[str], body: str) -> None:
        outbox = self.root / "outbox"
        outbox.mkdir(parents=True, exist_ok=True)
        subject = email.subject if email.subject.lower().startswith("re:") else f"Re: {email.subject}"
        (outbox / f"reply-{email.id}.txt").write_text(
            f"To: {', '.join(to)}\nCc: {', '.join(cc)}\nSubject: {subject}\n\n{body}\n"
        )

    def mark_processed(self, email: Email, category: str) -> None:
        state = self._state()
        state[email.id] = category
        self.state_path.write_text(json.dumps(state, indent=2))
