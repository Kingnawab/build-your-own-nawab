#!/usr/bin/env python
from __future__ import annotations

import json
import logging
import sys
from datetime import datetime
from typing import Callable, Optional

from dotenv import load_dotenv
from pydantic import BaseModel, Field

from crewai.flow import Flow, listen, start

from duty_swap_automation.config import Settings
from duty_swap_automation.mailbox import GraphMailbox, LocalMailbox, Mailbox
from duty_swap_automation.models import Email, ParsedEmail, SwapOutcome
from duty_swap_automation.notifications import compose_reply
from duty_swap_automation.roster import (
    GraphWorkbook,
    LocalWorkbook,
    Roster,
    SwapError,
    Workbook,
    apply_to_roster,
    load_roster,
    normalize_name,
    plan_changes,
)

log = logging.getLogger("duty_swap")

ParseFn = Callable[[Email, Roster, Settings], ParsedEmail]


def crew_parse(email: Email, roster: Roster, settings: Settings) -> ParsedEmail:
    from duty_swap_automation.crews.swap_parser_crew.swap_parser_crew import SwapParserCrew

    today = settings.today()
    result = SwapParserCrew().crew().kickoff(
        inputs={
            "today": today.isoformat(),
            "weekday": today.strftime("%A"),
            "people": ", ".join(roster.people) or "(none)",
            "duties": ", ".join(roster.duties) or "(not specified)",
            "sender_name": email.sender_name,
            "sender_email": email.sender_email,
            "cc": ", ".join(email.cc) or "(none)",
            "received_at": email.received_at.isoformat() if email.received_at else "",
            "subject": email.subject,
            "body": email.body[:8000],
        }
    )
    if result.pydantic is None:
        raise RuntimeError(f"Parser crew returned no structured output: {result.raw[:300]}")
    return result.pydantic  # type: ignore[return-value]


def build_services(settings: Settings) -> tuple[Mailbox, Workbook]:
    if settings.backend == "local":
        return LocalMailbox(settings), LocalWorkbook(settings)
    from duty_swap_automation.graph_client import GraphClient

    client = GraphClient(settings)
    return GraphMailbox(client, settings), GraphWorkbook(client, settings)


class DutySwapState(BaseModel):
    started_at: str = ""
    emails: list[Email] = Field(default_factory=list)
    parsed: dict[str, ParsedEmail] = Field(default_factory=dict)
    outcomes: list[SwapOutcome] = Field(default_factory=list)
    summary: str = ""


class DutySwapFlow(Flow[DutySwapState]):
    """Every morning: read inbox -> CrewAI extracts swap requests -> validate + edit roster -> email confirmations."""

    def __init__(
        self,
        settings: Optional[Settings] = None,
        mailbox: Optional[Mailbox] = None,
        workbook: Optional[Workbook] = None,
        parse_fn: Optional[ParseFn] = None,
        **kwargs,
    ):
        super().__init__(**kwargs)
        self._settings = settings
        self._mailbox = mailbox
        self._workbook = workbook
        self._parse_fn = parse_fn or crew_parse
        self._roster: Optional[Roster] = None

    @start()
    def fetch_emails(self, crewai_trigger_payload: dict = None):
        if self._settings is None:
            self._settings = Settings.from_env()
        if crewai_trigger_payload and "dry_run" in crewai_trigger_payload:
            self._settings.dry_run = bool(crewai_trigger_payload["dry_run"])
        if self._mailbox is None or self._workbook is None:
            mailbox, workbook = build_services(self._settings)
            self._mailbox = self._mailbox or mailbox
            self._workbook = self._workbook or workbook

        self.state.started_at = datetime.now(self._settings.tz).isoformat(timespec="seconds")
        self.state.emails = self._mailbox.fetch_candidates()
        self._roster = load_roster(self._workbook, self._settings)
        log.info("Found %d candidate email(s); roster has %d rows", len(self.state.emails), len(self._roster.rows))

    @listen(fetch_emails)
    def parse_emails(self):
        for email in self.state.emails:
            try:
                self.state.parsed[email.id] = self._parse_fn(email, self._roster, self._settings)
            except Exception:
                log.exception("Could not parse email %s (%s); it stays unread for the next run", email.id, email.subject)

    def _authorized(self, email: Email, people: list[str]) -> bool:
        s = self._settings
        if not s.require_requester_is_party or email.sender_email.lower() in s.manager_emails:
            return True
        sender_names = {normalize_name(email.sender_name)}
        mapped = self._roster.name_for_email(email.sender_email)
        if mapped:
            sender_names.add(normalize_name(mapped))
        return any(normalize_name(p) in sender_names for p in people if p)

    def _recipients(self, email: Email, outcomes: list[SwapOutcome]) -> tuple[list[str], list[str]]:
        to = [email.sender_email]
        cc: list[str] = []
        for o in outcomes:
            if o.status not in ("applied", "dry_run") or not o.request:
                continue
            for person in (o.request.person_a, o.request.person_b):
                addr = self._roster.email_for(person or "")
                if addr:
                    cc.append(addr)
        cc.extend(self._settings.notify_cc)
        seen = {to[0].lower()}
        deduped = []
        for a in cc:
            if a.lower() not in seen:
                seen.add(a.lower())
                deduped.append(a)
        return to, deduped

    @listen(parse_emails)
    def apply_and_notify(self):
        s = self._settings
        today = s.today()
        for email in self.state.emails:
            parsed = self.state.parsed.get(email.id)
            if parsed is None or not parsed.is_swap_request:
                continue

            outcomes: list[SwapOutcome] = []
            if parsed.needs_clarification or not parsed.requests:
                outcomes.append(SwapOutcome(
                    email_id=email.id, status="needs_clarification",
                    reason=parsed.clarification_question or "Please confirm who is swapping and on which dates.",
                ))
            for req in parsed.requests:
                if not self._authorized(email, [req.person_a, req.person_b]):
                    outcomes.append(SwapOutcome(
                        email_id=email.id, request=req, status="rejected",
                        reason="Only someone involved in the swap (or a roster manager) can request it.",
                    ))
                    continue
                try:
                    changes, warnings = plan_changes(self._roster, req, today, s.allow_past_dates)
                except SwapError as exc:
                    outcomes.append(SwapOutcome(email_id=email.id, request=req, status="rejected", reason=str(exc)))
                    continue
                if not s.dry_run:
                    self._workbook.write_cells(changes)
                    if s.log_sheet:
                        self._workbook.append_row(
                            s.log_sheet,
                            ["Processed At", "Requested By", "Change", "Cells", "Email Subject"],
                            [
                                datetime.now(s.tz).strftime("%Y-%m-%d %H:%M"),
                                email.sender_email,
                                req.summary or f"{req.request_type}: {req.person_a} <-> {req.person_b}",
                                "; ".join(f"row {c.row}: {c.old_value} -> {c.new_value}" for c in changes),
                                email.subject,
                            ],
                        )
                apply_to_roster(self._roster, changes)
                outcomes.append(SwapOutcome(
                    email_id=email.id, request=req, status="dry_run" if s.dry_run else "applied",
                    changes=changes, warnings=warnings,
                ))

            to, cc = self._recipients(email, outcomes)
            body = compose_reply(email.sender_name, outcomes)
            for o in outcomes:
                o.notify = to + cc
            self.state.outcomes.extend(outcomes)
            if s.dry_run:
                log.info("[dry run] would reply to %s cc %s:\n%s", to, cc, body)
                continue
            self._mailbox.reply(email, to, cc, body)
            all_done = all(o.status == "applied" for o in outcomes)
            self._mailbox.mark_processed(email, s.processed_category if all_done else s.attention_category)

    @listen(apply_and_notify)
    def finish(self):
        if self._workbook is not None:
            self._workbook.close()
        counts: dict[str, int] = {}
        for o in self.state.outcomes:
            counts[o.status] = counts.get(o.status, 0) + 1
        swap_emails = sum(1 for p in self.state.parsed.values() if p.is_swap_request)
        self.state.summary = (
            f"Checked {len(self.state.emails)} email(s), {swap_emails} contained swap requests. "
            + (", ".join(f"{k}: {v}" for k, v in sorted(counts.items())) or "No roster changes needed.")
        )
        print(self.state.summary)
        return self.state.summary


def _setup() -> None:
    load_dotenv()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")


def kickoff():
    _setup()
    DutySwapFlow().kickoff()


def plot():
    DutySwapFlow().plot()


def run_with_trigger():
    _setup()
    if len(sys.argv) < 2:
        raise Exception("No trigger payload provided. Please provide JSON payload as argument.")
    try:
        trigger_payload = json.loads(sys.argv[1])
    except json.JSONDecodeError:
        raise Exception("Invalid JSON payload provided as argument")
    DutySwapFlow().kickoff({"crewai_trigger_payload": trigger_payload})


if __name__ == "__main__":
    kickoff()
