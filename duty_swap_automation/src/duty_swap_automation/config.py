from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import date, datetime
from zoneinfo import ZoneInfo


def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def _env_bool(name: str, default: bool) -> bool:
    raw = _env(name)
    if not raw:
        return default
    return raw.lower() in {"1", "true", "yes", "y", "on"}


def _env_list(name: str, default: str = "") -> list[str]:
    return [item.strip() for item in _env(name, default).split(",") if item.strip()]


@dataclass
class Settings:
    # "graph" talks to Microsoft 365; "local" uses a folder of JSON emails + a local .xlsx (for testing).
    backend: str = "graph"
    dry_run: bool = False
    timezone: str = "America/Los_Angeles"

    # Microsoft Graph auth. With a client secret the app runs unattended (app-only) against GRAPH_USER;
    # without one it uses delegated device-code sign-in cached in GRAPH_TOKEN_CACHE.
    tenant_id: str = ""
    client_id: str = ""
    client_secret: str = ""
    graph_user: str = ""
    token_cache_path: str = ".graph_token_cache.json"

    # Where the roster workbook lives. Use one of: drive id + item id, site id + path, or a path in GRAPH_USER's OneDrive.
    roster_drive_id: str = ""
    roster_item_id: str = ""
    roster_site_id: str = ""
    roster_file_path: str = ""
    local_workbook_path: str = "sample_data/duty_roster.xlsx"
    local_mail_dir: str = "sample_data/mail"

    roster_sheet: str = "Roster"
    date_column: str = "Date"
    name_column: str = "Assigned"
    duty_column: str = "Duty"
    staff_sheet: str = "Staff"
    staff_name_column: str = "Name"
    staff_email_column: str = "Email"
    log_sheet: str = "Swap Log"

    lookback_hours: int = 72
    subject_keywords: list[str] = field(default_factory=list)
    processed_category: str = "Duty Swap Processed"
    attention_category: str = "Duty Swap Needs Attention"

    require_requester_is_party: bool = True
    manager_emails: list[str] = field(default_factory=list)
    notify_cc: list[str] = field(default_factory=list)
    allow_past_dates: bool = False

    @classmethod
    def from_env(cls) -> "Settings":
        return cls(
            backend=_env("DUTY_SWAP_BACKEND", "graph").lower(),
            dry_run=_env_bool("DRY_RUN", False),
            timezone=_env("TIMEZONE", "America/Los_Angeles"),
            tenant_id=_env("AZURE_TENANT_ID"),
            client_id=_env("AZURE_CLIENT_ID"),
            client_secret=_env("AZURE_CLIENT_SECRET"),
            graph_user=_env("GRAPH_USER"),
            token_cache_path=_env("GRAPH_TOKEN_CACHE", ".graph_token_cache.json"),
            roster_drive_id=_env("ROSTER_DRIVE_ID"),
            roster_item_id=_env("ROSTER_ITEM_ID"),
            roster_site_id=_env("ROSTER_SITE_ID"),
            roster_file_path=_env("ROSTER_FILE_PATH"),
            local_workbook_path=_env("LOCAL_WORKBOOK_PATH", "sample_data/duty_roster.xlsx"),
            local_mail_dir=_env("LOCAL_MAIL_DIR", "sample_data/mail"),
            roster_sheet=_env("ROSTER_SHEET", "Roster"),
            date_column=_env("ROSTER_DATE_COLUMN", "Date"),
            name_column=_env("ROSTER_NAME_COLUMN", "Assigned"),
            duty_column=_env("ROSTER_DUTY_COLUMN", "Duty"),
            staff_sheet=_env("STAFF_SHEET", "Staff"),
            staff_name_column=_env("STAFF_NAME_COLUMN", "Name"),
            staff_email_column=_env("STAFF_EMAIL_COLUMN", "Email"),
            log_sheet=_env("SWAP_LOG_SHEET", "Swap Log"),
            lookback_hours=int(_env("LOOKBACK_HOURS", "72")),
            subject_keywords=_env_list("SUBJECT_KEYWORDS"),
            processed_category=_env("PROCESSED_CATEGORY", "Duty Swap Processed"),
            attention_category=_env("ATTENTION_CATEGORY", "Duty Swap Needs Attention"),
            require_requester_is_party=_env_bool("REQUIRE_REQUESTER_IS_PARTY", True),
            manager_emails=[e.lower() for e in _env_list("MANAGER_EMAILS")],
            notify_cc=_env_list("NOTIFY_CC"),
            allow_past_dates=_env_bool("ALLOW_PAST_DATES", False),
        )

    @property
    def tz(self) -> ZoneInfo:
        return ZoneInfo(self.timezone)

    def today(self) -> date:
        return datetime.now(self.tz).date()
