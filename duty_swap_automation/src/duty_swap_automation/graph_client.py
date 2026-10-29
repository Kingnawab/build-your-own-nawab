from __future__ import annotations

import os
import time
from typing import Any
from urllib.parse import quote

import msal
import requests

from duty_swap_automation.config import Settings

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DELEGATED_SCOPES = ["Mail.ReadWrite", "Mail.Send", "Files.ReadWrite.All", "User.Read"]
APP_SCOPES = ["https://graph.microsoft.com/.default"]


class GraphError(RuntimeError):
    pass


class GraphClient:
    def __init__(self, settings: Settings):
        if not settings.client_id:
            raise GraphError("AZURE_CLIENT_ID must be set for the graph backend.")
        self.settings = settings
        self.app_only = bool(settings.client_secret)
        if self.app_only and not (settings.tenant_id and settings.graph_user):
            raise GraphError("AZURE_TENANT_ID and GRAPH_USER (e.g. you@company.com) are required with AZURE_CLIENT_SECRET.")
        self._session = requests.Session()
        self._token: str | None = None
        self._token_expires = 0.0
        self._cache = msal.SerializableTokenCache()
        # "organizations" lets any work/school account sign in without knowing the tenant id.
        authority = f"https://login.microsoftonline.com/{settings.tenant_id or 'organizations'}"
        if self.app_only:
            self._app: msal.ClientApplication = msal.ConfidentialClientApplication(
                settings.client_id, authority=authority, client_credential=settings.client_secret
            )
        else:
            if os.path.exists(settings.token_cache_path):
                with open(settings.token_cache_path) as fh:
                    self._cache.deserialize(fh.read())
            self._app = msal.PublicClientApplication(
                settings.client_id, authority=authority, token_cache=self._cache
            )

    @property
    def user_base(self) -> str:
        if self.app_only:
            return f"{GRAPH_ROOT}/users/{quote(self.settings.graph_user)}"
        return f"{GRAPH_ROOT}/me"

    def _save_cache(self) -> None:
        if not self.app_only and self._cache.has_state_changed:
            with open(self.settings.token_cache_path, "w") as fh:
                fh.write(self._cache.serialize())
            os.chmod(self.settings.token_cache_path, 0o600)

    def _acquire_token(self, interactive: bool = False) -> str:
        if self._token and time.time() < self._token_expires - 60:
            return self._token
        if self.app_only:
            result = self._app.acquire_token_for_client(scopes=APP_SCOPES)
        else:
            result = None
            accounts = self._app.get_accounts()
            if accounts:
                result = self._app.acquire_token_silent(DELEGATED_SCOPES, account=accounts[0])
            if not result:
                if not interactive:
                    raise GraphError(
                        "No cached Microsoft sign-in. Run `authorize` once (or set AZURE_CLIENT_SECRET for app-only)."
                    )
                flow = self._app.initiate_device_flow(scopes=DELEGATED_SCOPES)
                if "user_code" not in flow:
                    raise GraphError(f"Could not start device sign-in: {flow}")
                print(flow["message"], flush=True)
                result = self._app.acquire_token_by_device_flow(flow)
            self._save_cache()
        if not result or "access_token" not in result:
            raise GraphError(f"Token acquisition failed: {result.get('error_description') if result else result}")
        self._token = result["access_token"]
        self._token_expires = time.time() + int(result.get("expires_in", 3600))
        return self._token

    def request(
        self,
        method: str,
        url: str,
        *,
        json: Any = None,
        params: dict | None = None,
        headers: dict | None = None,
        expect_json: bool = True,
    ) -> Any:
        if not url.startswith("http"):
            url = f"{GRAPH_ROOT}{url}"
        for attempt in range(5):
            hdrs = {"Authorization": f"Bearer {self._acquire_token()}"}
            if headers:
                hdrs.update(headers)
            resp = self._session.request(method, url, json=json, params=params, headers=hdrs, timeout=60)
            if resp.status_code in (429, 500, 502, 503, 504):
                time.sleep(float(resp.headers.get("Retry-After", 2 ** attempt)))
                continue
            if resp.status_code >= 400:
                raise GraphError(f"{method} {url} -> {resp.status_code}: {resp.text[:500]}")
            if not expect_json or resp.status_code == 204 or not resp.content:
                return resp
            return resp.json()
        raise GraphError(f"{method} {url} kept failing after retries")


def authorize_cli() -> None:
    """One-time delegated sign-in that caches a refresh token for unattended daily runs."""
    from dotenv import load_dotenv  # crewai depends on python-dotenv

    load_dotenv()
    settings = Settings.from_env()
    if settings.client_secret:
        print("AZURE_CLIENT_SECRET is set, so the app runs app-only; no interactive sign-in needed.")
        return
    client = GraphClient(settings)
    client._acquire_token(interactive=True)
    me = client.request("GET", "/me", params={"$select": "displayName,mail,userPrincipalName"})
    print(f"Signed in as {me.get('displayName')} <{me.get('mail') or me.get('userPrincipalName')}>")
    print(f"Token cache saved to {os.path.abspath(settings.token_cache_path)}")

    from duty_swap_automation.mailbox import GraphMailbox
    from duty_swap_automation.roster import GraphWorkbook, load_roster

    ok = True
    try:
        emails = GraphMailbox(client, settings).fetch_candidates()
        print(f"[ok] Inbox readable: {len(emails)} matching unread email(s) in the last {settings.lookback_hours}h")
    except Exception as exc:
        ok = False
        print(f"[fail] Could not read inbox: {exc}")
    workbook = None
    try:
        workbook = GraphWorkbook(client, settings)
        roster = load_roster(workbook, settings)
        print(f"[ok] Roster '{settings.roster_sheet}' readable: {len(roster.rows)} duty rows, "
              f"{len(roster.people)} people, {len(roster.staff_emails)} staff emails")
    except Exception as exc:
        ok = False
        print(f"[fail] Could not read roster workbook: {exc}")
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass
    if not ok:
        raise SystemExit("Fix the items above (usually ROSTER_* settings in .env), then run `authorize` again.")
    print("Setup looks good. Try `DRY_RUN=true kickoff` next.")
