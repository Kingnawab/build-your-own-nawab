#!/usr/bin/env python3
"""One-command setup for sign-in-once mode. Run from this folder: python3 setup_wizard.py

Uses only the standard library so it can run before anything is installed.
"""
from __future__ import annotations

import getpass
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
ENV_PATH = ROOT / ".env"
EXAMPLE_PATH = ROOT / ".env.example"
IS_WINDOWS = os.name == "nt"
VENV_BIN = ROOT / ".venv" / ("Scripts" if IS_WINDOWS else "bin")
EXE = ".exe" if IS_WINDOWS else ""
TASK_NAME = "DutySwapAssistant"

GRAPH_API = "00000003-0000-0000-c000-000000000000"
DELEGATED_SCOPES = {
    "User.Read": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
    "Mail.ReadWrite": "024d486e-b451-40bb-833d-3e66d98c5c73",
    "Mail.Send": "e383f46e-2787-4529-855e-0e479a3ffac0",
    "Files.ReadWrite.All": "863451e7-0667-486c-a5d6-d135439485f0",
}

MANUAL_REGISTRATION = """
Register the app in your browser:
  1. Open https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
  2. New registration -> name "Duty Swap Assistant" -> "Accounts in this organizational directory only"
     -> leave Redirect URI empty -> Register
  3. Authentication -> Advanced settings -> Allow public client flows: Yes -> Save
  4. API permissions -> Add a permission -> Microsoft Graph -> Delegated permissions ->
     Mail.ReadWrite, Mail.Send, Files.ReadWrite.All, User.Read -> Add permissions
  5. Copy the "Application (client) ID" from the Overview page.
"""


def step(title: str) -> None:
    print(f"\n=== {title} ===")


def ask(prompt: str, default: str = "", secret: bool = False) -> str:
    suffix = f" [{default}]" if default and not secret else ""
    reader = getpass.getpass if secret else input
    value = reader(f"{prompt}{suffix}: ").strip()
    return value or default


def yes(prompt: str, default: bool = True) -> bool:
    answer = input(f"{prompt} [{'Y/n' if default else 'y/N'}]: ").strip().lower()
    return default if not answer else answer.startswith("y")


def read_env(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if path.exists():
        for line in path.read_text().splitlines():
            m = re.match(r"^([A-Z0-9_]+)=(.*)$", line)
            if m:
                values[m.group(1)] = m.group(2)
    return values


def write_env(updates: dict[str, str]) -> None:
    template = (ENV_PATH if ENV_PATH.exists() else EXAMPLE_PATH).read_text().splitlines()
    remaining = dict(updates)
    out = []
    for line in template:
        m = re.match(r"^([A-Z0-9_]+)=", line)
        if m and m.group(1) in remaining:
            out.append(f"{m.group(1)}={remaining.pop(m.group(1))}")
        else:
            out.append(line)
    out.extend(f"{k}={v}" for k, v in remaining.items())
    ENV_PATH.write_text("\n".join(out) + "\n")
    if not IS_WINDOWS:
        os.chmod(ENV_PATH, 0o600)


def run(cmd: list[str], check: bool = True, **kwargs) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, cwd=ROOT, check=check, **kwargs)


def install() -> None:
    step("1/6 Installing")
    if sys.version_info < (3, 10) or sys.version_info >= (3, 14):
        sys.exit(f"Python 3.10-3.13 is required (you have {platform.python_version()}). Install it from python.org.")
    if not (VENV_BIN / f"python{EXE}").exists():
        run([sys.executable, "-m", "venv", ".venv"])
    run([str(VENV_BIN / f"python{EXE}"), "-m", "pip", "install", "-q", "--upgrade", "pip"])
    run([str(VENV_BIN / f"python{EXE}"), "-m", "pip", "install", "-q", "-e", "."])
    print("Installed.")


def register_with_azure_cli() -> str | None:
    az = shutil.which("az")
    if not az or not yes("Azure CLI found. Register the Microsoft app automatically?"):
        return None
    if run([az, "account", "show"], check=False, capture_output=True).returncode != 0:
        print("Signing in to Azure CLI (a browser window or code will appear)...")
        if run([az, "login", "--allow-no-subscriptions"], check=False).returncode != 0:
            return None
    access = [{
        "resourceAppId": GRAPH_API,
        "resourceAccess": [{"id": i, "type": "Scope"} for i in DELEGATED_SCOPES.values()],
    }]
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as fh:
        json.dump(access, fh)
        perms = fh.name
    try:
        result = run(
            [az, "ad", "app", "create", "--display-name", "Duty Swap Assistant",
             "--sign-in-audience", "AzureADMyOrg", "--is-fallback-public-client", "true",
             "--required-resource-accesses", f"@{perms}", "--query", "appId", "-o", "tsv"],
            check=False, capture_output=True, text=True,
        )
    finally:
        os.unlink(perms)
    if result.returncode != 0 or not result.stdout.strip():
        print(f"Automatic registration failed:\n{result.stderr.strip()}")
        return None
    app_id = result.stdout.strip()
    print(f"Registered app {app_id}.")
    return app_id


def configure(existing: dict[str, str]) -> None:
    step("2/6 Microsoft app")
    app_id = existing.get("AZURE_CLIENT_ID", "")
    if app_id and yes(f"Use the existing app ID {app_id}?"):
        pass
    else:
        app_id = register_with_azure_cli() or ""
        if not app_id:
            print(MANUAL_REGISTRATION)
            while not re.fullmatch(r"[0-9a-fA-F-]{36}", app_id):
                app_id = ask("Paste the Application (client) ID")

    step("3/6 Settings")
    openai_key = existing.get("OPENAI_API_KEY", "")
    if not openai_key or not yes("Keep the saved OpenAI API key?"):
        openai_key = ask("OpenAI API key (input hidden; from https://platform.openai.com/api-keys)", secret=True)
    print("Where is the roster? Use the path inside your OneDrive, e.g. Documents/Duty Roster.xlsx")
    print("(For a SharePoint/Teams file, see ROSTER_SITE_ID in .env after setup.)")
    updates = {
        "AZURE_CLIENT_ID": app_id,
        "AZURE_CLIENT_SECRET": "",
        "GRAPH_USER": "",
        "OPENAI_API_KEY": openai_key,
        "DUTY_SWAP_BACKEND": "graph",
        "DRY_RUN": "true",
        "ROSTER_FILE_PATH": ask("Roster file path", existing.get("ROSTER_FILE_PATH", "Documents/Duty Roster.xlsx")),
        "ROSTER_SHEET": ask("Roster sheet (tab) name", existing.get("ROSTER_SHEET", "Roster")),
        "ROSTER_DATE_COLUMN": ask("Date column header", existing.get("ROSTER_DATE_COLUMN", "Date")),
        "ROSTER_DUTY_COLUMN": ask("Duty column header (blank if none)", existing.get("ROSTER_DUTY_COLUMN", "Duty")),
        "ROSTER_NAME_COLUMN": ask("Assigned-person column header", existing.get("ROSTER_NAME_COLUMN", "Assigned")),
        "NOTIFY_CC": ask("Always CC these emails on confirmations (comma-separated, optional)", existing.get("NOTIFY_CC", "")),
        "GRAPH_TOKEN_CACHE": str(ROOT / ".graph_token_cache.json"),
    }
    write_env(updates)
    print(f"Saved settings to {ENV_PATH}")


def sign_in() -> None:
    step("4/6 Sign in to Microsoft")
    print("Open the link below, enter the code, and sign in with your WORK account.")
    while run([str(VENV_BIN / f"authorize{EXE}")], check=False).returncode != 0:
        print(f"\nIf a check failed, fix the value in {ENV_PATH} (e.g. ROSTER_FILE_PATH) and save it.")
        if not yes("Try again?"):
            sys.exit("Stopped. Re-run `python3 setup_wizard.py` when ready; your answers are kept.")


def test_run() -> None:
    step("5/6 Test run (no changes are made)")
    run([str(VENV_BIN / f"kickoff{EXE}")], check=False)
    if yes("Does that look right? Switch on live mode (edit the roster and send emails)?", default=False):
        write_env({"DRY_RUN": "false"})
        print("Live mode on.")
    else:
        print("Staying in test mode. Set DRY_RUN=false in .env when you're happy.")


def schedule() -> None:
    step("6/6 Daily schedule")
    if not yes("Run automatically every weekday morning?"):
        return
    when = ask("Time (24h, your computer's time zone)", "07:00")
    hour, minute = (int(x) for x in when.split(":"))
    kickoff = VENV_BIN / f"kickoff{EXE}"
    log = ROOT / "run.log"
    system = platform.system()

    if system == "Windows":
        script = ROOT / "run_duty_swap.cmd"
        script.write_text(f'@echo off\r\ncd /d "{ROOT}"\r\n"{kickoff}" >> "{log}" 2>&1\r\n')
        run(["schtasks", "/Create", "/F", "/TN", TASK_NAME, "/SC", "WEEKLY", "/D", "MON,TUE,WED,THU,FRI",
             "/ST", f"{hour:02d}:{minute:02d}", "/TR", str(script)])
        print(f"Created Task Scheduler task '{TASK_NAME}'.")
    elif system == "Darwin":
        plist = Path.home() / "Library" / "LaunchAgents" / "com.dutyswap.assistant.plist"
        plist.parent.mkdir(parents=True, exist_ok=True)
        intervals = "".join(
            f"<dict><key>Weekday</key><integer>{d}</integer><key>Hour</key><integer>{hour}</integer>"
            f"<key>Minute</key><integer>{minute}</integer></dict>" for d in range(1, 6)
        )
        plist.write_text(f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
<key>Label</key><string>com.dutyswap.assistant</string>
<key>ProgramArguments</key><array><string>{kickoff}</string></array>
<key>WorkingDirectory</key><string>{ROOT}</string>
<key>StartCalendarInterval</key><array>{intervals}</array>
<key>StandardOutPath</key><string>{log}</string>
<key>StandardErrorPath</key><string>{log}</string>
</dict></plist>
""")
        run(["launchctl", "unload", str(plist)], check=False, capture_output=True)
        run(["launchctl", "load", str(plist)])
        print(f"Created launchd job {plist.name} (runs at the next wake if the Mac was asleep).")
    else:
        line = f'{minute} {hour} * * 1-5 cd "{ROOT}" && "{kickoff}" >> "{log}" 2>&1  # {TASK_NAME}'
        current = run(["crontab", "-l"], check=False, capture_output=True, text=True).stdout
        kept = [l for l in current.splitlines() if TASK_NAME not in l]
        run(["crontab", "-"], input="\n".join(kept + [line]) + "\n", text=True)
        print("Added cron job.")
    print(f"Output from each run is appended to {log}")


def main() -> None:
    os.chdir(ROOT)
    print("Duty Swap Assistant setup. Press Ctrl+C at any time to stop; re-running is safe.")
    install()
    configure(read_env(ENV_PATH))
    sign_in()
    test_run()
    schedule()
    print("\nAll done. Swap requests will be handled every weekday morning.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped.")
