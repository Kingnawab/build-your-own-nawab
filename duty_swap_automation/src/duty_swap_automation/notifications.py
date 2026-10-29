from __future__ import annotations

from duty_swap_automation.models import SwapOutcome, SwapRequest

SIGNATURE = "\n\n-- \nThis is an automated message from the duty roster assistant."


def describe(req: SwapRequest) -> str:
    duty_a = f" ({req.duty_a})" if req.duty_a else ""
    if req.request_type == "cover":
        return f"{req.person_b} covers {req.person_a}'s duty{duty_a} on {req.date_a:%A %d %B %Y}"
    duty_b = f" ({req.duty_b})" if req.duty_b else ""
    date_b = req.date_b or req.date_a
    return (
        f"{req.person_a}'s duty{duty_a} on {req.date_a:%A %d %B %Y} swapped with "
        f"{req.person_b}'s duty{duty_b} on {date_b:%A %d %B %Y}"
    )


def compose_reply(sender_name: str, outcomes: list[SwapOutcome]) -> str:
    greeting = f"Hi {sender_name.split()[0]}," if sender_name else "Hi,"
    applied = [o for o in outcomes if o.status in ("applied", "dry_run")]
    failed = [o for o in outcomes if o.status == "rejected"]
    unclear = [o for o in outcomes if o.status == "needs_clarification"]
    parts = [greeting, ""]

    if applied:
        parts.append("The following duty roster changes have been completed:")
        for o in applied:
            parts.append(f"  - {describe(o.request)}")
            for c in o.changes:
                parts.append(f"      row {c.row}: {c.old_value} -> {c.new_value}")
            for w in o.warnings:
                parts.append(f"      note: {w}")
        parts.append("")
    if failed:
        parts.append("These changes could not be made, so the roster was left as it was:")
        for o in failed:
            label = describe(o.request) if o.request and o.request.date_a else "Requested change"
            parts.append(f"  - {label}: {o.reason}")
        parts.append("")
    if unclear:
        parts.append("I need a bit more information before I can update the roster:")
        for o in unclear:
            parts.append(f"  - {o.reason}")
        parts.append("Please reply to this email with the details.")
        parts.append("")
    if failed and not unclear:
        parts.append("Please reply with corrected details, or contact the roster owner.")
    return "\n".join(parts).rstrip() + SIGNATURE
