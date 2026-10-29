from __future__ import annotations

from datetime import date, datetime
from typing import Literal, Optional

from pydantic import BaseModel, Field


class Email(BaseModel):
    id: str
    subject: str = ""
    sender_name: str = ""
    sender_email: str = ""
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    received_at: Optional[datetime] = None
    body: str = ""
    categories: list[str] = Field(default_factory=list)


class SwapRequest(BaseModel):
    """One requested roster change, as extracted from an email by the parser crew."""

    request_type: Literal["swap", "cover"] = Field(
        description="'swap' = two people exchange duties; 'cover' = person_b takes person_a's duty, nothing in return."
    )
    person_a: Optional[str] = Field(None, description="Person currently rostered on date_a who is giving up that duty.")
    date_a: Optional[date] = Field(None, description="Date (YYYY-MM-DD) of person_a's duty.")
    duty_a: Optional[str] = Field(None, description="Duty name for person_a's shift if the email mentions it.")
    person_b: Optional[str] = Field(None, description="The other person (takes date_a; for a swap, gives up date_b).")
    date_b: Optional[date] = Field(
        None, description="Date (YYYY-MM-DD) of person_b's duty being given to person_a. Null for 'cover'."
    )
    duty_b: Optional[str] = Field(None, description="Duty name for person_b's shift if the email mentions it.")
    summary: str = Field("", description="One-sentence plain-English description of the change.")


class ParsedEmail(BaseModel):
    is_swap_request: bool = Field(description="True only if the email asks for a duty swap/cover on the roster.")
    requests: list[SwapRequest] = Field(default_factory=list)
    needs_clarification: bool = Field(
        False, description="True if it is a swap request but names or dates are missing/ambiguous."
    )
    clarification_question: Optional[str] = Field(
        None, description="What to ask the sender if needs_clarification is true."
    )


class CellChange(BaseModel):
    sheet: str
    row: int
    col: int
    old_value: str
    new_value: str


class SwapOutcome(BaseModel):
    email_id: str
    request: Optional[SwapRequest] = None
    status: Literal["applied", "rejected", "needs_clarification", "dry_run"]
    reason: str = ""
    changes: list[CellChange] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    notify: list[str] = Field(default_factory=list)
