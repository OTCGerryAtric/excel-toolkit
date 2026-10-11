"""Phase B request/callback correlation contract — pure logic, mirrored by PhaseBBridge.bas.

No Excel, no COM, no pywin32: this is the executable specification of how the VBA front end
decides whether an incoming callback is the current, authentic result for a request it sent.
``PhaseBBridge.PhaseB_OnResult`` implements exactly these rules; the offline tests exercise
this module, so the rules are verified even though VBA itself cannot be run here.

What it guards against:
  * a callback for a request we never issued                      -> unknown request;
  * a duplicate / late callback for one already answered          -> already <status>;
  * a STALE callback for an earlier request the user has since
    superseded (its id is not the current one)                    -> stale;
    so an old audit's findings can never overwrite a newer request's results;
  * a callback whose Excel identity (pid / workbook / full name)
    or correlation token does not match what we recorded          -> identity / token mismatch.

Only an "accepted" verdict for a response whose own ``status == "ok"`` may be rendered as a
successful audit; any other verdict, or ``status == "failed"``, must surface as an error and
never as an empty-but-successful result.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class PendingRequest:
    """What the front end records the moment it sends a request (before any callback)."""

    id: str
    token: str
    pid: int
    workbook: str
    full_name: str
    status: str = "pending"


def classify_callback(registry, current_id, *, id, token, pid, workbook, full_name) -> str:
    """Return "accepted" or a specific "rejected: ..." reason for an incoming callback.

    ``registry`` maps request id -> PendingRequest; ``current_id`` is the most recent request
    the front end issued (callbacks for anything older are stale).
    """
    rec = registry.get(id)
    if rec is None:
        return "rejected: unknown request"
    if rec.status != "pending":
        return f"rejected: already {rec.status}"
    if id != current_id:
        return "rejected: stale (superseded by a newer request)"
    if (pid, workbook, full_name) != (rec.pid, rec.workbook, rec.full_name):
        return "rejected: identity mismatch"
    if token != rec.token:
        return "rejected: correlation token mismatch"
    return "accepted"
