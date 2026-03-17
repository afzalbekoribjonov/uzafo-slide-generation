from __future__ import annotations


def parse_inviter_id(payload: str | None) -> int | None:
    if not payload:
        return None
    payload = payload.strip()
    if not payload.isdigit():
        return None
    return int(payload)
