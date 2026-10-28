"""Pending sends: entries saved while Tally was closed, sent later with
one click. Rows are stored with their corrected values and party, then
rebuilt with the ledger names that are current when they are sent."""
from __future__ import annotations

import datetime as dt
import json
from typing import Dict, List, Optional

from sqlalchemy.orm import Session

from database.models import Company, PendingSend
from services import tally_connector as tc
from services import tally_sync_service as sync
from services.bilty_tally_service import DEFAULT_LEDGER_NAMES, ProcessedRow, build_row, number_occurrences


def _jsonable(value):
    if isinstance(value, dt.datetime):
        return {"__dt__": value.isoformat()}
    if isinstance(value, dt.date):
        return {"__dt__": dt.datetime(value.year, value.month, value.day).isoformat()}
    if isinstance(value, float) and value != value:
        return None
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _from_json(value):
    if isinstance(value, dict) and "__dt__" in value:
        return dt.datetime.fromisoformat(value["__dt__"])
    return value


def rows_to_json(rows: List[ProcessedRow]) -> str:
    out = []
    for r in rows:
        if r.problem:
            continue
        out.append({
            "source_row": r.source_row,
            "source": r.source,
            "party": r.party,
            "edited": r.edited,
            "raw": {k: _jsonable(v) for k, v in (r.raw or {}).items()},
        })
    return json.dumps(out)


def rows_from_json(text: str, ledger_names: Dict[str, str], style: str = "transporter") -> List[ProcessedRow]:
    rows = []
    for item in json.loads(text):
        values = {k: _from_json(v) for k, v in item["raw"].items()}
        names = {**DEFAULT_LEDGER_NAMES, **ledger_names, "party": item.get("party", "")}
        row = build_row(item["source_row"], values, names, style, edited=item.get("edited", False))
        row.source = item.get("source", "")
        rows.append(row)
    number_occurrences(rows)
    return rows


def save(session: Session, company_id: int, rows: List[ProcessedRow], label: str, mode: str) -> PendingSend:
    usable = [r for r in rows if not r.problem]
    item = PendingSend(
        company_id=company_id,
        label=label[:300],
        rows_json=rows_to_json(usable),
        row_count=len(usable),
        total_amount=f"{sum(r.amount for r in usable):.2f}",
        mode=mode,
    )
    session.add(item)
    session.commit()
    return item


def waiting(session: Session, company_id: int) -> List[PendingSend]:
    return (
        session.query(PendingSend)
        .filter(PendingSend.company_id == company_id, PendingSend.status == "pending")
        .order_by(PendingSend.created_at)
        .all()
    )


def count_waiting(session: Session, company_id: int) -> int:
    return (
        session.query(PendingSend)
        .filter(PendingSend.company_id == company_id, PendingSend.status == "pending")
        .count()
    )


def cancel(session: Session, item: PendingSend) -> None:
    item.status = "cancelled"
    session.commit()


def send_one(session: Session, company: Company, item: PendingSend, ledger_names: Dict[str, str],
             host: str, port: int) -> sync.SendOutcome:
    """Send one pending item. Raises tc.TallyConnectionError (item stays
    pending) when Tally cannot be reached or refuses a replace."""
    rows = rows_from_json(item.rows_json, ledger_names)
    problems = [r for r in rows if r.problem]
    try:
        out = sync.send(session, company, rows, ledger_names, item.label, host, port,
                        mode=item.mode, skipped_count=len(problems))
    except tc.TallyConnectionError as exc:
        item.last_error = str(exc)[:2000]
        session.commit()
        raise
    if (out.created + out.altered) > 0 or out.skipped_bilties or out.replaced_deleted:
        item.status = "sent"
        item.sent_at = dt.datetime.now(dt.timezone.utc)
        item.batch_id = out.batch_id
        item.last_error = ""
    else:
        item.last_error = "Tally made no entries. " + "; ".join(out.messages[:3])
    session.commit()
    return out
