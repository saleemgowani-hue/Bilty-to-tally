"""Direct send to Tally with Replace and Undo.

Every voucher posted straight into Tally carries a REMOTEID (see
bilty_tally_service.voucher_remote_id) and is recorded in SentVoucher.
That record is what makes these safe:

* Replace - a corrected Excel re-sent for bilties that were sent before:
  the old vouchers of those bilties are deleted from Tally first, then the
  corrected ones created. No duplicates.
* Skip    - alternatively, bilties already in Tally are left alone and only
  new bilties are sent.
* Undo    - delete every voucher of one earlier send from Tally.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from sqlalchemy.orm import Session

from database.models import Company, ImportBatch, SentVoucher
from services import tally_connector as tc
from services.bilty_tally_service import (
    ProcessedRow,
    bilty_keys,
    build_transporter_xml,
    company_tag,
    enabled_kinds,
    transporter_sent_list,
    voucher_remote_id,
)
from utils.config import settings


@dataclass
class SendOutcome:
    created: int = 0
    errors: int = 0
    ledgers_created: int = 0
    altered: int = 0
    raw: str = ""
    replaced_deleted: int = 0
    skipped_bilties: int = 0
    batch_id: Optional[int] = None
    messages: List[str] = field(default_factory=list)


def _now() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc)


def _row_key(r: ProcessedRow, ctag: str) -> str:
    return voucher_remote_id(ctag, r, "X")[:-2]


def already_sent(session: Session, company_id: int, rows: List[ProcessedRow],
                 tally_company_name: str) -> Dict[str, List[SentVoucher]]:
    """For each usable row already in Tally (sent before, not undone):
    row key -> its live SentVoucher records."""
    keys = set(bilty_keys(rows, tally_company_name))
    if not keys:
        return {}
    ctag = company_tag(tally_company_name)
    live = (
        session.query(SentVoucher)
        .filter(SentVoucher.company_id == company_id,
                SentVoucher.deleted_at.is_(None),
                SentVoucher.remote_id.like(f"SNBT-{ctag}-%"))
        .all()
    )
    out: Dict[str, List[SentVoucher]] = {}
    for sv in live:
        key = sv.remote_id.rsplit("-", 1)[0]
        if key in keys:
            out.setdefault(key, []).append(sv)
    return out


def send(session: Session, company: Company, rows: List[ProcessedRow],
         ledger_names: Dict[str, str], filename: str, host: str, port: int,
         mode: str = "replace", skipped_count: int = 0, mismatch_count: int = 0) -> SendOutcome:
    """mode: 'replace' (re-sent bilties overwrite their old vouchers) or
    'skip' (bilties already in Tally are not sent again)."""
    tally_name = company.tally_company_name or ""
    ok, message = tc.check_connection(host, port, tally_name)
    if not ok:
        raise tc.TallyConnectionError(message)

    usable = [r for r in rows if not r.problem]
    ctag = company_tag(tally_name)
    previous = already_sent(session, company.id, usable, tally_name)
    outcome = SendOutcome()

    if previous and mode == "skip":
        outcome.skipped_bilties = len(previous)
        usable = [r for r in usable if _row_key(r, ctag) not in previous]
    elif previous:
        # Only vouchers of the kinds being posted now are replaced; a kind
        # switched off in Settings is left in Tally as it is.
        kinds = enabled_kinds(ledger_names)
        old = [sv for svs in previous.values() for sv in svs if sv.remote_id.rsplit("-", 1)[1] in kinds]
        result = tc.delete_vouchers(
            [{"remote_id": sv.remote_id, "vch_type": sv.vch_type, "vch_date": sv.vch_date} for sv in old],
            host, port, tally_name,
        )
        if old and result.deleted == 0:
            raise tc.TallyConnectionError(
                f"{len(previous)} bilty(s) in this file were sent to Tally before, but Tally "
                "did not delete their old vouchers, so nothing was sent (to avoid duplicates). "
                "If they really are in Tally, delete them in Day Book yourself. If they are NOT in "
                "Tally, go to History → that send → 'Forget this send', then send again."
                + (" Tally said: " + "; ".join(result.messages[:3]) if result.messages else "")
            )
        when = _now()
        for sv in old:
            sv.deleted_at = when
        session.commit()
        outcome.replaced_deleted = result.deleted
        if result.deleted < len(old):
            outcome.messages.append(
                f"Only {result.deleted} of {len(old)} old vouchers were found in Tally to delete "
                "(some may already have been deleted by hand)."
            )

    if not usable or not transporter_sent_list(usable, tally_name, ledger_names):
        outcome.messages.append("Nothing to send - no voucher types are switched on, or no usable rows.")
        return outcome

    xml = build_transporter_xml(usable, tally_name, ledger_names)
    result = tc.push_xml(xml, host, port, tally_name)
    outcome.created = result.created
    outcome.altered = result.altered
    outcome.raw = result.raw
    outcome.errors = result.errors
    outcome.ledgers_created = result.ledgers_created
    outcome.messages.extend(result.messages)

    settings.EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    path = settings.EXPORT_DIR / f"SentToTally_{dt.datetime.now():%Y%m%d_%H%M%S}.xml"
    path.write_text(xml, encoding="utf-8")
    posted = (result.created + result.altered) > 0
    batch = ImportBatch(
        company_id=company.id,
        filename=f"{filename} (sent to Tally)" if posted else f"{filename} (Tally made NO entries)",
        row_count=len(usable),
        skipped_count=skipped_count,
        mismatch_count=mismatch_count,
        total_amount=f"{sum(r.amount for r in usable):.2f}",
        xml_path=str(path),
    )
    session.add(batch)
    session.commit()
    # Only remember vouchers as "in Tally" when Tally actually took them -
    # otherwise a later send would wrongly skip/replace bilties that are not there.
    if posted:
        for item in transporter_sent_list(usable, tally_name, ledger_names):
            session.add(SentVoucher(company_id=company.id, batch_id=batch.id, **item))
        session.commit()
    outcome.batch_id = batch.id
    return outcome


def live_vouchers(session: Session, batch_id: int) -> List[SentVoucher]:
    return (
        session.query(SentVoucher)
        .filter(SentVoucher.batch_id == batch_id, SentVoucher.deleted_at.is_(None))
        .all()
    )


def undo_batch(session: Session, company: Company, batch_id: int, host: str, port: int) -> tc.PushResult:
    """Delete every still-live voucher of one earlier send from Tally."""
    vouchers = live_vouchers(session, batch_id)
    result = tc.delete_vouchers(
        [{"remote_id": v.remote_id, "vch_type": v.vch_type, "vch_date": v.vch_date} for v in vouchers],
        host, port, company.tally_company_name or "",
    )
    if result.deleted:
        when = _now()
        for v in vouchers:
            v.deleted_at = when
        session.commit()
    return result


def forget_batch(session: Session, batch_id: int) -> int:
    """Mark one send's vouchers as not in Tally, without contacting Tally -
    for when they were never really created or were deleted by hand."""
    vouchers = live_vouchers(session, batch_id)
    when = _now()
    for v in vouchers:
        v.deleted_at = when
    session.commit()
    return len(vouchers)
