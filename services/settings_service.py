"""Per-company key/value settings, and saved column-mapping profiles."""
from __future__ import annotations

import json
from typing import Dict, List, Optional

from sqlalchemy.orm import Session

from database.models import ColumnMapping, Setting


def get_setting(session: Session, company_id: int, key: str, default: str = "") -> str:
    row = (
        session.query(Setting)
        .filter(Setting.company_id == company_id, Setting.key == key)
        .one_or_none()
    )
    return row.value if row else default


def set_setting(session: Session, company_id: int, key: str, value: str) -> None:
    row = (
        session.query(Setting)
        .filter(Setting.company_id == company_id, Setting.key == key)
        .one_or_none()
    )
    if row:
        row.value = value
    else:
        session.add(Setting(company_id=company_id, key=key, value=value))
    session.commit()


LEDGER_NAME_KEYS = {
    "freight": "ledger_freight",
    "tds": "ledger_tds",
    "commission": "ledger_commission",
    "discount": "ledger_discount",
    "cash": "ledger_cash",
    "other1": "ledger_other1",
    "other2": "ledger_other2",
    "party": "ledger_party",
    "diesel_exp": "ledger_diesel_exp",
    "other_exp": "ledger_other_exp",
    "transport_exp": "ledger_transport_exp",
    "sales_mode": "sales_mode",
    "item_name": "sales_item_name",
    "item_unit": "sales_item_unit",
    "vchno_sales": "vchno_sales",
    "vchno_others": "vchno_others",
}


def get_posting_style(session: Session, company_id: int) -> str:
    from services.bilty_tally_service import DEFAULT_POSTING_STYLE, POSTING_STYLES

    value = get_setting(session, company_id, "posting_style", DEFAULT_POSTING_STYLE)
    return value if value in POSTING_STYLES else DEFAULT_POSTING_STYLE


def save_posting_style(session: Session, company_id: int, style: str) -> None:
    set_setting(session, company_id, "posting_style", style)


def get_ledger_names(session: Session, company_id: int) -> Dict[str, str]:
    from services.bilty_tally_service import DEFAULT_LEDGER_NAMES

    names = dict(DEFAULT_LEDGER_NAMES)
    for field_name, key in LEDGER_NAME_KEYS.items():
        value = get_setting(session, company_id, key, "")
        if value:
            names[field_name] = value
    from services.bilty_tally_service import POST_KINDS_KEY
    kinds = get_setting(session, company_id, "post_kinds", "")
    if kinds:
        names[POST_KINDS_KEY] = kinds
    return names


def save_post_kinds(session: Session, company_id: int, kinds) -> None:
    """Which transporter-style vouchers to post (codes from VOUCHER_KINDS).
    Stored as e.g. 'PAY,TDS,DSL,TRN,SAL'; 'NONE' when all are off."""
    set_setting(session, company_id, "post_kinds", ",".join(kinds) if kinds else "NONE")


def save_ledger_names(session: Session, company_id: int, names: Dict[str, str]) -> None:
    for field_name, key in LEDGER_NAME_KEYS.items():
        if field_name in names:
            set_setting(session, company_id, key, names[field_name])


def list_mappings(session: Session, company_id: int) -> List[ColumnMapping]:
    return (
        session.query(ColumnMapping)
        .filter(ColumnMapping.company_id == company_id)
        .order_by(ColumnMapping.created_at.desc())
        .all()
    )


def save_mapping(session: Session, company_id: int, name: str, mapping: Dict[str, str]) -> ColumnMapping:
    existing = (
        session.query(ColumnMapping)
        .filter(ColumnMapping.company_id == company_id, ColumnMapping.name == name)
        .one_or_none()
    )
    payload = json.dumps(mapping)
    if existing:
        existing.mapping_json = payload
        session.commit()
        return existing
    record = ColumnMapping(company_id=company_id, name=name, mapping_json=payload)
    session.add(record)
    session.commit()
    return record


def load_mapping(record: ColumnMapping) -> Dict[str, str]:
    return json.loads(record.mapping_json)


def get_tally_connection(session: Session, company_id: int) -> tuple:
    from services.tally_connector import DEFAULT_HOST, DEFAULT_PORT

    host = get_setting(session, company_id, "tally_host", DEFAULT_HOST) or DEFAULT_HOST
    try:
        port = int(get_setting(session, company_id, "tally_port", str(DEFAULT_PORT)))
    except ValueError:
        port = DEFAULT_PORT
    return host, port


def save_tally_connection(session: Session, company_id: int, host: str, port: int) -> None:
    set_setting(session, company_id, "tally_host", host.strip())
    set_setting(session, company_id, "tally_port", str(int(port)))


def delete_mapping(session: Session, company_id: int, name: str) -> None:
    row = (
        session.query(ColumnMapping)
        .filter(ColumnMapping.company_id == company_id, ColumnMapping.name == name)
        .one_or_none()
    )
    if row:
        session.delete(row)
        session.commit()


def all_mappings(session: Session, company_id: int) -> Dict[str, Dict[str, str]]:
    """{layout name: mapping}, newest first."""
    return {m.name: load_mapping(m) for m in list_mappings(session, company_id)}
