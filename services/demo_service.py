"""Demo mode: one ready-made demo company, logged in automatically, with a
column layout and ledger names already set for the bundled sample Excel."""
from __future__ import annotations

import secrets
from pathlib import Path

from sqlalchemy.orm import Session

from database.models import ColumnMapping, Company, ImportBatch, PendingSend, SentVoucher, Setting, User
from services import settings_service
from utils.security import hash_password

ROOT = Path(__file__).resolve().parent.parent
SAMPLE_FILE = ROOT / "demo" / "Demo_Freight_Register.xlsx"
DEMO_COMPANY = "SN DEMO TRANSPORT"
DEMO_EMAIL = "demo@snsoftech.demo"

DEMO_MAPPING = {
    "date": "Date", "voucher_no": "BUILTY", "truck_no": "TRUCK", "owner": "OWNER",
    "shipment_no": "SHIPMENT NO", "from_place": "FROM", "to_place": "TO", "item_name": "ITEM",
    "weight": "WT", "rate": "RATE", "amount": "AMT", "tds": "TDS", "commission": "COMM",
    "pump_name": "PETROL PUMP", "diesel": "DIESEL", "slip_no": "SLIP NO",
    "cash_pump": "CASH FROM PUMP", "cash_office": "CASH FROM OFFICE", "balance": "TRANSPORT EXPENSES",
}
DEMO_NAMES = {"party": "DEMO CEMENT LTD", "tds": "TDS", "freight": "FREIGHT INCOME"}


def ensure(session: Session) -> tuple:
    """(company_id, user_id) of the demo company, creating it if needed."""
    company = session.query(Company).filter(Company.name == DEMO_COMPANY).one_or_none()
    if company is None:
        company = Company(name=DEMO_COMPANY, tally_company_name=DEMO_COMPANY)
        session.add(company)
        session.commit()
    user = session.query(User).filter(User.email == DEMO_EMAIL).one_or_none()
    if user is None:
        user = User(company_id=company.id, email=DEMO_EMAIL,
                    password_hash=hash_password(secrets.token_hex(16)), role="ADMIN")
        session.add(user)
        session.commit()
    if not settings_service.list_mappings(session, company.id):
        settings_service.save_mapping(session, company.id, "Demo Register", DEMO_MAPPING)
        settings_service.save_ledger_names(session, company.id, DEMO_NAMES)
    return company.id, user.id


def reset(session: Session) -> None:
    """Put the demo back to its starting state (settings, history, pending)."""
    company = session.query(Company).filter(Company.name == DEMO_COMPANY).one_or_none()
    if company is None:
        return
    for model in (SentVoucher, PendingSend, ImportBatch, ColumnMapping, Setting):
        session.query(model).filter(model.company_id == company.id).delete()
    session.commit()
    ensure(session)
