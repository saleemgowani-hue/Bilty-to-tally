"""SQLAlchemy models for the Freight Bilty-to-Tally Agent.

One company per installation (same model as the other SN Softech desktop
products): a company signs up, activates a licence key, saves a column
mapping for their Excel register once, then every later upload reuses it.
"""
from __future__ import annotations

import datetime as dt

from sqlalchemy import (
    Boolean, Column, DateTime, ForeignKey, Integer, String, Text,
    UniqueConstraint,
)
from sqlalchemy.orm import declarative_base, relationship

Base = declarative_base()


def _now() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc)


class Company(Base):
    __tablename__ = "companies"

    id = Column(Integer, primary_key=True)
    name = Column(String(200), nullable=False)
    tally_company_name = Column(String(200), default="")
    created_at = Column(DateTime, default=_now)

    license_plan = Column(String(20))
    license_expires_at = Column(DateTime)
    license_key_display = Column(String(40))

    users = relationship("User", back_populates="company", cascade="all, delete-orphan")


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    email = Column(String(200), nullable=False, unique=True)
    password_hash = Column(String(200), nullable=False)
    role = Column(String(20), default="ADMIN")
    created_at = Column(DateTime, default=_now)

    company = relationship("Company", back_populates="users")


class License(Base):
    __tablename__ = "licenses"
    __table_args__ = (UniqueConstraint("key_hash", name="uq_license_key_hash"),)

    id = Column(Integer, primary_key=True)
    key_hash = Column(String(64), nullable=False)
    key_display = Column(String(40), nullable=False)
    plan = Column(String(20), nullable=False)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    activated_at = Column(DateTime, default=_now)
    expires_at = Column(DateTime, nullable=False)
    activated_by = Column(Integer, ForeignKey("users.id"), nullable=True)


class Setting(Base):
    """Generic per-company key/value store: ledger name overrides, etc."""
    __tablename__ = "settings"
    __table_args__ = (UniqueConstraint("company_id", "key", name="uq_setting_company_key"),)

    id = Column(Integer, primary_key=True)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    key = Column(String(100), nullable=False)
    value = Column(Text, default="")


class ColumnMapping(Base):
    """A saved Excel column -> field mapping, so a customer maps their
    register layout once and reuses it on every later upload."""
    __tablename__ = "column_mappings"

    id = Column(Integer, primary_key=True)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    name = Column(String(100), nullable=False, default="Default")
    mapping_json = Column(Text, nullable=False)   # {"field": "Excel header or column letter"}
    created_at = Column(DateTime, default=_now)


class ImportBatch(Base):
    """One uploaded Excel file, processed into Tally vouchers."""
    __tablename__ = "import_batches"

    id = Column(Integer, primary_key=True)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    filename = Column(String(300), nullable=False)
    mapping_id = Column(Integer, ForeignKey("column_mappings.id"), nullable=True)
    row_count = Column(Integer, default=0)
    skipped_count = Column(Integer, default=0)
    mismatch_count = Column(Integer, default=0)
    total_amount = Column(String(40), default="0")
    xml_path = Column(String(500), default="")           # combined (both voucher types)
    sales_xml_path = Column(String(500), default="")      # Sales vouchers only
    journal_xml_path = Column(String(500), default="")    # Journal vouchers only
    created_at = Column(DateTime, default=_now)


class SentVoucher(Base):
    """Every voucher this app posted straight into Tally, keyed by the
    REMOTEID it was tagged with - so it can later be replaced (corrected
    Excel re-sent) or undone (deleted from Tally) without duplicates."""
    __tablename__ = "sent_vouchers"

    id = Column(Integer, primary_key=True)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    batch_id = Column(Integer, ForeignKey("import_batches.id"), nullable=True)
    remote_id = Column(String(200), nullable=False, index=True)
    vch_type = Column(String(40), nullable=False)
    vch_date = Column(String(8), nullable=False)      # YYYYMMDD
    bilty = Column(String(100), default="")
    amount = Column(String(40), default="0")
    sent_at = Column(DateTime, default=_now)
    deleted_at = Column(DateTime, nullable=True)       # set when undone/replaced


class PendingSend(Base):
    """Entries checked and saved while Tally was closed, waiting to be sent
    when Tally is open. rows_json holds each row's (possibly corrected)
    values, so nothing has to be uploaded or corrected again."""
    __tablename__ = "pending_sends"

    id = Column(Integer, primary_key=True)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    label = Column(String(300), nullable=False)
    rows_json = Column(Text, nullable=False)
    row_count = Column(Integer, default=0)
    total_amount = Column(String(40), default="0")
    mode = Column(String(20), default="replace")          # replace | skip
    status = Column(String(20), default="pending")        # pending | sent | cancelled
    created_at = Column(DateTime, default=_now)
    sent_at = Column(DateTime, nullable=True)
    batch_id = Column(Integer, ForeignKey("import_batches.id"), nullable=True)
    last_error = Column(Text, default="")
