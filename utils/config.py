"""Central configuration. Everything comes from environment variables
(optionally loaded from a .env file), so no secret is hard-coded."""
from __future__ import annotations

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent


def _load_dotenv() -> None:
    env_file = BASE_DIR / ".env"
    if not env_file.exists():
        return
    for raw in env_file.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), _clean_value(value))


def _clean_value(value: str) -> str:
    value = value.strip()
    if value[:1] in ('"', "'"):
        quote = value[0]
        end = value.find(quote, 1)
        return value[1:end] if end > 0 else value[1:]
    for marker in (" #", "\t#"):
        pos = value.find(marker)
        if pos >= 0:
            value = value[:pos]
    return value.strip()


_load_dotenv()


def _b(name: str, default: bool) -> bool:
    return os.getenv(name, str(default)).strip().lower() in ("1", "true", "yes", "on")


def _i(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, default))
    except (TypeError, ValueError):
        return default


class Settings:
    APP_NAME = os.getenv("APP_NAME", "Freight Bilty-to-Tally Agent")
    COMPANY_BRAND = os.getenv("COMPANY_BRAND", "SN Softech Solutions")
    BRAND_TAGLINE = os.getenv(
        "BRAND_TAGLINE", "Smart Software Solutions for Growing Businesses"
    )
    SUPPORT_CONTACT = os.getenv("SUPPORT_CONTACT", "9993199719")
    APP_VERSION = "1.7.0"
    BUILD_DATE = "2026-10-26"
    CHANGELOG = [
        ("1.7.0", "Other1/Other2 as owner advances: Transporting Exp = Amt-Comm-Diesel-Diesel cash-Office cash-Other1-Other2"),
        ("1.6.0", "Online demo build with sample Excel; customer package now includes all modules"),
        ("1.5.1", "Bilty number as voucher number (Sales by default; Payment/Journal optional)"),
        ("1.5.0", "Sales as item invoice: Name of Item, Qty = weight, Rate = rate"),
        ("1.4.0", "Pending: save checked entries while Tally is closed, send all with one click later"),
        ("1.3.0", "Several Excel files and all their sheets in one go; duplicate bilty warning"),
        ("1.2.1", "Party for this file box on Import page, remembered per layout"),
        ("1.2.0", "Several column layouts; the right one is picked automatically from the file headings"),
        ("1.1.2", "Rows where advances exceed freight are posted (owner Dr) instead of skipped"),
        ("1.1.1", "Failed sends are no longer recorded as sent; Forget option; Tally reply shown"),
        ("1.1.0", "Transporter-style posting (Payment/Journal/Sales), send directly to Tally, "
                  "Replace & Undo, on-screen corrections, choose which entries to post"),
        ("1.0.0", "First release"),
    ]

    # ---- licensing ----
    LICENSE_SIGNING_SECRET = os.getenv("LICENSE_SIGNING_SECRET", "").strip()
    LICENSE_SECRET_MIN_LENGTH = 16
    LICENSE_WARN_DAYS = _i("LICENSE_WARN_DAYS", 10)
    LICENSE_REQUIRED = _b("LICENSE_REQUIRED", True)
    # Demo build for showing customers (no licence, auto login, sample data,
    # no Tally connection). The demo packager switches this on.
    DEMO_MODE = _b("DEMO_MODE", True)

    # ---- storage ----
    DATA_DIR = Path(os.getenv("DATA_DIR", BASE_DIR / "data"))
    EXPORT_DIR = DATA_DIR / "exports"
    LOG_DIR = DATA_DIR / "logs"
    DATABASE_URL = os.getenv("DATABASE_URL", f"sqlite:///{DATA_DIR / 'biltytotally.db'}")

    # ---- security ----
    SECRET_KEY = os.getenv("SECRET_KEY", "dev-only-insecure-key-change-me")
    MAX_UPLOAD_MB = _i("MAX_UPLOAD_MB", 15)

    @classmethod
    def ensure_dirs(cls) -> None:
        for path in (cls.DATA_DIR, cls.EXPORT_DIR, cls.LOG_DIR):
            Path(path).mkdir(parents=True, exist_ok=True)


settings = Settings()
settings.ensure_dirs()
