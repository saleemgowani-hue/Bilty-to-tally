"""Freight Bilty-to-Tally Agent — main Streamlit app."""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st

from database.db import get_session, init_db
from database.models import Company, ImportBatch, User
import hashlib
import io
from types import SimpleNamespace

from services import demo_service, license_service, pending_service, settings_service, tally_connector, tally_sync_service
from services.bilty_tally_service import (
    FIELDS,
    NOT_MAPPED,
    REQUIRED_FIELDS,
    build_journal_xml,
    build_mapping_options,
    build_sales_xml,
    build_tally_xml,
    process_workbook,
    read_headers,
    safe_default_index,
    validate_mapping,
    MappingError,
    POSTING_STYLES,
    build_transporter_xml,
    count_transporter_vouchers,
    transporter_preview,
    editor_records,
    apply_edits,
    open_workbook,
    data_sheets,
    number_occurrences,
    guess_mapping,
    match_mappings,
    VOUCHER_KINDS,
    enabled_kinds,
)
from utils.config import settings
from utils.security import hash_password, verify_password

st.set_page_config(page_title=settings.APP_NAME, page_icon="🚚", layout="wide")
init_db()

# Colours for the dashboard's quick-action buttons and KPI cards. Keyed so
# each st.container(key=...) below can be targeted by its own CSS class
# (Streamlit gives every keyed container a stable `st-key-<key>` class).
_ACCENTS = {
    "import": "#2563eb",     # blue
    "history": "#059669",    # green
    "settings": "#d97706",   # amber
    "dashboard": "#7c3aed",  # violet
    "pending": "#dc2626",    # red
}


def _inject_css() -> None:
    rules = ["""
    /* Narrow, centered "card" look for the sign-in/sign-up/licence forms.
       st.container(key="auth_card") is used (not a hand-written <div>,
       which Streamlit renders as a sibling rather than a real wrapper)
       so this class actually encloses the widgets inside it. */
    .st-key-auth_card {
        max-width: 440px;
        margin: 0 auto;
        padding: 1.5rem 2rem;
        border: 1px solid rgba(49, 51, 63, 0.15);
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
    }
    .st-key-auth_card h1, .st-key-auth_card h2, .st-key-auth_card h3 { text-align: center; }
    .st-key-auth_card [data-testid="stForm"] { border: none; padding: 0; }

    /* KPI stat cards on the dashboard. */
    .kpi-card {
        border-radius: 10px;
        padding: 1rem 1.2rem;
        color: white;
    }
    .kpi-card .kpi-label { font-size: 0.85rem; opacity: 0.9; }
    .kpi-card .kpi-value { font-size: 1.8rem; font-weight: 700; }
    """]
    for key, color in _ACCENTS.items():
        rules.append(f"""
        .st-key-nav_{key} button {{
            background: linear-gradient(135deg, {color}, {color}cc);
            color: white;
            border: none;
            font-weight: 600;
            height: 3rem;
        }}
        .st-key-nav_{key} button:hover {{
            filter: brightness(1.1);
            color: white;
        }}
        /* Sidebar page-navigation buttons - same accent colours, left-aligned
           and a little shorter so they read as a menu, not a toolbar. */
        .st-key-navside_{key} button {{
            background: linear-gradient(135deg, {color}, {color}cc);
            color: white;
            border: none;
            font-weight: 600;
            justify-content: flex-start;
            height: 2.6rem;
            margin-bottom: 0.3rem;
        }}
        .st-key-navside_{key} button:hover {{
            filter: brightness(1.1);
            color: white;
        }}
        .st-key-navside_{key}_active button {{
            background: linear-gradient(135deg, {color}, {color}cc);
            color: white;
            border: 3px solid rgba(0, 0, 0, 0.35);
            font-weight: 700;
            justify-content: flex-start;
            height: 2.6rem;
            margin-bottom: 0.3rem;
        }}
        """)
    st.markdown(f"<style>{''.join(rules)}</style>", unsafe_allow_html=True)


def _kpi_card(label: str, value: str, color: str) -> str:
    return (
        f'<div class="kpi-card" style="background:{color}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div></div>'
    )


def _company_count() -> int:
    with get_session() as session:
        return session.query(Company).count()


def _require_login() -> None:
    if "user_id" not in st.session_state:
        st.stop()


def _current_company(session) -> Company:
    return session.query(Company).filter(Company.id == st.session_state["company_id"]).one()


# ---------------------------------------------------------------- Signup ---
def signup_page() -> None:
    _inject_css()
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        with st.container(key="auth_card"):
            st.markdown(f"### 🚚 {settings.APP_NAME}")
            st.caption(settings.BRAND_TAGLINE)
            st.markdown("**First-time setup — create your company**")
            with st.form("signup"):
                company_name = st.text_input("Company name")
                tally_company_name = st.text_input(
                    "Exact Tally company name",
                    help="This must match the company name in Tally exactly, so the import lands in the right books.",
                )
                email = st.text_input("Admin email")
                password = st.text_input("Password", type="password")
                submitted = st.form_submit_button("Create company", use_container_width=True)
    if submitted:
        if not (company_name and email and password):
            st.error("Fill in all fields.")
            return
        with get_session() as session:
            company = Company(name=company_name, tally_company_name=tally_company_name)
            session.add(company)
            session.commit()
            user = User(
                company_id=company.id,
                email=email.strip().lower(),
                password_hash=hash_password(password),
                role="ADMIN",
            )
            session.add(user)
            session.commit()
            st.session_state["user_id"] = user.id
            st.session_state["company_id"] = company.id
        st.success("Company created. Activate your licence key to continue.")
        st.rerun()


# ----------------------------------------------------------------- Login ---
def login_page() -> None:
    _inject_css()
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        with st.container(key="auth_card"):
            st.markdown(f"### 🚚 {settings.APP_NAME}")
            st.caption(settings.BRAND_TAGLINE)
            with st.form("login"):
                email = st.text_input("Email")
                password = st.text_input("Password", type="password")
                submitted = st.form_submit_button("Sign in", use_container_width=True)
    if submitted:
        with get_session() as session:
            user = (
                session.query(User)
                .filter(User.email == email.strip().lower())
                .one_or_none()
            )
            if user and verify_password(password, user.password_hash):
                st.session_state["user_id"] = user.id
                st.session_state["company_id"] = user.company_id
                st.rerun()
            else:
                st.error("Incorrect email or password.")


# ---------------------------------------------------------- Licence gate ---
def license_gate() -> bool:
    if settings.DEMO_MODE:
        return True
    with get_session() as session:
        company = _current_company(session)
        status = license_service.status(company)
        if status.is_valid:
            if status.is_expiring_soon:
                st.warning(status.message)
            return True

    _inject_css()
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        with st.container(key="auth_card"):
            st.markdown("### Activate your licence")
            st.info(status.message)
            with st.form("activate"):
                key = st.text_input("Licence key", placeholder="SNBT-XXXX-XXXX-XXXX-XXXX")
                submitted = st.form_submit_button("Activate", use_container_width=True)
    if submitted:
        with get_session() as session:
            company = _current_company(session)
            try:
                license_service.activate(session, company, key, activated_by=st.session_state["user_id"])
                st.success("Licence activated.")
                st.rerun()
            except license_service.LicenseError as exc:
                st.error(str(exc))
    return False


# --------------------------------------------------------------- Mapping ---
NEW_LAYOUT = "➕ New layout"


def mapping_tab(session, company: Company) -> None:
    st.markdown(
        "Tell the app which column in **your** Excel register holds each piece "
        "of information. Required fields are marked with *. If you keep registers "
        "in more than one format, save one **layout** per format - on Import the "
        "app picks the right layout by itself from the file's column headings."
    )
    layouts = settings_service.all_mappings(session, company.id)
    choices = list(layouts) + [NEW_LAYOUT]
    picked = st.selectbox("Layout", choices, index=0, key="layout_pick")

    uploaded = st.file_uploader(
        "Upload a sample file to pick columns from (or skip and type header names below)",
        type=["xlsx", "xlsm"],
        key="mapping_sample",
    )
    headers = []
    if uploaded is not None:
        try:
            wb = open_workbook(uploaded)
            sheets = data_sheets(wb) or wb.sheetnames[:1]
            sample_sheet = sheets[0]
            if len(sheets) > 1:
                sample_sheet = st.selectbox("Sheet", sheets, key="mapping_sheet")
            headers = read_headers(wb, sample_sheet)
            if layouts:
                best = match_mappings(wb, layouts, sheet=sample_sheet)[0]
                if best[1]:
                    st.success(f"This file already fits layout **{best[0]}** - nothing to do, unless you want to change it.")
                else:
                    st.info("This file does not fit any saved layout yet - choose “➕ New layout”, check the columns below and save.")
        except Exception as exc:
            st.error(f"Could not read that file: {exc}")

    if picked == NEW_LAYOUT:
        current = guess_mapping(headers) if headers else {}
        default_name = f"Layout {len(layouts) + 1}"
        if uploaded is not None:
            default_name = Path(uploaded.name).stem[:60]
        if headers and current:
            st.caption("Columns below were pre-filled by matching the headings - check each one.")
    else:
        current = layouts.get(picked, {})
        default_name = picked
    layout_name = st.text_input("Layout name", value=default_name, key=f"layout_name_{picked}")

    new_mapping = {}
    options = build_mapping_options(headers)
    # The widget key is scoped to the layout and to the sample file's
    # identity so switching between them never reuses a remembered value
    # that isn't valid any more.
    file_token = (f"{uploaded.name}_{uploaded.size}_{st.session_state.get('mapping_sheet', '')}"
                  if uploaded is not None else "nofile")
    for field_name, (label, required) in FIELDS.items():
        star = " *" if required else ""
        default = current.get(field_name, "")
        if options:
            choice = st.selectbox(
                f"{label}{star}", options, index=safe_default_index(options, default),
                key=f"map_select_{picked}_{file_token}_{field_name}",
            )
            new_mapping[field_name] = "" if choice == NOT_MAPPED else choice
        else:
            new_mapping[field_name] = st.text_input(
                f"{label}{star}", value=default, key=f"map_text_{picked}_{field_name}",
            )

    c1, c2 = st.columns([3, 1])
    with c1:
        if st.button("Save column mapping", type="primary", use_container_width=True):
            missing = validate_mapping(new_mapping)
            name = (layout_name or "").strip()
            if missing:
                st.error("Still missing a column for: " + ", ".join(missing))
            elif not name or name == NEW_LAYOUT:
                st.error("Give this layout a name.")
            else:
                if picked != NEW_LAYOUT and name != picked:
                    settings_service.delete_mapping(session, company.id, picked)   # renamed
                settings_service.save_mapping(session, company.id, name, new_mapping)
                st.success(f"Layout “{name}” saved. Files with these column headings will use it automatically.")
    with c2:
        if picked != NEW_LAYOUT and st.button("Delete layout", use_container_width=True):
            settings_service.delete_mapping(session, company.id, picked)
            st.rerun()


# ----------------------------------------------------------- Ledger tab ---
def ledger_tab(session, company: Company) -> None:
    st.markdown(
        "Choose how the register is posted, then rename the ledgers it posts to "
        "so they match **exactly** what already exists in your Tally."
    )
    names = settings_service.get_ledger_names(session, company.id)
    current_style = settings_service.get_posting_style(session, company.id)
    style_keys = list(POSTING_STYLES)
    style = st.radio(
        "Posting style",
        style_keys,
        index=style_keys.index(current_style),
        format_func=lambda k: POSTING_STYLES[k],
    )
    if style == "transporter":
        st.caption(
            "Per row: Payment (Other Expenses Dr / Cash Cr - cash from office) · "
            "Journal (TDS Dr / Party Cr) · Journal (Diesel Expenses + Other Expenses "
            "for cash from pump Dr / Petrol pump Cr) · Journal (Transporting Expenses "
            "Dr / Truck owner Cr) · Sales (Party Dr / Freight Cr)."
        )
    tally_name = st.text_input(
        "Exact Tally company name", value=company.tally_company_name or ""
    )
    cols = st.columns(2)
    if style == "transporter":
        labels = {
            "party": "Party / customer ledger (used when no Party column is mapped)",
            "freight": "Freight sales ledger",
            "tds": "TDS ledger",
            "diesel_exp": "Diesel expenses ledger",
            "other_exp": "Other expenses ledger (cash from office / pump)",
            "transport_exp": "Transporting expenses ledger",
            "cash": "Cash ledger",
            "discount": "Discount ledger",
            "other1": "Other deduction 1 ledger",
            "other2": "Other deduction 2 ledger",
        }
    else:
        labels = {
            "freight": "Freight income ledger",
            "tds": "TDS receivable ledger",
            "commission": "Commission ledger",
            "discount": "Discount ledger",
            "cash": "Cash ledger",
            "other1": "Other deduction 1 ledger",
            "other2": "Other deduction 2 ledger",
        }
    new_names = {}
    for i, (field_name, label) in enumerate(labels.items()):
        with cols[i % 2]:
            new_names[field_name] = st.text_input(
                label, value=names.get(field_name, ""), key=f"ledger_{style}_{field_name}"
            )

    if style == "transporter":
        st.markdown("**Sales voucher format**")
        sm_opts = ["invoice", "accounting"]
        sales_mode = st.radio(
            "Sales voucher format", sm_opts,
            index=sm_opts.index(names.get("sales_mode", "invoice")) if names.get("sales_mode", "invoice") in sm_opts else 0,
            format_func=lambda m: {
                "invoice": "Item invoice - Name of Item, Qty (= Excel weight) and Rate (= Excel rate)",
                "accounting": "Accounting invoice - only Party Dr / Freight ledger Cr",
            }[m],
            label_visibility="collapsed", key="sales_mode_radio",
        )
        new_names["sales_mode"] = sales_mode
        if sales_mode == "invoice":
            i1, i2 = st.columns(2)
            new_names["item_name"] = i1.text_input(
                "Item name (when the Excel has no item column)", value=names.get("item_name", "FREIGHT"),
                key="sales_item_name",
            )
            new_names["item_unit"] = i2.text_input(
                "Unit of the weight / Qty (e.g. MT, Ton, Qtl, Kg)", value=names.get("item_unit", "MT"),
                key="sales_item_unit",
            )
            st.caption("Map the Weight and Rate columns (and Item, if your Excel has one) in Column Mapping.")

    if style == "transporter":
        st.markdown("**Other1 / Other2 columns**")
        om_opts = ["advance", "deduction"]
        other_mode = st.radio(
            "Other1 / Other2 columns", om_opts,
            index=om_opts.index(names.get("other_mode", "advance")) if names.get("other_mode", "advance") in om_opts else 0,
            format_func=lambda m: {
                "advance": "Advance already given to the truck owner - reduces Transporting Expenses "
                           "(Transporting Exp. = Amt − Comm − Diesel − Diesel cash − Office cash − Other1 − Other2) "
                           "and is posted as Payment: Other Expenses Dr / paid-from ledger Cr",
                "deduction": "Deduction kept back from the owner - credited to the Other deduction 1/2 ledgers",
            }[m],
            label_visibility="collapsed", key="other_mode_radio",
        )
        new_names["other_mode"] = other_mode
        if other_mode == "advance":
            o1, o2 = st.columns(2)
            new_names["other1_from"] = o1.text_input(
                "Other1 paid from (Cash / Bank ledger)", value=names.get("other1_from", "Cash"), key="other1_from")
            new_names["other2_from"] = o2.text_input(
                "Other2 paid from (Cash / Bank ledger)", value=names.get("other2_from", "Cash"), key="other2_from")

    if style == "transporter":
        st.markdown("**Voucher number**")
        v1, v2 = st.columns(2)
        new_names["vchno_sales"] = "yes" if v1.checkbox(
            "Sales: voucher no. = Bilty no.", value=names.get("vchno_sales", "yes") == "yes", key="vchno_sales_cb",
        ) else "no"
        new_names["vchno_others"] = "yes" if v2.checkbox(
            "Payment & Journal: voucher no. = Bilty no.", value=names.get("vchno_others", "no") == "yes",
            key="vchno_others_cb",
        ) else "no"
        st.caption(
            "Unticked = Tally numbers the vouchers itself (1, 2, 3...). For bilty numbers to show, the "
            "voucher type's numbering in Tally should be 'Manual' or 'Automatic (Manual Override)'."
        )

    post_kinds = None
    if style == "transporter":
        st.markdown("**Which entries should go to Tally?**")
        st.caption(
            "Untick an entry to leave it out of Tally. Only that voucher is dropped - "
            "the other entries keep their amounts (e.g. without TDS the Sales bill is still the full freight)."
        )
        on = enabled_kinds(names)
        post_kinds = []
        kcols = st.columns(2)
        for i, (code, label) in enumerate(VOUCHER_KINDS.items()):
            with kcols[i % 2]:
                if st.checkbox(label, value=code in on, key=f"kind_{code}"):
                    post_kinds.append(code)

    if st.button("Save ledger names", type="primary"):
        settings_service.save_ledger_names(session, company.id, new_names)
        if post_kinds is not None:
            settings_service.save_post_kinds(session, company.id, post_kinds)
        settings_service.save_posting_style(session, company.id, style)
        company.tally_company_name = tally_name
        session.commit()
        st.success("Saved.")


# --------------------------------------------------------- Licence tab ----
def licence_tab(session, company: Company) -> None:
    if settings.DEMO_MODE:
        st.info(f"Demo version - no licence needed. To buy: {settings.COMPANY_BRAND}, {settings.SUPPORT_CONTACT}.")
        return
    status = license_service.status(company)
    st.info(status.message)
    with st.form("renew"):
        key = st.text_input("New / renewal licence key", placeholder="SNBT-XXXX-XXXX-XXXX-XXXX")
        submitted = st.form_submit_button("Activate")
    if submitted:
        try:
            license_service.activate(session, company, key, activated_by=st.session_state["user_id"])
            st.success("Activated.")
            st.rerun()
        except license_service.LicenseError as exc:
            st.error(str(exc))

    st.subheader("History")
    for lic in license_service.company_licenses(session, company.id):
        st.write(f"{lic.plan.title()} · {lic.key_display} · activated {lic.activated_at:%d %b %Y} · "
                  f"expires {lic.expires_at:%d %b %Y}")


def tally_connection_tab(session, company: Company) -> None:
    if settings.DEMO_MODE:
        st.info(
            "In the full software this connects to TallyPrime on the customer's computer "
            "(port 9000) and posts the entries straight into Tally. The online demo has no Tally, "
            "so here you can download the Tally XML file instead."
        )
        return
    st.markdown(
        "Send entries **straight into Tally** - no XML import needed - when Tally "
        "is open on this computer (or on the office network)."
    )
    st.info(
        "One-time setting in TallyPrime: **F1 (Help) → Settings → Connectivity → "
        "Client/Server configuration** → *TallyPrime acts as* = **Both**, "
        "*Port* = **9000**. Then restart Tally and open your company."
    )
    host, port = settings_service.get_tally_connection(session, company.id)
    c1, c2 = st.columns([3, 1])
    with c1:
        new_host = st.text_input(
            "Tally computer", value=host,
            help="'localhost' when Tally is on this same computer; otherwise that computer's IP, e.g. 192.168.1.10",
        )
    with c2:
        new_port = st.number_input("Port", value=int(port), min_value=1, max_value=65535, step=1)
    st.caption(f"Tally company name (from Ledger Names tab): **{company.tally_company_name or '— not set —'}**")
    b1, b2 = st.columns(2)
    with b1:
        if st.button("Save", type="primary", use_container_width=True, key="tally_save"):
            settings_service.save_tally_connection(session, company.id, new_host, int(new_port))
            st.success("Saved.")
    with b2:
        if st.button("Test connection", use_container_width=True, key="tally_test"):
            with st.spinner("Contacting Tally..."):
                ok, message = tally_connector.check_connection(
                    new_host, int(new_port), company.tally_company_name or ""
                )
            (st.success if ok else st.error)(message)


def _send_to_tally_section(session, company: Company, xml: str, uploaded, usable, problems, mismatches,
                           rows=None, style: str = "transporter", ledger_names=None) -> None:
    """'Send directly to Tally'. In the transporter style every voucher is
    tagged and recorded, so a re-sent bilty replaces its old vouchers and a
    send can be undone from History."""
    if settings.DEMO_MODE:
        st.info(
            "🚀 **Send directly to Tally** - in the full software this button posts all these "
            "Payment / Journal / Sales entries straight into TallyPrime running on your computer, "
            "with Replace, Undo and a Pending list for when Tally is closed. "
            "In this online demo, use **Generate Tally XML file** below to see the Tally data."
        )
        return
    host, port = settings_service.get_tally_connection(session, company.id)

    if style != "transporter":
        # Old owner style: plain send, guarded only within this session.
        fingerprint = hashlib.sha256(xml.encode("utf-8")).hexdigest()
        sent = st.session_state.setdefault("_sent_to_tally", set())
        if fingerprint in sent:
            st.warning("Already sent in this session - sending again creates duplicates.")
            if not st.checkbox("Send again anyway", key=f"resend_{fingerprint[:8]}"):
                return
        if st.button("🚀 Send directly to Tally", type="primary", use_container_width=True, key="send_tally"):
            with st.spinner("Sending to Tally ..."):
                try:
                    result = tally_connector.push_xml(xml, host, port, company.tally_company_name or "")
                except tally_connector.TallyConnectionError as exc:
                    st.error(str(exc))
                    return
            sent.add(fingerprint)
            (st.success if result.errors == 0 else st.error)(
                f"Tally created {result.created} voucher(s), errors: {result.errors}."
            )
            for msg in result.messages[:50]:
                st.write(f"• {msg}")
        return

    previous = tally_sync_service.already_sent(
        session, company.id, usable, company.tally_company_name or ""
    )
    mode = "replace"
    if previous:
        st.warning(
            f"**{len(previous)} bilty(s)** in this file were already sent to Tally earlier."
        )
        mode = st.radio(
            "What should happen to them?",
            ["replace", "skip"],
            format_func=lambda m: {
                "replace": "Replace - delete their old vouchers in Tally and post these (use after correcting the Excel)",
                "skip": "Skip - leave them as they are, send only the new bilties",
            }[m],
            key="send_mode",
        )
    b_send, b_pend = st.columns([3, 2])
    with b_pend:
        if st.button("📥 Save to Pending - send when Tally is open", use_container_width=True, key="save_pending"):
            item = pending_service.save(session, company.id, rows or usable, uploaded.name, mode)
            st.success(
                f"Saved to Pending ({item.row_count} rows, with your on-screen corrections). "
                "Open Tally later and press “Send all to Tally” on the Pending page."
            )
            return
    with b_send:
        send_clicked = st.button("🚀 Send directly to Tally", type="primary", use_container_width=True, key="send_tally")
    if send_clicked:
        with st.spinner(f"Sending to Tally at {host}:{port} ..."):
            try:
                out = tally_sync_service.send(
                    session, company, rows or usable, ledger_names or {}, uploaded.name, host, port,
                    mode=mode, skipped_count=len(problems), mismatch_count=len(mismatches),
                )
            except tally_connector.TallyConnectionError as exc:
                st.error(str(exc))
                st.info("Tally not open right now? Press **📥 Save to Pending** - the entries wait in the "
                        "software and go to Tally with one click once Tally is open.")
                return
        parts = []
        if out.replaced_deleted:
            parts.append(f"deleted {out.replaced_deleted} old voucher(s)")
        parts.append(f"created {out.created} voucher(s)")
        if out.altered:
            parts.append(f"updated {out.altered} voucher(s)")
        if out.ledgers_created:
            parts.append(f"{out.ledgers_created} new ledger(s)")
        if out.skipped_bilties:
            parts.append(f"skipped {out.skipped_bilties} bilty(s) already in Tally")
        made = out.created + out.altered
        if out.errors == 0 and (made or out.skipped_bilties):
            first = min((r.date for r in usable if r.date), default=None)
            st.success(
                "Done - Tally " + ", ".join(parts) + "."
                + (f" In Day Book press F2 and enter {first:%d-%m-%Y} (entries are on their Excel dates)." if first and made else "")
            )
        elif made == 0 and not out.skipped_bilties and out.raw:
            st.error(
                "Tally received the data but made NO entries. Nothing was recorded as sent, "
                "so you can fix the cause and send again. Tally's reply is below - "
                "send a screenshot of it to SN Softech support if the reason is not clear."
            )
        else:
            st.error(
                "Tally " + ", ".join(parts) + f" but reported {out.errors} error(s). Fix the cause "
                "below, then use History → Undo for this send and send again."
            )
        for msg in out.messages[:50]:
            st.write(f"• {msg}")
        if out.raw:
            with st.expander("Tally's reply (for troubleshooting)"):
                st.code(out.raw[:6000], language="xml")
        if out.batch_id:
            st.caption("Made a mistake? History → this send → **Undo** removes these vouchers from Tally.")


def settings_page(session, company: Company) -> None:
    st.header("⚙️ Settings")
    tab1, tab2, tab4, tab3 = st.tabs(
        ["Column Mapping", "Ledger Names & Company", "Tally Connection", "Licence"]
    )
    with tab1:
        mapping_tab(session, company)
    with tab2:
        ledger_tab(session, company)
    with tab4:
        tally_connection_tab(session, company)
    with tab3:
        licence_tab(session, company)


# ---------------------------------------------------------- Import page ---
def import_page(session, company: Company) -> None:
    st.header("📤 Import & Generate")

    mappings = settings_service.list_mappings(session, company.id)
    if not mappings:
        st.warning("No column mapping saved yet. Go to Settings → Column Mapping first.")
        return
    layouts = settings_service.all_mappings(session, company.id)
    ledger_names = settings_service.get_ledger_names(session, company.id)
    style = settings_service.get_posting_style(session, company.id)
    st.caption(f"Posting style: **{POSTING_STYLES[style]}** (change in Settings → Ledger Names)")
    if style == "transporter":
        off = [VOUCHER_KINDS[k] for k in VOUCHER_KINDS if k not in enabled_kinds(ledger_names)]
        if off:
            st.info("Switched off - **not** posted to Tally: " + "; ".join(off)
                    + "  (change in Settings → Ledger Names)")

    files = st.file_uploader(
        "Upload your freight/bilty register(s) - you can select several Excel files at once",
        type=["xlsx", "xlsm"], accept_multiple_files=True,
    )
    if settings.DEMO_MODE:
        d1, d2 = st.columns(2)
        if d1.button("📄 Use the demo sample Excel", type="primary", use_container_width=True):
            st.session_state["demo_use_sample"] = True
        if demo_service.SAMPLE_FILE.exists():
            d2.download_button("⬇️ Download the sample Excel", data=demo_service.SAMPLE_FILE.read_bytes(),
                               file_name=demo_service.SAMPLE_FILE.name, use_container_width=True)
        if not files and st.session_state.get("demo_use_sample") and demo_service.SAMPLE_FILE.exists():
            sample = io.BytesIO(demo_service.SAMPLE_FILE.read_bytes())
            sample.name = demo_service.SAMPLE_FILE.name
            sample.size = len(sample.getvalue())
            files = [sample]
    if not files:
        return

    # One "unit" per sheet that has data, across all uploaded files.
    units = []
    for f in files:
        try:
            wb = open_workbook(f)
        except Exception as exc:
            st.error(f"Could not read {f.name}: {exc}")
            continue
        for sheet in data_sheets(wb):
            ranking = match_mappings(wb, layouts, sheet=sheet)
            label = f.name if len(wb.sheetnames) == 1 else f"{f.name} / {sheet}"
            units.append({"file": f, "wb": wb, "sheet": sheet, "label": label,
                          "ranking": ranking, "token": f"{f.name}_{f.size}_{sheet}"})
    if not units:
        st.error("No sheet with data was found in the uploaded file(s).")
        return

    st.markdown(f"**{len(units)} sheet(s) found** - each one is matched to its column layout automatically.")
    rows = []
    used_units = []
    for u in units:
        ranking = u["ranking"]
        fits = [r[0] for r in ranking if r[1]]
        with st.container(border=True):
            c_inc, c_lab = st.columns([1, 9])
            include = c_inc.checkbox("Use", value=bool(fits), key=f"inc_{u['token']}",
                                     label_visibility="collapsed")
            c_lab.markdown(f"**{u['label']}**")
            if not include:
                if not fits:
                    best = ranking[0]
                    c_lab.caption(
                        f"Not used - headings don't match any saved layout (closest “{best[0]}”, "
                        f"missing: {', '.join(best[3][:6])}). Add a layout for it in Settings → Column Mapping."
                    )
                else:
                    c_lab.caption("Not used (unticked).")
                continue
            cols = st.columns([2, 3]) if style == "transporter" else st.columns([1])
            names_order = [r[0] for r in ranking]
            if len(layouts) > 1 or not fits:
                layout = cols[0].selectbox(
                    "Column layout", names_order, index=0, key=f"layout_{u['token']}",
                    format_func=lambda n, rk=ranking: n + ("  ✓ fits" if any(r[0] == n and r[1] for r in rk) else ""),
                )
            else:
                layout = names_order[0]
                cols[0].caption(f"Column layout: **{layout}**")
            info = next(r for r in ranking if r[0] == layout)
            if not info[1]:
                st.error(f"Headings don't match layout “{layout}” - missing: {', '.join(info[3])}.")
                continue
            mapping = layouts[layout]
            unit_names = dict(ledger_names)
            if style == "transporter":
                # Party for this sheet: starts from the party last used with this
                # layout (or the Settings default) and is remembered for next time.
                remember_key = f"party_for_layout::{layout}"
                last_party = (settings_service.get_setting(session, company.id, remember_key, "")
                              or ledger_names.get("party", ""))
                party = cols[1].text_input(
                    "Party (Sales bill & TDS go to this ledger)", value=last_party,
                    key=f"party_{u['token']}", help="Type the party's ledger name exactly as in Tally.",
                ).strip()
                if mapping.get("party"):
                    cols[1].caption(f"Rows with a name in the “{mapping['party']}” column use that name.")
                if party != last_party:
                    settings_service.set_setting(session, company.id, remember_key, party)
                unit_names["party"] = party
            try:
                unit_rows = process_workbook(u["wb"], mapping, unit_names, style=style,
                                             sheet=u["sheet"], source=u["label"])
            except MappingError as exc:
                st.error(str(exc))
                continue
            except Exception as exc:
                st.error(f"Could not read {u['label']}: {exc}")
                continue
            st.caption(f"{len(unit_rows)} row(s) read.")
            rows.extend(unit_rows)
            used_units.append(u)

    if not rows:
        st.error("Nothing to process - tick at least one sheet that matches a layout.")
        return
    number_occurrences(rows)   # a bilty+date repeated across files stays distinct
    labels = [u["label"] for u in used_units]
    uploaded = SimpleNamespace(
        name=labels[0] if len(labels) == 1 else f"{len(labels)} sheets: " + ", ".join(labels)[:200],
        size=hashlib.sha1("|".join(u["token"] for u in used_units).encode()).hexdigest()[:10],
    )
    multi = len(used_units) > 1

    summary = st.container()

    if style == "transporter":
        st.subheader("✏️ Check & correct before sending")
        st.caption(
            "Click any cell to correct it right here - party, owner, pump, amounts, date. "
            "The vouchers below update straight away. Your Excel file itself is not changed, "
            "so correct it there too for your records."
        )
        editor_key = f"edit_{uploaded.name}_{uploaded.size}"
        base = editor_records(rows)
        for rec in base:
            rec.pop("Transporting Exp.", None)
            rec.pop("Problem", None)
        money = st.column_config.NumberColumn(format="%.2f", min_value=0.0)
        edited_df = st.data_editor(
            pd.DataFrame(base),
            key=editor_key,
            hide_index=True,
            num_rows="fixed",
            use_container_width=True,
            disabled=["File / Sheet", "Excel row"],
            column_order=None if multi else [c for c in base[0] if c != "File / Sheet"] if base else None,
            column_config={
                "Date": st.column_config.DateColumn(format="DD-MM-YYYY"),
                "Freight Amt": money, "TDS": money, "Comm": money, "Diesel": money,
                "Cash from Pump": money, "Cash from Office": money, "Other1": money, "Other2": money,
            },
        )
        records = edited_df.to_dict("records")
        for rec in records:
            if isinstance(rec.get("Date"), pd.Timestamp):
                rec["Date"] = None if pd.isna(rec["Date"]) else rec["Date"].to_pydatetime()
        rows = apply_edits(rows, records, ledger_names, style)
        n_edited = sum(1 for r in rows if r.edited)
        if n_edited:
            e1, e2 = st.columns([4, 1])
            e1.info(f"✏️ {n_edited} row(s) corrected on screen - these corrected values will be sent.")
            if e2.button("Undo my edits", use_container_width=True):
                st.session_state.pop(editor_key, None)
                st.rerun()

    problems = [r for r in rows if r.problem]
    mismatches = [r for r in rows if r.balance_mismatch]
    usable = [r for r in rows if not r.problem]

    with summary:
        c1, c2, c3 = st.columns(3)
        c1.metric("Rows read", len(rows))
        c2.metric("Rows skipped", len(problems))
        c3.metric("Balance mismatches", len(mismatches))

        if problems:
            fix_where = "in the table below or in Excel" if style == "transporter" else "in Excel and re-upload"
            with st.expander(f"⚠️ {len(problems)} row(s) skipped — fix these {fix_where}", expanded=True):
                for r in problems:
                    where = f"{r.source} row {r.source_row}" if multi else f"Row {r.source_row}"
                    st.write(f"{where} (Bilty {r.voucher_no or '-'}): {r.problem}")

        repeats = [r for r in usable if r.occurrence > 1]
        if repeats:
            with st.expander(f"⚠️ {len(repeats)} bilty(s) appear more than once (same bilty no. and date) - "
                             "check you did not upload the same data twice", expanded=True):
                for r in repeats:
                    firsts = [x for x in usable if x.occurrence == 1 and x.voucher_no.upper() == r.voucher_no.upper()
                              and x.date == r.date]
                    first = f" - first seen in {firsts[0].source or 'this file'} row {firsts[0].source_row}" if firsts else ""
                    where = f"{r.source} row {r.source_row}" if multi else f"Row {r.source_row}"
                    st.write(f"{where}: Bilty {r.voucher_no} ({r.date:%d-%m-%Y}){first}. "
                             "If it is a duplicate, untick that sheet above.")

        warned = [r for r in usable if r.warning]
        if warned:
            with st.expander(f"ℹ️ {len(warned)} row(s) where the owner owes you (advances more than freight) - will be posted, please check"):
                for r in warned:
                    where = f"{r.source} row {r.source_row}" if multi else f"Row {r.source_row}"
                    st.write(f"{where} (Bilty {r.voucher_no}): {r.warning}")

        if mismatches:
            with st.expander(f"⚠️ {len(mismatches)} row(s) where our computed balance doesn't match your Balance column"):
                st.caption(
                    "This usually means a column is mapped to the wrong field, or this "
                    "row has an extra adjustment not covered by the standard columns."
                )
                for r in mismatches[:20]:
                    st.write(
                        f"Row {r.source_row} ({r.voucher_no}): computed {r.computed_balance:,.2f}, "
                        f"your sheet says {r.excel_balance:,.2f}"
                    )

    st.subheader("Preview")
    if style == "transporter":
        vouchers = transporter_preview(usable, ledger_names)
        counts = {t: sum(1 for v in vouchers if v["Voucher"] == t) for t in ("Payment", "Journal", "Sales")}
        st.markdown(
            f"**Tally entries that will be made (like Tally Day Book):** {counts['Payment']} Payment · "
            f"{counts['Journal']} Journal · {counts['Sales']} Sales"
        )
        only = st.multiselect(
            "Show voucher types", ["Payment", "Journal", "Sales"],
            default=["Payment", "Journal", "Sales"], key="pv_types",
        )
        shown = [v for v in vouchers if v["Voucher"] in only]
        if not multi:
            shown = [{k: val for k, val in v.items() if k != "File / Sheet"} for v in shown]
        st.dataframe(
            shown[:1000], use_container_width=True, hide_index=True,
            column_config={"Amount": st.column_config.NumberColumn(format="%.2f")},
        )
        if len(shown) > 1000:
            st.caption(f"Showing the first 1000 of {len(shown)} vouchers.")
    else:
        preview_rows = [
            {
                "Row": r.source_row,
                "Voucher": r.voucher_no,
                "Owner": r.owner,
                "Amount": r.amount,
                "TDS": r.tds,
                "Deductions": r.total_deduction,
                "Balance (computed)": r.computed_balance,
            }
            for r in usable[:200]
        ]
        st.dataframe(preview_rows, use_container_width=True, hide_index=True)

    if not usable:
        st.error("No usable rows — nothing to generate.")
        return

    if style == "transporter":
        direct_xml = build_transporter_xml(usable, company.tally_company_name, ledger_names)
    else:
        direct_xml = build_tally_xml(usable, company.tally_company_name, ledger_names)
    st.subheader("Post to Tally")
    _send_to_tally_section(
        session, company, direct_xml, uploaded, usable, problems, mismatches,
        rows=rows, style=style, ledger_names=ledger_names,
    )
    st.caption("…or download an XML file and import it in Tally yourself:")

    if style == "transporter" and st.button("Generate Tally XML file"):
        xml = direct_xml
        counts = count_transporter_vouchers(usable, ledger_names)
        settings.EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = f"{dt.datetime.now():%Y%m%d_%H%M%S}"
        name = f"TallyVouchers_{stamp}.xml"
        path = settings.EXPORT_DIR / name
        path.write_text(xml, encoding="utf-8")
        batch = ImportBatch(
            company_id=company.id,
            filename=uploaded.name,
            row_count=len(usable),
            skipped_count=len(problems),
            mismatch_count=len(mismatches),
            total_amount=f"{sum(r.amount for r in usable):.2f}",
            xml_path=str(path),
        )
        session.add(batch)
        session.commit()
        st.success(
            f"Generated {counts['Payment']} Payment, {counts['Journal']} Journal and "
            f"{counts['Sales']} Sales vouchers in one file."
        )
        st.download_button(
            "⬇️ Download Tally XML",
            data=xml,
            file_name=name,
            mime="application/xml",
            use_container_width=True,
        )
        st.caption(
            "In Tally: Gateway of Tally → Import → Transactions → pick this file. "
            "Ledger masters are inside the same file."
        )
        return

    if style == "owner" and st.button("Generate Tally XML file"):
        sales_xml = build_sales_xml(usable, company.tally_company_name, ledger_names)
        journal_xml = build_journal_xml(usable, company.tally_company_name, ledger_names)
        settings.EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = f"{dt.datetime.now():%Y%m%d_%H%M%S}"
        sales_name = f"TallySales_{stamp}.xml"
        journal_name = f"TallyJournal_{stamp}.xml"
        sales_path = settings.EXPORT_DIR / sales_name
        journal_path = settings.EXPORT_DIR / journal_name
        sales_path.write_text(sales_xml, encoding="utf-8")
        journal_path.write_text(journal_xml, encoding="utf-8")

        journal_voucher_count = sum(1 for r in usable if r.deductions)

        batch = ImportBatch(
            company_id=company.id,
            filename=uploaded.name,
            row_count=len(usable),
            skipped_count=len(problems),
            mismatch_count=len(mismatches),
            total_amount=f"{sum(r.amount for r in usable):.2f}",
            sales_xml_path=str(sales_path),
            journal_xml_path=str(journal_path),
        )
        session.add(batch)
        session.commit()

        st.success(
            f"Generated {len(usable)} Sales vouchers and {journal_voucher_count} "
            f"Journal vouchers as two separate files."
        )
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Download Sales XML",
                data=sales_xml,
                file_name=sales_name,
                mime="application/xml",
                use_container_width=True,
            )
            st.caption("Columns I–L: freight bills (Owner Dr, Freight + TDS Cr).")
        with d2:
            st.download_button(
                "⬇️ Download Journal XML",
                data=journal_xml,
                file_name=journal_name,
                mime="application/xml",
                use_container_width=True,
            )
            st.caption("Columns M onward: commission/discount/diesel/cash deductions.")
        st.caption(
            "In Tally: Gateway of Tally → Import Data → XML, import either file first — "
            "each one carries its own ledger masters."
        )


# --------------------------------------------------------------- History --
def _undo_section(session, company: Company, batch: ImportBatch) -> None:
    live = tally_sync_service.live_vouchers(session, batch.id)
    if not live:
        if "(sent to Tally)" in (batch.filename or ""):
            st.caption("These vouchers are no longer in Tally (undone or replaced).")
        return
    st.write(f"**{len(live)} voucher(s) from this send are in Tally.**")
    confirm = st.checkbox(
        "Yes, delete these vouchers from Tally", key=f"undo_ok_{batch.id}"
    )
    with st.popover("Entries not in Tally?"):
        st.caption(
            "Use this only if these entries are NOT in Tally (never created, or you deleted "
            "them by hand). The software then stops treating these bilties as already sent."
        )
        if st.button("Forget this send", key=f"forget_{batch.id}"):
            n = tally_sync_service.forget_batch(session, batch.id)
            st.success(f"Done - {n} voucher(s) forgotten. You can send these bilties again.")
            st.rerun()
    if st.button("↩️ Undo - delete from Tally", key=f"undo_{batch.id}", disabled=not confirm):
        host, port = settings_service.get_tally_connection(session, company.id)
        with st.spinner("Deleting from Tally ..."):
            try:
                result = tally_sync_service.undo_batch(session, company, batch.id, host, port)
            except tally_connector.TallyConnectionError as exc:
                st.error(str(exc))
                return
        if result.deleted:
            st.success(f"Deleted {result.deleted} voucher(s) from Tally.")
            if result.deleted < len(live):
                st.info("Some vouchers were not found - they may have been deleted in Tally already.")
        else:
            st.error(
                "Tally did not delete anything. Delete these vouchers from Day Book by hand."
                + (" Tally said: " + "; ".join(result.messages[:3]) if result.messages else "")
            )



def history_page(session, company: Company) -> None:
    st.header("🕑 History")
    batches = (
        session.query(ImportBatch)
        .filter(ImportBatch.company_id == company.id)
        .order_by(ImportBatch.created_at.desc())
        .all()
    )
    if not batches:
        st.info("No imports yet.")
        return
    for b in batches:
        with st.expander(f"{b.created_at:%d %b %Y %H:%M} — {b.filename} ({b.row_count} rows)"):
            st.write(f"Skipped: {b.skipped_count}  ·  Balance mismatches: {b.mismatch_count}  ·  Total: ₹{b.total_amount}")
            _undo_section(session, company, b)
            d1, d2 = st.columns(2)
            with d1:
                sales_path = Path(b.sales_xml_path) if b.sales_xml_path else None
                if sales_path and sales_path.exists():
                    st.download_button(
                        "⬇️ Sales XML",
                        data=sales_path.read_text(encoding="utf-8"),
                        file_name=sales_path.name,
                        mime="application/xml",
                        key=f"dl_sales_{b.id}",
                        use_container_width=True,
                    )
                elif b.xml_path and Path(b.xml_path).exists():
                    # Older batch, made before the Sales/Journal split — only
                    # the combined file exists for it.
                    old_path = Path(b.xml_path)
                    st.download_button(
                        "⬇️ Download (combined)",
                        data=old_path.read_text(encoding="utf-8"),
                        file_name=old_path.name,
                        mime="application/xml",
                        key=f"dl_combined_{b.id}",
                        use_container_width=True,
                    )
                else:
                    st.caption("Sales file no longer on disk.")
            with d2:
                journal_path = Path(b.journal_xml_path) if b.journal_xml_path else None
                if journal_path and journal_path.exists():
                    st.download_button(
                        "⬇️ Journal XML",
                        data=journal_path.read_text(encoding="utf-8"),
                        file_name=journal_path.name,
                        mime="application/xml",
                        key=f"dl_journal_{b.id}",
                        use_container_width=True,
                    )
                elif not (b.xml_path and Path(b.xml_path).exists()):
                    st.caption("Journal file no longer on disk.")


# --------------------------------------------------------------- Pending --
def pending_page(session, company: Company) -> None:
    st.header("📥 Pending - waiting for Tally")
    items = pending_service.waiting(session, company.id)
    if not items:
        st.info("Nothing pending. On Import & Generate, “📥 Save to Pending” keeps checked "
                "entries here when Tally is not open.")
        return
    host, port = settings_service.get_tally_connection(session, company.id)
    ok, message = tally_connector.check_connection(host, port, company.tally_company_name or "")
    if ok:
        st.success(f"Tally is open - {message}")
    else:
        st.warning(f"Tally is not reachable right now. {message}")
        st.button("🔄 Check again", key="pending_recheck")

    total_rows = sum(i.row_count for i in items)
    total_amt = sum(float(i.total_amount or 0) for i in items)
    st.markdown(f"**{len(items)} upload(s) waiting · {total_rows} rows · freight ₹{total_amt:,.2f}**")

    ledger_names = settings_service.get_ledger_names(session, company.id)
    if st.button("🚀 Send all to Tally", type="primary", use_container_width=True, disabled=not ok):
        sent = failed = 0
        for item in items:
            with st.spinner(f"Sending {item.label} ..."):
                try:
                    out = pending_service.send_one(session, company, item, ledger_names, host, port)
                except tally_connector.TallyConnectionError as exc:
                    st.error(f"{item.label}: {exc}")
                    failed += 1
                    break
            if item.status == "sent":
                sent += 1
                st.write(f"✅ {item.label}: created {out.created} voucher(s)"
                         + (f", replaced {out.replaced_deleted}" if out.replaced_deleted else "")
                         + (f", skipped {out.skipped_bilties} bilty(s) already in Tally" if out.skipped_bilties else ""))
            else:
                failed += 1
                st.error(f"{item.label}: {item.last_error}")
        if sent and not failed:
            st.success(f"All {sent} pending upload(s) sent to Tally. They now appear in History (with Undo).")
        return

    for item in items:
        with st.container(border=True):
            c1, c2, c3 = st.columns([6, 2, 2])
            saved = item.created_at.strftime("%d %b %Y %H:%M") if item.created_at else ""
            c1.markdown(f"**{item.label}**  \n{item.row_count} rows · ₹{float(item.total_amount or 0):,.2f} · "
                        f"saved {saved} · {'Replace' if item.mode == 'replace' else 'Skip'} if already in Tally")
            if item.last_error:
                c1.caption(f"Last try: {item.last_error[:200]}")
            if c2.button("Send", key=f"pend_send_{item.id}", disabled=not ok, use_container_width=True):
                try:
                    out = pending_service.send_one(session, company, item, ledger_names, host, port)
                except tally_connector.TallyConnectionError as exc:
                    st.error(str(exc))
                else:
                    if item.status == "sent":
                        st.success(f"Sent - created {out.created} voucher(s).")
                        st.rerun()
                    else:
                        st.error(item.last_error)
            if c3.button("Delete", key=f"pend_del_{item.id}", use_container_width=True):
                pending_service.cancel(session, item)
                st.rerun()


# ------------------------------------------------------------- Dashboard --
def dashboard_page(session, company: Company) -> None:
    _inject_css()
    st.header(f"📊 Dashboard — {company.name}")
    n_wait = pending_service.count_waiting(session, company.id)
    if n_wait:
        p1, p2 = st.columns([4, 1])
        p1.warning(f"📥 {n_wait} upload(s) are waiting in Pending - open Tally and send them.")
        if p2.button("Open Pending", use_container_width=True):
            st.session_state["_nav_request"] = "Pending"
            st.rerun()

    batches = (
        session.query(ImportBatch)
        .filter(ImportBatch.company_id == company.id)
        .order_by(ImportBatch.created_at.desc())
        .all()
    )
    total_rows = sum(b.row_count for b in batches)
    total_amount = sum(float(b.total_amount or 0) for b in batches)
    total_files = len(batches)
    # Every usable row makes one Sales voucher; a row only makes a Journal
    # voucher when it actually has a deduction, so this is an estimate from
    # row counts alone (good enough for a dashboard tile, not for Tally).
    status = license_service.status(company)
    days_left = ""
    if company.license_expires_at:
        # SQLite round-trips datetimes as naive; license_service._aware()
        # does this same normalization before comparing, so match it here.
        expires = company.license_expires_at
        if expires.tzinfo is None:
            expires = expires.replace(tzinfo=dt.timezone.utc)
        delta = expires - dt.datetime.now(dt.timezone.utc)
        days_left = str(max(delta.days, 0))

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(_kpi_card("Files imported", str(total_files), "#2563eb"), unsafe_allow_html=True)
    with k2:
        st.markdown(_kpi_card("Rows processed", str(total_rows), "#059669"), unsafe_allow_html=True)
    with k3:
        st.markdown(_kpi_card("Total freight value", f"₹{total_amount:,.0f}", "#d97706"), unsafe_allow_html=True)
    with k4:
        st.markdown(
            _kpi_card("Licence days left", "Demo" if settings.DEMO_MODE else (days_left or "—"), "#7c3aed"),
            unsafe_allow_html=True,
        )

    st.write("")
    if settings.DEMO_MODE:
        pass
    elif not status.is_valid:
        st.warning(status.message)
    elif status.is_expiring_soon:
        st.warning(status.message)

    st.write("")
    st.subheader("Quick actions")
    b1, b2, b3 = st.columns(3)
    with b1:
        with st.container(key="nav_import"):
            if st.button("📤  Import & Generate", use_container_width=True):
                st.session_state["_nav_request"] = "Import & Generate"
                st.rerun()
    with b2:
        with st.container(key="nav_history"):
            if st.button("🕑  History", use_container_width=True):
                st.session_state["_nav_request"] = "History"
                st.rerun()
    with b3:
        with st.container(key="nav_settings"):
            if st.button("⚙️  Settings", use_container_width=True):
                st.session_state["_nav_request"] = "Settings"
                st.rerun()

    st.write("")
    st.subheader("Recent imports")
    if not batches:
        st.info("No imports yet — use Import & Generate to process your first register.")
        return
    report_rows = [
        {
            "Date": b.created_at.strftime("%d %b %Y %H:%M"),
            "File": b.filename,
            "Rows": b.row_count,
            "Skipped": b.skipped_count,
            "Mismatches": b.mismatch_count,
            "Total Amount": f"₹{float(b.total_amount or 0):,.2f}",
        }
        for b in batches[:10]
    ]
    st.dataframe(report_rows, use_container_width=True, hide_index=True)


# ----------------------------------------------------------------- Main ---
def main() -> None:
    if settings.DEMO_MODE and "user_id" not in st.session_state:
        with get_session() as session:
            cid, uid = demo_service.ensure(session)
        st.session_state["user_id"] = uid
        st.session_state["company_id"] = cid
    if "user_id" not in st.session_state:
        if _company_count() == 0:
            signup_page()
        else:
            login_page()
        return

    with get_session() as session:
        company = _current_company(session)
        n_pending = pending_service.count_waiting(session, company.id)
        # (label, accent key into _ACCENTS, icon)
        nav_items = [
            ("Dashboard", "dashboard", "📊"),
            ("Import & Generate", "import", "📤"),
            *([] if settings.DEMO_MODE else [("Pending", "pending", "📥")]),
            ("History", "history", "🕑"),
            ("Settings", "settings", "⚙️"),
        ]
        if "nav_choice" not in st.session_state:
            st.session_state["nav_choice"] = "Dashboard"
        # The dashboard's own quick-action buttons also want to change the
        # page, but they render *after* this sidebar in the same run, so
        # they can't flip "nav_choice" directly without a false click
        # appearing on the button that's rendered here before it. They drop
        # their request in "_nav_request" instead; this runs first and
        # applies it before any nav button below is drawn.
        if "_nav_request" in st.session_state:
            st.session_state["nav_choice"] = st.session_state.pop("_nav_request")

        _inject_css()
        with st.sidebar:
            st.markdown(f"### 🚚 {settings.APP_NAME}")
            st.caption(company.name)
            st.write("")
            for label, accent_key, icon in nav_items:
                is_active = st.session_state["nav_choice"] == label
                container_key = f"navside_{accent_key}_active" if is_active else f"navside_{accent_key}"
                with st.container(key=container_key):
                    text = f"{icon}  {label}"
                    if label == "Pending" and n_pending:
                        text += f" ({n_pending})"
                    if st.button(text, use_container_width=True, key=f"navbtn_{accent_key}"):
                        st.session_state["nav_choice"] = label
                        st.rerun()
            st.write("")
            if settings.DEMO_MODE:
                st.caption("🎯 DEMO VERSION")
                if st.button("↺ Reset demo"):
                    demo_service.reset(session)
                    st.session_state.clear()
                    st.rerun()
            elif st.button("Sign out"):
                st.session_state.clear()
                st.rerun()
        page = st.session_state["nav_choice"]
        if settings.DEMO_MODE:
            st.info(f"🎯 **DEMO VERSION** of {settings.APP_NAME} - {settings.COMPANY_BRAND}, "
                    f"📞 {settings.SUPPORT_CONTACT}. Sample data only; demo data is reset from time to time.")

        if not license_gate():
            return

        if page == "Dashboard":
            dashboard_page(session, company)
        elif page == "Import & Generate":
            import_page(session, company)
        elif page == "History":
            history_page(session, company)
        elif page == "Settings":
            settings_page(session, company)
        elif page == "Pending":
            pending_page(session, company)


if __name__ == "__main__":
    main()
