"""
Core engine: Excel freight/bilty register -> Tally XML.

Two posting styles (see POSTING_STYLES). The default "transporter" style is
documented on transporter_vouchers() near the end of this file. The older
"owner" style, per row, produces two Tally vouchers:

  1. SALES voucher  (the freight bill itself)
       Owner (party)  Dr   Amt
       Freight ledger Cr   Amt - TDS
       TDS ledger     Cr   TDS

  2. JOURNAL voucher  (deductions/advances adjusted against the owner)
       Owner (party)  Dr   Comm + Disc + Diesel + CashFromPump + CashFromOffice
                            + Other1 + Other2
       Commission ledger       Cr   Comm
       Discount ledger         Cr   Disc
       Petrol-pump ledger       Cr   Diesel   (falls back to the Cash ledger
                                               if no pump name is given)
       Cash ledger              Cr   CashFromPump + CashFromOffice
       Other-1 / Other-2 ledger Cr   Other1 / Other2

This mirrors exactly what was verified against a real customer's register
(ABIS SOFT, Apr 2026): the owner ledger's resulting balance equalled the
register's own BALANCE column on every row, with zero mismatches.

Every field is read through a column mapping (see FIELDS below) rather than
a fixed column letter, because different customers keep this register with
different column orders — only the accounting logic is fixed.
"""
from __future__ import annotations

import hashlib
import html
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional

import openpyxl

# field_name -> (label shown to the user, is this field required to process a row)
FIELDS: Dict[str, tuple] = {
    "date":         ("Date",                         True),
    "voucher_no":   ("Bilty / Voucher No.",          True),
    "owner":        ("Truck Owner name",             True),
    "truck_no":     ("Truck number",                 False),
    "party":        ("Party / Customer name (billed party)", False),
    "shipment_no":  ("Shipment No.",                 False),
    "from_place":   ("From (place)",                 False),
    "to_place":     ("To (place)",                   False),
    "item_name":    ("Goods carried (e.g. Poultry Feed - narration only)", False),
    "weight":       ("Weight (Qty in Sales invoice)", False),
    "rate":         ("Rate (Rate in Sales invoice)",  False),
    "amount":       ("Freight Amount",                True),
    "tds":          ("TDS",                          False),
    "commission":   ("Commission",                  False),
    "discount":     ("Discount",                     False),
    "pump_name":    ("Petrol Pump name",             False),
    "diesel":       ("Diesel amount",                False),
    "cash_pump":    ("Cash paid from pump",          False),
    "cash_office":  ("Cash paid from office",        False),
    "other1":       ("Other deduction 1",            False),
    "other2":       ("Other deduction 2",            False),
    "slip_no":      ("Diesel slip no.",              False),
    "balance":      ("Balance / Transport Exp. (for checking only)", False),
}
REQUIRED_FIELDS = [name for name, (_, req) in FIELDS.items() if req]

DEFAULT_LEDGER_NAMES = {
    "freight": "Freight Receipts",
    "tds": "TDS Receivable",
    "commission": "Commission Received",
    "discount": "Discount Received",
    "cash": "Cash",
    "other1": "Other Deduction 1",
    "other2": "Other Deduction 2",
    # used by the "transporter" posting style only
    "party": "",                              # default billed party, if no column
    "diesel_exp": "DIESEL EXPENSES",
    "other_exp": "OTHER EXPENSES",
    "transport_exp": "TRANSPORTING EXPENSES",
    # Sales voucher format (transporter style): "invoice" = item invoice with
    # stock item, Qty (= weight) and Rate, like Tally's Sales item invoice;
    # "accounting" = party Dr / freight ledger Cr only.
    "sales_mode": "invoice",
    "item_name": "FREIGHT",      # stock item when the Excel has no item column
    "item_unit": "MT",           # unit of measure for the weight
    # Voucher number = bilty number ("yes"/"no"), for Sales and for the
    # Payment/Journal vouchers of the same bilty.
    # Other1 / Other2 columns: "advance" = money already given to the truck
    # owner (reduces Transporting Expenses, posted as a Payment like cash
    # from office); "deduction" = kept back from the owner (credited to the
    # Other1/Other2 ledgers as income).
    "other_mode": "advance",
    # Dr ledger for cash paid from office and the Other1/Other2 advances.
    # Blank = the Transporting Expenses ledger (they are part of the truck hire).
    "office_exp": "",
    "other1_exp": "",
    "other2_exp": "",
    "other1_from": "Cash",       # ledger the Other1 advance was paid from
    "other2_from": "Cash",
    "vchno_sales": "yes",
    "vchno_others": "no",
}

# Two ways of posting the same register:
#   "transporter" - the transport company's own books (verified against
#                   JAIN TRANSPORT TRAIL in Tally, Oct 2026). Per row:
#                   Payment (cash from office), Journal (TDS), Journal
#                   (diesel + cash from pump), Journal (transporting exp.
#                   payable to truck owner), Sales (freight bill to party).
#   "owner"       - the original style: owner as the party, one Sales and
#                   one Journal voucher per row.
POSTING_STYLES = {
    "transporter": "Transporter books - Payment + Journal + Sales (Jain Transport style)",
    "owner": "Owner ledger - Sales + one deduction Journal (old style)",
}
DEFAULT_POSTING_STYLE = "transporter"

# The five vouchers the transporter style can make per row. The user can
# switch any of them off (Settings); the others are not affected - e.g.
# switching off "TDS" makes no TDS journal but the Sales bill stays the
# full freight amount. The choice travels inside the ledger-names dict
# under POST_KINDS_KEY so every builder sees the same setting.
VOUCHER_KINDS = {
    "PAY": "Payment - Cash from office, Other1/Other2 advances (Other Expenses Dr / Cash Cr)",
    "TDS": "Journal - TDS (TDS Dr / Party Cr)",
    "DSL": "Journal - Diesel + Cash from pump (Diesel/Other Exp. Dr / Pump Cr)",
    "TRN": "Journal - Transporting Expenses (Transporting Exp. Dr / Owner Cr)",
    "SAL": "Sales - Freight bill (Party Dr / Freight Cr)",
}
POST_KINDS_KEY = "_post_kinds"


def enabled_kinds(names: Dict[str, str]) -> set:
    value = (names or {}).get(POST_KINDS_KEY)
    if value is None or value == "":
        return set(VOUCHER_KINDS) | {"OT1", "OT2", "OTH"}
    kinds = {k for k in str(value).split(",") if k in VOUCHER_KINDS}
    if "PAY" in kinds:
        kinds |= {"OT1", "OT2", "OTH"}   # Other1/Other2 advance payments follow the Payment switch
    return kinds
DEFAULT_OWNER_GROUP = "Sundry Debtors"
DEFAULT_PUMP_GROUP = "Sundry Creditors"


class MappingError(Exception):
    """A column mapping that cannot be used, with a message safe to show."""


@dataclass
class ProcessedRow:
    source_row: int
    date: Optional[datetime]
    voucher_no: str
    owner: str
    truck_no: str
    amount: float
    tds: float
    sale_amount: float          # amount - tds, credited to the freight ledger
    deductions: List[tuple] = field(default_factory=list)   # (ledger, amount)
    total_deduction: float = 0.0
    computed_balance: float = 0.0
    excel_balance: Optional[float] = None
    balance_mismatch: bool = False
    problem: Optional[str] = None   # set when the row could not be processed
    # extra detail used by the "transporter" posting style
    party: str = ""
    shipment_no: str = ""
    from_place: str = ""
    to_place: str = ""
    slip_no: str = ""
    pump: str = ""
    commission: float = 0.0
    diesel: float = 0.0
    cash_pump: float = 0.0
    cash_office: float = 0.0
    other_deductions: List[tuple] = field(default_factory=list)  # (ledger, amount)
    transport_amount: float = 0.0   # Dr Transporting Expenses
    owner_payable: float = 0.0      # Cr truck owner
    occurrence: int = 1             # 2, 3.. when the same bilty+date repeats in one file
    raw: Dict[str, object] = field(default_factory=dict)   # cell values as read
    edited: bool = False            # corrected on screen before posting
    other_advances: List[tuple] = field(default_factory=list)  # ("other1"/"other2", paid-from ledger, amount)
    weight: float = 0.0
    rate: float = 0.0
    item_name: str = ""
    source: str = ""                # "file / sheet" when several are processed together
    warning: Optional[str] = None   # posted, but worth a look (e.g. owner balance negative)


def _num(value) -> float:
    try:
        if value is None or (isinstance(value, str) and not value.strip()):
            return 0.0
        number = float(value)
        if number != number:      # NaN - an emptied cell in the on-screen editor
            return 0.0
        return round(number, 2)
    except (TypeError, ValueError):
        return 0.0


def _text(value) -> str:
    if value is None or (isinstance(value, float) and value != value):
        return ""
    text = str(value).strip()
    return "" if text in ("nan", "None", "NaT") else text


_DATE_FORMATS = ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y", "%d/%m/%y")


def _parse_date(value) -> Optional[datetime]:
    """openpyxl hands back a real datetime for a date-typed cell. Some
    registers keep the date column as typed text instead, so a few common
    Indian date formats are tried before giving up."""
    if isinstance(value, datetime):
        return value
    text = _text(value)
    if not text:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


_HEADER_SCAN_ROWS = 20


def open_workbook(file_like):
    """An openpyxl workbook (values, not formulas). Accepts a path, an
    uploaded file, or an already-open workbook so a file read once can be
    looked at sheet by sheet without re-reading it."""
    if isinstance(file_like, openpyxl.Workbook):
        return file_like
    if hasattr(file_like, "seek"):
        file_like.seek(0)
    wb = openpyxl.load_workbook(file_like, data_only=True)
    if hasattr(file_like, "seek"):
        file_like.seek(0)
    return wb


def _sheet(wb, sheet: Optional[str]):
    return wb[sheet] if sheet else wb.worksheets[0]


def data_sheets(file_like) -> List[str]:
    """Names of the sheets that hold anything more than a heading row."""
    wb = open_workbook(file_like)
    return [ws.title for ws in wb.worksheets if (ws.max_row or 0) >= 2]


def _norm(text) -> str:
    return "".join(str(text).lower().split()) if text is not None else ""


def _find_header_row(ws, wanted: Optional[List[str]] = None) -> int:
    """Row number holding the column headers.

    Many registers have a title row (or a merged "SALE ENTRY DETAILS /
    JOURNAL ENTRY DETAILS" band) above the real headers, so the header row
    is not always row 1. With a mapping, the row matching the most mapped
    header names wins; without one, the row with the most text cells."""
    wanted_norm = {_norm(w) for w in (wanted or []) if w}
    best_row, best_score = 1, -1
    max_r = min(ws.max_row or 1, _HEADER_SCAN_ROWS)
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_r, values_only=True), start=1):
        texts = [v for v in row if isinstance(v, str) and v.strip()]
        if wanted_norm:
            score = len({_norm(v) for v in texts} & wanted_norm)
        else:
            score = len(texts)
        if score > best_score:
            best_row, best_score = r, score
    return best_row


def read_headers(file_like, sheet: Optional[str] = None) -> List[str]:
    """The header row of a sheet (the first by default), as header strings."""
    wb = open_workbook(file_like)
    ws = _sheet(wb, sheet)
    header_row = _find_header_row(ws)
    return [_text(c.value) for c in ws[header_row]]


def _column_index(ws, header_name: str, header_row: int = 1) -> Optional[int]:
    """1-based column index of a header.

    Exact text is tried first, so two headers that differ only in case
    ("Date" for the bilty date and "DATE" for the diesel-slip date in the
    ABIS register) still map to their own columns. Then a case/space-
    insensitive match, then a plain column letter such as "K"."""
    cells = list(ws[header_row])
    for cell in cells:
        if cell.value is not None and str(cell.value).strip() == header_name.strip():
            return cell.column
    target = _norm(header_name)
    for cell in cells:
        if cell.value and _norm(cell.value) == target:
            return cell.column
    letters = header_name.strip().upper()
    if letters.isalpha() and len(letters) <= 3:
        from openpyxl.utils import column_index_from_string
        try:
            return column_index_from_string(letters)
        except ValueError:
            return None
    return None


# Common header spellings, used to pre-fill the mapping for a new layout.
HEADER_SYNONYMS: Dict[str, List[str]] = {
    "date": ["date", "bilty date", "lr date"],
    "voucher_no": ["builty", "bilty", "m-bl/no", "bl no", "blno", "lr no", "lrno", "bilty no", "gr no"],
    "owner": ["owner", "truck owner", "owner name"],
    "truck_no": ["truck", "truckno", "truck no", "vehicle", "vehicle no"],
    "party": ["party", "customer", "consignor", "party name"],
    "shipment_no": ["shipment no", "shipment", "gpno", "gp no"],
    "from_place": ["from"],
    "to_place": ["to", "destination"],
    "item_name": ["item", "item name", "goods", "material"],
    "weight": ["wt", "weight"],
    "rate": ["rate"],
    "amount": ["amt", "amount", "freight", "freight amount"],
    "tds": ["tds"],
    "commission": ["comm", "commission", "commision", "comission"],
    "discount": ["disc", "discount"],
    "pump_name": ["petrol pump", "pump", "pump name"],
    "diesel": ["diesel"],
    "slip_no": ["slip no", "slipno", "slip"],
    "cash_pump": ["cash from pump", "pump cash", "diesel-cash", "diesel cash", "cash diesel"],
    "cash_office": ["cash from office", "office cash"],
    "other1": ["other1", "other 1"],
    "other2": ["other2", "other 2"],
    "balance": ["balance", "transport expenses", "transporting expenses"],
}


def guess_mapping(headers: List[str]) -> Dict[str, str]:
    """Best-guess field -> header for a register not mapped before. The
    first matching header wins, so "Date" (bilty date) beats a later
    "DATE" (diesel-slip date)."""
    guess: Dict[str, str] = {}
    used = set()
    for field_name, names in HEADER_SYNONYMS.items():
        wanted = {_norm(n) for n in names}
        for h in headers:
            if h and h not in used and _norm(h) in wanted:
                guess[field_name] = h
                used.add(h)
                break
    return guess


def match_mappings(file_like, mappings: Dict[str, Dict[str, str]], sheet: Optional[str] = None) -> List[tuple]:
    """How well each saved layout fits this file (or one sheet of it), best
    first: [(name, fits, matched columns, [missing column headers])].
    A layout "fits" when every column it names is in the sheet."""
    wb = open_workbook(file_like)
    ws = _sheet(wb, sheet)
    results = []
    for name, mapping in mappings.items():
        headers = [h for f, h in mapping.items() if h and f in FIELDS]
        header_row = _find_header_row(ws, headers)
        missing = [h for h in headers if _column_index(ws, h, header_row) is None
                   or (len(h.strip()) <= 3 and h.strip().isalpha() and not any(
                       _norm(c.value) == _norm(h) for c in ws[header_row] if c.value))]
        results.append((name, not missing and not validate_mapping(mapping), len(headers) - len(missing), missing))
    if hasattr(file_like, "seek"):
        file_like.seek(0)
    results.sort(key=lambda t: (not t[1], -t[2]))
    return results


NOT_MAPPED = "(not in this file)"


def build_mapping_options(headers: List[str]) -> Optional[List[str]]:
    """The dropdown choices for a column-mapping field, or None if no file
    is loaded yet (the caller then falls back to a plain text box).

    A column with no header text at all - a blank cell in the header row,
    which a real customer register (ABIS SOFT's) has - is dropped: nobody
    could pick it by name anyway, and keeping it caused a worse bug. "Not
    yet mapped" is represented internally as "", and a blank header is also
    "" - so every unmapped field matched that blank column and every
    dropdown looked empty, even though the right headers were right there.
    """
    named = [h for h in headers if h]
    return [NOT_MAPPED] + named if named else None


def safe_default_index(options: List[str], default: str) -> int:
    """Index to pre-select in a mapping dropdown.

    `default` is "" for a field that has no saved mapping yet, and must
    never be looked up with `in`/`index` against `options` in that case -
    see build_mapping_options' docstring for why that lookup can silently
    land on the wrong column.
    """
    if default and default in options:
        return options.index(default)
    return 0


def duplicate_columns(mapping: Dict[str, str]) -> List[tuple]:
    """Excel columns picked for more than one field, e.g. the "diesel"
    column chosen for both Diesel and Cash from pump - which makes that
    amount count twice. [(column, [field labels])]."""
    used: Dict[str, List[str]] = {}
    for field_name, header in mapping.items():
        if header and field_name in FIELDS:
            used.setdefault(header.strip(), []).append(FIELDS[field_name][0])
    return [(h, labels) for h, labels in used.items() if len(labels) > 1]


def duplicate_message(dups: List[tuple]) -> str:
    return "; ".join(f"column “{h}” is chosen for {' AND '.join(labels)}" for h, labels in dups)


def validate_mapping(mapping: Dict[str, str]) -> List[str]:
    """Which required fields have no column assigned."""
    missing = []
    for field_name in REQUIRED_FIELDS:
        if not mapping.get(field_name):
            missing.append(FIELDS[field_name][0])
    return missing


def process_workbook(
    file_like,
    mapping: Dict[str, str],
    ledger_names: Optional[Dict[str, str]] = None,
    style: str = "owner",
    sheet: Optional[str] = None,
    source: str = "",
) -> List[ProcessedRow]:
    """Read every data row of one sheet (the first by default) and work out
    its vouchers' entries. `source` labels the rows (e.g. "file / sheet")
    when several files or sheets are processed together."""
    names = {**DEFAULT_LEDGER_NAMES, **(ledger_names or {})}
    missing = validate_mapping(mapping)
    if missing:
        raise MappingError(
            "These required fields are not mapped to a column yet: " + ", ".join(missing)
        )

    wb = open_workbook(file_like)
    ws = _sheet(wb, sheet)
    header_row = _find_header_row(ws, [h for h in mapping.values() if h])

    columns: Dict[str, int] = {}
    for field_name, header in mapping.items():
        if not header or field_name not in FIELDS:
            continue
        idx = _column_index(ws, header, header_row)
        if idx is None:
            raise MappingError(f"Column '{header}' (for {FIELDS[field_name][0]}) was not found in this file.")
        columns[field_name] = idx

    results: List[ProcessedRow] = []
    for r in range(header_row + 1, ws.max_row + 1):
        values = {}
        for field_name, idx in columns.items():
            values[field_name] = ws.cell(r, idx).value
        if not _text(values.get("voucher_no")):
            continue  # a wholly blank trailing row
        row = build_row(r, values, names, style)
        row.source = source
        results.append(row)

    number_occurrences(results)
    return results


def build_row(source_row: int, values: Dict[str, object], names: Dict[str, str],
              style: str = "owner", edited: bool = False) -> ProcessedRow:
    """Work out one register row's figures from its raw field values
    ({field name: cell value}). Used for rows read from Excel and again
    for rows corrected on screen."""
    def cell(field_name):
        return values.get(field_name)

    date_val = cell("date")
    owner = _text(cell("owner"))
    amount = _num(cell("amount"))
    tds = _num(cell("tds"))
    row = ProcessedRow(
        source_row=source_row,
        date=_parse_date(date_val),
        voucher_no=_text(cell("voucher_no")),
        owner=owner,
        truck_no=_text(cell("truck_no")),
        amount=amount,
        tds=tds,
        sale_amount=round(amount - tds, 2),
    )

    row.party = _text(cell("party")) or names.get("party", "")
    row.shipment_no = _text(cell("shipment_no"))
    row.from_place = _text(cell("from_place"))
    row.to_place = _text(cell("to_place"))
    row.slip_no = _text(cell("slip_no"))
    row.weight = _num(cell("weight"))
    row.rate = _num(cell("rate"))
    row.item_name = _text(cell("item_name"))

    if not row.voucher_no:
        row.problem = "Bilty / voucher number is missing."
    elif not owner:
        row.problem = "No owner/party name — this row cannot be posted."
    elif row.date is None:
        row.problem = "Date is missing or not a recognisable date."
    elif amount <= 0:
        row.problem = "Freight amount is zero or missing."

    comm = _num(cell("commission"))
    disc = _num(cell("discount"))
    diesel = _num(cell("diesel"))
    pump = _text(cell("pump_name"))
    cash_pump = _num(cell("cash_pump"))
    cash_office = _num(cell("cash_office"))
    other1 = _num(cell("other1"))
    other2 = _num(cell("other2"))

    if comm:
        row.deductions.append((names["commission"], comm))
    if disc:
        row.deductions.append((names["discount"], disc))
    if diesel:
        row.deductions.append((pump if pump else names["cash"], diesel))
    cash_total = round(cash_pump + cash_office, 2)
    if cash_total:
        row.deductions.append((names["cash"], cash_total))
    if other1:
        row.deductions.append((names["other1"], other1))
    if other2:
        row.deductions.append((names["other2"], other2))

    row.total_deduction = round(sum(a for _, a in row.deductions), 2)
    row.computed_balance = round(amount - row.total_deduction, 2)

    # transporter-style figures (same numbers, different posting)
    row.pump = pump
    row.commission = comm
    row.diesel = diesel
    row.cash_pump = cash_pump
    row.cash_office = cash_office
    advance_mode = names.get("other_mode", "advance") == "advance"
    if disc:
        row.other_deductions.append((names["discount"], disc))
    for key, amt in (("other1", other1), ("other2", other2)):
        if not amt:
            continue
        if advance_mode:
            row.other_advances.append((key, names.get(f"{key}_from") or names["cash"], amt))
        else:
            row.other_deductions.append((names[key], amt))
    advances = sum(a for _, _, a in row.other_advances)
    row.transport_amount = round(amount - comm - diesel - cash_pump - cash_office - advances, 2)
    row.owner_payable = round(
        row.transport_amount - sum(a for _, a in row.other_deductions), 2
    )
    if style == "transporter" and not row.problem:
        if not row.party:
            row.problem = (
                "No party for the Sales bill - type it in “Party for this file” above "
                "(or in the Party column of the table)."
            )
        elif row.owner_payable < 0:
            # Advances (diesel/cash) were more than the hire: the owner now
            # owes the transporter. Posted, not skipped - see transporter_vouchers.
            row.warning = (
                f"Advances are more than the freight - owner {row.owner} owes you "
                f"{-row.owner_payable:,.2f} (posted as Owner Dr)."
            )

    excel_balance = cell("balance")
    if excel_balance not in (None, ""):
        row.excel_balance = _num(excel_balance)
        row.balance_mismatch = abs(row.excel_balance - row.computed_balance) > 1.0


    row.raw = dict(values)
    row.edited = edited
    if edited:
        # The sheet's own balance no longer applies to a hand-corrected row.
        row.excel_balance = None
        row.balance_mismatch = False
    return row


def number_occurrences(rows: List[ProcessedRow]) -> None:
    """2, 3.. for a bilty+date that repeats in one file (kept distinct in Tally)."""
    seen: Dict[tuple, int] = {}
    for row in rows:
        row.occurrence = 1
        if row.date is not None:
            key = (row.date.strftime("%Y%m%d"), row.voucher_no.upper())
            seen[key] = seen.get(key, 0) + 1
            row.occurrence = seen[key]


# Fields the user may correct on the preview screen, in display order.
EDITABLE_FIELDS = [
    ("date", "Date"), ("voucher_no", "Bilty"), ("truck_no", "Truck"), ("party", "Party"),
    ("owner", "Owner"), ("item_name", "Goods"), ("weight", "Wt (Qty)"), ("rate", "Rate"),
    ("amount", "Freight Amt"), ("tds", "TDS"), ("commission", "Comm"),
    ("pump_name", "Petrol Pump"), ("diesel", "Diesel"), ("slip_no", "Slip No"),
    ("cash_pump", "Cash from Pump"), ("cash_office", "Cash from Office"),
    ("other1", "Other1"), ("other2", "Other2"),
]


def editor_records(rows: List[ProcessedRow]) -> List[dict]:
    """Rows as plain dicts for the on-screen editor (current values)."""
    out = []
    for r in rows:
        out.append({
            "File / Sheet": r.source,
            "Excel row": r.source_row,
            "Date": r.date.date() if r.date else None,
            "Bilty": r.voucher_no,
            "Truck": r.truck_no,
            "Party": r.party,
            "Owner": r.owner,
            "Goods": r.item_name,
            "Wt (Qty)": r.weight,
            "Rate": r.rate,
            "Freight Amt": r.amount,
            "TDS": r.tds,
            "Comm": r.commission,
            "Petrol Pump": r.pump,
            "Diesel": r.diesel,
            "Slip No": r.slip_no,
            "Cash from Pump": r.cash_pump,
            "Cash from Office": r.cash_office,
            "Other1": _num((r.raw or {}).get("other1")),
            "Other2": _num((r.raw or {}).get("other2")),
            "Transporting Exp.": r.transport_amount,
            "Problem": r.problem or "",
        })
    return out


def _same(a, b) -> bool:
    if isinstance(a, (int, float)) or isinstance(b, (int, float)):
        return abs(_num(a) - _num(b)) < 0.005
    if hasattr(a, "strftime") and hasattr(b, "strftime"):
        return a.strftime("%Y%m%d") == b.strftime("%Y%m%d")
    return _text(a) == _text(b)


def apply_edits(rows: List[ProcessedRow], edited: List[dict], names: Optional[Dict[str, str]] = None,
                style: str = "owner") -> List[ProcessedRow]:
    """Rebuild rows from the on-screen editor's records (same order as
    editor_records). A row whose values were not changed is kept as read."""
    names = {**DEFAULT_LEDGER_NAMES, **(names or {})}
    original = editor_records(rows)
    out = []
    for r, before, after in zip(rows, original, edited):
        changed = [f for f, label in EDITABLE_FIELDS if not _same(before.get(label), after.get(label))]
        if not changed:
            out.append(r)
            continue
        values = dict(getattr(r, "raw", {}) or {})
        for f, label in EDITABLE_FIELDS:
            v = after.get(label)
            if f == "date" and v is not None and not isinstance(v, datetime):
                v = datetime(v.year, v.month, v.day) if hasattr(v, "year") else v
            values[f] = v
        new = build_row(r.source_row, values, names, style, edited=True)
        new.source = r.source
        out.append(new)
    number_occurrences(out)
    return out


def _esc(value) -> str:
    return html.escape(str(value), quote=False)


def _ledger_master(name: str, group: str, billwise: bool = False) -> str:
    return f"""
      <LEDGER NAME="{_esc(name)}" ACTION="Create">
        <NAME.LIST><NAME>{_esc(name)}</NAME></NAME.LIST>
        <PARENT>{_esc(group)}</PARENT>
        <ISBILLWISEON>{"Yes" if billwise else "No"}</ISBILLWISEON>
        <OPENINGBALANCE>0</OPENINGBALANCE>
      </LEDGER>"""


def _dr(ledger: str, amount: float) -> str:
    return f"""
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>{_esc(ledger)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
          <AMOUNT>-{amount:.2f}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>"""


def _cr(ledger: str, amount: float) -> str:
    return f"""
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>{_esc(ledger)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
          <AMOUNT>{amount:.2f}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>"""


def _sales_voucher_xml(r: "ProcessedRow", names: Dict[str, str]) -> str:
    d = r.date.strftime("%Y%m%d")
    entries = _dr(r.owner, r.amount)
    entries += _cr(names["freight"], r.sale_amount)
    if r.tds:
        entries += _cr(names["tds"], r.tds)
    narration = f"Freight - {r.truck_no}".strip(" -") if r.truck_no else "Freight"
    return f"""
      <VOUCHER VCHTYPE="Sales" ACTION="Create">
        <DATE>{d}</DATE>
        <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
        <VOUCHERNUMBER>{_esc(r.voucher_no)}</VOUCHERNUMBER>
        <PARTYLEDGERNAME>{_esc(r.owner)}</PARTYLEDGERNAME>
        <NARRATION>{_esc(narration)}</NARRATION>
        {entries}
      </VOUCHER>"""


def _journal_voucher_xml(r: "ProcessedRow") -> Optional[str]:
    if not r.deductions:
        return None
    d = r.date.strftime("%Y%m%d")
    entries = _dr(r.owner, r.total_deduction)
    for ledger, amt in r.deductions:
        entries += _cr(ledger, amt)
    return f"""
      <VOUCHER VCHTYPE="Journal" ACTION="Create">
        <DATE>{d}</DATE>
        <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>
        <VOUCHERNUMBER>{_esc(r.voucher_no)}-J</VOUCHERNUMBER>
        <NARRATION>Deductions/advance against {_esc(r.voucher_no)}</NARRATION>
        {entries}
      </VOUCHER>"""


def _ledger_masters_xml(
    rows: List["ProcessedRow"], names: Dict[str, str], *, include_sales: bool, include_journal: bool
) -> List[str]:
    """Ledger masters for the requested voucher type(s). Both voucher types
    debit the owner, so the owner ledger is included whenever either type
    is requested - importing just one file on its own still creates every
    ledger that file's vouchers reference."""
    owners = sorted({r.owner for r in rows})
    pumps = sorted({
        ledger for r in rows for ledger, _ in r.deductions
        if ledger not in names.values() and ledger != names["cash"]
    }) if include_journal else []

    used_ledgers = set()
    if include_sales:
        used_ledgers.add(names["freight"])
        if any(r.tds for r in rows):
            used_ledgers.add(names["tds"])
    if include_journal:
        used_ledgers |= {ledger for r in rows for ledger, _ in r.deductions}

    fixed = [
        (names["freight"], "Sales Accounts"),
        (names["tds"], "Current Assets"),
        (names["commission"], "Indirect Incomes"),
        (names["discount"], "Indirect Incomes"),
        (names["cash"], "Cash-in-Hand"),
        (names["other1"], "Indirect Incomes"),
        (names["other2"], "Indirect Incomes"),
    ]
    ledger_xml = [
        _ledger_master(name, group) for name, group in fixed if name in used_ledgers
    ]
    for owner in owners:
        ledger_xml.append(_ledger_master(owner, DEFAULT_OWNER_GROUP, billwise=True))
    for pump in pumps:
        ledger_xml.append(_ledger_master(pump, DEFAULT_PUMP_GROUP, billwise=True))
    return ledger_xml


def _envelope(report_name: str, inner_xml: List[str], company: str) -> str:
    company_esc = _esc(company or "")
    return f"""<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>{report_name}</REPORTNAME>
        <STATICVARIABLES><SVCURRENTCOMPANY>{company_esc}</SVCURRENTCOMPANY></STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">{''.join(inner_xml)}
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>
"""


def build_tally_xml(
    rows: List[ProcessedRow],
    tally_company_name: str,
    ledger_names: Optional[Dict[str, str]] = None,
) -> str:
    """Both voucher types in one file: masters, then every Sales and Journal
    voucher together - ready to import via Gateway of Tally -> Import Data
    -> XML in a single pass."""
    names = {**DEFAULT_LEDGER_NAMES, **(ledger_names or {})}
    usable = [r for r in rows if not r.problem]
    ledger_xml = _ledger_masters_xml(usable, names, include_sales=True, include_journal=True)
    voucher_xml = []
    for r in usable:
        voucher_xml.append(_sales_voucher_xml(r, names))
        journal = _journal_voucher_xml(r)
        if journal:
            voucher_xml.append(journal)
    return (
        _envelope("All Masters", ledger_xml, tally_company_name)
        + _envelope("Vouchers", voucher_xml, tally_company_name)
    )


def build_sales_xml(
    rows: List[ProcessedRow],
    tally_company_name: str,
    ledger_names: Optional[Dict[str, str]] = None,
) -> str:
    """Just the freight-bill (Sales) vouchers - the I-to-L columns."""
    names = {**DEFAULT_LEDGER_NAMES, **(ledger_names or {})}
    usable = [r for r in rows if not r.problem]
    ledger_xml = _ledger_masters_xml(usable, names, include_sales=True, include_journal=False)
    voucher_xml = [_sales_voucher_xml(r, names) for r in usable]
    return (
        _envelope("All Masters", ledger_xml, tally_company_name)
        + _envelope("Vouchers", voucher_xml, tally_company_name)
    )


def build_journal_xml(
    rows: List[ProcessedRow],
    tally_company_name: str,
    ledger_names: Optional[Dict[str, str]] = None,
) -> str:
    """Just the deductions/advance (Journal) vouchers - the M-onward columns.
    Rows with no deductions at all produce no journal voucher, same as in
    the combined file."""
    names = {**DEFAULT_LEDGER_NAMES, **(ledger_names or {})}
    usable = [r for r in rows if not r.problem]
    ledger_xml = _ledger_masters_xml(usable, names, include_sales=False, include_journal=True)
    voucher_xml = [j for r in usable if (j := _journal_voucher_xml(r))]
    return (
        _envelope("All Masters", ledger_xml, tally_company_name)
        + _envelope("Vouchers", voucher_xml, tally_company_name)
    )


# ------------------------------------------------- transporter posting style
def company_tag(tally_company_name: str) -> str:
    """Short stable code for a Tally company, so REMOTEIDs never collide
    between two companies kept in the same Tally."""
    return hashlib.sha1((tally_company_name or "").strip().upper().encode()).hexdigest()[:6].upper()


def voucher_remote_id(ctag: str, r: "ProcessedRow", kind: str) -> str:
    """REMOTEID that marks a voucher as made by this app for one bilty row:
    SNBT-<company>-<date>-<bilty>[#n]-<kind>. Tally stores it with the
    voucher, which is how a corrected re-send replaces the old vouchers and
    'Undo' finds them again."""
    bilty = re.sub(r"[^A-Za-z0-9/_.-]", "", r.voucher_no) or "NA"
    suffix = f"#{r.occurrence}" if r.occurrence > 1 else ""
    return f"SNBT-{ctag}-{r.date:%Y%m%d}-{bilty}{suffix}-{kind}"


def _voucher(vtype: str, r: "ProcessedRow", narration: str, entries: str,
             party: Optional[str] = None, reference: Optional[str] = None,
             remote_id: Optional[str] = None, invoice: bool = False,
             number: Optional[str] = None) -> str:
    """One voucher. No VOUCHERNUMBER is sent, so Tally numbers each voucher
    type itself (Payment 1, Journal 1/2/3, Sales 1 ...) exactly as when
    entered by hand; the bilty number goes in the narration/reference."""
    d = r.date.strftime("%Y%m%d")
    party_xml = f"\n        <PARTYLEDGERNAME>{_esc(party)}</PARTYLEDGERNAME>" if party else ""
    ref_xml = f"\n        <REFERENCE>{_esc(reference)}</REFERENCE>" if reference else ""
    remote_attr = f' REMOTEID="{html.escape(remote_id, quote=True)}"' if remote_id else ""
    number_xml = f"\n        <VOUCHERNUMBER>{_esc(number)}</VOUCHERNUMBER>" if number else ""
    view = "Invoice Voucher View" if invoice else "Accounting Voucher View"
    invoice_xml = "\n        <ISINVOICE>Yes</ISINVOICE>" if invoice else ""
    return f"""
      <VOUCHER{remote_attr} VCHTYPE="{vtype}" ACTION="Create" OBJVIEW="{view}">
        <DATE>{d}</DATE>
        <EFFECTIVEDATE>{d}</EFFECTIVEDATE>
        <VOUCHERTYPENAME>{vtype}</VOUCHERTYPENAME>{number_xml}{party_xml}{ref_xml}
        <PERSISTEDVIEW>{view}</PERSISTEDVIEW>{invoice_xml}
        <NARRATION>{_esc(narration)}</NARRATION>
        {entries}
      </VOUCHER>"""


def sales_item(r: "ProcessedRow", names: Dict[str, str]) -> tuple:
    """(stock item, qty, rate, unit) for the Sales invoice line.
    Qty is the Excel weight and Rate the Excel rate; the amount is always
    the Excel amount (weight x rate is often rounded in the register).
    With no weight the line is 1 x amount."""
    # Always the freight item (e.g. FREIGHT) - the register's own item
    # column holds the goods carried (Poultry Feed, DOC...), which only
    # goes into the narration.
    item = names.get("item_name") or "FREIGHT"
    unit = names.get("item_unit") or "MT"
    if r.weight > 0:
        rate = r.rate if r.rate > 0 else round(r.amount / r.weight, 2)
        return item, r.weight, rate, unit
    return item, 1.0, r.amount, unit


def _qty(q: float) -> str:
    return f"{q:.3f}".rstrip("0").rstrip(".")


def _invoice_entries(r: "ProcessedRow", names: Dict[str, str]) -> str:
    item, qty, rate, unit = sales_item(r, names)
    return f"""
        <LEDGERENTRIES.LIST>
          <LEDGERNAME>{_esc(r.party)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
          <AMOUNT>-{r.amount:.2f}</AMOUNT>
          <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
        </LEDGERENTRIES.LIST>
        <ALLINVENTORYENTRIES.LIST>
          <STOCKITEMNAME>{_esc(item)}</STOCKITEMNAME>
          <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
          <RATE>{rate:.2f}/{_esc(unit)}</RATE>
          <AMOUNT>{r.amount:.2f}</AMOUNT>
          <ACTUALQTY> {_qty(qty)} {_esc(unit)}</ACTUALQTY>
          <BILLEDQTY> {_qty(qty)} {_esc(unit)}</BILLEDQTY>
          <ACCOUNTINGALLOCATIONS.LIST>
            <LEDGERNAME>{_esc(names["freight"])}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <AMOUNT>{r.amount:.2f}</AMOUNT>
          </ACCOUNTINGALLOCATIONS.LIST>
        </ALLINVENTORYENTRIES.LIST>"""


def _unit_master(unit: str) -> str:
    return f"""
      <UNIT NAME="{_esc(unit)}" ACTION="Create">
        <NAME>{_esc(unit)}</NAME>
        <ISSIMPLEUNIT>Yes</ISSIMPLEUNIT>
        <DECIMALPLACES>3</DECIMALPLACES>
      </UNIT>"""


def _stock_item_master(item: str, unit: str) -> str:
    return f"""
      <STOCKITEM NAME="{_esc(item)}" ACTION="Create">
        <NAME.LIST><NAME>{_esc(item)}</NAME></NAME.LIST>
        <BASEUNITS>{_esc(unit)}</BASEUNITS>
      </STOCKITEM>"""


def _row_tag(r: "ProcessedRow") -> str:
    parts = [f"Bilty {r.voucher_no}"]
    if r.truck_no:
        parts.append(f"Truck {r.truck_no}")
    if r.shipment_no:
        parts.append(f"Shipment {r.shipment_no}")
    if r.from_place or r.to_place:
        parts.append(f"{r.from_place} to {r.to_place}".strip())
    if r.item_name:
        parts.append(f"Goods {r.item_name}")
    return " / ".join(parts)


def transporter_vouchers(r: "ProcessedRow", names: Dict[str, str], ctag: str = "") -> List[tuple]:
    """(voucher type, xml, remote id, amount) for one row, in the same order they appear in
    the Tally Day Book when entered by hand:

      Payment  TRANSPORTING EXPENSES Dr / Cash Cr     - cash from office
      Payment  TRANSPORTING EXPENSES Dr / Cash Cr     - Other1 / Other2 advances
               (own ledgers can be set instead in Settings)
      Journal  TDS Dr / Party Cr                       - TDS deducted by party
      Journal  DIESEL EXPENSES Dr (+ OTHER EXPENSES Dr - cash from pump)
               / Petrol pump Cr                        - diesel slip
      Journal  TRANSPORTING EXPENSES Dr / Truck owner Cr - net hire payable
      Sales    Party Dr / Freight Cr                   - freight bill
    Zero amounts produce no voucher.
    """
    tag = _row_tag(r)
    out: List[tuple] = []

    kinds = enabled_kinds(names)

    bilty_no = r.voucher_no + (f"-{r.occurrence}" if r.occurrence > 1 else "")

    def add(vtype, kind, amount, narration, entries, **kw):
        if kind not in kinds:
            return
        use = names.get("vchno_sales" if kind == "SAL" else "vchno_others", "no") == "yes"
        if use:
            kw["number"] = bilty_no
        rid = voucher_remote_id(ctag, r, kind) if ctag else None
        out.append((vtype, _voucher(vtype, r, narration, entries, remote_id=rid, **kw), rid, amount))

    if r.cash_office:
        add("Payment", "PAY", r.cash_office, f"Cash paid from office to {r.owner} - {tag}",
            _dr(names.get("office_exp") or names["transport_exp"], r.cash_office) + _cr(names["cash"], r.cash_office))

    for key, paid_from, amt in r.other_advances:
        exp = names.get(f"{key}_exp") or names["transport_exp"]
        label = "Other1" if key == "other1" else "Other2"
        add("Payment", "OT1" if key == "other1" else "OT2", amt,
            f"{label} advance paid to {r.owner} - {tag}",
            _dr(exp, amt) + _cr(paid_from, amt))

    if r.tds:
        add("Journal", "TDS", r.tds, f"TDS deducted by {r.party} on freight - {tag}",
            _dr(names["tds"], r.tds) + _cr(r.party, r.tds))

    if r.diesel or r.cash_pump:
        pump = r.pump or names["cash"]
        entries = ""
        if r.diesel:
            entries += _dr(names["diesel_exp"], r.diesel)
        if r.cash_pump:
            entries += _dr(names["other_exp"], r.cash_pump)
        entries += _cr(pump, round(r.diesel + r.cash_pump, 2))
        slip = f" Slip {r.slip_no}" if r.slip_no else ""
        detail = f"Diesel {r.diesel:.2f}" + (f" + cash from pump {r.cash_pump:.2f}" if r.cash_pump else "")
        add("Journal", "DSL", round(r.diesel + r.cash_pump, 2),
            f"{detail} from {pump}{slip} - {tag}", entries)

    if r.transport_amount or r.owner_payable or r.other_deductions:
        # Signed lines (+ = debit). Normally Transporting Exp. Dr / Owner Cr;
        # when advances exceed the hire the signs flip and the owner is debited.
        lines = [(names["transport_exp"], r.transport_amount), (r.owner, -r.owner_payable)]
        lines += [(ledger, -amt) for ledger, amt in r.other_deductions]
        lines = [(l, round(a, 2)) for l, a in lines if round(a, 2)]
        lines.sort(key=lambda la: la[1] < 0)          # debits first, like Tally shows
        entries = "".join(_dr(l, a) if a > 0 else _cr(l, -a) for l, a in lines)
        total = round(sum(a for _, a in lines if a > 0), 2)
        if r.owner_payable >= 0:
            narration = (f"Truck hire payable to {r.owner} (freight {r.amount:.2f} less commission "
                         f"{r.commission:.2f}, diesel, cash advances) - {tag}")
        else:
            narration = (f"Advances to {r.owner} exceed truck hire (freight {r.amount:.2f} less "
                         f"commission {r.commission:.2f}) - {-r.owner_payable:.2f} recoverable - {tag}")
        add("Journal", "TRN", total, narration, entries)

    if names.get("sales_mode", "invoice") == "invoice":
        add("Sales", "SAL", r.amount, f"Freight bill - {tag}",
            _invoice_entries(r, names), party=r.party, reference=r.voucher_no, invoice=True)
    else:
        add("Sales", "SAL", r.amount, f"Freight bill - {tag}",
            _dr(r.party, r.amount) + _cr(names["freight"], r.amount),
            party=r.party, reference=r.voucher_no)
    return out


def _transporter_masters(rows: List["ProcessedRow"], names: Dict[str, str]) -> List[str]:
    """Ledger masters for exactly the ledgers the built vouchers use."""
    used = []
    for r in rows:
        for _, xml, _, _ in transporter_vouchers(r, names):
            for m in re.finditer(r"<LEDGERNAME>(.*?)</LEDGERNAME>", xml):
                name = html.unescape(m.group(1))
                if name not in used:
                    used.append(name)
    groups = {
        names["freight"]: "Sales Accounts",
        names["tds"]: "Loans & Advances (Asset)",
        names["diesel_exp"]: "Direct Expenses",
        names["other_exp"]: "Indirect Expenses",
        **{names[k]: "Indirect Expenses" for k in ("office_exp", "other1_exp", "other2_exp") if names.get(k)},
        names["transport_exp"]: "Direct Expenses",
        names["discount"]: "Indirect Incomes",
        names["other1"]: "Indirect Incomes",
        names["other2"]: "Indirect Incomes",
    }
    for r in rows:
        groups.setdefault(r.party, "Sundry Debtors")
    for r in rows:
        if r.pump:
            groups.setdefault(r.pump, "Sundry Creditors")
        groups.setdefault(r.owner, "Sundry Creditors")
    xml = []
    invoice = names.get("sales_mode", "invoice") == "invoice" and "SAL" in enabled_kinds(names)
    if invoice and rows:
        unit = names.get("item_unit") or "MT"
        xml.append(_unit_master(unit))
        for item in sorted({sales_item(r, names)[0] for r in rows}):
            xml.append(_stock_item_master(item, unit))
    fixed_order = list(dict.fromkeys(groups))
    for name in sorted(used, key=lambda n: (fixed_order.index(n) if n in fixed_order else 999, n)):
        if name == names["cash"]:
            continue   # "Cash" exists in every Tally company already
        master = _ledger_master(name, groups.get(name, "Sundry Creditors"))
        if invoice and name == names["freight"]:
            # A sales ledger used in an item invoice must affect inventory.
            master = master.replace("<OPENINGBALANCE>", "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>\n        <OPENINGBALANCE>")
        xml.append(master)
    return xml


def build_transporter_xml(
    rows: List[ProcessedRow],
    tally_company_name: str,
    ledger_names: Optional[Dict[str, str]] = None,
) -> str:
    """One file: ledger masters, then every row's Payment/Journal/Sales
    vouchers in Day Book order."""
    names = {**DEFAULT_LEDGER_NAMES, **(ledger_names or {})}
    usable = [r for r in rows if not r.problem]
    ledger_xml = _transporter_masters(usable, names)
    ctag = company_tag(tally_company_name)
    voucher_xml = [v[1] for r in usable for v in transporter_vouchers(r, names, ctag)]
    return (
        _envelope("All Masters", ledger_xml, tally_company_name)
        + _envelope("Vouchers", voucher_xml, tally_company_name)
    )


def count_transporter_vouchers(rows: List[ProcessedRow], ledger_names: Optional[Dict[str, str]] = None) -> Dict[str, int]:
    names = {**DEFAULT_LEDGER_NAMES, **(ledger_names or {})}
    counts: Dict[str, int] = {"Payment": 0, "Journal": 0, "Sales": 0}
    for r in rows:
        if r.problem:
            continue
        for v in transporter_vouchers(r, names):
            counts[v[0]] += 1
    return counts


def transporter_sent_list(rows: List[ProcessedRow], tally_company_name: str,
                          ledger_names: Optional[Dict[str, str]] = None) -> List[dict]:
    """What build_transporter_xml will post, one dict per voucher - saved
    in the database after a direct send so it can be replaced or undone."""
    names = {**DEFAULT_LEDGER_NAMES, **(ledger_names or {})}
    ctag = company_tag(tally_company_name)
    out = []
    for r in rows:
        if r.problem:
            continue
        for vtype, _, rid, amount in transporter_vouchers(r, names, ctag):
            out.append({"remote_id": rid, "vch_type": vtype, "vch_date": r.date.strftime("%Y%m%d"),
                        "bilty": r.voucher_no, "amount": f"{amount:.2f}"})
    return out


def bilty_keys(rows: List[ProcessedRow], tally_company_name: str) -> List[str]:
    """REMOTEID prefix per usable row (everything except the -KIND part)."""
    ctag = company_tag(tally_company_name)
    return [voucher_remote_id(ctag, r, "X")[:-2] for r in rows if not r.problem]


def transporter_preview(rows: List[ProcessedRow], ledger_names: Optional[Dict[str, str]] = None) -> List[dict]:
    """Day Book-style list of every voucher that will be posted - one line
    per voucher with its debit and credit ledgers - for the preview screen."""
    names = {**DEFAULT_LEDGER_NAMES, **(ledger_names or {})}
    out = []
    for r in rows:
        if r.problem:
            continue
        for vtype, xml, _, amount in transporter_vouchers(r, names):
            dr, cr = [], []
            for m in re.finditer(
                r"<LEDGERNAME>(.*?)</LEDGERNAME>\s*<ISDEEMEDPOSITIVE>(\w+)</ISDEEMEDPOSITIVE>\s*<AMOUNT>(-?[\d.]+)</AMOUNT>",
                xml,
            ):
                text = f"{html.unescape(m.group(1))} {abs(float(m.group(3))):,.2f}"
                (dr if m.group(2) == "Yes" else cr).append(text)
            narration = re.search(r"<NARRATION>(.*?)</NARRATION>", xml, flags=re.S)
            inv = re.search(r"<STOCKITEMNAME>(.*?)</STOCKITEMNAME>.*?<RATE>([\d.]+)/.*?<BILLEDQTY>\s*([\d.]+)\s*(.*?)</BILLEDQTY>", xml, flags=re.S)
            item_text = (f"{html.unescape(inv.group(1))}: {inv.group(3)} {html.unescape(inv.group(4))} "
                         f"@ {float(inv.group(2)):,.2f}") if inv else ""
            out.append({
                "File / Sheet": r.source,
                "Excel row": r.source_row,
                "Date": r.date.strftime("%d-%m-%Y"),
                "Bilty": r.voucher_no,
                "Voucher": vtype,
                "Vch No.": (re.search(r"<VOUCHERNUMBER>(.*?)</VOUCHERNUMBER>", xml) or [None, "(auto)"])[1],
                "Debit (Dr)": " + ".join(dr),
                "Credit (Cr)": " + ".join(cr),
                "Amount": amount,
                "Item (Qty @ Rate)": item_text,
                "Narration": html.unescape(narration.group(1)) if narration else "",
            })
    return out
