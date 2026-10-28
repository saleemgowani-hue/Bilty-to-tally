"""Direct connection to a running TallyPrime / Tally.ERP 9 on this computer
(or the office network), so vouchers go straight into Tally without the
user importing an XML file by hand.

Tally has a built-in XML server. Turn it on once in TallyPrime:
    F1 (Help) -> Settings -> Connectivity -> Client/Server configuration
        TallyPrime acts as : Both   (or Server)
        Port              : 9000
then restart Tally and keep the company open.

The same XML this app writes to a file is POSTed to http://host:port and
Tally's reply (created / altered / errors) is read back and shown.
Only the Python standard library is used - nothing extra to install.
"""
from __future__ import annotations

import re
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from html import escape, unescape
from typing import List, Optional, Set

DEFAULT_HOST = "localhost"
DEFAULT_PORT = 9000
TIMEOUT_SECONDS = 120

NOT_RUNNING_HELP = (
    "Could not reach Tally at {url}. Check that: (1) Tally is open with your "
    "company loaded; (2) in TallyPrime F1 Help -> Settings -> Connectivity, "
    "'TallyPrime acts as' is 'Both' or 'Server' and Port is {port}; "
    "(3) Tally was restarted after changing that setting."
)


class TallyConnectionError(Exception):
    """Tally could not be reached, with a message safe to show the user."""


@dataclass
class PushResult:
    created: int = 0
    altered: int = 0
    errors: int = 0
    ignored: int = 0
    deleted: int = 0
    ledgers_created: int = 0
    raw: str = ""                   # Tally's own replies, for troubleshooting
    messages: List[str] = field(default_factory=list)

    def add(self, other: "PushResult") -> None:
        self.created += other.created
        self.altered += other.altered
        self.errors += other.errors
        self.ignored += other.ignored
        self.deleted += other.deleted
        self.messages.extend(other.messages)
        self.raw += other.raw


def _url(host: str, port: int) -> str:
    return f"http://{(host or DEFAULT_HOST).strip()}:{int(port or DEFAULT_PORT)}"


def _post(host: str, port: int, body: str) -> str:
    url = _url(host, port)
    req = urllib.request.Request(
        url,
        data=body.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT_SECONDS) as resp:
            raw = resp.read()
    except (urllib.error.URLError, ConnectionError, TimeoutError, OSError) as exc:
        raise TallyConnectionError(NOT_RUNNING_HELP.format(url=url, port=port)) from exc
    # Tally sometimes answers in UTF-16; detect by BOM / NUL bytes.
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in raw[:20]:
        return raw.decode("utf-16", errors="replace")
    return raw.decode("utf-8", errors="replace")


def _strip_invalid_xml_chars(text: str) -> str:
    # Tally can emit control characters (e.g. &#4;) that ElementTree rejects.
    text = re.sub(r"&#(?:x0*[0-8bcef]|x0*1[0-9a-f]|0*[0-8]|0*1[1-9]|0*2[0-9]|0*3[01]);", "", text, flags=re.I)
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)


def _tag_int(text: str, tag: str) -> int:
    total = 0
    for m in re.finditer(rf"<{tag}>\s*(-?\d+)\s*</{tag}>", text):
        total += int(m.group(1))
    return total


def parse_import_response(text: str) -> PushResult:
    """Read Tally's reply to an Import Data request. Works for both the
    older <RESPONSE> form and TallyPrime's <IMPORTRESULT> form."""
    result = PushResult(
        raw=text.strip()[:4000] + "\n",
        created=_tag_int(text, "CREATED"),
        altered=_tag_int(text, "ALTERED"),
        errors=_tag_int(text, "ERRORS"),
        ignored=_tag_int(text, "IGNORED"),
        deleted=_tag_int(text, "DELETED"),
    )
    for m in re.finditer(r"<LINEERROR>(.*?)</LINEERROR>", text, flags=re.S):
        msg = unescape(m.group(1)).strip()
        if msg:
            result.messages.append(msg)
    if not result.messages and result.errors == 0 and "<RESPONSE>" in text and result.created == 0:
        # e.g. "<RESPONSE>Unknown Request, cannot be processed</RESPONSE>"
        m = re.search(r"<RESPONSE>(.*?)</RESPONSE>", text, flags=re.S)
        if m and m.group(1).strip() and "<" not in m.group(1):
            result.messages.append(unescape(m.group(1)).strip())
    return result


def _collection_request(name: str, obj_type: str, company: Optional[str]) -> str:
    company_xml = (
        f"<SVCURRENTCOMPANY>{escape(company, quote=False)}</SVCURRENTCOMPANY>" if company else ""
    )
    return f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>{name}</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        {company_xml}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="{name}" ISMODIFY="No">
            <TYPE>{obj_type}</TYPE>
            <FETCH>NAME</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""


def _names_from_collection(text: str, obj_tag: str) -> List[str]:
    names = []
    for m in re.finditer(rf'<{obj_tag}\s+NAME="([^"]*)"', text):
        names.append(unescape(m.group(1)).strip())
    if not names:
        for m in re.finditer(rf"<{obj_tag}\b[^>]*>.*?<NAME[^>]*>(.*?)</NAME>", text, flags=re.S):
            names.append(unescape(m.group(1)).strip())
    return [n for n in names if n]


def open_companies(host: str, port: int) -> List[str]:
    """Companies currently loaded in Tally. Raises TallyConnectionError if
    Tally is not reachable."""
    text = _post(host, port, _collection_request("SNOpenCompanies", "Company", None))
    return _names_from_collection(text, "COMPANY")


def existing_masters(host: str, port: int, company: str) -> dict:
    """{"LEDGER": {...}, "STOCKITEM": {...}, "UNIT": {...}} upper-cased names."""
    out = {"LEDGER": existing_ledgers(host, port, company)}
    for tag, tally_type in (("STOCKITEM", "StockItem"), ("UNIT", "Unit")):
        text = _post(host, port, _collection_request(f"SN{tally_type}Names", tally_type, company))
        out[tag] = {n.upper() for n in _names_from_collection(text, tag)}
    return out


def existing_ledgers(host: str, port: int, company: str) -> Set[str]:
    text = _post(host, port, _collection_request("SNLedgerNames", "Ledger", company))
    return {n.upper() for n in _names_from_collection(text, "LEDGER")}


def check_connection(host: str, port: int, company: str) -> tuple:
    """(ok, message) for the Settings page's Test button."""
    try:
        companies = open_companies(host, port)
    except TallyConnectionError as exc:
        return False, str(exc)
    if not companies:
        return False, "Tally is running but no company is open. Open your company in Tally."
    if company and company.strip().upper() not in {c.upper() for c in companies}:
        return False, (
            f"Tally is running, but company '{company}' is not open. "
            f"Open companies: {', '.join(companies)}. The name in Settings must match exactly."
        )
    return True, f"Connected. Open companies: {', '.join(companies)}"


def _split_envelopes(xml: str) -> List[str]:
    parts = [p.strip() for p in xml.split("</ENVELOPE>") if p.strip()]
    return [p + "\n</ENVELOPE>" for p in parts]


def _drop_existing_ledgers(envelope: str, existing) -> tuple:
    """Remove masters (ledgers, stock items, units) that Tally already has,
    so re-sending never shows 'already exists' errors. `existing` is a set
    of ledger names or a {tag: set} dict. Returns (envelope, how many left)."""
    if not isinstance(existing, dict):
        existing = {"LEDGER": existing}
    envelope = envelope.replace('xmlns:UDF="TallyUDF"', "")   # ElementTree needs no prefix decl here
    root = ET.fromstring(_strip_invalid_xml_chars(envelope))
    left = 0
    for parent in list(root.iter()):
        for tag in ("UNIT", "STOCKITEM", "LEDGER"):
            for master in list(parent.findall(tag)):
                if (master.get("NAME") or "").strip().upper() in existing.get(tag, set()):
                    parent.remove(master)
                else:
                    left += 1
    return ET.tostring(root, encoding="unicode"), left


def push_xml(xml: str, host: str, port: int, company: str) -> PushResult:
    """Send one generated XML (masters envelope + vouchers envelope) to Tally.

    Masters first - minus ledgers Tally already has - then the vouchers.
    Raises TallyConnectionError when Tally is not reachable or the company
    is not open, before anything is posted."""
    ok, message = check_connection(host, port, company)
    if not ok:
        raise TallyConnectionError(message)

    total = PushResult()
    for envelope in _split_envelopes(xml):
        if "<REPORTNAME>All Masters</REPORTNAME>" in envelope:
            existing = existing_masters(host, port, company)
            envelope, left = _drop_existing_ledgers(envelope, existing)
            if left == 0:
                continue
            masters = parse_import_response(_post(host, port, envelope))
            # Masters are reported separately so voucher counts stay clear.
            masters.messages = [f"Ledger: {m}" for m in masters.messages]
            total.messages.extend(masters.messages)
            total.errors += masters.errors
            total.ledgers_created += masters.created
            total.raw += "[ledgers] " + masters.raw
            continue
        total.add(parse_import_response(_post(host, port, envelope)))
    return total


def _delete_envelope(vouchers: List[dict], company: str) -> str:
    items = []
    for v in vouchers:
        items.append(
            f"""
          <VOUCHER REMOTEID="{escape(v['remote_id'], quote=True)}" VCHTYPE="{escape(v['vch_type'], quote=True)}" ACTION="Delete">
            <DATE>{v['vch_date']}</DATE>
            <VOUCHERTYPENAME>{escape(v['vch_type'], quote=False)}</VOUCHERTYPENAME>
          </VOUCHER>"""
        )
    return f"""<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES><SVCURRENTCOMPANY>{escape(company or '', quote=False)}</SVCURRENTCOMPANY></STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">{''.join(items)}
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>
"""


def delete_vouchers(vouchers: List[dict], host: str, port: int, company: str) -> PushResult:
    """Delete vouchers this app created earlier, found by their REMOTEID.
    `vouchers` items need remote_id, vch_type and vch_date (YYYYMMDD)."""
    if not vouchers:
        return PushResult()
    ok, message = check_connection(host, port, company)
    if not ok:
        raise TallyConnectionError(message)
    return parse_import_response(_post(host, port, _delete_envelope(vouchers, company)))
