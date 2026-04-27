"""
Morgan Stanley PDF Statement Parser.

Parses the annual PDF statement downloaded from StockPlanConnect.

PDF Structure (for Alphabet / Google employees):
  Page 1:      Cover, Price History, Account Summary, GSU grant table, start of GSU Activity
  Pages 2-4:   Continuation of GSU Activity table (tracks unvested RSU units, NOT owned shares)
  Pages 5-15:  Per-release receipts ("Share Units - Release (RBxxxxxx)") — tax detail per event
  Page 16:     ESPP / Long Share Savings Plan summary
  Page 17:     ESPP Activity table — THIS is the primary transaction source:
                 Release (RBx)       → DEPOSIT (net shares issued after tax withholding)
                 Sale                → SELL
                 Dividend (Cash)     → DIVIDEND
                 IRS Withholding     → TAX
               + Withdrawal block(s) at the bottom of the page
  Page 18+:    More withdrawal blocks

Key insight:
  The employee's ACTUAL owned shares are held in the GOOG savings plan account,
  NOT in the GOOGL GSU tracking system. Vested GSU (GOOGL) units are automatically
  converted to GOOG and deposited to the savings plan after tax withholding.
  The ESPP Activity table (page 17) reflects the true economic transactions:
    - Release entries = NET shares deposited (gross - tax-withheld shares)
    - Sale entries    = user-initiated sales of GOOG shares
"""

import re
import logging
from decimal import Decimal, InvalidOperation
from datetime import date, datetime
from typing import Optional, List

import pdfplumber
from pydantic import TypeAdapter

from espp2.datamodels import Transactions, EntryTypeEnum, Entry

logger = logging.getLogger(__name__)

_entry_adapter = TypeAdapter(Entry)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _clean(s) -> str:
    """Strip zero-width spaces, newlines, and surrounding whitespace."""
    if s is None:
        return ""
    return str(s).replace("\u200b", "").replace("\n", " ").replace("\r", "").strip()


def _parse_date(s: str) -> Optional[date]:
    """Parse date strings: '29-Apr-2025', 'April 29, 2025', '2025-01-25'."""
    s = _clean(s)
    if not s:
        return None
    for fmt in ("%d-%b-%Y", "%B %d, %Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _parse_decimal(s) -> Optional[Decimal]:
    """Parse a numeric string, stripping $, commas, spaces."""
    s = _clean(s)
    if not s or s in ("-", "—", "N/A", ""):
        return None
    cleaned = s.replace("$", "").replace(",", "").replace(" ", "").replace("USD", "").replace("NOK", "")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def _make_amount(currency: str, value: Decimal, amountdate: date) -> dict:
    return {"currency": currency, "value": value, "amountdate": amountdate}


def _make_entry(d: dict) -> Entry:
    return _entry_adapter.validate_python(d)


def _fmv(symbol: str, entry_date: date) -> Decimal:
    try:
        from espp2.fmv import FMV
        return FMV()[symbol, entry_date]
    except Exception as e:
        logger.warning("FMV lookup failed for %s on %s: %s", symbol, entry_date, e)
        return Decimal("0")


# ---------------------------------------------------------------------------
# Page 1: Metadata extraction
# ---------------------------------------------------------------------------

def _extract_metadata(page_text: str) -> dict:
    """Extract account number, symbols, and period from page 1 text."""
    result = {}

    # Account number: "Account Number: MS-627432-82"
    m = re.search(r"Account Number[:\s]+([A-Z0-9\-]+)", page_text)
    if m:
        result["account_id"] = m.group(1).strip()

    # Summary period: "01-Jan-2025 to 31-Dec-2025"
    m = re.search(r"(\d{2}-[A-Za-z]{3}-\d{4})\s+to\s+(\d{2}-[A-Za-z]{3}-\d{4})", page_text)
    if m:
        result["fromdate"] = _parse_date(m.group(1))
        result["todate"] = _parse_date(m.group(2))

    # Symbols from Price History: "GOOGL - NASDAQ", "GOOG - NASDAQ"
    symbols = re.findall(r"([A-Z]{2,6})\s*-\s*NASDAQ", page_text)
    result["symbols"] = list(dict.fromkeys(symbols))  # dedupe, preserve order

    return result


# ---------------------------------------------------------------------------
# ESPP Activity table parser (page 17 and similar pages)
# ---------------------------------------------------------------------------

# ESPP Activity table column layout:
#   0: Entry Date
#   1: Activity
#   2: Type of Money
#   3: Cash           (USD amount for cash entries)
#   4: Number of Shares
#   5: Share Price
#   6: Book Value     (total cost basis of shares)
#   7: Market Value   (qty × price, negative for sales)

_ESPP_COL = {
    "date": 0,
    "activity": 1,
    "type_of_money": 2,
    "cash": 3,
    "shares": 4,
    "share_price": 5,
    "book_value": 6,
    "market_value": 7,
}


def _is_espp_header(row: list) -> bool:
    """Detect the ESPP Activity table header row."""
    row_text = " ".join(_clean(c).lower() for c in row if c)
    return "entry date" in row_text and "activity" in row_text and "number of shares" in row_text


def _parse_espp_activity(table: list, symbol: str, entries: list, source: str = "morgan_pdf"):
    """
    Parse the ESPP Activity table and generate transaction entries.

    Activities handled:
      Release (RBxxxxxx)     → DEPOSIT (net shares issued after tax withholding)
      Sale                   → SELL (user-initiated sale)
      Dividend (Cash)        → DIVIDEND (cash dividend)
      IRS Nonresident...     → TAX (withholding on dividend)
      Closing Value          → skip (year-end summary)
      Cash Transfer Out      → skip (internal cash movement after dividend)
    """
    header_found = False

    for row in table:
        if not row:
            continue
        row_clean = [_clean(c) for c in row]

        # Skip header rows
        if _is_espp_header(row):
            header_found = True
            continue
        if not header_found:
            continue

        # Pad row to expected columns
        while len(row_clean) < 8:
            row_clean.append("")

        raw_date = row_clean[_ESPP_COL["date"]]
        raw_activity = row_clean[_ESPP_COL["activity"]]

        if not raw_date or not raw_activity:
            continue

        entry_date = _parse_date(raw_date)
        if entry_date is None:
            continue

        activity = raw_activity.strip().lower()

        # ── Release (RBxxxxxx) → DEPOSIT ───────────────────────────────────
        if activity.startswith("release"):
            qty = _parse_decimal(row_clean[_ESPP_COL["shares"]])
            if qty is None or qty <= 0:
                continue
            book_value = _parse_decimal(row_clean[_ESPP_COL["book_value"]])
            if book_value and book_value > 0 and qty > 0:
                price_per_share = book_value / qty
            else:
                price_per_share = _fmv(symbol, entry_date)

            entry = _make_entry({
                "type": EntryTypeEnum.DEPOSIT,
                "date": entry_date,
                "qty": qty,
                "symbol": symbol,
                "purchase_price": _make_amount("USD", price_per_share, entry_date),
                "description": raw_activity,
                "source": source,
            })
            entries.append(entry)
            logger.debug("DEPOSIT %s %s qty=%s price=%s", symbol, entry_date, qty, price_per_share)

        # ── Sale → SELL ─────────────────────────────────────────────────────
        elif activity == "sale":
            qty_raw = _parse_decimal(row_clean[_ESPP_COL["shares"]])
            if qty_raw is None:
                continue
            qty = abs(qty_raw)
            share_price = _parse_decimal(row_clean[_ESPP_COL["share_price"]])
            market_value = _parse_decimal(row_clean[_ESPP_COL["market_value"]])

            if market_value is not None:
                amount = abs(market_value)
            elif share_price is not None:
                amount = qty * share_price
            else:
                amount = Decimal("0")

            if share_price is None:
                share_price = Decimal("0")

            entry = _make_entry({
                "type": EntryTypeEnum.SELL,
                "date": entry_date,
                "qty": -qty,   # negative = sale
                "symbol": symbol,
                "amount": _make_amount("USD", amount, entry_date),
                "fee": _make_amount("USD", Decimal("0"), entry_date),
                "description": "Sale",
                "source": source,
            })
            entries.append(entry)
            logger.debug("SELL %s %s qty=%s amount=%s", symbol, entry_date, qty, amount)

        # ── Dividend (Cash) → DIVIDEND ──────────────────────────────────────
        elif "dividend" in activity and "cash" in activity:
            cash = _parse_decimal(row_clean[_ESPP_COL["cash"]])
            if cash is None or cash <= 0:
                continue
            div_entry: dict = {
                "type": EntryTypeEnum.DIVIDEND,
                "date": entry_date,
                "symbol": symbol,
                "amount": _make_amount("USD", cash, entry_date),
                "description": "Dividend",
                "source": source,
            }
            # Per-share price is optional; only add if FMV lookup succeeds
            price = _fmv(symbol, entry_date)
            try:
                if price and price > 0 and price == price:  # NaN check: NaN != NaN
                    div_entry["amount_ps"] = _make_amount("USD", price, entry_date)
            except Exception:
                pass
            entry = _make_entry(div_entry)
            entries.append(entry)
            logger.debug("DIVIDEND %s amount=%s", entry_date, cash)

        # ── IRS Nonresident Alien Withholding → TAX ─────────────────────────
        elif "irs" in activity or "withholding" in activity or "nonresident" in activity:
            cash = _parse_decimal(row_clean[_ESPP_COL["cash"]])
            if cash is None or cash == 0:
                continue
            entry = _make_entry({
                "type": EntryTypeEnum.TAX,
                "date": entry_date,
                "symbol": symbol,
                "amount": _make_amount("USD", -abs(cash), entry_date),
                "description": "IRS withholding",
                "source": source,
            })
            entries.append(entry)
            logger.debug("TAX %s amount=%s", entry_date, cash)

        # ── Closing Value / Cash Transfer Out → skip ────────────────────────
        elif activity in ("closing value", "cash transfer out", ""):
            continue

        else:
            logger.debug("Unhandled ESPP activity: %r on %s", raw_activity, entry_date)


# ---------------------------------------------------------------------------
# Withdrawal block parser (text-based, pages 17-18)
# ---------------------------------------------------------------------------

_WIRE_PATTERN = re.compile(
    r"Withdrawal\s+on\s+([A-Za-z]+ \d+,? \d{4})"   # "Withdrawal on April 29, 2025"
    r".*?"
    r"Reference Number:\s*([A-Z0-9\-]+)"             # "Reference Number: WRC914B1A53-1EE"
    r".*?"
    r"Net Proceeds[:\s]+\$?([\d,]+\.?\d*)\s*USD",    # "Net Proceeds: $4,945.93 USD"
    re.DOTALL | re.IGNORECASE,
)

_WIRE_DATE_PATTERN = re.compile(
    r"Settlement Date:\s*(\d{2}-[A-Za-z]{3}-\d{4})",
    re.IGNORECASE,
)


def _parse_withdrawal_blocks(full_text: str, entries: list, source: str = "morgan_pdf"):
    """
    Parse withdrawal/wire blocks from page text.
    Creates WIRE entries with USD Net Proceeds (NOK amount to be entered by user).
    Deduplicates by base reference number (stripping suffix like '-1EE').
    """
    seen_refs = set()

    for m in _WIRE_PATTERN.finditer(full_text):
        raw_date_str = m.group(1).strip()
        ref_no = m.group(2).strip()
        net_proceeds_str = m.group(3).strip()

        # Deduplicate by base ref (prefix before first '-' after the main code)
        base_ref = ref_no.split("-")[0]
        if base_ref in seen_refs:
            continue
        seen_refs.add(base_ref)

        # Try settlement date from nearby text for more accuracy
        block = m.group(0)
        sd_match = _WIRE_DATE_PATTERN.search(block)
        if sd_match:
            entry_date = _parse_date(sd_match.group(1))
        else:
            entry_date = _parse_date(raw_date_str)

        if entry_date is None:
            logger.warning("Could not parse wire date: %r", raw_date_str)
            continue

        net_proceeds = _parse_decimal(net_proceeds_str)
        if net_proceeds is None or net_proceeds <= 0:
            continue

        entry = _make_entry({
            "type": EntryTypeEnum.WIRE,
            "date": entry_date,
            "amount": _make_amount("USD", -net_proceeds, entry_date),  # negative = cash out
            "description": f"Wire {ref_no}",
            "source": source,
        })
        entries.append(entry)
        logger.debug("WIRE %s amount=%s ref=%s", entry_date, net_proceeds, ref_no)


# ---------------------------------------------------------------------------
# Main read() function
# ---------------------------------------------------------------------------

def read(fd, filename: str = "", **kwargs) -> Transactions:  # noqa: C901
    """
    Parse a Morgan Stanley annual PDF statement from StockPlanConnect.

    Returns a Transactions object containing:
      - DEPOSIT entries for each net share release to the savings plan
      - SELL entries for each user-initiated sale
      - DIVIDEND entries for cash dividends received
      - TAX entries for IRS withholding on dividends
      - WIRE entries for each USD withdrawal/bank transfer (USD amount only;
        the NOK amount is entered by the user via the web UI)
    """
    import io

    if hasattr(fd, "read"):
        data = fd.read()
        if isinstance(data, str):
            data = data.encode("utf-8")
        stream = io.BytesIO(data)
    else:
        stream = fd

    source = f"morgan_pdf:{filename}" if filename else "morgan_pdf"
    entries: List[Entry] = []
    metadata = {}

    # Accumulate all text across pages (for withdrawal block parsing)
    all_withdrawal_text = ""

    # Flags to detect ESPP activity table pages
    espp_activity_found = False

    with pdfplumber.open(stream) as pdf:
        for page_num, page in enumerate(pdf.pages):
            page_text = _clean(page.extract_text() or "")

            # ── Page 1: extract metadata ───────────────────────────────────
            if page_num == 0:
                metadata = _extract_metadata(page_text)
                logger.info("Account: %s, Symbols: %s, Period: %s → %s",
                            metadata.get("account_id"),
                            metadata.get("symbols"),
                            metadata.get("fromdate"),
                            metadata.get("todate"))

            # ── Detect ESPP activity page ──────────────────────────────────
            lower_text = page_text.lower()
            is_espp_activity = (
                "entry date" in lower_text
                and "number of" in lower_text
                and "shares" in lower_text
                and "activity" in lower_text
            )

            # ── Parse ESPP activity table ──────────────────────────────────
            if is_espp_activity:
                espp_activity_found = True
                # Determine symbol from page text (prefer GOOG for savings plan)
                page_symbols = re.findall(r"([A-Z]{2,6})\s*-\s*NASDAQ", page_text)
                if page_symbols:
                    symbol = page_symbols[0]
                elif metadata.get("symbols"):
                    # Last symbol in Price History is typically GOOG (Class C, savings plan)
                    symbol = metadata["symbols"][-1]
                else:
                    symbol = "GOOG"

                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    # Only process tables that have the ESPP activity header
                    has_header = any(_is_espp_header(row) for row in table[:3])
                    if has_header or (espp_activity_found and _looks_like_espp_data(table)):
                        _parse_espp_activity(table, symbol, entries, source=source)

            # ── Collect withdrawal text ────────────────────────────────────
            if "withdrawal on" in lower_text or "net proceeds" in lower_text:
                all_withdrawal_text += "\n" + page_text

    # ── Parse withdrawal/wire blocks ────────────────────────────────────────
    if all_withdrawal_text:
        _parse_withdrawal_blocks(all_withdrawal_text, entries, source=source)

    if not entries:
        logger.warning("No transactions found in PDF: %s — check PDF format", filename)
        # Return empty but valid Transactions
        today = date.today()
        return Transactions(transactions=[], fromdate=today, todate=today)

    # Sort by date
    entries.sort(key=lambda e: e.date)

    fromdate = metadata.get("fromdate") or entries[0].date
    todate = metadata.get("todate") or entries[-1].date

    logger.info(
        "Parsed %d transactions from '%s' (%s → %s) account=%s",
        len(entries), filename, fromdate, todate, metadata.get("account_id"),
    )

    t = Transactions(transactions=entries, fromdate=fromdate, todate=todate)
    # Store account_id as a dynamic attribute for the web server to pick up
    t.__dict__["account_id"] = metadata.get("account_id", "")
    return t


def _looks_like_espp_data(table: list) -> bool:
    """Heuristic: does this table look like it has ESPP activity data rows?"""
    for row in table[:5]:
        row_text = " ".join(_clean(c).lower() for c in row if c)
        if any(kw in row_text for kw in ("release", "sale", "dividend", "closing value")):
            return True
    return False
