# TaxLedger Pro — Norwegian Tax Reporting for Morgan Stanley ESPP/RSU

> **This is a fork of [ESPP2](https://github.com/otroan/ESPP2) by
> [Ole Trøan](https://github.com/otroan)**, licensed under the MIT License.
> The original project supports multiple brokers (Morgan Stanley, Schwab, and
> others) and is the right choice if you use **Schwab** or need the CLI tools.
> This fork is scoped exclusively to **Morgan Stanley StockPlan Connect** and
> adds a web-based interface, automatic Norges Bank exchange rate fetching, and
> an extended Skattemeldingen Excel report with detailed FX gain/loss
> calculations.
>
> **Using Schwab?** Go to the original app instead:
> [github.com/otroan/ESPP2](https://github.com/otroan/ESPP2)

A local web application that reads your Morgan Stanley annual statement PDF and
produces a completed **Skattemeldingen** Excel file ready to copy into
[skatteetaten.no](https://skatteetaten.no).

Covers everything a Norwegian tax resident needs to report from ESPP / RSU
share programmes held at Morgan Stanley:

- **Formue** — market value of shares on 31 December
- **Skattepliktig gevinst / Fradragsberettiget tap** — realised gain or loss on
  share sales (FIFO, NOK cost basis using Norges Bank rates)
- **Skattepliktig utbytte** — dividend income in NOK
- **Kreditfradrag** — credit for US withholding tax (15 % under the tax treaty)
- **Valutakursgevinst/-tap** — FX gain/loss on USD → NOK conversions, split
  correctly between what is already embedded in the share gain and what must be
  reported separately under *Finans → Valuta*
- **USD-kontobeholdning** — remaining USD cash balance as taxable wealth

---

## Requirements

| Dependency | Minimum version |
|---|---|
| Python | 3.11 |
| pip | any recent version |

No database, no cloud service, no account.  Everything runs on your machine.

---

## Installation

```bash
# 1. Clone or download the repository
git clone https://github.com/eddiehaug/morgan-norwegian-taxes.git
cd morgan-norwegian-taxes

# 2. Create and activate a virtual environment (recommended)
python3 -m venv .venv
source .venv/bin/activate        # macOS / Linux
# .venv\Scripts\activate         # Windows

# 3. Install the package and all dependencies
pip install -e .
```

The first `pip install` downloads roughly 20 MB of dependencies (FastAPI,
pdfplumber, openpyxl, etc.).  Subsequent runs are instant.

---

## Running the app

```bash
taxledger
```

The command starts a local web server at **http://127.0.0.1:8000** and opens
your default browser automatically.  The terminal window must remain open while
you use the app; press `Ctrl-C` to stop it when finished.

> **No data leaves your machine.** The server only listens on `127.0.0.1`
> (localhost) and exchange rates are fetched from the publicly accessible
> Norges Bank API.

---

## Step-by-step usage

### Step 1 — Select the reporting year

The year selector in the top-right of the *Financial Reporting* page defaults
to the previous calendar year.  Change it if needed.  The app immediately
fetches the correct USD/NOK exchange rates from Norges Bank (you will see a
green "✓ N rate days loaded" indicator).

### Step 2 — Upload your Morgan Stanley PDF statements

You need up to two PDFs, both downloadable from the Morgan Stanley StockPlan
Connect portal:

| PDF | Where to find it | Purpose |
|---|---|---|
| **Reporting year statement** (required) | Statements → Annual → select the tax year | Parses all transactions for the year you are reporting |
| **Prior year statement** (optional but recommended) | Statements → Annual → select the year before | Used to compute your opening share balance (cost basis of shares already held at 1 January) |

Drop each file onto the corresponding drop zone, or click to browse.

When the reporting year PDF is uploaded the app will:
- Confirm the year (and update the selector automatically if different)
- Pre-fill the **wire transfer table** with every USD bank transfer found in the
  PDF, including the exact USD amount and date

### Step 3 — Enter NOK amounts for each bank transfer

The wire table shows each USD transfer extracted from the PDF.  For every row,
enter the **NOK amount you actually received in your Norwegian bank account**.
You will find this in your DNB (or other bank) transaction history — look for
incoming international transfers in the same time period as each entry.

The date field is editable in case your bank settled the transfer 1–2 days
after the date shown in the PDF.

Click **+ Add Row** if you have a transfer that does not appear in the PDF
(e.g. a manual transfer or a dividend-only payment).

### Step 4 — Calculate

Click **Calculate Tax Report**.  The app will:

1. Parse all transactions from the reporting year PDF
2. Compute the opening balance from the prior year PDF
3. Match your wire transfers against the transaction history
4. Run FIFO cost-basis calculations in NOK
5. Produce a downloadable Excel report

Progress is shown on screen; the calculation typically takes 5–15 seconds.

### Step 5 — Download the Excel file

Click **Download Skattemeldingen.xlsx** when the results appear.

---

## Understanding the results

### Results screen

The results page summarises every value you need to enter in the
Skattemeldingen, organised by section.

### Fradragsberettiget tap / Skattepliktig gevinst

This is the net result of your share sales for the year in NOK.  It is
calculated as:

```
Sale price (NOK) − Purchase price (NOK) − Skjermingsfradrag
```

Where NOK values use **Norges Bank's mid-rate on the relevant transaction
date**.  This figure also includes the FX result on any sale proceeds that were
wired to Norway within 14 days of the sale (see *Type A* below).

Enter this under **Aksjer og fond → Gevinst/tap ved salg** in Skattemeldingen.

### Kreditfradrag

If you have a valid W-8BEN form on file with Morgan Stanley, US withholding tax
is 15 % of gross dividends.  The app calculates the credit you can claim to
avoid double taxation.  Enter the three values shown under
**Utenlandske inntekter → Kreditfradrag** in Skattemeldingen.

### Valutakursgevinst/-tap — Type A vs Type B

The FX section is split into two distinct types:

#### Type A — Aggregated (already included in the share gain/loss above)

When USD from a share sale is transferred to your Norwegian bank account
**within 14 days** of the sale, the FX result (difference between Norges Bank
rate on the sale date and the actual NOK you received) is treated as part of
the same transaction and is already embedded in the *Fradragsberettiget tap* /
*Skattepliktig gevinst* figure.

**Do not report this separately.**  The xlsx notes this explicitly in amber.

#### Type B — Separate (report under Finans → Valuta)

FX gain or loss on USD that came from dividends, or from sale proceeds held for
more than 14 days before being wired, is reported separately under
**Finans → Valuta**.  In most cases this will be zero or very small.

The xlsx shows a green confirmation when there is nothing to report separately.

#### Per-transfer breakdown

The Excel file contains a detailed table showing every bank transfer with:

| Column | Meaning |
|---|---|
| Dato | Date of the wire transfer |
| Kostpris (NOK) | USD amount × Norges Bank rate on the sale date — your cost basis |
| Mottatt (NOK) | Actual NOK credited to your Norwegian bank account |
| Gevinst/Tap (NOK) | Mottatt − Kostpris |
| Behandling | Type A (in share gain) or Type B (report separately) |

This table is your audit trail if Skatteetaten asks questions.

### USD-kontobeholdning

Any USD remaining in your Morgan Stanley account on 31 December is taxable
wealth.  Report it under **Formue → Bankinnskudd, kontanter og lignende** (or
the foreign asset section, depending on the year's Skattemeldingen layout).

---

## Supported broker

This app supports **Morgan Stanley StockPlan Connect only** — ESPP, RSU, and
GSU (Google Stock Unit) programmes.

It reads the **ESPP Activity** table on page 17 of the Morgan Stanley annual
PDF statement and handles DEPOSIT (ESPP purchase / RSU vest), SELL, DIVIDEND,
TAX (withholding), and WIRE entries.

> **Schwab users:** this app will not work for Schwab statements.  Use the
> original [ESPP2 by Ole Trøan](https://github.com/otroan/ESPP2) instead,
> which has full Schwab support.

---

## Data and privacy

- **PDFs** are processed locally and stored only in a temporary folder that is
  deleted when your operating system cleans temp files.
- **Exchange rates** are fetched once per session from the Norges Bank public
  API (`data.norges-bank.no`).  No other external requests are made.
- **Settings** (reporting year) are saved to `~/.espp2/settings.json`.
- Nothing is sent to any third party.

---

## Troubleshooting

**"No transactions found"**
The app expects the standard Morgan Stanley annual PDF (typically 15–25 pages).
Make sure you are uploading the *Annual Statement*, not a monthly statement or
trade confirmation.

**"Rates not loaded"**
The Norges Bank API requires an internet connection.  If you are behind a
corporate proxy, set `HTTPS_PROXY` in your environment before running
`taxledger`.

**Empty wire table after PDF upload**
Wire transfers only appear if you actually wired USD to a Norwegian bank account
during the year.  If all your proceeds stayed in the Morgan Stanley account, the
table will be empty — add rows manually for any transfers you made.

**Wrong year detected**
The app reads the statement period from the PDF.  If the auto-detected year is
wrong, change it manually in the year selector before calculating.

**Opening balance is zero / cost basis seems wrong**
Make sure you upload the *prior year* statement (e.g. the 2024 statement when
filing for 2025).  Without it the app assumes you held no shares at the start of
the year, which will produce an incorrect cost basis for shares bought in
previous years.

---

## Disclaimer

This tool is provided as-is for personal use.  The calculations follow standard
Norwegian tax rules for foreign shares as of the 2024/2025 tax year.  Always
verify the output against official guidance from Skatteetaten before submitting
your return.  The author is not a tax advisor.

---

## License

MIT — same license as the upstream [ESPP2](https://github.com/otroan/ESPP2)
project by Ole Trøan.  See [LICENSE](LICENSE).
