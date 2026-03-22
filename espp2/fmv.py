#!/usr/bin/env python3

"""
The "Fair Market Value" module downloads and stock prices and exchange rates and
caches them in a set of JSON files.
"""

# pylint: disable=invalid-name,line-too-long

import json
from importlib import resources
from enum import Enum
from datetime import date, datetime, timedelta
from typing import Union, Tuple
import logging
from decimal import Decimal
import math
import urllib3
from pydantic import BaseModel

from .vault import Vault


class Fundamentals(BaseModel):
    """Fundamentals"""

    name: str
    isin: str
    country: str
    symbol: str


class FMVException(Exception):
    """Exception class for FMV module"""


logger = logging.getLogger(__name__)

# Load manually maintained exchange rates / tax deduction rates
with resources.files("espp2").joinpath("data.json").open("r", encoding="utf-8") as f:
    MANUALRATES = json.load(f)


def get_espp_exchange_rate(ratedate):
    """Return the 6 month P&L average. Manually maintained for now."""
    return Decimal(str(MANUALRATES["espp"][ratedate]))


def get_tax_deduction_rate(year):
    """Return tax deduction rate for year"""
    #
    # Remember to add the new tax-free deduction rates for a new year
    #
    if year < 2006:
        logger.error(
            "The tax deduction rate was introduced in 2006, no support for years prior to that. %s",
            year,
        )
        return Decimal("0.0")

    yearstr = str(year)
    if yearstr not in MANUALRATES["tax_deduction_rates"]:
        raise FMVException(f"No tax deduction rate for year {year}")

    return Decimal(str(MANUALRATES["tax_deduction_rates"][yearstr][0]))


class FMVTypeEnum(Enum):
    """Enum for FMV types"""

    STOCK = "STOCK"
    CURRENCY = "CURRENCY"
    DIVIDENDS = "DIVIDENDS"
    FUNDAMENTALS = "FUNDAMENTALS"

    def __str__(self):
        return str(self.value)


# Store downloaded files in cache directory under current directory
# CACHE_DIR = "cache"
# Find data directory from resources
DATA_DIR = resources.files("espp2").joinpath("data")


def todate(datestr: str) -> date:
    """Convert string to datetime"""
    return datetime.strptime(datestr, "%Y-%m-%d").date()


class FMV:
    """Class implementing the Fair Market Value module. Singleton"""

    _instance = None
    _local_rates: dict = {}  # currency → {date_str → Decimal}

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(FMV, cls).__new__(cls)
            # Put any initialization here.
            # if not os.path.exists(CACHE_DIR):
            #     os.makedirs(CACHE_DIR)

            cls.fetchers = {
                FMVTypeEnum.STOCK: cls.fetch_stock2,
                FMVTypeEnum.CURRENCY: cls.fetch_currency,
                FMVTypeEnum.DIVIDENDS: cls.fetch_dividends,
                FMVTypeEnum.FUNDAMENTALS: cls.fetch_fundamentals,
            }
            cls.table = {
                FMVTypeEnum.STOCK: {},
                FMVTypeEnum.CURRENCY: {},
                FMVTypeEnum.DIVIDENDS: {},
                FMVTypeEnum.FUNDAMENTALS: {},
            }
        return cls._instance

    @classmethod
    def load_local_exchange_rates(cls, csv_path: str) -> int:
        """
        Load Norges Bank daily exchange rate CSV.

        Format (semicolon-delimited):
          FREQ;...;BASE_CUR;...;QUOTE_CUR;...;TIME_PERIOD;OBS_VALUE
        Example row:
          B;Business;USD;US dollar;NOK;Norwegian krone;SP;Spot;4;false;0;Units;C;ECB...;2025-01-02;11.3529

        Returns the number of rate entries loaded.
        """
        import csv as _csv
        loaded = 0
        new_rates: dict = {}
        try:
            with open(csv_path, "r", encoding="utf-8-sig") as f:
                reader = _csv.reader(f, delimiter=";")
                headers = None
                for row in reader:
                    if headers is None:
                        headers = [h.strip() for h in row]
                        continue
                    if len(row) < 2:
                        continue
                    # Map header names to indices
                    try:
                        base_idx = headers.index("BASE_CUR")
                        quote_idx = headers.index("QUOTE_CUR")
                        date_idx = headers.index("TIME_PERIOD")
                        val_idx = headers.index("OBS_VALUE")
                    except ValueError:
                        # Fallback: last two columns are date and value
                        date_idx = len(row) - 2
                        val_idx = len(row) - 1
                        base_idx = None
                        quote_idx = None

                    date_str = row[date_idx].strip()
                    val_str = row[val_idx].strip()
                    if not date_str or not val_str:
                        continue

                    base_cur = row[base_idx].strip() if base_idx is not None else "USD"
                    if base_cur not in new_rates:
                        new_rates[base_cur] = {}
                    try:
                        new_rates[base_cur][date_str] = Decimal(val_str)
                        loaded += 1
                    except Exception:
                        continue
        except Exception as e:
            logger.error("Failed to load exchange rates from %s: %s", csv_path, e)
            return 0

        cls._local_rates.update(new_rates)
        logger.info("Loaded %d local exchange rate entries for currencies: %s",
                    loaded, list(new_rates.keys()))
        return loaded

    @classmethod
    def fetch_norges_bank_rates(cls, year: int) -> int:
        """Fetch USD/NOK daily rates from the Norges Bank SDMX-JSON API.

        Fetches from December 1 of the prior year through December 31 of the
        reporting year (covers the Dec 31 opening balance of the prior year plus
        all transactions in the reporting year).

        Returns the number of trading days loaded.
        """
        start = f"{year - 1}-12-01"
        end = f"{year}-12-31"
        url = (
            f"https://data.norges-bank.no/api/data/EXR/B.USD.NOK.SP"
            f"?format=sdmx-json&startPeriod={start}&endPeriod={end}&locale=en"
        )
        logger.info("Fetching Norges Bank rates: %s → %s", start, end)
        try:
            http = urllib3.PoolManager()
            r = http.request("GET", url, timeout=15)
            if r.status != 200:
                raise FMVException(f"Norges Bank API returned HTTP {r.status}")
            data = json.loads(r.data.decode("utf-8"))
            # SDMX-JSON structure (all payload under data["data"]):
            #   data.structure.dimensions.observation[0].values → [{id: "YYYY-MM-DD"}, ...]
            #   data.dataSets[0].series["0:0:0:0"].observations → {"0": [rate], "1": [rate], ...}
            payload = data["data"]
            dates = [
                v["id"]
                for v in payload["structure"]["dimensions"]["observation"][0]["values"]
            ]
            observations = payload["dataSets"][0]["series"]["0:0:0:0"]["observations"]
            rates: dict = {}
            for idx_str, obs in observations.items():
                date_str = dates[int(idx_str)]
                rates[date_str] = Decimal(str(obs[0]))
            cls._local_rates["USD"] = rates
            logger.info("Loaded %d Norges Bank rate entries for USD (%s → %s)", len(rates), start, end)
            return len(rates)
        except Exception as e:
            logger.error("Failed to fetch Norges Bank rates for year %d: %s", year, e)
            raise

    def fetch_stock(self, symbol):
        """Returns a dictionary of date and closing value from AlphaVantage"""
        http = urllib3.PoolManager()
        # The REST api is described here: https://www.alphavantage.co/documentation/
        url = (
            f"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol={symbol}&outputsize=full&"
            "apikey={apikey}"
        )
        r = http.request("GET", url)
        if r.status != 200:
            raise FMVException(f"Fetching stock data for {symbol} failed {r.status}")
        raw = json.loads(r.data.decode("utf-8"))
        return {k: float(v["4. close"]) for k, v in raw["Time Series (Daily)"].items()}

    def fetch_stock2(self, symbol):
        """Returns a dictionary of date and closing value from EOD Historical Data"""
        vault = Vault()
        EODHDKEY = vault["EODHD"]
        url = f"https://eodhd.com/api/eod/{symbol}.US?api_token={EODHDKEY}&fmt=json"
        http = urllib3.PoolManager()
        r = http.request("GET", url)
        if r.status != 200:
            raise FMVException(f"Fetching stock data for {symbol} failed {r.status}")
        raw = json.loads(r.data.decode("utf-8"))
        return {r["date"]: r["close"] for r in raw}

    def fetch_currency(self, currency):
        """Returns a dictionary of date and closing value"""
        http = urllib3.PoolManager()
        # The REST api is described here: https://app.norges-bank.no/query/index.html#/no/
        # url = f'https://data.norges-bank.no/api/data/EXR/B.{currency}.NOK.SP?startPeriod=2000&format=sdmx-json'
        # url = f'https://data.norges-bank.no/api/data/EXR/B.{currency}.NOK.SP?startPeriod=1998&format=csv-:-comma-false-y'
        url = f"https://data.norges-bank.no/api/data/EXR/B.{currency}.NOK.SP?format=csv&startPeriod=1998&locale=us&bom=include"
        r = http.request("GET", url)
        # B;Business;USD;US dollar;NOK;Norwegian krone;SP;Spot;4;false;0;Units;
        # C;ECB concertation time 14:15 CET;2022-05-24;9.5979
        if r.status != 200:
            raise FMVException(
                f"Fetching currency data for {currency} failed {r.status}"
            )
        cur = {}
        for i, line in enumerate(r.data.decode("utf-8").split("\n")):
            if i == 0 or ";" not in line:
                continue  # Skip header and blank lines
            fields = line.strip().split(";")
            d = fields[-2]
            cur[d] = float(fields[-1])
        return cur

    def fetch_dividends(self, symbol):
        """Returns a dividends object keyed on payment date"""
        http = urllib3.PoolManager()
        # url = f'https://eodhistoricaldata.com/api/div/{symbol}.US?fmt=json&from=2000-01-01&api_token={EODHDKEY}'
        vault = Vault()
        EODHDKEY = vault["EODHD"]
        url = f"https://eodhistoricaldata.com/api/div/{symbol}.US?fmt=json&api_token={EODHDKEY}"
        r = http.request("GET", url)
        if r.status != 200:
            raise FMVException(
                f"Fetching dividends data for {symbol} failed {r.status}"
            )
        raw = json.loads(r.data.decode("utf-8"))
        r = {}
        for element in raw:
            d = element["paymentDate"] if element["paymentDate"] else element["date"]
            r[d] = element
        return r

    def fetch_fundamentals(self, symbol):
        """Returns a fundamentals object for symbol"""
        http = urllib3.PoolManager()
        vault = Vault()
        EODHDKEY = vault["EODHD"]
        url = f"https://eodhistoricaldata.com/api/fundamentals/{symbol}.US?api_token={EODHDKEY}"
        r = http.request("GET", url)
        if r.status != 200:
            raise FMVException(
                f"Fetching fundamentals data for {symbol} failed {r.status}"
            )
        raw = json.loads(r.data.decode("utf-8"))
        return raw

    def get_filename(self, fmvtype: FMVTypeEnum, symbol):
        """Get filename for symbol"""
        return f"{DATA_DIR}/{fmvtype}_{symbol}.json"

    def load(self, fmvtype: FMVTypeEnum, symbol):
        """Load data for symbol"""
        filename = self.get_filename(fmvtype, symbol)
        with open(filename, "r", encoding="utf-8") as f:
            self.table[fmvtype][symbol] = json.load(f)

    def need_refresh(self, fmvtype: FMVTypeEnum, symbol, d: datetime.date):
        """Check if we need to refresh data for symbol"""
        if symbol not in self.table[fmvtype]:
            return True
        fetched = self.table[fmvtype][symbol]["fetched"]
        fetched = datetime.strptime(fetched, "%Y-%m-%d").date()
        if d and d > fetched:
            return True
        return False

    def refresh(self, symbol: str, d: datetime.date, fmvtype: FMVTypeEnum):
        """Refresh data for symbol if needed"""
        if not self.need_refresh(fmvtype, symbol, d):
            return

        filename = self.get_filename(fmvtype, symbol)

        # Try loading from cache
        try:
            with open(filename, "r", encoding="utf-8") as f:
                self.table[fmvtype][symbol] = json.load(f)
                if not self.need_refresh(fmvtype, symbol, d):
                    return
        except IOError:
            pass

        data = self.fetchers[fmvtype](self, symbol)

        logging.info("Caching data for %s to %s", symbol, filename)
        data["fetched"] = str(date.today())
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f)

        self.table[fmvtype][symbol] = data

    def extract_date(
        self, input_date: Union[str, datetime, datetime.date]
    ) -> Tuple[datetime.date, str]:
        """Extract date component from input string or datetime object"""
        if isinstance(input_date, str):
            try:
                date_obj = datetime.strptime(input_date, "%Y-%m-%d").date()
            except ValueError as exc:
                raise ValueError(
                    f"Invalid date format '{input_date}'. Use 'YYYY-MM-DD' format."
                ) from exc
        elif isinstance(input_date, datetime):
            date_obj = input_date.date()
        elif isinstance(input_date, date):
            date_obj = input_date
        else:
            raise TypeError(
                f"Input must be string or datetime object, not {type(input_date)}"
            )
        date_str = date_obj.isoformat()
        return date_obj, date_str

    def __getitem__(self, item):
        symbol, itemdate = item
        fmvtype = FMVTypeEnum.STOCK
        itemdate, date_str = self.extract_date(itemdate)
        self.refresh(symbol, itemdate, fmvtype)
        for _ in range(5):
            try:
                return Decimal(str(self.table[fmvtype][symbol][date_str]))
            except KeyError:
                # Might be a holiday, iterate backwards
                itemdate -= timedelta(days=1)
                date_str = str(itemdate)
        return math.nan

    def get_currency(
        self,
        currency: str,
        date_union: Union[str, datetime],
        target_currency: str = "NOK",
    ) -> float:
        """Get currency value. If not found, iterate backwards until found.
        Checks locally uploaded Norges Bank CSV data first."""
        itemdate, date_str = self.extract_date(date_union)
        if currency == "ESPPUSD":
            try:
                return get_espp_exchange_rate(date_str)
            except KeyError:
                # Missing ESPP data for this date
                logger.error("Missing ESPP exchange rate for %s", date_str)
                # Fall-back to USD
                currency = "USD"

        # Check local rates first (from uploaded Norges Bank CSV)
        if currency in FMV._local_rates and FMV._local_rates[currency]:
            check_date = itemdate
            for _ in range(6):
                check_str = str(check_date)
                if check_str in FMV._local_rates[currency]:
                    return FMV._local_rates[currency][check_str]
                check_date -= timedelta(days=1)

        self.refresh(currency, itemdate, FMVTypeEnum.CURRENCY)

        for _ in range(6):
            try:
                return Decimal(
                    str(self.table[FMVTypeEnum.CURRENCY][currency][date_str])
                )
            except KeyError:
                # Might be a holiday, iterate backwards
                itemdate -= timedelta(days=1)
                date_str = str(itemdate)
        raise FMVException(f"No currency data for {currency} on {date_str}")

    def get_dividend(
        self, dividend: str, payment_date: Union[str, datetime]
    ) -> Tuple[date, date, Decimal]:
        """Lookup a dividends record given the paydate."""
        itemdate, date_str = self.extract_date(payment_date)
        self.refresh(dividend, itemdate, FMVTypeEnum.DIVIDENDS)
        for _ in range(5):
            try:
                divinfo = self.table[FMVTypeEnum.DIVIDENDS][dividend][date_str]
                exdate = todate(divinfo["date"])
                declarationdate = (
                    todate(divinfo["declarationDate"])
                    if divinfo["declarationDate"]
                    else exdate
                )
                return exdate, declarationdate, Decimal(str(divinfo["value"]))
            except KeyError:
                # Might be a holiday, iterate backwards
                itemdate -= timedelta(days=1)
                date_str = str(itemdate)
        raise FMVException(f"No dividends data for {dividend} on {date_str}")

    def get_dividends(self, symbol: str) -> dict:
        """Lookup a symbol and return dividends"""
        self.refresh(symbol, None, FMVTypeEnum.DIVIDENDS)

        try:
            return self.table[FMVTypeEnum.DIVIDENDS][symbol]

        except KeyError as e:
            raise FMVException(f"No dividends data for {symbol}") from e

    def get_fundamentals(self, symbol: str) -> dict:
        """Lookup a symbol and return fundamentals"""
        self.refresh(symbol, None, FMVTypeEnum.FUNDAMENTALS)

        try:
            return self.table[FMVTypeEnum.FUNDAMENTALS][symbol]
        except KeyError as e:
            raise FMVException(f"No fundamentals data for {symbol}") from e

    def get_fundamentals2(self, symbol: str) -> dict:
        f = self.get_fundamentals(symbol)
        isin = f.get("General", {}).get("ISIN", None)
        if not isin:
            isin = f.get("ETF_Data", {}).get("ISIN", "")

        return Fundamentals(
            name=f["General"]["Name"],
            isin=isin,
            country=f["General"]["CountryName"],
            symbol=f["General"]["Code"],
        )


if __name__ == "__main__":
    fmv = FMV()
    print("LOOKING UP DATA", fmv[FMVTypeEnum.STOCK, "CSCO", "2021-12-31"])
    # print('LOOKING UP DATA', f['CSCO', '2022-12-31'])
    print("LOOKING UP DATA", fmv[FMVTypeEnum.STOCK, "SLT", "2021-12-31"])
    # f.fetch_currency('USD')
    print("LOOKING UP DATA USD2NOK", fmv.get_currency("USD", "2021-12-31"))

    print("LOOKING UP DIVIDENDS", fmv.get_dividend("CSCO", "2023-01-25"))

    print("CISCO FUNDAMETNALS", fmv.get_fundamentals("CSCO"))
    fundamentals = fmv.get_fundamentals("CSCO")
    print("CISCO FUNDAMETNALS", fundamentals["ISIN"])
