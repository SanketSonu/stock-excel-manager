from __future__ import annotations

import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta

try:
    from .config import Settings  # package import (e.g. `python -m stock_excel_manager.cli`)
except ImportError:
    from config import Settings   # flat import (Streamlit runs the script directly)


@dataclass(frozen=True)
class DailyClose:
    symbol: str
    trading_date: date
    close: float
    open_price: float

    @property
    def change_pct(self) -> float:
        """Calculates intra-day change percentage (Close vs Open)."""
        if self.open_price == 0:
            return 0.0
        return ((self.close - self.open_price) / self.open_price) * 100.0


class KiteService:
    def __init__(self, settings: Settings) -> None:
        try:
            from kiteconnect import KiteConnect
        except ImportError as exc:
            raise RuntimeError("kiteconnect is required. Install stock_excel_manager/requirements.txt") from exc

        self.settings = settings
        self.client = KiteConnect(api_key=settings.kite_api_key)
        if settings.kite_access_token:
            self.client.set_access_token(settings.kite_access_token)
        self._instrument_token_by_symbol: dict[str, int] | None = None

    def login_url(self) -> str:
        return self.client.login_url()

    def generate_access_token(self, request_token: str) -> str:
        if not self.settings.kite_api_secret:
            raise RuntimeError("KITE_API_SECRET is required to generate an access token")
        data = self.client.generate_session(request_token, api_secret=self.settings.kite_api_secret)
        return str(data["access_token"])

    def instruments(self) -> list[dict]:
        return self.client.instruments("NSE")

    def daily_close(self, symbol: str, target_date: date) -> DailyClose | None:
        token = self.instrument_token(symbol)
        from_date = target_date - timedelta(days=14)
        rows = self.client.historical_data(
            instrument_token=token,
            from_date=datetime.combine(from_date, datetime.min.time()),
            to_date=datetime.combine(target_date, datetime.max.time()),
            interval="day",
        )
        time.sleep(self.settings.request_delay_seconds)
        candles = [
            row for row in rows
            if _as_date(row["date"]) <= target_date and float(row["close"]) > 0
        ]
        if len(candles) < 2:
            return None
        latest = candles[-1]
        if _as_date(latest["date"]) != target_date:
            return None
        return DailyClose(
            symbol=symbol,
            trading_date=_as_date(latest["date"]),
            close=float(latest["close"]),
            open_price=float(latest["open"]),
        )

    def instrument_token(self, symbol: str) -> int:
        if self._instrument_token_by_symbol is None:
            print("Loading instrument cache...")
            # Fetch both NSE and BSE to be safe
            nse_instruments = self.client.instruments("NSE")
            bse_instruments = self.client.instruments("BSE")
            
            self._instrument_token_by_symbol = {
                str(item["tradingsymbol"]).upper(): int(item["instrument_token"])
                for item in (nse_instruments + bse_instruments)
                if item.get("tradingsymbol") and item.get("instrument_token")
            }
            
        symbol_upper = symbol.upper()
        if symbol_upper in self._instrument_token_by_symbol:
            return self._instrument_token_by_symbol[symbol_upper]
        
        # Try common variations if direct lookup fails
        variations = [f"{symbol_upper}-EQ", f"{symbol_upper}-BE"]
        for v in variations:
            if v in self._instrument_token_by_symbol:
                return self._instrument_token_by_symbol[v]
                
        raise LookupError(f"Instrument token not found for {symbol}")


def _as_date(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.fromisoformat(str(value)).date()
