from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


SRC = Path(__file__).resolve().parent
ROOT = SRC.parent


def load_dotenv(path: Path = ROOT / ".env") -> None:
    if not path.exists():
        return
    for raw_line in path.read_text().splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ[key.strip()] = value.strip().strip('"').strip("'")


@dataclass(frozen=True)
class Settings:
    kite_api_key: str
    kite_api_secret: str | None
    kite_access_token: str | None
    workbook_path: Path
    sheet_name: str
    run_time: str
    request_delay_seconds: float
    overrides_path: Path


def load_settings() -> Settings:
    load_dotenv()
    api_key = os.environ.get("KITE_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("KITE_API_KEY is required in environment or .env")
    workbook = Path(os.environ.get("STOCK_EXCEL_PATH", str(SRC / "output" / "Stock Excel Manager.xlsx"))).expanduser()
    return Settings(
        kite_api_key=api_key,
        kite_api_secret=os.environ.get("KITE_API_SECRET") or None,
        kite_access_token=os.environ.get("KITE_ACCESS_TOKEN") or None,
        workbook_path=workbook,
        sheet_name=os.environ.get("STOCK_SHEET_NAME", "Stock Report 2026"),
        run_time=os.environ.get("STOCK_RUN_TIME", "16:00"),
        request_delay_seconds=float(os.environ.get("KITE_REQUEST_DELAY_SECONDS", "0.35")),
        overrides_path=Path(os.environ.get("STOCK_SYMBOL_OVERRIDES", str(SRC / "symbol_overrides.json"))),
    )
