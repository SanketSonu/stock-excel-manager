# Stock Excel Manager

Streamlit apps that update a daily stock report Excel file using
[Zerodha Kite Connect](https://kite.trade/) historical data.

Two flavours:

| File | Purpose |
| --- | --- |
| `src/app_local.py` | Reads & writes a local `.xlsx` (path set via `STOCK_EXCEL_PATH`). |
| `src/app_deploy.py` | Reads & writes an `.xlsx` stored in Google Drive (intended for Streamlit Cloud). |

## Highlights

- One-click **Login with Zerodha** via OAuth — request token is exchanged for an access token automatically (no manual paste each morning).
- Auto-fills missing weekday columns from the last filled date through today.
- Skips holidays / dates with no candle data — those columns aren't created, so the next run retries them.
- Colour-codes positive / negative / flat days.

## Local setup

```bash
git clone https://github.com/SanketSonu/stock-excel-manager.git
cd stock-excel-manager
python -m venv venv && source venv/bin/activate
pip install -r requirements.txt

cp .env.example .env
# edit .env: set KITE_API_KEY, KITE_API_SECRET, STOCK_EXCEL_PATH

streamlit run src/app_local.py
```

In your Kite Connect app on https://developers.kite.trade/apps, set the
**Redirect URL** to `http://localhost:8501/` for the auto-login flow to work.

## Deploy on Streamlit Cloud

1. Push this repo to GitHub.
2. https://share.streamlit.io → **New app** → main file `src/app_deploy.py`.
3. Paste the secrets template below into **Advanced settings → Secrets**.
4. Click Deploy. Once you have the live URL, update the Kite Connect Redirect URL to that URL.

### Secrets template

```toml
[gcp_service_account]
type                        = "service_account"
project_id                  = "..."
private_key_id              = "..."
private_key                 = """-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----
"""
client_email                = "...@...iam.gserviceaccount.com"
client_id                   = "..."
auth_uri                    = "https://accounts.google.com/o/oauth2/auth"
token_uri                   = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url        = "..."

[app]
GDRIVE_FILE_ID  = "..."
KITE_API_KEY    = "..."
KITE_API_SECRET = "..."
SHEET_NAME      = "Stock Report 2026"
OVERRIDES_PATH  = "src/symbol_overrides.json"
```

### Drive setup

The Drive file must be a real `.xlsx`, not a Google Sheet. Share it with the
service account's `client_email` as **Editor**.

## Daily flow

1. Open the Streamlit app.
2. Click **🔐 Login with Zerodha** → enter creds + 2FA.
3. Click **Update Excel now** (local) or **Update & Upload to Drive** (cloud).
