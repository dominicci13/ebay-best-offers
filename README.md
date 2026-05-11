# ebay-pending-offers

Processes pending eBay Best Offers across all seller accounts: evaluates each offer against configurable margin thresholds, accepts or declines automatically, syncs results to SQL Server, updates an Excel workbook, and emails a daily summary. Runs on a daily schedule via APScheduler.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
pip install git+https://github.com/dominicci13/shared-python-utils.git
```

### 2. Configure environment

```bash
cp .env.example .env
```

Edit `.env` with your credentials, SQL table names, and offer thresholds.

## Run

```bash
python run_ebay_pending_offers.py
```

The script runs automatically at 17:30 daily via APScheduler.

## Environment Variables

| Variable | Description |
|---|---|
| `eBay_pass` | eBay account password |
| `CHROME_USER_DATA_DIR` | Path to Chrome automation profile directory |
| `SENDER_EMAIL` | Outlook account used to send the report email |
| `TO_EMAIL` | Comma-separated list of recipient email addresses |
| `CC_EMAIL` | Comma-separated list of CC email addresses |
| `DB_TABLE_PENDING` | SQL table for pending offers (default: `PendingOffers`) |
| `DB_TABLE_AGED` | SQL table for aged inventory (default: `AgedInventory`) |
| `EBAY_COMMISSION` | eBay fee rate as a decimal (default: `0.091`) |
| `MIN_DISCOUNT` | Maximum price as fraction of list price to accept (default: `0.95`) |
| `MAX_DISCOUNT` | Floor price as fraction of list price (default: `0.9`) |
| `MIN_PROFIT` | Minimum profit margin required to accept an offer (default: `0.1`) |
| `BRANDS` | Comma-separated list of brand names to process |
