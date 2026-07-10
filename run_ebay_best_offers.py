"""Daily eBay Best Offer automation (SQL-first rebuild).

Runs once a day at 17:30 local time. Each run:

1. Reads the pricing settings (commission, minimum profit floors, counteroffer
   discount band, shipping floor) from the control workbook. If a setting is
   missing or invalid, the run stops and emails the business team what to fix,
   so no offer is ever sent using a wrong number.
2. Scrapes each seller account's pending Best Offers (Seller Hub grid), reads
   every buyer offer in one eBay Trading API call (GetBestOffers), and enriches
   from SQL (site cost, aged status).                               [Step 3]
3. Decides Accept / Counteroffer / Decline (or skips) per offer.    [Step 2]
4. Answers each offer (Accept / Counter / Decline) via RespondToBestOffer when
   ACT_ON_OFFERS is enabled; otherwise a dry run just logs the intent.  [Step 6]
5. Writes the full result set to SQL and drafts a summary email with the
   read-only report workbook attached.                               [Step 5]

Import-safe: the prompt and the scheduler run only under
``if __name__ == "__main__"``, so tests can import this module freely.
"""
from __future__ import annotations

import html
import re
import traceback
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from seller_automation_utils import accounts, alert_utils, chrome, custom_functions, database_utils, ebay, excel_utils
from seller_automation_utils.accounts import EBAY_PROFILES
from seller_automation_utils.config_utils import get_env, load_config_safe
from seller_automation_utils.greeting import greeting_for
from seller_automation_utils.logging_utils import setup_logging
from seller_automation_utils.outlook import send_email
from seller_automation_utils.schedule_utils import run_on_schedule
from seller_automation_utils.screenshot_utils import save_debug_screenshot
from seller_automation_utils.ui_utils import ask_user

log = setup_logging("ebay_best_offers")

CONFIG_DIR = Path(__file__).resolve().parent / "config"


# =============================================================================
# CONFIG — read the pricing settings from the control workbook
# =============================================================================
# The control workbook is a simple spreadsheet the business team edits, so the
# pricing rules can change without touching the code. The workbook holds the
# actual numbers; the table below only says which labels to read and the range
# each value must fall in — so a typo or a wrong number stops the run and gets
# reported instead of driving a bad counteroffer.

SETTINGS_SHEET = "Settings"

# workbook label -> (name used in code, lowest ok, highest ok, kind). Percentage
# bounds are fractions (0.30 = 30%); "minimum estimated shipping" is a dollar
# amount. The "expected" hint in error messages is derived from the bounds, so no
# workbook value is ever hardcoded here.
SETTINGS = {
    "ebay commission":                       ("commission",                 0.01, 0.30, "pct"),
    "minimum profit margin":                 ("min_profit",                 0.00, 0.50, "pct"),
    "slow item minimum profit margin":       ("slow_min_profit",            0.00, 0.50, "pct"),
    "dead item minimum profit margin":       ("dead_min_profit",            0.00, 0.50, "pct"),
    "sell below cost minimum profit margin": ("sell_below_cost_min_profit", 0.00, 0.50, "pct"),
    "min counteroffer discount":             ("min_discount",               0.00, 0.50, "pct"),
    "max counteroffer discount":             ("max_discount",               0.00, 0.50, "pct"),
    "minimum estimated shipping":            ("shipping_floor",             0.00, 100.0, "usd"),
}


def read_settings(path: str) -> tuple[dict, list[str]]:
    """Read the settings workbook with pandas and check the values.

    pandas reads the .xlsx straight from disk, so we never open Excel here.

    Args:
        path: Full path to the settings workbook.

    Returns:
        ``(settings, problems)`` — see :func:`check_settings`.
    """
    try:
        sheet = pd.read_excel(path, sheet_name=SETTINGS_SHEET, header=None)
    except FileNotFoundError:
        return {}, [f"The settings file was not found at:\n{path}\n"
                    "Make sure the file exists there, then run again."]
    except Exception as error:
        return {}, [f"The settings file could not be opened:\n{path}\n"
                    f"Reason: {error}\nClose it if it is open, then run again."]

    # Column A holds the labels, column B the values -> build a label:value map.
    entered = {}
    for label, value in zip(sheet[0], sheet[1]):
        if isinstance(label, str) and label.strip():
            entered[label.strip().lower()] = value

    return check_settings(entered)


def check_settings(entered: dict) -> tuple[dict, list[str]]:
    """Turn the raw workbook cells into settings, collecting any problems.

    Pure (no file access) so it is easy to test.

    Args:
        entered: label (lower-cased) -> value, read from the Setting/Value columns.

    Returns:
        ``(settings, problems)``:
          settings - the tunables by name (profit / discount floors as fractions,
                     ``shipping_floor`` in dollars).
          problems - plain-language issues; an empty list means everything is fine.
    """
    settings = {}
    problems = []

    for label, (name, low, high, kind) in SETTINGS.items():
        if kind == "usd":
            expected = f"a dollar amount between ${low:,.0f} and ${high:,.0f}"
        else:
            expected = f"a percentage between {low:.0%} and {high:.0%}"
        value = entered.get(label)
        if value is None or pd.isna(value) or (isinstance(value, str) and not value.strip()):
            problems.append(f"'{label}' is missing or blank — enter {expected} in the Value column.")
            continue
        try:
            value = float(str(value).replace("$", "").replace(",", "").strip())
        except (TypeError, ValueError):
            problems.append(f"'{label}' should be {expected}, but it says '{value}'.")
            continue
        if not (low <= value <= high):
            shown = f"${value:,.2f}" if kind == "usd" else f"{value:.1%}"
            problems.append(f"'{label}' is {shown}, outside the allowed range — enter {expected}.")
            continue
        settings[name] = value

    if "min_discount" in settings and "max_discount" in settings and settings["max_discount"] < settings["min_discount"]:
        problems.append("'Max counteroffer discount' must be the same as or larger than 'Min counteroffer discount'.")

    return settings, problems


def email_settings_problems(problems: list[str], path: str) -> None:
    """Email the business team that the settings workbook needs a fix.

    Sends to SETTINGS_ALERT_TO and blind-copies SETTINGS_ALERT_BCC. A send
    failure is logged, never raised, so it can't hide the original problem.
    """
    # Escape: the problem text and path echo raw workbook cells, so a stray
    # angle bracket or ampersand there must not break (or inject into) the email.
    fixes = "\n".join(f"      <li>{html.escape(problem)}</li>" for problem in problems)
    body = f"""
    <p>Hi,</p>
    <p>The eBay Best Offers automation did not run today because the settings
       file needs a quick fix. No offers were sent to any buyer.</p>
    <p><b>Settings file:</b> {html.escape(path)}</p>
    <p><b>Please fix the following, then save the file:</b></p>
    <ul>
{fixes}
    </ul>
    <p><b>How to edit:</b> open the settings workbook, go to the {SETTINGS_SHEET}
       tab, and change only the yellow Value cells. Type each rate as a number and
       Excel shows it as a percent automatically. The automation will use your
       changes on its next run.</p>
    <p>Thank you!</p>
    """
    to = _split_emails(get_env("SETTINGS_ALERT_TO", default=""))
    bcc = _split_emails(get_env("SETTINGS_ALERT_BCC", default=""))
    if not to:
        log.error("SETTINGS_ALERT_TO is not set — cannot email the settings problem.")
        return
    try:
        send_email(
            account=get_env("SENDER_EMAIL", required=True),
            subject="eBay Best Offers — settings need a quick fix",
            body=body,
            to=to,
            bcc=bcc or None,
            show=False,
            send=True,
        )
        log.warning(f"Settings problem emailed to {', '.join(to)}.")
    except Exception:
        log.error("Failed to email the settings problem.")
        traceback.print_exc()


def _split_emails(raw: str | None) -> list[str]:
    """Turn a comma-separated env value into a clean list of addresses."""
    return [item.strip() for item in (raw or "").split(",") if item.strip()]


# =============================================================================
# SCRAPE — read each account's pending offers from the Seller Hub grid  [Step 3]
# =============================================================================
# eBay renders the offers as a data grid. We read each value from its own cell
# in one JavaScript call, so the badges and links eBay injects between cells are
# never mistaken for data. Every real offer row is `<tr class="grid-row"
# data-id="{item number}">`, so the item number comes straight off the row.

PENDING_OFFERS_URL = (
    "https://www.ebay.com/sh/lst/active"
    "?action=search&format=ALL_FORMATS&status=PENDING_OFFERS&offset=0&limit=200"
)

# One value per column, read from the grid cell it belongs to. `.clipped` spans
# are screen-reader text inside the cells — removed on a clone so the live page
# is untouched. An out-of-stock listing has no "Respond" link, so an empty href
# flags it. The order here must match the columns in `offers_to_frame`.
EXTRACT_OFFERS_JS = """
const strip = (el) => {
    if (!el) return "";
    const c = el.cloneNode(true);
    c.querySelectorAll(".clipped").forEach(n => n.remove());
    return c.textContent.trim();
};
return Array.from(document.querySelectorAll("tr.grid-row[data-id]")).map(r => [
    r.dataset.id,
    strip(r.querySelector(".shui-dt-column__title .column-title__text")),
    strip(r.querySelector(".shui-dt-column__listingSKU .shui-dt--text-column")),
    strip(r.querySelector(".shui-dt-column__price .col-price__current")),
    (r.querySelector(".shui-dt-column__lineActions a.primary-action__button") || {}).href || "",
]);
"""


def offers_to_frame(rows: list, account: str, today: str) -> pd.DataFrame:
    """Shape the raw grid rows into the offers table.

    Pure (no browser) so it is easy to test. Columns are ordered date, account,
    then the offer fields. An empty respond link means the listing is out of
    stock (no way to respond), flagged in ``out_of_stock``.

    Args:
        rows: ``[item_number, title, sku, current_price, respond_href]`` per offer.
        account: eBay account display name.
        today: Report date string (YYYY-MM-DD).

    Returns:
        DataFrame with columns date, account, title, sku, current_price,
        item_number, out_of_stock (current_price as a number).
    """
    df = pd.DataFrame(rows, columns=["item_number", "title", "sku", "current_price", "respond_href"])
    df["out_of_stock"] = df["respond_href"] == ""
    df["current_price"] = pd.to_numeric(
        df["current_price"].astype(str).str.replace(r"[$,]", "", regex=True), errors="coerce"
    )
    df.insert(0, "account", account)
    df.insert(0, "date", today)
    return df[["date", "account", "title", "sku", "current_price", "item_number", "out_of_stock"]]


def scrape_pending_offers(driver: object, account: str, today: str) -> pd.DataFrame:
    """Open one account's pending offers, reset the columns, and read the grid.

    Calls the shared ``customize_offers_table`` first so the table is reset to a
    known layout (the core Title/SKU/Price columns are always in that set), then
    re-applies the pending-offers view and reads every row in one JS call.

    Args:
        driver: Active browser already logged into the account.
        account: eBay account display name.
        today: Report date string (YYYY-MM-DD).

    Returns:
        The account's offers as a DataFrame (see :func:`offers_to_frame`); empty
        if the account has no pending offers.

    Raises:
        RuntimeError: If the grid renders but no rows are read (DOM changed).
    """
    driver.get(PENDING_OFFERS_URL)
    driver.switch_to_window(0)
    ebay.customize_offers_table(driver)  # reset columns; shared accounts drift

    driver.get(PENDING_OFFERS_URL)  # re-apply the pending-offers view after the reload
    driver.switch_to_window(0)
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".shui-dt")))
    except TimeoutException:
        log.info(f"No pending offers for [cyan]{account}[/cyan].")
        return offers_to_frame([], account, today)

    rows = driver.execute_script(EXTRACT_OFFERS_JS)
    if not rows:
        save_debug_screenshot(driver, root=account, section="scrape", description="no_offer_rows")
        raise RuntimeError("Pending-offers grid rendered but no rows were read — the eBay layout may have changed.")

    log.info(f"Read [cyan]{len(rows)}[/cyan] pending offers for [cyan]{account}[/cyan].")
    return offers_to_frame(rows, account, today)


# --- Enrich with SQL data ----------------------------------------------------
# The eBay grid gives us the item and its price; the SKU's cost and aged status
# come from SQL (the same two tables the old workbook read). We pull the two
# small reference tables once, then match them onto the offers by SKU in pandas.

def load_reference_data(conn: object) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Read the site-cost and aged-status reference tables from the Reports database.

    Two plain SELECTs, kept to one row per SKU:
      - ``SellerCloud``   -> the SKU's site cost and its SellBelowCost flag
      - ``AgedInventory`` -> the SKU's aged status (e.g. Dead / Slow)

    ``SellBelowCost`` is optional: the SellerCloud sync may not have added the
    column yet. If it's absent we still run — every SKU is treated as not
    sell-below-cost, so the 4% floor stays dormant — and the script picks the flag
    up automatically the next run after the column exists and is populated. A dropped
    column can therefore never crash the run.

    Args:
        conn: An open pyodbc connection to the ``Reports`` database.

    Returns:
        ``(site_costs, aged)`` DataFrames with columns [sku, site_cost,
        sell_below_cost] and [sku, aged_status].
    """
    def read(query: str, columns: list[str]) -> pd.DataFrame:
        cursor = conn.cursor()
        cursor.execute(query)
        rows = [tuple(row) for row in cursor.fetchall()]
        return pd.DataFrame(rows, columns=columns)

    probe = conn.cursor()
    probe.execute(
        "SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME = 'SellerCloud' AND COLUMN_NAME = 'SellBelowCost'"
    )
    if probe.fetchone() is not None:
        site_costs = read("SELECT SKU, SiteCost, SellBelowCost FROM dbo.SellerCloud",
                          ["sku", "site_cost", "sell_below_cost"])
    else:
        log.warning("SellerCloud.SellBelowCost not found — treating every SKU as not "
                    "sell-below-cost until the column exists and is populated.")
        site_costs = read("SELECT SKU, SiteCost FROM dbo.SellerCloud", ["sku", "site_cost"])
        site_costs["sell_below_cost"] = False

    aged = read("SELECT SKU, Status FROM dbo.AgedInventory", ["sku", "aged_status"])
    return site_costs, aged


def enrich_offers(offers: pd.DataFrame, site_costs: pd.DataFrame, aged: pd.DataFrame) -> pd.DataFrame:
    """Match site cost and aged status onto the offers by SKU.

    Pure (no database) so it is easy to test. A SKU with no site-cost row gets
    0 — the decision treats 0 (and 0.01) as "cost unknown" and skips it rather
    than risk a bad counteroffer. A SKU with no aged row gets "N/A"; a SKU with no
    (or NULL) SellBelowCost flag is treated as False.

    Args:
        offers: The scraped offers table (see :func:`offers_to_frame`).
        site_costs: [sku, site_cost, sell_below_cost] reference data.
        aged: [sku, aged_status] reference data.

    Returns:
        The offers table with ``site_cost``, ``aged_status`` and
        ``sell_below_cost`` columns added.
    """
    site_costs = site_costs.drop_duplicates("sku")  # one cost per SKU, never multiply rows
    aged = aged.drop_duplicates("sku")
    df = offers.merge(site_costs, on="sku", how="left")
    df = df.merge(aged, on="sku", how="left")
    df["site_cost"] = pd.to_numeric(df["site_cost"], errors="coerce").fillna(0).round(2)
    df["aged_status"] = df["aged_status"].fillna("N/A")
    df["sell_below_cost"] = df["sell_below_cost"].fillna(False).astype(bool)
    return df


AGED_TABLE = "AgedInventory"


def upload_aged_inventory(aged_inv_path: str, today: str) -> None:
    """Refresh the AgedInventory SQL table from its Excel report, if it changed today.

    The Aged Inventory report arrives as an Excel file every so often. On each
    run we check the file's modified date: if it was updated today, we replace
    the SQL table's contents with its rows (the same three sheets the old script
    read); otherwise we leave the table alone.

    Args:
        aged_inv_path: Full path to the Aged Inventory .xlsx file.
        today: Today's date string (YYYY-MM-DD).
    """
    aged_file = Path(aged_inv_path)
    if not aged_file.exists():
        log.warning(f"Aged Inventory file not found at [cyan]{aged_inv_path}[/cyan] — skipping.")
        return

    modified = datetime.fromtimestamp(aged_file.stat().st_mtime).strftime("%Y-%m-%d")
    if modified != today:
        log.info("Aged Inventory not updated today — skipping its upload.")
        return

    log.info("Aged Inventory updated today — refreshing the SQL table.")
    frames = [
        pd.read_excel(aged_file, sheet_name=sheet, skiprows=4)[["SKU", "Status"]]
        for sheet in ("Raw data", "Dead", "Slow")
    ]
    aged = pd.concat(frames, ignore_index=True)

    conn = custom_functions.sql_connection("Reports")
    try:
        cursor = conn.cursor()
        cursor.execute(f"DELETE FROM {AGED_TABLE}")
        database_utils.insert_dataframe(cursor, AGED_TABLE, aged, ["SKU", "Status"])
        conn.commit()
    finally:
        conn.close()
    log.success(f"Aged Inventory refreshed — [cyan]{len(aged)}[/cyan] rows.")


# =============================================================================
# DECIDE — profit-margin decision                                     [Step 2]
# =============================================================================
# Pure functions from the offer numbers to an action — no browser, no database,
# so every branch is easy to test. Counter with the biggest discount in the
# allowed band that still clears the minimum profit (aged / SellBelowCost items
# ease that floor) — the best price for the buyer that still protects us.

def est_shipping(site_cost: float, floor: float) -> float:
    """Estimate what it costs us to ship an item, from its site cost.

    A tiered rate on the site cost (cheaper items ship at a higher %), with a
    ``floor`` (the least we assume) that comes from the settings workbook.
    """
    if site_cost < 1000:
        rate = 0.05
    elif site_cost < 2000:
        rate = 0.03
    else:
        rate = 0.02
    return round(max(site_cost * rate, floor), 2)


def margin(price: float, total_cost: float, commission: float) -> float:
    """Our profit as a fraction of the sale price, after cost and commission.

    Shared by :func:`decide_offer` (which price clears the min profit?) and
    :func:`build_results` (what margin does a given price actually yield?), so the
    formula lives in one place.
    """
    return (price - total_cost - price * commission) / price


def effective_min_profit(aged_status: str, sell_below_cost: bool, settings: dict) -> float:
    """The lowest minimum-profit floor that applies to one item.

    Starts at the default and eases to the more permissive floor for a Slow / Dead
    aged item or a SellBelowCost SKU, so flagged inventory can be countered deeper.
    When several apply, the lowest (most permissive) wins. All floors come from the
    settings workbook.
    """
    floors = [settings["min_profit"]]
    aged = str(aged_status).strip().lower()
    if aged == "slow":
        floors.append(settings["slow_min_profit"])
    elif aged == "dead":
        floors.append(settings["dead_min_profit"])
    if sell_below_cost:
        floors.append(settings["sell_below_cost_min_profit"])
    return min(floors)


def decide_offer(cx_offer: float, current_price: float, site_cost: float,
                 aged_status: str, sell_below_cost: bool, settings: dict,
                 out_of_stock: bool = False) -> tuple[str, float, float]:
    """Decide what to do with one buyer offer.

    Args:
        cx_offer: The buyer's offer amount (0 if there's no readable offer).
        current_price: Our current list price.
        site_cost: The SKU's site cost (0 or 0.01 means "cost unknown").
        aged_status: The SKU's aged status ("Slow" / "Dead" ease the min profit).
        sell_below_cost: The SKU's SellBelowCost flag (also eases the min profit).
        settings: The control-workbook settings.
        out_of_stock: The listing has no sellable quantity — eBay blocks Accept and
            Counter on it, so we skip it rather than send a doomed response.

    Returns:
        ``(action, counter_price, profit_pct)``:
          action - Expired Offer / Out of Stock / Missing Site Cost / Accepted /
                   Counteroffer / Declined.
          counter_price - the price to counter at (0.0 unless it's a Counteroffer).
          profit_pct - our margin at the acted price (0.0 for skip / decline).
    """
    # Cases where we can't safely act on the offer — skip with the reason why.
    if cx_offer <= 0:
        return ("Expired Offer", 0.0, 0.0)  # was pending at scrape time, now unreadable = expired/withdrawn
    if out_of_stock:
        return ("Out of Stock", 0.0, 0.0)   # can't fulfill; a counter/accept would fail at eBay (err 21922)
    if site_cost == 0 or site_cost == 0.01:
        return ("Missing Site Cost", 0.0, 0.0)

    commission = settings["commission"]
    min_profit = effective_min_profit(aged_status, sell_below_cost, settings)
    total_cost = site_cost + est_shipping(site_cost, settings["shipping_floor"])

    # The buyer's own offer already clears the margin — accept it.
    if margin(cx_offer, total_cost, commission) >= min_profit:
        return ("Accepted", 0.0, round(margin(cx_offer, total_cost, commission), 4))

    # Otherwise counter with the biggest discount (within the allowed band) that
    # still clears the margin. A lower price is better for the buyer but lowers
    # our margin, so we start at the deepest discount and ease back only as far
    # as we must to hold the margin.
    lowest_price = current_price * (1 - settings["max_discount"])   # deepest allowed discount, best for the buyer
    highest_price = current_price * (1 - settings["min_discount"])  # shallowest allowed discount, best for us
    break_even = total_cost / (1 - commission - min_profit)         # price where margin == min_profit

    target = max(lowest_price, break_even)  # cheapest price that still clears the margin
    if target <= highest_price:
        counter = round(target, 2)
        return ("Counteroffer", counter, round(margin(counter, total_cost, commission), 4))

    # Even the shallowest allowed discount can't clear the margin — decline.
    return ("Declined", 0.0, 0.0)


# =============================================================================
# READ OFFERS via the eBay Trading API (GetBestOffers)                 [Step 3]
# =============================================================================
# One GetBestOffers call returns every active best offer for an account (item,
# amount, BestOfferID) — no page navigation, so eBay's bot check never fires. This
# replaces the old per-item browser read. Good candidate to promote into
# seller_automation_utils. Buyer identity in the response is intentionally never
# parsed/stored (keeps our Marketplace-Account-Deletion exemption valid).

TRADING_ENDPOINT = "https://api.ebay.com/ws/api.dll"
TRADING_COMPAT_LEVEL = "1193"


def _token_env(account: str) -> str:
    """Env var name holding a seller account's Trading API token.

    ``Account4`` -> ``EBAY_AUTH_TOKEN_ACCOUNT4``.
    """
    return "EBAY_AUTH_TOKEN_" + re.sub(r"[^A-Za-z0-9]", "", account).upper()


def build_get_best_offers_xml(token: str) -> str:
    """Build the GetBestOffers request body — all active offers for the account.

    Pure (no HTTP) so it is easy to test. The token is XML-escaped defensively.
    """
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<GetBestOffersRequest xmlns="urn:ebay:apis:eBLBaseComponents">'
        f"<RequesterCredentials><eBayAuthToken>{html.escape(token)}</eBayAuthToken></RequesterCredentials>"
        "<BestOfferStatus>Active</BestOfferStatus>"
        "<DetailLevel>ReturnAll</DetailLevel>"
        "</GetBestOffersRequest>"
    )


def parse_best_offers(xml: bytes | str) -> dict:
    """Parse a GetBestOffers response into ack, errors, and the offers.

    Pure (no HTTP) so it is easy to test. Namespaces are stripped so elements are
    reachable by local name. The buyer's identity in the response is deliberately
    not read.

    Args:
        xml: The raw response (bytes preferred; a str is encoded first so an XML
            encoding declaration doesn't trip ElementTree).

    Returns:
        ``{"ack": str, "errors": list[str], "offers": list[dict]}`` where each
        offer is ``{item_number, cx_offer, best_offer_id, quantity, code, status}``.
    """
    if isinstance(xml, str):
        xml = xml.encode("utf-8")
    root = ET.fromstring(xml)
    for el in root.iter():
        el.tag = el.tag.split("}")[-1]

    errors = [f"{e.findtext('ErrorCode')}: {e.findtext('LongMessage')}" for e in root.findall(".//Errors")]
    offers = []
    for group in root.findall(".//ItemBestOffers"):
        item = group.find("Item")
        item_number = item.findtext("ItemID") if item is not None else None
        for bo in group.findall(".//BestOffer"):
            price = bo.find("Price")
            qty_text = bo.findtext("Quantity")
            offers.append({
                "item_number": item_number,
                "cx_offer": round(float(price.text), 2) if price is not None and price.text else 0.0,
                "best_offer_id": bo.findtext("BestOfferID"),
                "quantity": int(float(qty_text)) if qty_text else 1,
                "code": bo.findtext("BestOfferCodeType"),
                "status": bo.findtext("Status"),
            })
    return {"ack": root.findtext("Ack") or "", "errors": errors, "offers": offers}


def get_best_offers(token: str) -> list[dict]:
    """Call GetBestOffers and return the account's active offers.

    Args:
        token: The seller account's Trading API user token.

    Returns:
        A list of ``{item_number, cx_offer, best_offer_id, code, status}``.

    Raises:
        RuntimeError: If eBay returns a Failure Ack (auth/keys/other), with the
            eBay error messages attached.
    """
    headers = {
        "X-EBAY-API-CALL-NAME": "GetBestOffers",
        "X-EBAY-API-SITEID": "0",
        "X-EBAY-API-COMPATIBILITY-LEVEL": TRADING_COMPAT_LEVEL,
        "X-EBAY-API-APP-NAME": get_env("EBAY_APP_ID", required=True),
        "X-EBAY-API-DEV-NAME": get_env("EBAY_DEV_ID", required=True),
        "X-EBAY-API-CERT-NAME": get_env("EBAY_CERT_ID", required=True),
        "Content-Type": "text/xml",
    }
    body = build_get_best_offers_xml(token).encode("utf-8")
    resp = requests.post(TRADING_ENDPOINT, data=body, headers=headers, timeout=60)
    result = parse_best_offers(resp.content)
    if result["ack"] not in ("Success", "Warning"):
        raise RuntimeError(f"GetBestOffers failed: {result['errors'] or resp.status_code}")
    return result["offers"]


# =============================================================================
# RESPOND — send the decided Accept / Counter / Decline to eBay        [Step 6]
# =============================================================================
# One RespondToBestOffer call per answered offer. Safe by default: nothing is sent
# unless ACT_ON_OFFERS is truthy in the environment. With the flag off the run is a
# dry run — it logs the action it would take and sends nothing, so even a scheduled
# run can't answer a buyer before the business signs off. Once an offer is answered
# it leaves eBay's Active set, so a later GetBestOffers won't return it again — a
# rerun can't double-answer.

RESPOND_ACTIONS = {"Accepted": "Accept", "Counteroffer": "Counter", "Declined": "Decline"}

# eBay caps the buyer-facing message (SellerResponse) at 250 characters; a longer
# one gets the whole call rejected. These are trimmed to fit worst-case (the
# counter's {price} is filled at send time). Pending the business's final sign-off.
MESSAGE_LIMIT = 250
COUNTER_MESSAGE = (
    "Thanks for your offer! Our absolute final price is ${price}.\n\n"
    "This is unbeatable value! Fast, free shipping and worry-free returns "
    "from a trusted seller. We can't go lower on this item.\n\n"
    "Accept our counteroffer now to secure your deal!"
)
DECLINE_MESSAGE = (
    "Thank you for your offer!\n\n"
    "Unfortunately, we cannot accept it at this time.\n\n"
    "If you are still interested, we would love to fulfill your order, which "
    "includes fast, free shipping and worry-free returns."
)


def build_respond_offer_xml(token: str, item_id: str, best_offer_id: str, action: str,
                            counter_price: float | None = None, counter_quantity: int = 1,
                            message: str | None = None) -> str:
    """Build a RespondToBestOffer request body for one offer.

    Pure (no HTTP) so it is easy to test. Token and message are XML-escaped. Only a
    ``Counter`` carries the price and quantity; ``Accept`` and ``Decline`` omit them.
    A message (``SellerResponse``) is included when provided and is hard-capped at
    :data:`MESSAGE_LIMIT` so an over-long note can't get the call rejected.

    Args:
        token: The seller account's Trading API user token.
        item_id: The listing's eBay item number.
        best_offer_id: The offer to respond to.
        action: ``Accept``, ``Counter``, or ``Decline``.
        counter_price: The counter amount (required for ``Counter``).
        counter_quantity: The counter quantity (``Counter`` only).
        message: Optional buyer-facing note.

    Returns:
        The XML request body as a string.
    """
    parts = [
        '<?xml version="1.0" encoding="utf-8"?>',
        '<RespondToBestOfferRequest xmlns="urn:ebay:apis:eBLBaseComponents">',
        f"<RequesterCredentials><eBayAuthToken>{html.escape(token)}</eBayAuthToken></RequesterCredentials>",
        f"<ItemID>{html.escape(item_id)}</ItemID>",
        f"<BestOfferID>{html.escape(best_offer_id)}</BestOfferID>",
        f"<Action>{action}</Action>",
    ]
    if action == "Counter":
        parts.append(f'<CounterOfferPrice currencyID="USD">{counter_price:.2f}</CounterOfferPrice>')
        parts.append(f"<CounterOfferQuantity>{int(counter_quantity)}</CounterOfferQuantity>")
    if message:
        parts.append(f"<SellerResponse>{html.escape(message[:MESSAGE_LIMIT])}</SellerResponse>")
    parts.append("</RespondToBestOfferRequest>")
    return "".join(parts)


def parse_respond_response(xml: bytes | str) -> dict:
    """Parse a RespondToBestOffer response into ack and errors.

    Pure (no HTTP). Namespaces are stripped so elements are reachable by local name.

    Returns:
        ``{"ack": str, "errors": list[str]}``.
    """
    if isinstance(xml, str):
        xml = xml.encode("utf-8")
    root = ET.fromstring(xml)
    for el in root.iter():
        el.tag = el.tag.split("}")[-1]
    errors = [f"{e.findtext('ErrorCode')}: {e.findtext('LongMessage')}" for e in root.findall(".//Errors")]
    return {"ack": root.findtext("Ack") or "", "errors": errors}


def respond_to_best_offer(token: str, item_id: str, best_offer_id: str, action: str,
                          counter_price: float | None = None, counter_quantity: int = 1,
                          message: str | None = None) -> dict:
    """Send one RespondToBestOffer call and return its ack and errors.

    A non-Success ack is returned, not raised, so the caller can log it and carry on
    with the account's other offers.

    Returns:
        ``{"ack": str, "errors": list[str]}``.
    """
    headers = {
        "X-EBAY-API-CALL-NAME": "RespondToBestOffer",
        "X-EBAY-API-SITEID": "0",
        "X-EBAY-API-COMPATIBILITY-LEVEL": TRADING_COMPAT_LEVEL,
        "X-EBAY-API-APP-NAME": get_env("EBAY_APP_ID", required=True),
        "X-EBAY-API-DEV-NAME": get_env("EBAY_DEV_ID", required=True),
        "X-EBAY-API-CERT-NAME": get_env("EBAY_CERT_ID", required=True),
        "Content-Type": "text/xml",
    }
    body = build_respond_offer_xml(
        token, item_id, best_offer_id, action, counter_price, counter_quantity, message
    ).encode("utf-8")
    resp = requests.post(TRADING_ENDPOINT, data=body, headers=headers, timeout=60)
    return parse_respond_response(resp.content)


def _offer_log_line(row, settings: dict) -> str:
    """One human-readable line describing the action taken on an offer.

    Accepted / Counteroffer read straight from the decided row; Declined recomputes
    the best-case margin (at the shallowest allowed discount) to show why it failed.
    """
    item = row.item_number
    if row.action == "Accepted":
        return (f"Offer accepted for item {item} at ${row.cx_offer:,.2f} "
                f"with a profit of {row.buyer_margin * 100:.2f}%.")
    if row.action == "Counteroffer":
        return (f"Counteroffered for item {item} from ${row.cx_offer:,.2f} to ${row.counter:,.2f} "
                f"with a profit of {row.counter_margin * 100:.2f}%.")
    best_margin = margin(row.current_price * (1 - settings["min_discount"]), row.total_cost, settings["commission"])
    shape = "negative" if best_margin < 0 else "too low"
    return (f"Offer declined for item {item}. Margin was {shape} at {best_margin * 100:.2f}% "
            f"after the lowest discount {settings['min_discount'] * 100:.0f}% applied.")


def respond_to_offers(results: pd.DataFrame, api_by_item: dict, token: str, settings: dict, live: bool) -> dict:
    """Answer each decided offer on eBay (or log the intent in a dry run).

    Walks the decided rows and, for every Accept / Counter / Decline, sends one
    RespondToBestOffer using the live ``BestOfferID`` from the API read. Skip reasons
    (Expired Offer, Missing Site Cost) are never sent. With ``live`` false nothing is
    sent — each intended action is logged instead. A single offer's failure (a bad
    ack or a network error) is logged and skipped so the rest still go through.

    Args:
        results: The decided rows for one account (:func:`build_results` output).
        api_by_item: ``item_number -> offer dict`` from :func:`get_best_offers`
            (carries ``best_offer_id`` and ``quantity``).
        token: The seller account's Trading API user token.
        settings: The control-workbook settings (for the declined-offer log line).
        live: Send for real when True; dry-run (log only) when False.

    Returns:
        ``{item_number: answered_at}`` for offers answered successfully — used to
        stamp ``acted_at`` so a rerun never answers them again.
    """
    acted: dict = {}
    answered_at = datetime.now()
    for row in results.itertuples(index=False):
        ebay_action = RESPOND_ACTIONS.get(row.action)
        if ebay_action is None:
            continue  # Expired Offer / Missing Site Cost — nothing to send
        info = api_by_item.get(row.item_number)
        best_offer_id = info.get("best_offer_id") if info else None
        if not best_offer_id:
            log.warning(f"No BestOfferID for {row.account}/{row.item_number} — cannot {ebay_action}; skipped.")
            continue

        if ebay_action == "Counter":
            price, message = row.counter, COUNTER_MESSAGE.format(price=f"{row.counter:,.2f}")
        elif ebay_action == "Decline":
            price, message = None, DECLINE_MESSAGE
        else:
            price, message = None, None
        quantity = int(info.get("quantity", 1) or 1)
        line = _offer_log_line(row, settings)

        if not live:
            log.info(f"[DRY RUN] {line}")
            continue

        try:
            resp = respond_to_best_offer(token, row.item_number, best_offer_id, ebay_action, price, quantity, message)
        except Exception as exc:
            log.error(f"{ebay_action} call errored for {row.account}/{row.item_number}: {exc}")
            continue
        if resp["ack"] in ("Success", "Warning"):
            acted[row.item_number] = answered_at
            log.success(line)
        else:
            log.error(f"{ebay_action} rejected for {row.account}/{row.item_number}: {resp['errors']}")
    return acted


# =============================================================================
# RECORD — build the full result record and store it in SQL            [Step 4]
# =============================================================================
# Every offer becomes one archive row: the numbers we scraped, the cost we
# enriched, the margins we computed, and the decision we reached. `BestOffers`
# is a permanent history — we only ever add today's rows, never wipe past days.

RESULT_COLUMNS = [
    "report_date", "account", "title", "sku", "item_number",
    "current_price", "site_cost", "est_shipping", "total_cost",
    "aged_status", "cx_offer", "buyer_margin",
    "action", "counter", "discount", "counter_margin", "out_of_stock",
]


def build_results(offers: pd.DataFrame, settings: dict) -> pd.DataFrame:
    """Compute the full archive record for each enriched offer.

    Pure (no browser, no database) so every branch is testable. For each offer we
    ask :func:`decide_offer` for the action and counter price, then compute the
    shipping, total cost, and margins that go in the archive. Fields that don't
    apply to a row (a counter on a declined offer, a buyer margin on an expired
    one) are left as None so they store as SQL NULL — blank in the report, never a
    misleading 0.

    Args:
        offers: Enriched offers with a ``cx_offer`` column (see
            :func:`enrich_offers` and :func:`get_best_offers`).
        settings: The control-workbook settings.

    Returns:
        A DataFrame of ``RESULT_COLUMNS``, one row per offer, ready to store.
    """
    commission = settings["commission"]
    records = []
    for row in offers.itertuples(index=False):
        shipping = est_shipping(row.site_cost, settings["shipping_floor"])
        total_cost = round(row.site_cost + shipping, 2)

        action, counter_price, _ = decide_offer(
            row.cx_offer, row.current_price, row.site_cost,
            row.aged_status, row.sell_below_cost, settings, bool(row.out_of_stock)
        )
        # buyer_margin only means something when the offer is actually priceable.
        priceable = row.cx_offer > 0 and row.site_cost not in (0, 0.01)
        buyer_margin = round(margin(row.cx_offer, total_cost, commission), 4) if priceable else None

        is_counter = action == "Counteroffer"
        counter = counter_price if is_counter else None
        discount = round((row.current_price - counter_price) / row.current_price, 4) if is_counter else None
        counter_margin = round(margin(counter_price, total_cost, commission), 4) if is_counter else None

        records.append({
            "report_date": row.date,
            "account": row.account,
            "title": row.title,
            "sku": row.sku,
            "item_number": row.item_number,
            "current_price": row.current_price,
            "site_cost": row.site_cost,
            "est_shipping": shipping,
            "total_cost": total_cost,
            "aged_status": row.aged_status,
            "cx_offer": row.cx_offer,
            "buyer_margin": buyer_margin,
            "action": action,
            "counter": counter,
            "discount": discount,
            "counter_margin": counter_margin,
            "out_of_stock": bool(row.out_of_stock),
        })
    return pd.DataFrame(records, columns=RESULT_COLUMNS)


# --- Store in the permanent archive ------------------------------------------
# We add each account's rows as its own committed insert; a same-day rerun skips
# any account already recorded, so a day is never double-counted. If a later
# account fails, this run's inserts are removed and the team is alerted.

RESULTS_DB = "eBay"
RESULTS_TABLE = "BestOffers"

# The computed record plus acted_at — when the offer was answered on eBay (Step 6),
# NULL until then and throughout a dry run.
STORE_COLUMNS = RESULT_COLUMNS + ["acted_at"]


def account_fully_acted(conn: object, account: str, today: str) -> bool:
    """True if this account's offers for today are all recorded AND answered.

    A same-day rerun skips only accounts that are fully done — every offer recorded
    and answered (``acted_at`` set). An account recorded but not yet answered is
    reprocessed so its offers still get handled. No offers are answered until Step 6,
    so today this is always False and every account is reprocessed.
    """
    cursor = conn.cursor()
    cursor.execute(
        f"SELECT COUNT(*), COUNT(CASE WHEN acted_at IS NULL THEN 1 END) "
        f"FROM {RESULTS_TABLE} WHERE report_date = ? AND account = ?",
        today, account,
    )
    total, unacted = cursor.fetchone()
    return total > 0 and unacted == 0


def store_account_results(conn: object, results: pd.DataFrame, account: str, today: str) -> None:
    """Replace one account's not-yet-answered rows for today, in a single commit.

    Deletes the account's un-answered rows for today first, so a rerun refreshes the
    day's decisions with the latest scrape without duplicating; rows already answered
    (a real offer was sent) are kept as the permanent record. NaN becomes None before
    insert — ``insert_dataframe`` binds with pyodbc, which can't bind NaN to a nullable
    column, so not-applicable fields must arrive as None to land as SQL NULL.
    """
    cursor = conn.cursor()
    cursor.execute(
        f"DELETE FROM {RESULTS_TABLE} WHERE report_date = ? AND account = ? AND acted_at IS NULL",
        today, account,
    )
    if results.empty:
        conn.commit()
        return
    if "acted_at" not in results.columns:
        results = results.assign(acted_at=None)
    clean = results.astype(object).where(pd.notnull(results), None)
    database_utils.insert_dataframe(cursor, RESULTS_TABLE, clean, STORE_COLUMNS)


def remove_run_results(conn: object, accounts: list[str], today: str) -> None:
    """Undo this run's not-yet-answered inserts for the given accounts after a failure.

    Keeps a run all-or-nothing for work that wasn't finalized: if one account fails,
    the un-answered rows this run inserted are removed so the day isn't left
    half-recorded. Rows already answered (a real offer was sent) and past days are
    never touched.
    """
    if not accounts:
        return
    cursor = conn.cursor()
    for account in accounts:
        cursor.execute(
            f"DELETE FROM {RESULTS_TABLE} WHERE report_date = ? AND account = ? AND acted_at IS NULL",
            today, account,
        )
    conn.commit()


# =============================================================================
# REPORT — refresh the read-only workbook and email the summary        [Step 5]
# =============================================================================
# The workbook is a read-only view (a Power Query over today's BestOffers rows);
# the script only refreshes it, never writes cells. The email summary is built
# straight from the in-memory results. While the automation is paused it is
# drafted (opened, not sent) for review. Rates/brands come from `settings`.

REPORT_ACTIONS = [
    "Accepted", "Counteroffer", "Declined",
    "Expired Offer", "Out of Stock", "Missing Site Cost",
]

# Short column headers for the email summary table (the SQL/action names run long).
ACTION_LABELS = {
    "Accepted": "Accepted",
    "Counteroffer": "Counter",
    "Declined": "Declined",
    "Expired Offer": "Expired",
    "Out of Stock": "Out of Stock",
    "Missing Site Cost": "Missing Cost",
}


def build_summary_email(results: pd.DataFrame, settings: dict, greeting_text: str | None = None) -> str:
    """Build the HTML summary of today's decisions from the in-memory results.

    Reads only the current time (for the greeting) — no Outlook/Excel — so it is
    easy to test. Opens with a time-of-day greeting and a short intro, then a
    polished per-account summary of each outcome with totals, a footer echoing the
    settings the run used (all from ``settings`` — nothing hardcoded), and a plain
    closing. Per-offer counteroffer detail lives in the attached workbook, not
    here. Account names are HTML-escaped so a stray ``&`` or ``<`` can't break the
    email.

    Args:
        results: Today's combined result rows (see :func:`build_results`).
        settings: The control-workbook settings.
        greeting_text: Override the greeting (used by tests); defaults to the
            time-of-day greeting for the current hour.

    Returns:
        The email body as an HTML string.
    """
    greeting = greeting_text if greeting_text is not None else greeting_for()
    font = "font-family:Segoe UI,Arial,sans-serif"

    def esc(value: object) -> str:
        return html.escape(str(value))

    report_iso = str(results["report_date"].iloc[0]) if len(results) else ""
    try:
        report_date = datetime.strptime(report_iso, "%Y-%m-%d").strftime("%m/%d/%Y")
    except ValueError:
        report_date = report_iso

    def th(text: str, align: str = "center") -> str:
        return (f"<th style='{font};background-color:#334155;color:#ffffff;padding:7px 12px;"
                f"text-align:{align};font-weight:600;font-size:13px'>{text}</th>")

    def td(value: object, align: str = "center", bg: str = "#ffffff",
           bold: bool = False, color: str = "#1e293b", top: str = "") -> str:
        return (f"<td style='{font};padding:7px 12px;text-align:{align};background-color:{bg};"
                f"color:{color};font-weight:{'700' if bold else '400'};font-size:13px;"
                f"border-bottom:1px solid #e2e8f0;{top}'>{value}</td>")

    # Per-account count of each outcome. Zeros show blank for a calmer table.
    header = th("Account", "left") + "".join(th(ACTION_LABELS[a]) for a in REPORT_ACTIONS) + th("Total")
    body_rows = ""
    for i, account in enumerate(dict.fromkeys(results["account"])):
        acc = results[results["account"] == account]
        bg = "#f8fafc" if i % 2 else "#ffffff"
        cells = td(esc(account), "left", bg, bold=True)
        for action in REPORT_ACTIONS:
            cells += td(int((acc["action"] == action).sum()) or "", bg=bg)
        cells += td(len(acc), bg=bg, bold=True)
        body_rows += f"<tr>{cells}</tr>"

    top = "border-top:2px solid #334155"
    total_row = td("All accounts", "left", "#e2e8f0", bold=True, top=top)
    for action in REPORT_ACTIONS:
        total_row += td(int((results["action"] == action).sum()) or "", bg="#e2e8f0", bold=True, top=top)
    total_row += td(len(results), bg="#e2e8f0", bold=True, top=top)

    table = (
        "<table style='border-collapse:collapse;margin:8px 0 4px'>"
        f"<tr>{header}</tr>{body_rows}<tr>{total_row}</tr></table>"
    )

    footer = (
        "<p style='font-family:Segoe UI,Arial,sans-serif;font-size:12px;color:#555'>"
        f"Rules used: commission {settings['commission']:.1%}; minimum profit {settings['min_profit']:.1%} "
        f"(Slow {settings['slow_min_profit']:.1%}, Dead {settings['dead_min_profit']:.1%}, "
        f"sell-below-cost {settings['sell_below_cost_min_profit']:.1%}); counter discount band "
        f"{settings['min_discount']:.0%} to {settings['max_discount']:.0%}; estimated shipping floor "
        f"${settings['shipping_floor']:,.2f}.</p>"
    )
    intro = (
        f"<p style='{font}'>{esc(greeting)},</p>"
        f"<p style='{font}'>Here is today's summary of the pending Best Offers across all seller "
        "accounts. The attached workbook has the full detail for every offer.</p>"
    )
    title = f"<p style='{font};font-weight:600;font-size:14px;margin:14px 0 4px'>Today's Best Offers ({report_date})</p>"
    closing = (
        f"<p style='{font}'>Let me know if you have any questions.</p>"
        f"<p style='{font}'>Thanks,</p>"
    )
    return f"<div>{intro}{title}{table}{footer}{closing}</div>"


def draft_summary_email(results: pd.DataFrame, settings: dict, workbook_path: str | None) -> None:
    """Refresh the report workbook and draft the summary email for review.

    The automation is paused, so the email is drafted (opened in Outlook, not
    sent). The refreshed workbook is attached. A refresh failure is logged and the
    draft is still created — the SQL archive is the source of truth; the workbook
    is only a view. Recipients default to the sender (draft to Brian) via
    ``REPORT_DRAFT_TO``; switch to the stakeholder list and ``send=True`` when live.

    Args:
        results: Today's combined result rows.
        settings: The control-workbook settings.
        workbook_path: Path to the report workbook, or None to skip the attachment.
    """
    if workbook_path:
        try:
            excel_utils.refresh_workbook(workbook_path)
        except Exception:
            log.error("Could not refresh the report workbook — drafting the email without a fresh refresh.")
            traceback.print_exc()

    report_to = _split_emails(get_env("REPORT_DRAFT_TO", default="")) or [get_env("SENDER_EMAIL", required=True)]
    attachments = [workbook_path] if workbook_path and Path(workbook_path).exists() else None
    send_email(
        account=get_env("SENDER_EMAIL", required=True),
        subject=f"eBay Report - Best Offers - {datetime.now().strftime('%m/%d/%Y')}",  # fleet convention
        body=build_summary_email(results, settings),
        to=report_to,
        attachments=attachments,
        show=True,   # open the draft for review
        send=True,  # paused — never auto-send
    )
    log.success("Summary email drafted for review.")


# =============================================================================
# ORCHESTRATION
# =============================================================================

def main() -> None:
    """Read the settings, then run the daily best-offer flow.

    Settings are read first: if the workbook has a problem, email the business
    team and stop before any browser or database work. Then each account is
    scraped, enriched, decided, answered on eBay, and archived to SQL, and a
    summary email is drafted. Answering is gated behind ``ACT_ON_OFFERS``: with the
    flag off the run is a dry run that records decisions but sends nothing.
    """
    load_dotenv()

    paths = load_config_safe(CONFIG_DIR / "paths.json")
    control_wb_path = paths.get("control_wb_path")
    if not control_wb_path:
        email_settings_problems(
            ["'control_wb_path' is not set in config/paths.json — it must point to the settings file."],
            "config/paths.json",
        )
        raise SystemExit(1)

    settings, problems = read_settings(control_wb_path)
    if problems:
        email_settings_problems(problems, control_wb_path)
        log.error("Settings problem — emailed the business team; run stopped.")
        raise SystemExit(1)

    log.success(
        f"Settings loaded — commission {settings['commission']:.1%}, "
        f"min profit {settings['min_profit']:.1%} (Slow {settings['slow_min_profit']:.1%}, "
        f"Dead {settings['dead_min_profit']:.1%}, sell-below-cost {settings['sell_below_cost_min_profit']:.1%}), "
        f"discount band {settings['min_discount']:.0%}-{settings['max_discount']:.0%}, "
        f"shipping floor ${settings['shipping_floor']:,.2f}."
    )

    today = date.today().strftime("%Y-%m-%d")
    user_data_dir = get_env("CHROME_USER_DATA_DIR", required=True)
    password = get_env("eBay_pass", required=True)

    # Safety gate: only answer offers on eBay when explicitly enabled. Off => dry run.
    act_live = (get_env("ACT_ON_OFFERS") or "").strip().lower() in ("1", "true", "yes", "on")
    if act_live:
        log.warning("ACT_ON_OFFERS is ON — decided offers will be ANSWERED on eBay (Accept / Counter / Decline).")
    else:
        log.info("Dry run — offers are decided and recorded but not answered on eBay (ACT_ON_OFFERS off).")

    # Wrap the operational body so any unexpected failure emails a crash report.
    # Per-account failures already alert and stop inside the loop (they raise
    # SystemExit, which bypasses this handler); this catches everything else
    # (aged upload, reference load, SQL connection).
    ebay_conn = None
    injected: list[str] = []
    all_results: list[pd.DataFrame] = []
    try:
        # Refresh Aged Inventory from its Excel report if it was updated today.
        aged_inv_path = paths.get("aged_inv_path")
        if aged_inv_path:
            upload_aged_inventory(aged_inv_path, today)

        # Read the SQL reference tables once, then match them onto each account.
        reports_conn = custom_functions.sql_connection("Reports")
        try:
            site_costs, aged = load_reference_data(reports_conn)
        finally:
            reports_conn.close()
        log.info(f"Loaded [cyan]{len(site_costs)}[/cyan] site costs and [cyan]{len(aged)}[/cyan] aged-status rows.")

        # Scrape, enrich, decide, answer, and archive each account. A rerun refreshes
        # each account's not-yet-answered rows and skips accounts already fully answered.
        # If an account fails, this run's un-answered inserts are undone and the team is
        # alerted. Offers are answered on eBay only when ACT_ON_OFFERS is set (else dry run).
        ebay_conn = custom_functions.sql_connection(RESULTS_DB)
        for account, profile in EBAY_PROFILES.items():
            if account_fully_acted(ebay_conn, account, today):
                log.info(f"[cyan]{account}[/cyan] already recorded and answered today — skipping.")
                continue

            driver = None
            try:
                driver = chrome.start_browser(user_data_dir, profile, headless=True)
                driver.get("https://www.ebay.com/sh/ovw")
                driver.switch_to_window(0)
                try:
                    accounts.ebay(password=password, driver=driver)
                except TimeoutException:
                    pass

                offers = scrape_pending_offers(driver, account, today)
                offers = enrich_offers(offers, site_costs, aged)

                # Read every buyer offer in one Trading API call — no page navigation,
                # so eBay's bot check never fires. An item in the grid but not in the
                # API response has no active offer (cx_offer 0 -> Expired Offer).
                token = get_env(_token_env(account), required=True)
                api_offers = get_best_offers(token)
                cx_by_item = {o["item_number"]: o["cx_offer"] for o in api_offers}
                offers["cx_offer"] = offers["item_number"].map(cx_by_item).fillna(0.0)

                results = build_results(offers, settings)

                # Answer each decided offer on eBay (or just log the intent in a dry run),
                # then stamp acted_at on the ones answered so a rerun won't answer them again.
                api_by_item = {o["item_number"]: o for o in api_offers}
                acted = respond_to_offers(results, api_by_item, token, settings, act_live)
                results["acted_at"] = results["item_number"].map(acted)

                store_account_results(ebay_conn, results, account, today)
                injected.append(account)
                all_results.append(results)
                counts = ", ".join(f"{n} {a}" for a, n in results["action"].value_counts().items()) or "none"
                log.success(f"[cyan]{account}[/cyan] — recorded {len(results)} offers ({counts}).")
            except Exception:
                log.error(f"Failed on [cyan]{account}[/cyan] — undoing this run's inserts and alerting.")
                ebay_conn.rollback()  # drop the failed account's uncommitted partial insert
                remove_run_results(ebay_conn, injected, today)
                alert_utils.handle_crash(
                    driver, traceback.format_exc(),
                    automation_name=f"eBay Best Offers (failed on {account})",
                )
                raise SystemExit(1)
            finally:
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:
                        pass
    except Exception:
        alert_utils.handle_crash(None, traceback.format_exc(), automation_name="eBay Best Offers")
        raise SystemExit(1)
    finally:
        if ebay_conn is not None:
            ebay_conn.close()

    # Refresh the report workbook and draft the summary email for review. Non-fatal:
    # today's results are already stored, so a workbook or email hiccup must not fail
    # the run or roll back the archive.
    if all_results:
        try:
            draft_summary_email(pd.concat(all_results, ignore_index=True), settings, paths.get("report_wb_path"))
        except Exception:
            log.error("Report/email step failed — today's results are safely stored in SQL.")
            traceback.print_exc()


if __name__ == "__main__":
    if ask_user("Run eBay Best Offers now?", "eBay Best Offers"):
        main()
    run_on_schedule(main, hour=17, minute=30, day_of_week="mon-sun")
