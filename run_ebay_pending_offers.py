"""Daily eBay Best Offer automation.

Once per day at 17:30 local time, this script:

1. Logs into each configured eBay seller account (`EBAY_PROFILES`).
2. Scrapes the pending-offers list, cleans each row, and inserts the result
   into a SQL Server `PendingOffers` table.
3. Opens the `Pending-Offers.xlsm` workbook *hidden* and triggers the VBA
   `RefreshAll` macro so Power Query syncs the new rows into the worksheet.
4. Walks every pending offer and decides Accept / Counter / Removed using a
   profit-margin rule (see `_decide_offer`); for counteroffers, it drives
   the browser to submit the response.
5. Saves the workbook and emails a per-account summary via Outlook.
"""
import time
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xl
from dotenv import load_dotenv
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from fc_utils import accounts, alert_utils, chrome, custom_functions, database_utils, ebay, file_utils, greeting_for, outlook
from fc_utils.accounts import EBAY_PROFILES
from fc_utils.config_utils import get_env, load_config_safe
from fc_utils.logging_utils import setup_logging
from fc_utils.schedule_utils import run_on_schedule
from fc_utils.ui_utils import ask_user


_ebay_account_names: list[str] = list(EBAY_PROFILES.keys())

log = setup_logging("ebay_pending_offers")
load_dotenv()
password: str = get_env("eBay_pass", required=True)
user_data_dir: str = get_env("CHROME_USER_DATA_DIR", required=True)
sender_email: str = get_env("SENDER_EMAIL", required=True)
to_email: list[str] = [e.strip() for e in (get_env("TO_EMAIL", required=True) or "").split(",") if e.strip()]
cc_email: list[str] = [e.strip() for e in (get_env("CC_EMAIL", default="") or "").split(",") if e.strip()]
table_pending: str = get_env("DB_TABLE_PENDING", default="PendingOffers") or "PendingOffers"
table_aged: str = get_env("DB_TABLE_AGED", default="AgedInventory") or "AgedInventory"
for _t in (table_pending, table_aged):
    if not _t.replace("_", "").isalnum():
        raise ValueError(f"Invalid table name: {_t!r}")

ebay_commission: float = float(get_env("EBAY_COMMISSION", default="0.091") or "0.091")
min_discount: float = float(get_env("MIN_DISCOUNT", default="0.95") or "0.95")
max_discount: float = float(get_env("MAX_DISCOUNT", default="0.9") or "0.9")
min_profit_threshold: float = float(get_env("MIN_PROFIT", default="0.1") or "0.1")
brands: list[str] = [b.strip() for b in (get_env("BRANDS", default="") or "").split(",") if b.strip()]

_paths = load_config_safe(Path(__file__).resolve().parent / "config" / "paths.json")
offers_wb_path: str = _paths["offers_wb_path"]
aged_inv_path: str = _paths["aged_inv_path"]


def _email_body(account_blocks: list[str]) -> str:
    """Build the per-day summary email body.

    Wraps the time-of-day greeting (`greeting_for()`) around the workbook-
    derived per-account stat blocks. Called at email-send time so the greeting
    reflects the actual send hour rather than the script-import hour.

    Args:
        account_blocks: HTML `<ul>` blocks (one per included account) built
            from `_read_account_stats(...)` reads.

    Returns:
        Full HTML body ready to pass to `outlook.send_email(body=...)`.
    """
    return f"""{greeting_for()},
        <p>I hope you are doing well.</p>
        <p>Please find attached the eBay best offers cleared today for all accounts.</p>
        <p>Also, please find below a summary:</p>
        {account_blocks[0]}
        {account_blocks[1]}
        <p>Thank you.</p>
        Sincerely,
        """


# --- eBay row parser ---------------------------------------------------------
# Each row in eBay's Pending Offers grid is rendered as a <tbody> whose
# textContent — when split on newlines — contains the real data fields
# interleaved with screen-reader labels, action buttons, and promo links.
# We strip those out and normalize the price so each row collapses to the
# canonical 4-tuple [Title, SKU, CurrentPrice, ItemNumber] before insert.

# Prefixes of lines that should be dropped from the start of a row. Each one
# corresponds to a button, badge, or screen-reader label that eBay renders
# inline with the data cells.
JUNK_ROW_PREFIXES: tuple[str, ...] = (
    "Offer",          # offer-count badge ("Offer received")
    "Highest",        # "Highest offer" label
    "Respond",        # respond CTA
    "Out of stock",   # inventory status badge
    "Listing ",       # listing-status badge (trailing space is intentional)
    "Restock",        # restock CTA
    "Link. ",         # screen-reader prefix for embedded links
    "View message",   # messaging CTA
    "Send offer",     # send-offer CTA
    "Edit",           # edit-listing button
    "Visibility ",    # visibility badge (trailing space)
    "Boost ",         # boost-listing CTA (trailing space)
    "Promote ",       # promoted-listing CTA (trailing space)
)

# Suffixes of lines that should also be dropped from the start (offer-count
# annotations like "3 buyers" / "1 buyer" / "X offers received").
JUNK_ROW_SUFFIXES: tuple[str, ...] = ("received", " buyer", " buyers")


def _parse_offer_row(raw_lines: list[str]) -> list[str] | None:
    """Clean one raw eBay row into ``[Title, SKU, CurrentPrice, ItemNumber]``.

    The seller-hub list interleaves data cells with UI noise; this function
    trims that noise, extracts the item id where it's embedded in the price
    cell, and normalizes the price formatting. Returns ``None`` when the row
    is too short to parse — letting the caller skip it instead of crashing.

    Args:
        raw_lines (list[str]): Lines of one row's textContent split on "\\n".

    Returns:
        list[str] | None: ``[title, sku, current_price, item_number]`` or
            ``None`` if the row cannot be parsed.
    """
    row = list(raw_lines)  # copy so we don't mutate the caller's list

    # 1) Strip noise from the front while data_row[0] matches.
    while row and (
        row[0].startswith(JUNK_ROW_PREFIXES) or row[0].endswith(JUNK_ROW_SUFFIXES)
    ):
        row.pop(0)
    if not row:
        return None

    # 2) Some rows start with a screen-reader "owner." label and bury the
    #    real fields four lines deeper.
    if row[0].startswith("owner."):
        del row[:4]
    if len(row) < 2:
        return None

    # 3) The format cell sometimes reads "Buy It Now · {item_id}" — pull the
    #    id out and place it where ItemNumber belongs (index 3).
    if row[1].startswith("Buy It Now"):
        item_id = row[1].split(" · ")[-1]
        row.pop(1)
        row.insert(3, item_id)
    if len(row) < 5:
        return None

    # 4) Drop two consecutive UI labels at position 4 (after each pop, the
    #    next label slides into position 4).
    row.pop(4)
    row.pop(4)

    # 5) The "Research prices" link sometimes lingers at position 4.
    if len(row) > 4 and row[4] == "Research prices":
        row.pop(4)

    # 6) Normalize the current-price string: "$1,299.00" -> "1299.00".
    if len(row) > 2 and row[2].startswith("$"):
        row[2] = row[2].replace("$", "").replace(",", "")

    # 7) Some rows duplicate the "Buy It Now" label at position 3.
    if len(row) > 3 and row[3] == "Buy It Now":
        row.pop(3)

    if len(row) < 4:
        return None
    return row[:4]


def _read_account_stats(sheet: object, column: str) -> dict[str, int]:
    """Read the 6 stats (rows 4-9) under one account's column.

    Args:
        sheet (object): xlwings sheet for the Pending Offers tab.
        column (str): Column letter holding the account's counts (e.g. "C").

    Returns:
        dict[str, int]: ``{accepted, counter, expired, declined, removed, total}``.
    """
    accepted, counter, expired, declined, removed, total = (
        int(sheet.range(f"{column}{row}").value) for row in range(4, 10)
    )
    return {
        "accepted": accepted,
        "counter": counter,
        "expired": expired,
        "declined": declined,
        "removed": removed,
        "total": total,
    }


def _click_with_fallbacks(
    driver: object,
    locators: list[tuple[str, str]],
    timeout: int = 15,
) -> None:
    """Click the first element among ``locators`` that becomes clickable.

    eBay's best-offer modal occasionally shifts its action-button position
    by one container depth depending on whether an info banner is rendered.
    Callers list a primary CSS selector followed by positional XPath
    fallbacks that target the same logical button at alternate depths; the
    first locator to resolve wins.

    Args:
        driver (object): Active SeleniumBase WebDriver instance.
        locators (list[tuple[str, str]]): ``(By, selector)`` pairs to try
            in order. The first one whose element becomes clickable within
            ``timeout`` seconds is clicked.
        timeout (int): Per-locator wait, in seconds. Defaults to 15.

    Raises:
        TimeoutException: If none of the locators yield a clickable element.
    """
    for by, selector in locators:
        try:
            WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, selector))
            ).click()
            return
        except TimeoutException:
            continue
    raise TimeoutException(
        f"None of {len(locators)} dialog-button locators became clickable within {timeout}s"
    )


def last_update(cursor: object, account: str) -> str | None:
    """Return the most recent date offers were downloaded for the given account.

    Args:
        cursor (object): Active pyodbc cursor connected to the eBay database.
        account (str): eBay account name to check.

    Returns:
        str | None: The latest date string, or None if no records exist.
    """
    cursor.execute(
        "SELECT MAX(Date) AS MaxDate FROM PendingOffers WHERE Account = ?",
        (account,)
    )
    return cursor.fetchone()[0]


def aged_inventory() -> None:
    """Upload the Aged Inventory Excel file into the database.

    Reads the Raw data, Dead, and Slow sheets, concatenates them, clears
    the aged inventory table, and inserts the combined SKU/Status rows.
    """
    conn = custom_functions.sql_connection("Reports")
    cursor = conn.cursor()

    cursor.execute(f"DELETE FROM {table_aged}")
    log.info("Aged inventory table cleared.")

    dataframes = []
    for sheet in ["Raw data", "Dead", "Slow"]:
        df = pd.read_excel(
            f"{aged_inv_path}/Aged Inventory.xlsx",
            sheet_name=sheet,
            skiprows=4
        )
        dataframes.append(df[["SKU", "Status"]])

    master_df = pd.concat(dataframes, ignore_index=True)
    database_utils.insert_dataframe(cursor, table_aged, master_df, ["SKU", "Status"])
    conn.commit()
    log.info("Aged inventory data inserted.")


def get_pending_offers(driver: object, account: str, today: str, conn: object, cursor: object) -> None:
    """Scrape all pending offers for one account and insert them into the database.

    Paginates through the pending offers list, cleans each scraped row,
    builds a DataFrame, and inserts it into the pending offers table.

    Args:
        driver (object): Active SeleniumBase WebDriver instance.
        account (str): eBay account display name.
        today (str): Today's date string in YYYY-MM-DD format.
        conn (object): Active pyodbc connection.
        cursor (object): Active pyodbc cursor.
    """
    page = 1
    offset = 0
    quantity = 1
    total_qty = 0

    while quantity != total_qty:
        try:
            result_range: list[str] = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.item:nth-child(1)"))
            ).text.split(" ")
        except TimeoutException:
            result_range = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".result-range"))
            ).text.split(" ")

        total_qty = int(result_range[2])
        quantity = int(result_range[0].split("-")[1])

        parent = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".shui-dt")))
        children = parent.find_elements(By.CSS_SELECTOR, "[id]")
        for child in children:
            row_id: str = child.get_attribute("id").split("@")[0]
            break

        log.info(f"Retrieving {quantity} out of {total_qty} offers from eBay.")
        row = 3
        data_from_ebay = []
        getting_rows = True
        while getting_rows:
            try:
                data_row: list[str] = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, fr"#{row_id}\@gridData-\@grid-table > tbody:nth-child({row})"))
                ).text.split("\n")
            except TimeoutException:
                getting_rows = False
                continue

            parsed = _parse_offer_row(data_row)
            if parsed is not None:
                data_from_ebay.append(parsed)
            row += 1

        df = pd.DataFrame(data_from_ebay, columns=["Title", "SKU", "CurrentPrice", "ItemNumber"])
        df.insert(0, "Account", account)
        df.insert(0, "Date", today)
        df["CurrentPrice"] = pd.to_numeric(df["CurrentPrice"], errors="coerce").fillna(0)
        df = df.dropna(subset=["SKU"])
        if df.empty:
            log.warning(f"No valid rows for [cyan]{account}[/cyan]. Skipping.")
            continue

        columns = ["Date", "Account", "Title", "SKU", "CurrentPrice", "ItemNumber"]
        database_utils.insert_dataframe(cursor, table_pending, df, columns)
        conn.commit()

        if quantity != total_qty:
            page += 1
            offset += 25
            log.info(f"Navigating to page #{page}.")
            driver.get(f"https://www.ebay.com/sh/lst/active?status=PENDING_OFFERS&limit=25&offset={offset}")
            driver.switch_to_window(0)


def _decide_offer(
    cx_offer: float,
    site_cost: float,
    current_price: float,
    total_cost: float,
    brand: str,
    blocked_brands: list[str],
    commission: float,
    floor_factor: float,
    ceiling_factor: float,
) -> tuple[str, float, float]:
    """Decide the response to one pending offer.

    Args:
        cx_offer: The buyer's offered price.
        site_cost: Site cost of the SKU (0 or 0.01 are sentinels for "missing").
        current_price: The listing's current price.
        total_cost: Total landed cost (site cost + shipping + ...).
        brand: First word of the item title, title-cased.
        blocked_brands: Brands we refuse to counter on.
        commission: eBay's fee rate as a fraction (e.g. 0.091).
        floor_factor: Floor multiplier of current_price (e.g. 0.9).
        ceiling_factor: Ceiling multiplier of current_price (e.g. 0.95).

    Returns:
        (action, counteroffer_amount, profit_pct_for_log).
        action is one of "Removed By Buyer", "Accepted", "Counteroffer".
        counteroffer_amount is 0.0 unless action == "Counteroffer".
    """
    if cx_offer == 0 or site_cost == 0 or site_cost == 0.01 or brand in blocked_brands:
        return ("Removed By Buyer", 0.0, 0.0)

    cx_profit_pct = (cx_offer - (total_cost + cx_offer * commission)) / cx_offer
    if cx_profit_pct >= 0.11:
        return ("Accepted", 0.0, cx_profit_pct)

    price_min = current_price * floor_factor
    price_max = current_price * ceiling_factor
    profit_min_pct = (price_min - (total_cost + price_min * commission)) / price_min
    profit_max_pct = (price_max - (total_cost + price_max * commission)) / price_max

    if profit_min_pct >= 0.11 and profit_max_pct < 0.11:
        return ("Counteroffer", round(price_min, 2), profit_min_pct)
    if profit_max_pct >= 0.11 and cx_profit_pct < 0.11:
        return ("Counteroffer", round(price_max, 2), profit_max_pct)

    list_lowered = current_price - 0.01
    curr_profit_pct = (list_lowered - (total_cost + list_lowered * commission)) / list_lowered
    return ("Counteroffer", round(list_lowered, 2), curr_profit_pct)


def attend_pending_offers(driver: object, offers_sh: object, first_item: int) -> None:
    """Read one pending offer row, calculate the response, and act on it in the browser.

    Determines whether to accept, counteroffer, or mark as removed based on
    profit margin calculations using the configured commission and discount rates.

    Args:
        driver (object): Active SeleniumBase WebDriver instance.
        offers_sh (object): xlwings sheet for the Pending Offers tab.
        first_item (int): Row number of the current offer in the workbook.
    """
    try:
        cx_offer: float = round(float(WebDriverWait(driver, 15).until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            ".ui-component-offer-details > dl:nth-child(1) > div:nth-child(1) > dd:nth-child(2)"
        ))).text.split(" ")[0].replace("$", "").replace(",", "")), 2)
    except TimeoutException:
        cx_offer = 0

    offers_sh.range(f"J{first_item}").value = cx_offer

    if cx_offer == 0:
        log.info("Offer not received. Moving to next item.")
        return

    brand = offers_sh.range(f"D{first_item}").value.split(" ")[0].title()
    site_cost = float(offers_sh.range(f"H{first_item}").value)
    current_price = float(offers_sh.range(f"G{first_item}").value)
    total_cost = float(offers_sh.range(f"N{first_item}").value)

    action, counteroffer, log_pct = _decide_offer(
        cx_offer=cx_offer,
        site_cost=site_cost,
        current_price=current_price,
        total_cost=total_cost,
        brand=brand,
        blocked_brands=brands,
        commission=ebay_commission,
        floor_factor=max_discount,
        ceiling_factor=min_discount,
    )

    if action == "Removed By Buyer":
        offers_sh.range(f"L{first_item}").value = [action, 0]
    elif action == "Accepted":
        offers_sh.range(f"L{first_item}").value = [action, 0]
        log.info(f"Offer [cyan]accepted[/cyan] at ${cx_offer} ({round(log_pct * 100, 2)}% profit).")
    else:  # Counteroffer
        offers_sh.range(f"L{first_item}").value = [action, counteroffer]
        if counteroffer == round(current_price - 0.01, 2):
            log.info(f"Offer [cyan]countered[/cyan] at ${counteroffer} (price lowered by $0.01, {round(log_pct * 100, 2)}% profit).")
        else:
            log.info(f"Offer [cyan]countered[/cyan] at ${counteroffer} ({round(log_pct * 100, 2)}% profit).")

    if action == "Accepted":
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ui-component-button:nth-child(1)"))).click()
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ui-component-button:nth-child(1)"))).click()

    elif action == "Counteroffer":
        counteroffer_msg = f"""
        Thanks for your offer! We send our absolute final price of ${counteroffer}.

        This is unbeatable value! You get fast Free Shipping, worry Free Returns, from a trusted seller. We cannot go lower on this item.

        Accept our counteroffer now to secure your deal!
        """

        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ui-component-button:nth-child(2)"))).click()
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#bestoffer__priceInput"))).send_keys(str(counteroffer))
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#bestoffer__messageTextarea"))).send_keys(counteroffer_msg)

        # Click "Send counteroffer" in the price-input dialog.
        _click_with_fallbacks(driver, [
            (By.CSS_SELECTOR, "button.ui-component-button:nth-child(1)"),
            (By.XPATH, "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[4]/button[1]"),
        ])

        # Confirm the counteroffer in the follow-up dialog. The "Confirm"
        # button can sit at one of two container depths depending on whether
        # eBay also renders a fee/notice banner in the modal, so try the
        # primary CSS selector first then two positional XPaths.
        _click_with_fallbacks(driver, [
            (By.CSS_SELECTOR, "button.ui-component-button:nth-child(1)"),
            (By.XPATH, "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[3]/button[1]"),
            (By.XPATH, "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[4]/button[1]"),
        ])

        # If eBay rejects the counteroffer (e.g. below allowed minimum) it
        # renders an error banner in the same modal. This XPath reads that
        # banner's text; if it isn't present the click succeeded.
        try:
            alert_status: str = WebDriverWait(driver, 15).until(EC.presence_of_element_located((
                By.XPATH, "/html/body/div[5]/div[4]/div[2]/div/div[2]/div/div/div/div[2]/div/div[2]"
            ))).text
            log.error(f"[cyan]{alert_status}[/cyan]. Moving to next item.")
            offers_sh.range(f"J{first_item}").value = 0
        except TimeoutException:
            pass

    else:
        log.info(f"Item [cyan]{action}[/cyan]. Moving to next item.")


def main() -> None:
    """Download pending offers for each eBay account, process them, and email the summary.

    Uploads aged inventory if updated today, downloads new pending offers per
    account, refreshes the workbook, attends each pending offer (accept or
    counteroffer), saves the workbook, and emails the results summary.
    """
    driver = None
    xl_app = None
    try:
        today: str = datetime.now().strftime("%Y-%m-%d")
        date_str: str = datetime.now().strftime("%m/%d/%Y")

        last_modified = file_utils.latest_modified_date(aged_inv_path)

        if last_modified is not None and last_modified.strftime("%Y-%m-%d") == today:
            log.info("Uploading [cyan]Aged Inventory[/cyan] items to database.")
            aged_inventory()

        conn = custom_functions.sql_connection("eBay")
        cursor = conn.cursor()

        # Download pending offers for each account
        for account, profile in EBAY_PROFILES.items():
            if last_update(cursor, account) == today:
                log.info(f"Offers already retrieved today for [cyan]{account}[/cyan]. Skipping.")
                continue

            driver = chrome.start_browser(user_data_dir, profile, headless=True)

            driver.get("https://www.ebay.com/sh/ovw")
            driver.switch_to_window(0)

            try:
                accounts.ebay(password=password, driver=driver)
            except TimeoutException:
                pass

            log.info(f"Navigating to pending offers for [cyan]{account}[/cyan].")
            driver.get("https://www.ebay.com/sh/lst/active?status=PENDING_OFFERS&limit=25")
            driver.switch_to_window(0)

            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.fake-link:nth-child(3)"))).click()
            except (TimeoutException, ElementClickInterceptedException):
                pass

            ebay.customize_offers_table(driver)

            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".zeroResultsMessage")))
                log.info(f"No pending offers for [cyan]{account}[/cyan].")
            except TimeoutException:
                get_pending_offers(driver, account, today, conn, cursor)

            driver.quit()

        # Open workbook hidden, refresh, and sort. Excel runs invisible so it
        # never steals focus during the scheduled run.
        log.info("Opening [cyan]Pending Offers[/cyan] workbook (hidden).")
        xl_app = xl.App(visible=False, add_book=False)
        xl_app.display_alerts = False
        offers_wb = xl_app.books.open(offers_wb_path)
        offers_sh = offers_wb.sheets("Pending Offers")

        # modUtilities.refresh + sortCols + reorder now run synchronously
        # (every Power Query connection is forced to BackgroundQuery=False
        # inside VBA), so the previous `time.sleep(5)` waits are no longer
        # needed.
        log.info("Refreshing all queries.")
        offers_wb.macro("modUtilities.refresh")()
        offers_wb.save()

        offers_wb.macro("modUtilities.sortCols")()

        # Attend pending offers for each account
        for account, profile in EBAY_PROFILES.items():
            first_item: int = custom_functions.first_empty_row(offers_sh, "J", "B11")
            last_item: int = int(offers_sh.range(f"B{offers_sh.cells.last_cell.row}").end("up").row)

            if first_item > last_item:
                log.info("No new pending offers.")
                raw_data = None
            elif first_item == last_item:
                raw_data = [offers_sh.range(f"C{first_item}:F{last_item}").value]
            else:
                raw_data = offers_sh.range(f"C{first_item}:F{last_item}").value

            if raw_data is None:
                continue

            curr_account = raw_data[0][0]
            if curr_account != account:
                continue

            driver = chrome.start_browser(user_data_dir, profile, headless=True)

            log.info(f"Navigating to [cyan]{account}[/cyan] profile on eBay.")
            driver.get("https://www.ebay.com/")
            time.sleep(2)
            driver.switch_to_window(0)

            for item in raw_data:
                if item[0] != account:
                    log.info(f"{account} offers done. Moving to [cyan]{item[0]}[/cyan].")
                    driver.quit()
                    break

                log.info(f"Navigating to item {item[3]}.")
                driver.get(f"https://www.ebay.com/bo/seller/showOffers?itemid={item[3]}")
                driver.switch_to_window(0)

                attend_pending_offers(driver, offers_sh, first_item)
                first_item += 1

        # Email summary covers only the first two EBAY_PROFILES entries
        # (workbook columns C and E). Additional accounts are intentionally
        # excluded.
        account_blocks = []
        for name, column in zip(_ebay_account_names[:2], ("C", "E")):
            s = _read_account_stats(offers_sh, column)
            account_blocks.append(
                "<ul>\n"
                f"        <p><b><u>{name}</u></b></p>\n"
                f"        <li><b>Accepted: </b> {s['accepted']}</li>\n"
                f"        <li><b>Counteroffer: </b> {s['counter']}</li>\n"
                f"        <li><b>Declined: </b> {s['declined']}</li>\n"
                f"        <li><b>Removed by Buyer: </b> {s['removed']}</li>\n"
                f"        <li><b>Expired: </b> {s['expired']}</li>\n"
                f"        <li><b>Total: </b> {s['total']}</li>\n"
                "        </ul>"
            )

        log.info("Saving and closing workbook.")
        offers_wb.macro("modUtilities.reorder")()
        offers_wb.save()
        time.sleep(2)
        offers_wb.close()
        xl_app.quit()
        xl_app = None

        outlook.send_email(
            account=sender_email,
            subject=f"eBay Report - Best Offers - {date_str}",
            body=_email_body(account_blocks),
            to=to_email,
            cc=cc_email,
            attachments=[offers_wb_path],
            show=True,
            send=True,
        )
        log.info("Email sent.")

    except (KeyboardInterrupt, SystemExit):
        log.warning("Script interrupted by user.")
        raise SystemExit(0)

    except Exception:
        alert_utils.handle_crash(driver, traceback.format_exc(), "eBay Pending Offers")
        raise SystemExit(1)

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        custom_functions.kill_app("chrome")
        if xl_app is not None:
            try:
                xl_app.quit()
            except Exception:
                pass


if ask_user("Run now?", "eBay Pending Offers"):
    main()
run_on_schedule(main, hour=17, minute=30, day_of_week="mon-sun")
