import os
import time
import traceback
import pandas as pd
import xlwings as xl
from rich import print
from dotenv import load_dotenv
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fc_utils import accounts, chrome, database_utils, custom_functions, ebay, outlook, alert_utils
from fc_utils.config_utils import get_env
from fc_utils.schedule_utils import run_on_schedule
from fc_utils.ui_utils import ask_user
from fc_utils.accounts import EBAY_PROFILES
from selenium.common.exceptions import (
    TimeoutException, ElementClickInterceptedException, SessionNotCreatedException
)

_ebay_account_names: list[str] = list(EBAY_PROFILES.keys())

directory: str = os.getcwd()

load_dotenv()
password: str = os.getenv("eBay_pass")
sender_email: str = os.getenv("SENDER_EMAIL", "")
to_email: list[str] = [e.strip() for e in os.getenv("TO_EMAIL", "").split(",") if e.strip()]
cc_email: list[str] = [e.strip() for e in os.getenv("CC_EMAIL", "").split(",") if e.strip()]
table_pending: str = os.getenv("DB_TABLE_PENDING", "PendingOffers")
table_aged: str = os.getenv("DB_TABLE_AGED", "AgedInventory")
user_data_dir: str = get_env("CHROME_USER_DATA_DIR", required=True)

for _t in (table_pending, table_aged):
    if not _t.replace("_", "").isalnum():
        raise ValueError(f"Invalid table name: {_t!r}")

ebay_commission: float = float(os.getenv("EBAY_COMMISSION", "0.091"))
min_discount: float = float(os.getenv("MIN_DISCOUNT", "0.95"))
max_discount: float = float(os.getenv("MAX_DISCOUNT", "0.9"))
min_profit_threshold: float = float(os.getenv("MIN_PROFIT", "0.1"))
brands: list[str] = [b.strip() for b in os.getenv("BRANDS", "").split(",") if b.strip()]

offers_wb_path: str = f"{directory}/eBay/Pending-Offers.xlsm"
aged_inv_path: str = f"{directory}/eBay/Items/eBay Aged Inventory"


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
    print("[cyan][INFO][/cyan] Aged inventory table cleared.")

    dataframes = []
    for sheet in ["Raw data", "Dead", "Slow"]:
        df = pd.read_excel(
            f"{aged_inv_path}/Aged Inventory.xlsm",
            sheet_name=sheet,
            skiprows=4
        )
        dataframes.append(df[["SKU", "Status"]])

    master_df = pd.concat(dataframes, ignore_index=True)
    database_utils.insert_dataframe(cursor, table_aged, master_df, ["SKU", "Status"])
    conn.commit()
    print("[cyan][INFO][/cyan] Aged inventory data inserted.")


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

        print(f"[cyan][INFO][/cyan] Retrieving {quantity} out of {total_qty} offers from eBay.")
        row = 3
        data_from_ebay = []
        getting_rows = True
        while getting_rows:
            try:
                data_row: list[str] = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, fr"#{row_id}\@gridData-\@grid-table > tbody:nth-child({row})"))
                ).text.split("\n")

                while data_row[0].startswith((
                    "Offer", "Highest", "Respond", "Out of stock", "Listing ",
                    "Restock", "Link. ", "View message", "Send offer", "Edit",
                    "Visibility ", "Boost ", "Promote "
                )) or data_row[0].endswith(("received", " buyer", " buyers")):
                    data_row.pop(0)

                if data_row[0].startswith("owner."):
                    del data_row[:4]

                if data_row[1].startswith("Buy It Now"):
                    item_id: str = data_row[1].split(" · ")[-1]
                    data_row.pop(1)
                    data_row.insert(3, item_id)

                data_row.pop(4)
                data_row.pop(4)

                if data_row[4] == "Research prices":
                    data_row.pop(4)

                if data_row[2].startswith("$"):
                    price = data_row[2].replace("$", "").replace(",", "")
                    data_row.pop(2)
                    data_row.insert(2, price)

                if data_row[3] == "Buy It Now":
                    data_row.pop(3)

                data_from_ebay.append(data_row[:4])
                row += 1

            except TimeoutException:
                getting_rows = False

        df = pd.DataFrame(data_from_ebay, columns=["Title", "SKU", "CurrentPrice", "ItemNumber"])
        df.insert(0, "Account", account)
        df.insert(0, "Date", today)
        df["CurrentPrice"] = pd.to_numeric(df["CurrentPrice"], errors="coerce").fillna(0)
        df = df.dropna(subset=["SKU"])
        if df.empty:
            print(f"[yellow][WARNING][/yellow] No valid rows for [cyan]{account}[/cyan]. Skipping.")
            continue

        columns = ["Date", "Account", "Title", "SKU", "CurrentPrice", "ItemNumber"]
        database_utils.insert_dataframe(cursor, table_pending, df, columns)
        conn.commit()

        if quantity != total_qty:
            page += 1
            offset += 25
            print(f"[cyan][INFO][/cyan] Navigating to page #{page}.")
            driver.get(f"https://www.ebay.com/sh/lst/active?status=PENDING_OFFERS&limit=25&offset={offset}")
            driver.switch_to_window(0)


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
        print("[cyan][INFO][/cyan] Offer not received. Moving to next item.")
        return

    brand = offers_sh.range(f"D{first_item}").value.split(" ")[0].title()
    site_cost = float(offers_sh.range(f"H{first_item}").value)
    current_price = float(offers_sh.range(f"G{first_item}").value)
    current_price_lowered = current_price - 0.01
    total_cost = float(offers_sh.range(f"N{first_item}").value)

    price_min = current_price * max_discount
    price_max = current_price * min_discount
    commission_min = price_min * ebay_commission
    commission_max = price_max * ebay_commission
    commission_cx = cx_offer * ebay_commission
    commission_curr = current_price_lowered * ebay_commission

    real_cost_min = total_cost + commission_min
    real_cost_max = total_cost + commission_max
    real_cost_cx = total_cost + commission_cx
    real_cost_curr = total_cost + commission_curr

    profit_min = price_min - real_cost_min
    profit_max = price_max - real_cost_max
    cx_profit = cx_offer - real_cost_cx
    curr_profit = current_price_lowered - real_cost_curr

    profit_min_pct = profit_min / price_min
    profit_max_pct = profit_max / price_max
    cx_profit_pct = cx_profit / cx_offer
    curr_profit_pct = curr_profit / current_price_lowered

    if cx_offer == 0 or site_cost == 0 or site_cost == 0.01 or brand in brands:
        action = "Removed By Buyer"
        offers_sh.range(f"L{first_item}").value = [action, 0]

    elif cx_profit_pct >= 0.11:
        action = "Accepted"
        offers_sh.range(f"L{first_item}").value = [action, 0]
        print(f"[cyan][INFO][/cyan] Offer [cyan]accepted[/cyan] at ${cx_offer} ({round(cx_profit_pct * 100, 2)}% profit).")

    elif profit_min_pct >= 0.11 and profit_max_pct < 0.11:
        counteroffer = round(price_min, 2)
        action = "Counteroffer"
        offers_sh.range(f"L{first_item}").value = [action, counteroffer]
        print(f"[cyan][INFO][/cyan] Offer [cyan]countered[/cyan] at ${counteroffer} ({round(profit_min_pct * 100, 2)}% profit).")

    elif profit_max_pct >= 0.11 and cx_profit_pct < 0.11:
        counteroffer = round(price_max, 2)
        action = "Counteroffer"
        offers_sh.range(f"L{first_item}").value = [action, counteroffer]
        print(f"[cyan][INFO][/cyan] Offer [cyan]countered[/cyan] at ${counteroffer} ({round(profit_max_pct * 100, 2)}% profit).")

    else:
        counteroffer = round(current_price_lowered, 2)
        action = "Counteroffer"
        offers_sh.range(f"L{first_item}").value = [action, counteroffer]
        print(f"[cyan][INFO][/cyan] Offer [cyan]countered[/cyan] at ${counteroffer} (price lowered by $0.01, {round(curr_profit_pct * 100, 2)}% profit).")

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

        try:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ui-component-button:nth-child(1)"))).click()
        except TimeoutException:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[4]/button[1]"))).click()

        try:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ui-component-button:nth-child(1)"))).click()
        except TimeoutException:
            try:
                WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[3]/button[1]"))).click()
            except TimeoutException:
                WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[4]/button[1]"))).click()

        try:
            alert_status: str = WebDriverWait(driver, 15).until(EC.presence_of_element_located((
                By.XPATH, "/html/body/div[5]/div[4]/div[2]/div/div[2]/div/div/div/div[2]/div/div[2]"
            ))).text
            print(f"[bold red][ERROR][/bold red] [cyan]{alert_status}[/cyan]. Moving to next item.")
            offers_sh.range(f"J{first_item}").value = 0
        except TimeoutException:
            pass

    else:
        print(f"[cyan][INFO][/cyan] Item [cyan]{action}[/cyan]. Moving to next item.")


def main() -> None:
    """Download pending offers for each eBay account, process them, and email the summary.

    Uploads aged inventory if updated today, downloads new pending offers per
    account, refreshes the workbook, attends each pending offer (accept or
    counteroffer), saves the workbook, and emails the results summary.
    """
    driver = None
    try:
        today: str = datetime.now().strftime("%Y-%m-%d")
        date_str: str = datetime.now().strftime("%m/%d/%Y")

        info = custom_functions.files_info(aged_inv_path)
        for file in info:
            file_date = str(file["Date Modified"]).split(" ")[0]

        if file_date == today:
            print("[cyan][INFO][/cyan] Uploading [cyan]Aged Inventory[/cyan] items to database.")
            aged_inventory()

        conn = custom_functions.sql_connection("eBay")
        cursor = conn.cursor()

        # Download pending offers for each account
        for account, profile in EBAY_PROFILES.items():
            if last_update(cursor, account) == today:
                print(f"[cyan][INFO][/cyan] Offers already retrieved today for [cyan]{account}[/cyan]. Skipping.")
                continue

            driver = chrome.start_browser(user_data_dir, profile, headless=True)

            driver.get("https://www.ebay.com/sh/ovw")
            driver.switch_to_window(0)

            try:
                accounts.ebay(password=password, driver=driver)
            except TimeoutException:
                pass

            print(f"[cyan][INFO][/cyan] Navigating to pending offers for [cyan]{account}[/cyan].")
            driver.get("https://www.ebay.com/sh/lst/active?status=PENDING_OFFERS&limit=25")
            driver.switch_to_window(0)

            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.fake-link:nth-child(3)"))).click()
            except (TimeoutException, ElementClickInterceptedException):
                pass

            ebay.customize_offers_table(driver)

            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".zeroResultsMessage")))
                print(f"[cyan][INFO][/cyan] No pending offers for [cyan]{account}[/cyan].")
            except TimeoutException:
                get_pending_offers(driver, account, today, conn, cursor)

            driver.quit()

        # Open workbook, refresh, and sort
        print("[cyan][INFO][/cyan] Opening [cyan]Pending Offers[/cyan] workbook.")
        offers_wb = xl.Book(offers_wb_path)
        offers_sh = offers_wb.sheets("Pending Offers")
        refresh_all = offers_wb.macro("Module1.RefreshAll")
        sort_all = offers_wb.macro("Module1.SortAll")
        reorganize = offers_wb.macro("Module1.Reorganize")

        print("[cyan][INFO][/cyan] Refreshing all queries.")
        refresh_all()
        time.sleep(5)
        offers_wb.save()

        sort_all()
        time.sleep(5)

        # Attend pending offers for each account
        for account, profile in EBAY_PROFILES.items():
            first_item: int = custom_functions.first_empty_row(offers_sh, "J", "B11")
            last_item: int = int(offers_sh.range(f"B{offers_sh.cells.last_cell.row}").end("up").row)

            if first_item > last_item:
                print("[cyan][INFO][/cyan] No new pending offers.")
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

            print(f"[cyan][INFO][/cyan] Navigating to [cyan]{account}[/cyan] profile on eBay.")
            driver.get("https://www.ebay.com/")
            time.sleep(2)
            driver.switch_to_window(0)

            for item in raw_data:
                if item[0] != account:
                    print(f"[cyan][INFO][/cyan] {account} offers done. Moving to [cyan]{item[0]}[/cyan].")
                    driver.quit()
                    break

                print(f"[cyan][INFO][/cyan] Navigating to item {item[3]}.")
                driver.get(f"https://www.ebay.com/bo/seller/showOffers?itemid={item[3]}")
                driver.switch_to_window(0)

                attend_pending_offers(driver, offers_sh, first_item)
                first_item += 1

        # Build email summary from workbook cells
        current_hour = datetime.now().hour
        if 5 <= current_hour <= 11:
            greeting = "Good morning"
        elif 12 <= current_hour <= 17:
            greeting = "Good afternoon"
        else:
            greeting = "Good evening"

        fc_accepted = int(offers_sh.range("C4").value)
        fc_counter = int(offers_sh.range("C5").value)
        fc_expired = int(offers_sh.range("C6").value)
        fc_declined = int(offers_sh.range("C7").value)
        fc_removed = int(offers_sh.range("C8").value)
        fc_total = int(offers_sh.range("C9").value)
        ls_accepted = int(offers_sh.range("E4").value)
        ls_counter = int(offers_sh.range("E5").value)
        ls_expired = int(offers_sh.range("E6").value)
        ls_declined = int(offers_sh.range("E7").value)
        ls_removed = int(offers_sh.range("E8").value)
        ls_total = int(offers_sh.range("E9").value)

        body = f"""{greeting},
        <p>I hope you are doing well.</p>
        <p>Please find attached the eBay best offers cleared today for all accounts.</p>
        <p>Also, please find below a summary:</p>
        <ul>
        <p><b><u>{_ebay_account_names[0]}</u></b></p>
        <li><b>Accepted: </b> {fc_accepted}</li>
        <li><b>Counteroffer: </b> {fc_counter}</li>
        <li><b>Declined: </b> {fc_declined}</li>
        <li><b>Removed by Buyer: </b> {fc_removed}</li>
        <li><b>Expired: </b> {fc_expired}</li>
        <li><b>Total: </b> {fc_total}</li>
        </ul>
        <ul>
        <p><b><u>{_ebay_account_names[1]}</u></b></p>
        <li><b>Accepted: </b> {ls_accepted}</li>
        <li><b>Counteroffer: </b> {ls_counter}</li>
        <li><b>Declined: </b> {ls_declined}</li>
        <li><b>Removed by Buyer: </b> {ls_removed}</li>
        <li><b>Expired: </b> {ls_expired}</li>
        <li><b>Total: </b> {ls_total}</li>
        </ul>
        <p>Thank you.</p>
        Sincerely,
        """

        print("[cyan][INFO][/cyan] Saving and closing workbook.")
        reorganize()
        offers_wb.save()
        time.sleep(2)
        offers_wb.close()

        outlook.send_email(
            account=sender_email,
            subject=f"eBay Report - Best Offers - {date_str}",
            body=body,
            to=to_email,
            cc=cc_email,
            attachments=[offers_wb_path],
            show=True,
            send=True,
        )
        print("[cyan][INFO][/cyan] Email sent.")

    except (KeyboardInterrupt, SystemExit):
        print("[yellow][WARNING][/yellow] Script interrupted by user.")
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


if ask_user("Run now?", "eBay Pending Offers"):
    main()
run_on_schedule(main, hour=17, minute=30, day_of_week="mon-sun")
