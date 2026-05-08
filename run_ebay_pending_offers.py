import os
import time
import ctypes
import traceback
import pandas as pd
import xlwings as xl
from rich import print
from dotenv import load_dotenv
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from fc_utils import accounts, chrome, database_utils, custom_functions, ebay, outlook
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, ElementClickInterceptedException, SessionNotCreatedException
)

###############################################################################################################################################
#Get the user and working directory
directory: str = os.getcwd()
win_user: str = os.getlogin()

# Load environment credentials
load_dotenv()
password: str = os.getenv("eBay_pass")

sender_email: str = os.getenv("SENDER_EMAIL", "")
to_email: list[str] = [e.strip() for e in os.getenv("TO_EMAIL", "").split(",") if e.strip()]
cc_email: list[str] = [e.strip() for e in os.getenv("CC_EMAIL", "").split(",") if e.strip()]
table_pending: str = os.getenv("DB_TABLE_PENDING", "PendingOffers")
table_aged: str = os.getenv("DB_TABLE_AGED", "AgedInventory")
for _t in (table_pending, table_aged):
    if not _t.replace("_", "").isalnum():
        raise ValueError(f"Invalid table name: {_t!r}")

#Create list of accounts
Accounts: list[str] = ["SellerOrg", 
                        "SellerOrg3", 
                        "Account4",
                        "SellerOrg2"]

#Specify the path of files to be used
OffersWbStr: str = f"{directory}/eBay/Pending-Offers.xlsm"
AgedInventoryPath: str = f"{directory}/eBay/Items/eBay Aged Inventory"

#Set Chrome User Data Directory
user_data_dir: str = f"C:/ChromeAutomationProfile"

###############################################################################################################################################
#Set eBay default values
eBayCommission: float = float(os.getenv("EBAY_COMMISSION", "0.091"))
MinDiscount: float = float(os.getenv("MIN_DISCOUNT", "0.95"))
MaxDiscount: float = float(os.getenv("MAX_DISCOUNT", "0.9"))
MinProfit: float = float(os.getenv("MIN_PROFIT", "0.1"))

#Brands list
Brands: list[str] = [b.strip() for b in os.getenv("BRANDS", "").split(",") if b.strip()]

###############################################################################################################################################
def seconds_until_target(TargetTime: str):
    #Calculate the number of seconds until the target time
    now = datetime.now()
    TargetTime = datetime.strptime(TargetTime, "%H:%M:%S").replace(year=now.year, month=now.month, day=now.day)

    if TargetTime < now:
        TargetTime += timedelta(days=1)

    return (TargetTime - now).total_seconds()

###############################################################################################################################################
def LastUpdate() -> str:
    """
    Confirm if the pending offers have been downloaded today.\n
    Returns the latest date in the table database for the current account.

    returns:
        The latest date.
    """
    #Get the most recent date from the "Date" column in the database
    cursor.execute(
        """
        SELECT MAX(Date) AS MaxDate
        FROM PendingOffers
        WHERE Account = ?
        """,
        (Account,)
    )

    max_date = cursor.fetchone()[0]

    return max_date

###############################################################################################################################################
def AgedInventory() -> None:
    """
    Automated the process of uploading the Aged Inventory file to the database.
    """
    #Create SQL Database connection
    conn = custom_functions.SQLConnection("Reports")
    cursor = conn.cursor()

    cursor.execute(f"DELETE FROM {table_aged}")
    print("[cyan][INFO][/cyan] Table rows deleted successfully.")

    dataframes = []
    for sheet in ["Raw data", "Dead", "Slow"]:
        df = pd.read_excel(
            f"{AgedInventoryPath}/Aged Inventory.xlsm",
            sheet_name=sheet,
            skiprows=4
        )

        # Select only the SKU and Status columns
        FilteredRawData = df[["SKU", "Status"]]

        dataframes.append(FilteredRawData)

    MasterDF = pd.concat(dataframes, ignore_index=True)

    #Insert data into the SQL Database
    database_utils.insert_dataframe(cursor, table_aged, MasterDF, ["SKU", "Status"])

    # Commit the transaction
    conn.commit()

    print("[cyan][INFO][/cyan] Data inserted successfully.")

###############################################################################################################################################
def GetPendingOffers() -> None:
    """
    Automates the process of retrieving all Pending Offers.
    """
    page = 1
    offset = 0
    Quantity = 1
    TotalQty = 0
    while Quantity != TotalQty:
        #Retrieve how many offers are being shown
        try:
            result_range: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.item:nth-child(1)"))).text.split(" ")

        except TimeoutException:
            result_range: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".result-range"))).text.split(" ")

        TotalQty = int(result_range[2])
        Quantity = int(result_range[0].split("-")[1])

        #Get the main ID of the table offers
        parent = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".shui-dt")))
        children = parent.find_elements(By.CSS_SELECTOR, "[id]")

        for child in children:
            id: str = child.get_attribute("id").split("@")[0]
            break

        #Find table and retrieve its information
        print(f"[cyan][INFO][/cyan] Retrieving {Quantity} out of {TotalQty} offers from eBay.")
        row = 3
        DataFromEbay = []
        getting_rows = True
        while getting_rows:
            try:
                DataRow: list[str] = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, fr"#{id}\@gridData-\@grid-table > tbody:nth-child({row})"))).text.split("\n")
                #print(f"Row {row}:", DataRow, end="\n\n")

                #Clean data
                while DataRow[0].startswith(
                    (
                        "Offer",
                        "Highest",
                        "Respond",
                        "Out of stock",
                        "Listing ",
                        "Restock",
                        "Link. ",
                        "View message",
                        "Send offer",
                        "Edit",
                        "Visibility ",
                        "Boost ",
                        "Promote "
                    )
                ) or DataRow[0].endswith(
                    (
                        "received",
                        " buyer",
                        " buyers"
                    )
                ):
                    DataRow.pop(0)

                if DataRow[0].startswith("owner."):
                    del DataRow[:4]

                if DataRow[1].startswith("Buy It Now"):
                    ItemID = DataRow[1].split(" · ")[-1]
                    DataRow.pop(1)
                    DataRow.insert(3, ItemID)

#                if Account == "SellerOrg3":
#                    ItemID = DataRow[-2]
#                    DataRow.pop(-2)
#                    DataRow.insert(3, ItemID)

                DataRow.pop(4)
                DataRow.pop(4)

                if DataRow[4] == "Research prices":
                    DataRow.pop(4)

                if DataRow[2].startswith("$"):
                    price = DataRow[2].replace("$", "").replace(",", "")
                    DataRow.pop(2)
                    DataRow.insert(2, price)

                if DataRow[3] == "Buy It Now":
                    DataRow.pop(3)

                DataFromEbay.append(DataRow[:4])
                #print(f"Row {row}:", DataRow, end="\n\n")
                row += 1

            except TimeoutException:
                getting_rows = False

        #Create DataFrame and save it to an Excel file
        df = pd.DataFrame(
            DataFromEbay,
            columns=[
                "Title",
                "SKU",
                "CurrentPrice",
                "ItemNumber"
            ]
        )

        #Insert a column to the beginning of the table with the name of the account
        df.insert(0, "Account", Account)

        #Insert another column to the beginning of the table with the date
        df.insert(0, "Date", Today)

        #Convert the columns to the correct data types
        df["CurrentPrice"] = pd.to_numeric(df["CurrentPrice"], errors="coerce").fillna(0)
#        print(df)
#        driver.quit()
#        quit()

        df = df.dropna(subset=["SKU"])
        if df.empty:
            print(f"[yellow][WARNING][/yellow] No valid rows to insert for account [cyan]{Account}[/cyan]. Skipping.")
            continue

        #Insert dataframe into the table in the SQL Database
        columns = ["Date", "Account", "Title", "SKU", "CurrentPrice", "ItemNumber"]
        database_utils.insert_dataframe(cursor, table_pending, df, columns)

        # Commit the transaction
        conn.commit()

        #Confirm that all items have been downloaded
        if Quantity != TotalQty:
            page += 1
            offset += 25

            #Go to next page
            print(f"[cyan][INFO][/cyan] Navigating to page #{page}.")
            driver.get(f"https://www.ebay.com/sh/lst/active?status=PENDING_OFFERS&limit=25&offset={offset}")
            driver.switch_to_window(0)

###############################################################################################################################################
def AttendPendingOffers() -> None:
    """
    Automates the process of attending all pending offers.
    """
    try:
        print("[cyan][INFO][/cyan] Retrieving offers.")
        CxOffer: float = round(float(WebDriverWait(driver, 15).until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            ".ui-component-offer-details > dl:nth-child(1) > div:nth-child(1) > dd:nth-child(2)"
        ))).text.split(" ")[0].replace("$", "").replace(",", "")), 2)

    except TimeoutException:
        CxOffer = 0

    #Write the offer to the Excel file
    OffersSh.range(f"J{FirstItem}").value = CxOffer

    if CxOffer == 0:
        print("[cyan][INFO][/cyan] Offer not received. Moving to next item.")
        return

    #Get the offer details
    Brand = OffersSh.range(f"D{FirstItem}").value.split(" ")[0].title() #Get the brand name from the Excel file
    SiteCost = float(OffersSh.range(f"H{FirstItem}").value) #Get the site cost from the Excel file
    CurrentPrice = float(OffersSh.range(f"G{FirstItem}").value) #Get the current price from the Excel file
    CurrentPriceLowered = CurrentPrice - 0.01 #Get the current price lowered by 0.01
    TotalCost = float(OffersSh.range(f"N{FirstItem}").value) #Get the total cost from the Excel file
    CurrentPriceMax = CurrentPrice * MinDiscount #Get the price to counteroffer after applying a 5% discount
    CurrentPriceMin = CurrentPrice * MaxDiscount #Get the price to counteroffer after applying a 10% discount
    MinCommission = CurrentPriceMin * eBayCommission #Get the 9.1% commission from eBay for the minimum discount
    MaxCommission = CurrentPriceMax * eBayCommission #Get the 9.1% commission from eBay for the maximum discount
    CxOfferCommission = CxOffer * eBayCommission #Get the 9.1% commission from eBay from customer's offer
    CurrPriceCommission = CurrentPriceLowered * eBayCommission #Get the 9.1% commission from eBay from the current price
    RealCost1 = TotalCost + MinCommission #Get the real cost by adding the total cost and eBay's 9.1% commission from the minimum discount
    RealCost2 = TotalCost + MaxCommission #Get the real cost by adding the total cost and eBay's 9.1% commission from the maximum discount
    RealCost3 = TotalCost + CxOfferCommission #Get the real cost by adding the total cost and eBay's 9.1% commission from the customer's offer
    RealCost4 = TotalCost + CurrPriceCommission #Get the real cost by adding the total cost and eBay's 9.1% commission from the current price
    MinProfit = CurrentPriceMin - RealCost1 #Get the profit from the minimum discount
    MaxProfit = CurrentPriceMax - RealCost2 #Get the profit from the maximum discount
    CxProfit = CxOffer - RealCost3 #Get the profit from the customer's offer
    CurrentPriceProfit = CurrentPriceLowered - RealCost4 #Get the profit from the current price
    MinProfitPerc = MinProfit / CurrentPriceMin #Get the profit percentage from the minimum discount
    MaxProfitPerc = MaxProfit / CurrentPriceMax #Get the profit percentage from the maximum discount
    CxProfitPerc = CxProfit / CxOffer #Get the profit percentage from the customer's offer
    CurrentPricePerc = CurrentPriceProfit / CurrentPriceLowered #Get the profit percentage from the current price

    #Determine if the offer is accepted or countered
    if CxOffer == 0 or SiteCost == 0 or SiteCost == 0.01 or Brand in Brands:
        Action = "Removed By Buyer"
        OffersSh.range(f"L{FirstItem}").value = [Action, 0]

    elif CxProfitPerc >= 0.11:
        Action = "Accepted"

        OffersSh.range(f"L{FirstItem}").value = [Action, 0]
        print(f"[cyan][INFO][/cyan] Customer's offer [cyan]accepted[/cyan] for ${CxOffer}, with a {round(CxProfitPerc * 100, 2)}% profit/loss.")

    elif MinProfitPerc >= 0.11 and MaxProfitPerc < 0.11:
        Counteroffer = round(CurrentPriceMin, 2)
        Action = "Counteroffer"

        OffersSh.range(f"L{FirstItem}").value = [Action, Counteroffer]
        print(f"[cyan][INFO][/cyan] Customer's offer [cyan]countered[/cyan] for ${Counteroffer}, with a {round(MinProfitPerc * 100, 2)}% profit/loss.")

    elif MaxProfitPerc >= 0.11 and CxProfitPerc < 0.11:
        Counteroffer = round(CurrentPriceMax, 2)
        Action = "Counteroffer"

        OffersSh.range(f"L{FirstItem}").value = [Action, Counteroffer]
        print(f"[cyan][INFO][/cyan] Customer's offer [cyan]countered[/cyan] for ${Counteroffer}, with a {round(MaxProfitPerc * 100, 2)}% profit/loss.")

    else:
        Counteroffer = round(CurrentPriceLowered, 2)
        Action = "Counteroffer"

        OffersSh.range(f"L{FirstItem}").value = [Action, Counteroffer]
        print(f"[cyan][INFO][/cyan] Customer's offer [cyan]countered[/cyan] by lowering our current price by $0.01 to ${Counteroffer}, with a {round(CurrentPricePerc * 100, 2)}% profit/loss.")

    #Work the offer based on the action
    if Action == "Accepted":
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
            By.CSS_SELECTOR,
            "button.ui-component-button:nth-child(1)"
        ))).click()
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
            By.CSS_SELECTOR,
            "button.ui-component-button:nth-child(1)"
        ))).click()
        
    elif Action == "Counteroffer":
        #Build the counter offer message
        CounterofferMessage = f"""
        Thanks for your offer! We send our absolute final price of ${Counteroffer}.

        This is unbeatable value! You get fast Free Shipping, worry Free Returns, from a trusted seller. We cannot go lower on this item.

        Accept our counteroffer now to secure your deal!
        """

        #Click first counteroffer button, input offer number and counteroffer message
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
            By.CSS_SELECTOR,
            "button.ui-component-button:nth-child(2)"
        ))).click()

        WebDriverWait(driver, 15).until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "#bestoffer__priceInput"
        ))).send_keys(str(Counteroffer))

        WebDriverWait(driver, 15).until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "#bestoffer__messageTextarea"
        ))).send_keys(CounterofferMessage)

        #Review counteroffer
        try:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "button.ui-component-button:nth-child(1)"
            ))).click()

        except TimeoutException:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                By.XPATH,
                "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[4]/button[1]"
            ))).click()

        #Confirm counteroffer
        try:
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "button.ui-component-button:nth-child(1)"
            ))).click()

        except TimeoutException:
            try:
                WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                    By.XPATH,
                    "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[3]/button[1]"
                ))).click()

            except TimeoutException:
                WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                    By.XPATH,
                    "/html/body/div[5]/div[4]/div[2]/div/div[2]/div[1]/div/div/div[4]/button[1]"
                ))).click()

        #If error message pops up, then skip the item and write 0 in Excel
        try:
            AlertStatus: str = WebDriverWait(driver, 15).until(EC.presence_of_element_located((
                By.XPATH,
                "/html/body/div[5]/div[4]/div[2]/div/div[2]/div/div/div/div[2]/div/div[2]"
            ))).text
            
            print(f"[bold red][ERROR][/bold red] [cyan]{AlertStatus}[/cyan]. Moving to next item.")
            OffersSh.range(f"J{FirstItem}").value = 0

        except TimeoutException:
            pass

    else:
        print(f"[cyan][INFO][/cyan] Item [cyan]{Action}[/cyan]. Moving to next item.")

###############################################################################################################################################
#Ask the user if they want to start the process now
BtnPressed = ctypes.windll.user32.MessageBoxW(
    0,
    "Do you want to start the process now?",
    "eBay Pending Offers",
    4 | 0x20
)

while True:
    #Time to start
    StartTime = "17:30:00"
    StartHour = int(StartTime.split(":")[0])
    StartMin = StartTime.split(":")[1]
    nowHour = int(datetime.now().strftime("%H"))
    tomorrow: str = custom_functions.tomorrow()
    SleepTime = seconds_until_target(StartTime)

    #If the user pressed "Yes", then start the process
    if BtnPressed == 7:
        if nowHour >= StartHour:
            if StartHour > 12:
                print(f"[cyan][INFO][/cyan] eBay Pending Offers will be attended tomorrow {tomorrow} at {StartHour - 12}:{StartMin} PM.")
            else:
                print(f"[cyan][INFO][/cyan] eBay Pending Offers will be attended tomorrow {tomorrow} at {StartHour}:{StartMin} AM.")
        else:
            if StartHour > 12:
                print(f"[cyan][INFO][/cyan] eBay Pending Offers will be attended today at {StartHour - 12}:{StartMin} PM.")
            else:
                print(f"[cyan][INFO][/cyan] eBay Pending Offers will be attended today at {StartHour}:{StartMin} AM.")

        #Sleep until just before the Start time
        time.sleep(max(SleepTime - 1, 0))

        #Loop to ensure that we catch the exact time
        while datetime.now().strftime("%H:%M:%S") != StartTime:
            time.sleep(0.5)

    #Get today's date on two formats
    Today: str = datetime.now().strftime("%Y-%m-%d")
    Date: str = datetime.now().strftime("%m/%d/%Y")

    #Reset the value of the button
    BtnPressed = 7

    ##################################################################################################################################################
    #Check the date of the last time the "Aged Inventory" Excel file was updated
    info = custom_functions.files_info(AgedInventoryPath)

    for file in info:
        file_date = str(file["Date Modified"]).split(" ")[0]

    #If it was updated today, then upload the file to the database
    if file_date == Today:
        print("[cyan][INFO][/cyan] Uploading [cyan]Aged Inventory[/cyan] items to database.")
        AgedInventory()

    #Create SQL Database connection
    conn = custom_functions.SQLConnection("eBay")
    cursor = conn.cursor()

    ###############################################################################################################################################
    getting_offers = True
    while getting_offers:
        try:
            ###############################################################################################################################################
            for Account in Accounts:
                #Set the Google Chrome user folder name
                if Account == "SellerOrg":
                    profile = "Default"
                elif Account == "SellerOrg2":
                    profile = "Profile 1"
                elif Account == "SellerOrg3":
                    profile = "Profile 4"
                elif Account == "Account4":
                    profile = "Profile 5"

                #Confirm if pending offers have been already downloaded today
                if LastUpdate() == Today:
                    print(f"[cyan][INFO][/cyan] All pending offers have been retrieved today from [cyan]{Account}[/cyan] account. Moving to next account.")

                else:
                    ###############################################################################################################################################
                    #Initialize Chrome
                    opening_browser = True
                    while opening_browser:
                        try:
                            driver: object = chrome.start_browser(
                                user_data_dir,
                                profile,
                                headless=True
                            )
                            opening_browser = False
                        except (SessionNotCreatedException, RuntimeError):
                            print("[bold red][ERROR][/bold red] Failed to open the Chrome. It seems Chrome was already open. Killing the application and retrying.")
                            custom_functions.kill_app("chrome")

                    #Navigate to eBay's overview webpage to load cookies
                    driver.get("https://www.ebay.com/sh/ovw")
                    driver.switch_to_window(0)

                    #If not logged in, then login
                    try:
                        accounts.eBay(
                            password=password,
                            driver=driver
                        )

                    except TimeoutException:
                        pass

                    print(f"[cyan][INFO][/cyan] Navigating to [cyan]{Account}[/cyan] profile on eBay.")
                    driver.get("https://www.ebay.com/sh/lst/active?status=PENDING_OFFERS&limit=25")
                    driver.switch_to_window(0)

                    #If banner to send offers appears, close it 
                    try:
                        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.fake-link:nth-child(3)"))).click()
                    except (TimeoutException, ElementClickInterceptedException):
                        pass

                    #Select the desired columns
                    ebay.customize_offers_table(driver)

                    #Confirm if there are any pending offers
                    try:
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".zeroResultsMessage")))
                        print(f"[cyan][INFO][/cyan] No pending offers for [cyan]{Account}[/cyan] account.")
                    except TimeoutException:
                        #Retrieve and save pending offers
                        GetPendingOffers()

                    #Close browser
                    driver.quit()

            ###############################################################################################################################################
            #Open Pending Offers Workbook
            print("[cyan][INFO][/cyan] Opening [cyan]Pending Offers[/cyan] workbook.")
            OffersWb = xl.Book(OffersWbStr)
            OffersSh = OffersWb.sheets("Pending Offers")
            custom_functions.update_directory(OffersWb)

            #Load Macros
            RefreshAll = OffersWb.macro("Module1.RefreshAll")
            SortAll = OffersWb.macro("Module1.SortAll")
            Reorganize = OffersWb.macro("Module1.Reorganize")

            print("[cyan][INFO][/cyan] Refreshing all queries.")
            RefreshAll()
            time.sleep(5)
            OffersWb.save()

            #Sort table to get the missing offers at the bottom
            SortAll()
            time.sleep(5)

            for Account in Accounts:
                #Find the first and last empty row in the table
                FirstItem: int = custom_functions.first_empty_row(OffersSh, "J", "B11")
                LastItem = int(OffersSh.range(f"B{OffersSh.cells.last_cell.row}").end("up").row)

                if FirstItem > LastItem:
                    print("[cyan][INFO][/cyan] No new Pending Offers. Exiting.")
                    RawData = None

                elif FirstItem == LastItem:
                    RawData = []
                    RawData.append(OffersSh.range(f"C{FirstItem}:F{LastItem}").value)

                else:
                    RawData = OffersSh.range(f"C{FirstItem}:F{LastItem}").value

                if RawData is not None:
                    CurrAccount = RawData[0][0]

                    if CurrAccount != Account:
                        continue

                    #Set the Google Chrome user folder name
                    if Account == "SellerOrg":
                        profile = "Default"
                    elif Account == "SellerOrg2":
                        profile = "Profile 1"
                    elif Account == "SellerOrg3":
                        profile = "Profile 4"
                    elif Account == "Account4":
                        profile = "Profile 5"

                    ###############################################################################################################################################
                    #Initialize Chrome
                    opening_browser = True
                    while opening_browser:
                        try:
                            driver: object = chrome.start_browser(user_data_dir, profile, headless=True)
                            opening_browser = False
                        except SessionNotCreatedException:
                            print("[bold red][ERROR][/bold red] Failed to open the Chrome. It seems Chrome was already open. Killing the application and retrying.")
                            custom_functions.kill_app("chrome")

                    print(f"[cyan][INFO][/cyan] Navigating to [cyan]{Account}[/cyan] profile on eBay.")
                    driver.get("https://www.ebay.com/")
                    time.sleep(2)
                    driver.switch_to_window(0)

                    for item in RawData:
                        if item[0] != Account:
                            print(f"[cyan][INFO][/cyan] {Account} pending offers are completed. Moving to [cyan]{item[0]}[/cyan].")
                            driver.quit()
                            break

                        print(f"[cyan][INFO][/cyan] Navigating to item {item[3]}.")
                        driver.get(f"https://www.ebay.com/bo/seller/showOffers?itemid={item[3]}")
                        driver.switch_to_window(0)

                        #Attend pending offers
                        AttendPendingOffers()
                        
                        FirstItem += 1

            getting_offers = False

        except Exception:
            driver.save_screenshot("Error.png")
            print(f"[bold red][ERROR][/bold red] An error occurred. A screenshot of the error has been saved to the current working directory.\n\nError:")
            traceback.print_exc()

        finally:
            custom_functions.kill_app("chrome")

    ##################################################################################################################################################
    #Gather data from the Pending Offers workbook to create the email body
    FCAccepted = int(OffersSh.range("C4").value)
    FCCounter = int(OffersSh.range("C5").value)
    FCExpired = int(OffersSh.range("C6").value)
    FCDeclined = int(OffersSh.range("C7").value)
    FCRemoved = int(OffersSh.range("C8").value)
    FCTotalOff = int(OffersSh.range("C9").value)
    LSAccepted = int(OffersSh.range("E4").value)
    LSCounter = int(OffersSh.range("E5").value)
    LSExpired = int(OffersSh.range("E6").value)
    LSDeclined = int(OffersSh.range("E7").value)
    LSRemoved = int(OffersSh.range("E8").value)
    LSTotalOff = int(OffersSh.range("E9").value)

    #Get current hour
    current_hour = datetime.now().strftime("%H")

    #Create the greeting based on the current hour
    if 5 <= int(current_hour) <= 11:
        greeting = "Good morning"
    elif 12 <= int(current_hour) <= 17:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"

    #Create the body of the email
    body = f"""{greeting},
        <p>I hope you are doing well.</p>
        <p>Please find attached the eBay best offers cleared today for all accounts.</p>
        <p>Also, please find below a summary:</p>
        <ul>
        <p><b><u>SellerOrg</u></b></p>
        <li><b>Accepted: </b> {FCAccepted}</li>
        <li><b>Counteroffer: </b> {FCCounter}</li>
        <li><b>Declined: </b> {FCDeclined}</li>
        <li><b>Removed by Buyer: </b> {FCRemoved}</li>
        <li><b>Expired: </b> {FCExpired}</li>
        <li><b>Total: </b> {FCTotalOff}</li>
        </ul>
        <ul>
        <p><b><u>SellerOrg2</u></b></p>
        <li><b>Accepted: </b> {LSAccepted}</li>
        <li><b>Counteroffer: </b> {LSCounter}</li>
        <li><b>Declined: </b> {LSDeclined}</li>
        <li><b>Removed by Buyer: </b> {LSRemoved}</li>
        <li><b>Expired: </b> {LSExpired}</li>
        <li><b>Total: </b> {LSTotalOff}</li>
        </ul>
        <p>Thank you.</p>
        Sincerely,
        """

    ##################################################################################################################################################
    #Reorganize the table, save and close the workbook
    print("[cyan][INFO][/cyan] Saving and closing workbook.")
    Reorganize()
    OffersWb.save()
    time.sleep(2)
    OffersWb.close()

    #Send email notification
    outlook.send_email(
        account=sender_email,
        subject=f"eBay Report - Best Offers - {Date}",
        body=body,
        to=to_email,
        cc=cc_email,
        attachments=[OffersWbStr],
        show=True,
        send=True,
    )
    print("[cyan][INFO][/cyan] Email notification sent.")

    #Sleep 60 seconds before starting over
    time.sleep(60)