# --------------------------------------------------------------------------
# Author:           Erin Asilo
# Create date:      11/1/2022
# Ogranization:     RingCentral
# Email:            erin.asilo@ringcentral.com
#
# Purpose:          To automate the process of grabbing device warranty
#                   expiration dates and other things
# --------------------------------------------------------------------------

# --------------------------------------------------------------------------
# Imports
# --------------------------------------------------------------------------

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import chromedriver_autoinstaller
import time
import os
from dotenv import load_dotenv
from datetime import date
from datetime import datetime
import sys
import gspread
import pandas as pd
from googleapiclient.discovery import build
import string
from oauth2client.service_account import ServiceAccountCredentials

# --------------------------------------------------------------------------
# Load .env file
# --------------------------------------------------------------------------

load_dotenv()
EMAIL = os.environ.get("EMAIL")
PASSWORD = os.environ.get("PASSWORD")
DOWNLOAD_PATH = os.environ.get("DOWNLOAD_PATH")
SERVICE_ACCOUNT = os.environ.get("SERVICE_ACCOUNT")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
MASTER_ID = os.environ.get("MASTER_ID")
FILE_PATH = os.environ.get("FILE_PATH")
# --------------------------------------------------------------------------
# Create service with Google API
# --------------------------------------------------------------------------

gc = gspread.service_account(SERVICE_ACCOUNT)
mastersheet = gc.open_by_key(MASTER_ID)

site_list = mastersheet.worksheets()

CLIENT_SECRET_FILE = SERVICE_ACCOUNT
API_NAME = "sheets"
API_VERSION = "v4"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://spreadsheets.google.com/feeds",
]

creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT, scopes=SCOPES)
gc = gspread.authorize(creds)

spreadsheet = gc.open_by_key(SPREADSHEET_ID)

service = build(API_NAME, API_VERSION, credentials=creds)

# --------------------------------------------------------------------------
# Selenium chromedriver options
# --------------------------------------------------------------------------

chromedriver_autoinstaller.install()

options = webdriver.ChromeOptions()
# options.headless = True
options.add_argument("--window-size=1920,1080")
options.add_argument("start-maximized")
options.add_argument("--disable-gpu")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36"
)
# options.add_argument("--remote-debugging-port=9222")
options.add_argument("--no-sandbox")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--lang=en_US")

prefs = {"download.default_directory": DOWNLOAD_PATH}

options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)

wait = WebDriverWait(driver, 300)
action = ActionChains(driver)

sys_cls_clear = "Cls"

# --------------------------------------------------------------------------
# Date and time format
# --------------------------------------------------------------------------

today = datetime.now().strftime("%m/%d/%Y")
day = date.today().strftime("%A")

# --------------------------------------------------------------------------
# Function:     scrape()
# Purpose:      Log in to Mist and download the .csv file containing metrics
# --------------------------------------------------------------------------


def scrape():
    try:
        # Log in
        print("Executing Juniper Scrape ... ")
        print("    Logging into Juniper ... ")
        driver.get("https://entitlementsearch.juniper.net/")
        time.sleep(1.5)
        email = wait.until(
            EC.element_to_be_clickable((By.ID, "idp-discovery-username"))
        )
        email.send_keys(EMAIL)
        login = wait.until(EC.element_to_be_clickable((By.ID, "idp-discovery-submit")))
        login.click()
        time.sleep(3)
        password = wait.until(
            EC.element_to_be_clickable((By.ID, "okta-signin-password"))
        )
        password.send_keys(PASSWORD)
        login = wait.until(EC.element_to_be_clickable((By.ID, "okta-signin-submit")))
        login.click()
        time.sleep(2)
        # For each site download all data
        for i in range(len(site_list)):
            worksheet = mastersheet.get_worksheet(i)
            serialno = worksheet.find("Serial Number")
            hostname = worksheet.find("Device Name")
            hostname = worksheet.col_values(hostname.col)
            del hostname[:2]
            values = worksheet.col_values(serialno.col)
            del values[:2]
            values = [j for j in values if j]
            devicename = []
            for i in range(len(values)):
                devicename.append(hostname[i])
            textbox = wait.until(
                EC.element_to_be_clickable((By.ID, "textAreaSerialNos"))
            )
            textbox.send_keys(Keys.CONTROL + "a")
            textbox.send_keys(Keys.DELETE)
            for elem in values:
                textbox.send_keys(elem + "\n")
            print(values)
            submit = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="root"]/div/main/div/div[3]/button')
                )
            )
            submit.click()
            while True:
                try:
                    download = wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                '//*[@id="root"]/div/div/main/div[2]/div/button[2]',
                            )
                        )
                    )
                    download.click()
                    time.sleep(3)

                    download = wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.CSS_SELECTOR,
                                "#modal-export-excel > div.modal-footer > div.modal-actions-right > button.success",
                            )
                        )
                    )
                    print("Download Button Found")
                    try:
                        os.remove(FILE_PATH)
                    except:
                        pass
                    download.click()
                    while not os.path.exists(FILE_PATH):
                        time.sleep(1)
                    df = pd.read_excel(
                        FILE_PATH, sheet_name="ReportData", engine="openpyxl"
                    )
                    df = df.drop(df.columns[[1]], axis=1, inplace=False)
                    df = df.drop_duplicates(
                        subset=["Serial No."], keep="first", inplace=False
                    )
                    print(df)
                    # df = df.fillna("N/A")

                    df["Warranty Expiry Date"] = pd.to_datetime(
                        df["Warranty Expiry Date"], format="%m-%d-%Y", errors="coerce"
                    )
                    df["Start Date"] = pd.to_datetime(
                        df["Start Date"], format="%m-%d-%Y", errors="coerce"
                    )
                    df["End Date"] = pd.to_datetime(
                        df["End Date"], format="%m-%d-%Y", errors="coerce"
                    )

                    df["Warranty Expiry Date"] = df["Warranty Expiry Date"].dt.strftime(
                        "%m-%d-%Y"
                    )
                    df["Start Date"] = df["Start Date"].dt.strftime("%m-%d-%Y")
                    df["End Date"] = df["End Date"].dt.strftime("%m-%d-%Y")

                    df.insert(0, "Device Name", devicename)
                    df = df.fillna("N/A")
                    print("Updating " + worksheet.title + " worksheet ...")
                    values = [
                        ["Last updated: " + today, "", "", "", "", "", "", "", ""]
                    ]
                    values.extend([df.columns.values.tolist()])
                    values.extend(df.values.tolist())
                    spreadsheet.values_update(
                        worksheet.title,
                        params={"valueInputOption": "USER_ENTERED"},
                        body={"values": values},
                    )
                    os.remove(FILE_PATH)
                    back = wait.until(
                        EC.element_to_be_clickable(
                            (By.XPATH, '//*[@id="root"]/div/div/main/div[1]/a')
                        )
                    )
                    back.click()
                    break
                except Exception as e:
                    print(e)
                    continue

    except Exception as e:
        print(e)
        driver.quit()
        retry()


# --------------------------------------------------------------------------
# Function:     parse_and_upload()
# Purpose:      Parses through .csv list and saves the data to a variable
#               per site, then sends data to Google sheet through API
# --------------------------------------------------------------------------


def parse_and_upload():
    pass


# --------------------------------------------------------------------------
# Function:     update_sheet(worksheet, range, values)
# Purpose:      Updates cells on a worksheet
# --------------------------------------------------------------------------


def update_sheet(worksheet, range, values):
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        valueInputOption="USER_ENTERED",
        range=worksheet + range,
        body=values,
    ).execute()


# --------------------------------------------------------------------------
# Function:     append_sheet(worksheet, range, values)
# Purpose:      Appends cells on a worksheet
# --------------------------------------------------------------------------


def append_sheet(worksheet, range, values):
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        valueInputOption="USER_ENTERED",
        range=worksheet + range,
        body=values,
    ).execute()


# --------------------------------------------------------------------------
# Function:     next_available_row(worksheet)
# Purpose:      Finds the next row number of the given worksheet
# --------------------------------------------------------------------------


def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list) + 2)


# --------------------------------------------------------------------------
# Function:     n2a()
# Purpose:      Converts a number to its corresponding column (0 = A, 26 = AA)
# --------------------------------------------------------------------------


def n2a(n, b=string.ascii_uppercase):
    d, m = divmod(n, len(b))
    return n2a(d - 1, b) + b[m] if d else b[m]


# --------------------------------------------------------------------------
# Function:     bye()
# Purpose:      Closes the program
# --------------------------------------------------------------------------


def bye():
    try:
        sys.exit(1)
    except:
        os._exit(1)


# --------------------------------------------------------------------------
# Function:     retry()
# Purpose:      Starts a new instance of the program
# --------------------------------------------------------------------------


def retry():
    sys.stdout.flush()
    os.execv(sys.executable, [sys.executable, '"' + sys.argv[0] + '"'] + sys.argv[1:])


# --------------------------------------------------------------------------
# Function:     main()
# Purpose:      Run the previous functions and catches errors
# --------------------------------------------------------------------------


def main():
    if day == "Saturday" or day == "Sunday":
        driver.quit()
        bye()
    else:
        try:
            print("Today is " + day + " " + today + "\n")
            scrape()
            parse_and_upload()
            print("Done!\n")
            driver.quit()
        # Retry program if it fails
        except Exception as e:
            print(e)
            try:
                driver.quit()
            except:
                pass
            finally:
                retry()
        finally:
            bye()


main()
