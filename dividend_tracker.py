
# Import libraries
import yfinance as yf
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import logging
import traceback

# Mount Google Drive
from google.colab import drive
drive.mount('/content/drive')  # Mounts your Google Drive

# =========================
# Step 1: Configure Logging
# =========================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# =========================
# Step 2: Google Sheets API Setup
# =========================
#  Updated JSON key file path
json_key_file = '/content/drive/My Drive/keys/dividend-tracker-449904-45b1f3e4aebb.json'

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

try:
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_key_file, scope)
    client = gspread.authorize(creds)
    logging.info(" Successfully authenticated with Google Sheets API.")
except Exception as e:
    logging.error(f"❌ Failed to authenticate with Google Sheets API: {e}")
    logging.error(traceback.format_exc())
    exit()  # Exit if authentication fails

# Spreadsheet ID and sheet details
SPREADSHEET_ID = '1acdknVZlB5hWK5Zqk_rJA-V4IDvBMoFB3WFEOVUGPOs'  # Replace with your Spreadsheet ID
SHEET_NAME = 'Dividend Amount'  # Ensure this matches the exact sheet name

# =========================
# Step 3: Dividend Data Collection Function
# =========================

def fetch_and_upload_dividends():
    logging.info(f"📊 Starting dividend data fetch at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    ticker_names = {
        "IEMG": "NYSEARCA:IEMG",
        "IXJ": "NYSEARCA:IXJ",
        "IXC": "NYSEARCA:IXC",
        "VTV": "NYSEARCA:VTV",
        "VOOV": "NYSEARCA:VOOV",
        "UAE": "NASDAQ:UAE",
        "CCI": "NYSE:CCI",
        "ICSH": "BATS:ICSH",
        "IGSB": "NASDAQ:IGSB",
        "RDY": "NYSE:RDY",
        "GOOGL": "NASDAQ:GOOGL",
        "INDA": "BATS:INDA",
        "IEFA": "BATS:IEFA",
        "MSFT": "NASDAQ:MSFT",
        "QCOM": "NASDAQ:QCOM",
        "2330.TW": "TPE:2330",
        "2454.TW": "TPE:2454",
        "SMCI": "NASDAQ:SMCI",
        "SHY": "NASDAQ:SHY",
        "IHE": "NYSEARCA:IHE",
        "EXR": "NYSE:EXR",
        "HAL.NS": "NSE:HAL",
        "BEL.NS": "NSE:BEL",
        "GESHIP.NS": "NSE:GESHIP",
        "NAZARA.NS": "NSE.NAZARA",
        "FALN": "NASDAQ:FALN",
        "HYG": "NYSEARCA:HYG",
        "HYDB": "BATS:HYDB",
        "TSM": "NYSE:TSM"
    }

    start_date = "2024-01-01"
    end_date = datetime.today().strftime('%Y-%m-%d')

    dividend_records = []
    tickers_processed = 0

    # Fetching dividend data
    for ticker_symbol, name in ticker_names.items():
        logging.info(f"Fetching dividends for {ticker_symbol} ({name})...")
        try:
            stock = yf.Ticker(ticker_symbol)
            dividends = stock.dividends[start_date:end_date]

            if not dividends.empty:
                for date, amount in dividends.items():
                    dividend_records.append({
                        'Ticker': name,
                        'Ex-Dividend Date': date.strftime('%Y-%m-%d'),  #  Convert date to string
                        'Dividend Amount': amount
                    })
                logging.info(f"Found {len(dividends)} dividend records for {ticker_symbol}.")
            else:
                logging.warning(f"No dividend data found for {ticker_symbol}.")

            tickers_processed += 1

        except Exception as e:
            logging.error(f"Error fetching data for {ticker_symbol}: {e}")
            logging.error(traceback.format_exc())

    if not dividend_records:
        logging.warning(" No dividend records collected. Exiting function.")
        return

    # Convert to DataFrame
    dividends_df = pd.DataFrame(dividend_records)
    logging.info(f"📊 Collected dividend data for {tickers_processed} tickers.")

    # =========================
    # Upload Data to Google Sheet
    # =========================

    try:
        logging.info("📥 Connecting to Google Sheet...")
        spreadsheet = client.open_by_key(SPREADSHEET_ID)

        # Verify if the worksheet exists
        try:
            worksheet = spreadsheet.worksheet(SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            logging.error(f" Worksheet '{SHEET_NAME}' not found in the spreadsheet.")
            return

        logging.info(" Clearing existing data in the sheet...")
        worksheet.clear()

        logging.info(f" Uploading {len(dividends_df)} records to Google Sheets...")
        worksheet.update([dividends_df.columns.tolist()] + dividends_df.values.tolist(),
                          value_input_option='USER_ENTERED')

        logging.info("Dividend data uploaded successfully to Google Sheets!")

    except Exception as e:
        logging.error(f" Error uploading data to Google Sheets: {e}")
        logging.error(traceback.format_exc())

# =========================
# Step 4: Run the Function Directly for Testing
# =========================

fetch_and_upload_dividends()


print(" Data successfully updated in Google Sheets!")
