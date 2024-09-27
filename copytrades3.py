import os
import pdfplumber
import pandas as pd
from openpyxl import load_workbook, Workbook
import logging

# Setup logging configuration
logging.basicConfig(
    filename='parsing_log.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Function to extract trade data from the PDF using pdfplumber
def extract_pdf_data_with_pdfplumber(pdf_path):
    logging.info(f"Extracting data from {pdf_path} using PDFplumber")
    trades = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:  # Add a check if the PDF has any pages
                logging.error(f"No pages found in PDF: {pdf_path}")
                return []

            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()

                # Debug: Check if text was extracted successfully
                if not text:
                    logging.warning(f"No text extracted from page {page_num} of {pdf_path}")
                    continue

                logging.info(f"Raw extracted text from page {page_num}: {text}")

                # Check for "SECURITIES TRADING ACTIVITY" section
                if "SECURITIES TRADING ACTIVITY" in text:
                    logging.info(f"Found 'SECURITIES TRADING ACTIVITY' section on page {page_num}")

                    # Split the text into lines
                    lines = text.split('\n')

                    # Process each line and extract relevant data
                    for line in lines:
                        # Filter out lines that don't look like trade data
                        if not line or "MKT" in line or "Market" in line or "TRADING SUMMARY" in line:
                            continue  # Skip irrelevant lines

                        # Now, attempt to split and parse trade lines
                        trade_data = line.split()
                        
                        # Log parsed trade data for debugging
                        logging.debug(f"Parsed trade data from line: {trade_data}")

                        # Ensure the line contains sufficient trade information (expecting at least 11 fields)
                        if len(trade_data) >= 11:
                            try:
                                # Map fields based on expected structure
                                symbol = trade_data[0]
                                cusip = trade_data[1]
                                trade_date = trade_data[2]
                                settlement_date = trade_data[3]
                                account_type = trade_data[4]
                                buy_sell = trade_data[5]
                                quantity = float(trade_data[6].replace(',', ''))
                                price = float(trade_data[7].replace(',', ''))
                                gross_amount = float(trade_data[8].replace(',', ''))
                                commission = float(trade_data[9].replace(',', ''))
                                fee_tax = float(trade_data[10].replace(',', ''))
                                net_amount = float(trade_data[11].replace(',', ''))

                                # Add the extracted data to the trades list
                                trades.append({
                                    "Symbol": symbol,
                                    "Cusip": cusip,
                                    "Trade Date": trade_date,
                                    "Settlement Date": settlement_date,
                                    "Buy/Sell": buy_sell,
                                    "Quantity": quantity,
                                    "Price": price,
                                    "Gross Amount": gross_amount,
                                    "Commission": commission,
                                    "Fee/Tax": fee_tax,
                                    "Net Amount": net_amount
                                })
                            except ValueError as e:
                                logging.error(f"Error parsing trade data on line: {line} - {e}")
                        else:
                            logging.warning(f"Incomplete trade data on line: {line}")

            if not trades:
                logging.info(f"No trades extracted from {pdf_path}")

        logging.info(f"Successfully extracted {len(trades)} trades from {pdf_path}")
        return trades

    except Exception as e:
        logging.error(f"Error extracting data from {pdf_path}: {e}")
        return []

# Function to check or create sheets if they don't exist
def get_or_create_sheet(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        logging.info(f"Sheet '{sheet_name}' not found. Creating new sheet.")
        return wb.create_sheet(sheet_name)
    return wb[sheet_name]

# Function to append trade data to the Excel sheets
def append_trades_to_excel(trades, ws_trade_details, ws_trade_outcome):
    logging.info("Appending trades to Excel sheets")

    if not trades:
        logging.warning("No trades available to append to Excel")
        return

    for trade in trades:
        ws_trade_details.append([
            trade['Trade Date'],        # Date
            '',                         # Time (not available)
            '',                         # Ticker Symbol (not available in the PDF)
            trade['Quantity'],          # Position Size
            trade['Price'],             # Entry Price
            '',                         # Exit Price (not available)
            'Buy' if trade['Buy/Sell'] == 'B' else 'Sell',  # Order Type
            'Long'  # Assume all trades are Long since it's a cash account
        ])

        ws_trade_outcome.append([
            trade['Net Amount'],        # Profit/Loss
            trade['Commission']         # Commissions and Fees
        ])

    logging.info(f"Finished appending trades. Total trades appended: {len(trades)}")

# Main function to handle the entire process
def main():
    pdf_folder = input("Enter the folder path containing the PDFs: ").strip()
    excel_path = input("Enter the path to the Excel file: ").strip()

    logging.info(f"Starting to process PDF files in folder: {pdf_folder}")

    if not os.path.exists(pdf_folder):
        logging.error(f"The folder {pdf_folder} does not exist.")
        return
    if not os.path.exists(excel_path):
        logging.error(f"The file {excel_path} does not exist.")
        return

    try:
        wb = load_workbook(excel_path)
        logging.info(f"Successfully loaded Excel file: {excel_path}")
    except Exception as e:
        logging.error(f"Error opening Excel file {excel_path}: {e}")
        return

    ws_trade_details = get_or_create_sheet(wb, 'Trade Entry Details')
    ws_trade_outcome = get_or_create_sheet(wb, 'Trade Outcome')

    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

    if not pdf_files:
        logging.warning(f"No PDF files found in the folder {pdf_folder}.")
        return

    all_trades = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        logging.info(f"Processing {pdf_file}...")

        # Extract data from the PDF using pdfplumber
        trades = extract_pdf_data_with_pdfplumber(pdf_path)

        if trades:
            all_trades.extend(trades)
        else:
            logging.warning(f"No valid trade data found in {pdf_file}.")

    if all_trades:
        append_trades_to_excel(all_trades, ws_trade_details, ws_trade_outcome)
        wb.save(excel_path)
        logging.info(f"Trades have been updated and saved to {excel_path}")
    else:
        logging.warning("No trades to append to the Excel file.")

if __name__ == "__main__":
    main()
