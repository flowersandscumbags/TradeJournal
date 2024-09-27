import os
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import logging

# Setup logging configuration
logging.basicConfig(filename='parsing_log.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Function to extract trade data from the PDF using pdfplumber
def extract_pdf_data_with_pdfplumber(pdf_path):
    logging.info(f"Extracting data from {pdf_path} using PDFplumber")
    trades = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                
                # Look for the section that has "SECURITIES TRADING ACTIVITY" in it
                if "SECURITIES TRADING ACTIVITY" in text:
                    table = page.extract_table()
                    for row in table[1:]:
                        # Extracting relevant columns from the table row
                        try:
                            trade_date = row[2]  # Trade Date
                            buy_sell = row[4]  # Buy/Sell
                            quantity = float(row[5].replace(',', ''))  # Quantity
                            price = float(row[6].replace(',', ''))  # Price
                            commission = float(row[8].replace(',', ''))  # Commission
                            net_amount = float(row[10].replace(',', ''))  # Net Amount

                            # Append extracted trade data to the trades list
                            trades.append({
                                "Trade Date": trade_date,
                                "Buy/Sell": buy_sell,
                                "Quantity": quantity,
                                "Price": price,
                                "Commission": commission,
                                "Net Amount": net_amount
                            })
                        except Exception as e:
                            logging.error(f"Error parsing row {row}: {e}")
                            continue
        logging.info(f"Successfully extracted data from {pdf_path}")
        return trades
    except Exception as e:
        logging.error(f"Error extracting data from {pdf_path}: {e}")
        return []

# Function to append trade data to the Excel sheets
def append_trades_to_excel(trades, ws_trade_details, ws_trade_outcome):
    logging.info("Appending trades to Excel sheets")

    for trade in trades:
        # Append to 'Trade Entry Details' sheet
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

        # Append to 'Trade Outcome' sheet
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

    # Load the existing Excel workbook
    try:
        wb = load_workbook(excel_path)
        ws_trade_details = wb['Trade Entry Details']
        ws_trade_outcome = wb['Trade Outcome']
        logging.info(f"Successfully loaded Excel file: {excel_path}")
    except Exception as e:
        logging.error(f"Error opening Excel file {excel_path}: {e}")
        return

    # Scan the folder for PDFs
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

    # If trades were found, append them to the Excel file
    if all_trades:
        append_trades_to_excel(all_trades, ws_trade_details, ws_trade_outcome)
        wb.save(excel_path)
        logging.info(f"Trades have been updated and saved to {excel_path}")
    else:
        logging.warning("No trades to append to the Excel file.")

if __name__ == "__main__":
    main()
