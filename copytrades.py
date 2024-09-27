import os
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import logging
import tkinter as tk
from tkinter import filedialog

# Setup logging configuration
logging.basicConfig(
    filename='parsing_log.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Log to console as well
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger('').addHandler(console)

def extract_pdf_data_with_pdfplumber(pdf_path):
    logging.info(f"Extracting data from {pdf_path} using pdfplumber")
    trades = []
    columns = ["Symbol & Name", "Cusip", "Trade Date", "Settlement Date", "Account Type", "Buy/Sell", "Quantity", "Price", "Gross Amount", "Commission", "Fee/Tax", "Net Amount", "MKT", "Solicitation", "CAP"]

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages[2:], start=3):  # Start from the third page
                logging.info(f"Processing page {page_num}")
                tables = page.extract_tables()
                for table_num, table in enumerate(tables, start=1):
                    logging.info(f"Processing table {table_num} on page {page_num}")
                    if len(table[0]) == len(columns):
                        logging.info(f"Table {table_num} on page {page_num} matches expected column structure")
                        for row in table[1:]:  # Skip the header row
                            trade = {columns[i]: str(cell).strip() for i, cell in enumerate(row)}
                            trades.append(trade)
                            logging.info(f"Captured trade data: {trade}")
                    else:
                        logging.warning(f"Table {table_num} on page {page_num} does not match expected column structure")

    except Exception as e:
        logging.error(f"Error processing PDF {pdf_path}: {str(e)}")

    logging.info(f"Extracted {len(trades)} trades from {pdf_path}")
    return trades

def append_trades_to_excel(trades, ws_trade_details, ws_trade_outcome):
    logging.info(f"Appending {len(trades)} trades to Excel sheets")
    
    for trade in trades:
        # Mapping for Trade Entry Details
        trade_details_row = [
            trade['Trade Date'],
            '',  # Time (not available in your data)
            trade['Symbol & Name'].split()[0],  # Ticker Symbol
            abs(float(trade['Quantity'].replace(',', ''))),  # Shares (absolute value)
            abs(float(trade['Gross Amount'].replace(',', ''))),  # Position Size (absolute value)
            float(trade['Price'].replace(',', '')) if trade['Buy/Sell'] == 'B' else '',  # Entry Price
            float(trade['Price'].replace(',', '')) if trade['Buy/Sell'] == 'S' else '',  # Exit Price
            'Buy' if trade['Buy/Sell'] == 'B' else 'Sell',  # Order Type
            'Long'  # Assuming all trades are Long
        ]
        ws_trade_details.append(trade_details_row)
        logging.info(f"Appended trade details: {trade_details_row}")

        # Mapping for Trade Outcome
        outcome_row = [
            float(trade['Commission'].replace(',', '')),  # Commissions and Fees
            float(trade['Fee/Tax'].replace(',', '')),  # Tax
            float(trade['Net Amount'].replace(',', ''))  # Net
        ]
        ws_trade_outcome.append(outcome_row)
        logging.info(f"Appended trade outcome: {outcome_row}")

def get_or_create_sheet(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        logging.info(f"Creating new sheet: {sheet_name}")
        return wb.create_sheet(sheet_name)
    return wb[sheet_name]

def select_folder_or_file(prompt):
    root = tk.Tk()
    root.withdraw()
    
    if "folder" in prompt.lower():
        path = filedialog.askdirectory(title=prompt)
    else:
        path = filedialog.askopenfilename(title=prompt, filetypes=[("Excel files", "*.xlsx")])
    
    return path

def main():
    print("Please select the folder containing the PDFs.")
    pdf_folder = select_folder_or_file("Select the folder containing the PDFs")
    
    print("Please select the Excel file.")
    excel_path = select_folder_or_file("Select the Excel file")

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
        logging.error(f"Error opening Excel file {excel_path}: {str(e)}")
        return

    ws_trade_details = get_or_create_sheet(wb, 'Trade Entry Details')
    ws_trade_outcome = get_or_create_sheet(wb, 'Trade Outcome')

    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

    if not pdf_files:
        logging.warning(f"No PDF files found in the folder {pdf_folder}.")
    else:
        all_trades = []
        for pdf_file in pdf_files:
            pdf_path = os.path.join(pdf_folder, pdf_file)
            logging.info(f"Processing {pdf_file}...")
            trades = extract_pdf_data_with_pdfplumber(pdf_path)
            all_trades.extend(trades)

        if all_trades:
            append_trades_to_excel(all_trades, ws_trade_details, ws_trade_outcome)
        else:
            logging.warning("No trades found in any PDF.")

    wb.save(excel_path)
    logging.info(f"Workbook saved to {excel_path}")

    logging.info("Script execution completed.")

if __name__ == "__main__":
    main()