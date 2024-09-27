import os
import PyPDF2
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

# Function to extract text from PDF
def extract_pdf_text(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in range(len(reader.pages)):
                text += reader.pages[page].extract_text()
        return text
    except Exception as e:
        print(f"Error extracting text from {pdf_path}: {e}")
        return ""

# Function to parse PDF content for trade data
def parse_pdf_for_trades(text):
    trades = []
    
    # Find all trade data for securities (BNZI, MU, etc.)
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if "SECURITIES TRADING ACTIVITY" in line:
            trade_section = lines[i+1:]
            for j in range(0, len(trade_section), 5):
                try:
                    symbol = trade_section[j].split()[0]
                    trade_date = trade_section[j+1].split()[1]
                    buy_sell = trade_section[j+2].split()[1]
                    quantity = float(trade_section[j+3].split()[1])
                    price = float(trade_section[j+4].split()[1])
                    commission = float(trade_section[j+4].split()[3])  # Added commission extraction
                    trades.append({
                        "Symbol": symbol,
                        "Trade Date": trade_date,
                        "Buy/Sell": buy_sell,
                        "Quantity": quantity,
                        "Price": price,
                        "Commission": commission  # Added commission to the trade data
                    })
                except (IndexError, ValueError):
                    break
    return trades

# Function to load Excel and check for existing dates
def check_existing_dates(ws_trade_details):
    existing_dates = []
    for row in ws_trade_details.iter_rows(min_row=2, max_col=1, values_only=True):
        try:
            date = pd.to_datetime(row[0]).date()
            existing_dates.append(date)
        except Exception:
            continue
    return existing_dates

# Function to append or overwrite trade data in the Excel sheet
def append_trades_to_excel(trades, ws_trade_details, ws_trade_outcome, existing_dates):
    for trade in trades:
        trade_date = pd.to_datetime(trade['Trade Date']).date()

        # Check if the date exists in the spreadsheet
        if trade_date in existing_dates:
            action = input(f"Date {trade_date} already exists. Would you like to overwrite or append? (O/A): ").strip().upper()
            if action == 'O':
                # Overwrite logic
                for row in ws_trade_details.iter_rows(min_row=2, max_col=1):
                    if pd.to_datetime(row[0].value).date() == trade_date:
                        ws_trade_details.delete_rows(row[0].row)
                        break
            elif action != 'A':
                print(f"Invalid input for {trade_date}. Skipping...")
                continue

        # Insert new trade entry in 'Trade Entry Details' sheet
        ws_trade_details.append([
            trade['Trade Date'],        # Date
            '',                         # Time (not available)
            trade['Symbol'],            # Ticker Symbol
            trade['Quantity'],          # Position Size
            trade['Price'],             # Entry Price
            '',                         # Exit Price (not available)
            trade['Buy/Sell'],          # Order Type
            'Long' if trade['Buy/Sell'] == 'Buy' else 'Short'  # Long/Short
        ])
        
        # Calculate the Net Amount and append to 'Trade Outcome'
        net_amount = trade['Quantity'] * trade['Price']
        ws_trade_outcome.append([
            net_amount,  # Profit/Loss (simplified)
            trade['Commission']  # Using actual commission from trade data
        ])

# Function to insert trades in chronological order
def insert_trades_chronologically(trades, ws_trade_details):
    trades = sorted(trades, key=lambda x: pd.to_datetime(x['Trade Date']).date())

    for trade in trades:
        trade_date = pd.to_datetime(trade['Trade Date']).date()
        for row in ws_trade_details.iter_rows(min_row=2, max_col=1):
            row_date = pd.to_datetime(row[0].value).date()
            if trade_date < row_date:
                ws_trade_details.insert_rows(row[0].row)
                ws_trade_details.cell(row=row[0].row, column=1).value = trade['Trade Date']
                ws_trade_details.cell(row=row[0].row, column=3).value = trade['Symbol']
                ws_trade_details.cell(row=row[0].row, column=4).value = trade['Quantity']
                ws_trade_details.cell(row=row[0].row, column=5).value = trade['Price']
                break
        else:
            # If no matching or earlier date found, append at the end
            ws_trade_details.append([
                trade['Trade Date'],        # Date
                '',                         # Time (not available)
                trade['Symbol'],            # Ticker Symbol
                trade['Quantity'],          # Position Size
                trade['Price'],             # Entry Price
                '',                         # Exit Price (not available)
                trade['Buy/Sell'],          # Order Type
                'Long' if trade['Buy/Sell'] == 'Buy' else 'Short'  # Long/Short
            ])

# Main function to scan folder for PDFs and process trades
def main():
    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()

    # Open file dialog for PDF folder selection
    pdf_folder = filedialog.askdirectory(title="Select folder containing PDFs")
    if not pdf_folder:
        print("No folder selected. Exiting.")
        return

    # Open file dialog for Excel file selection
    excel_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        print("No Excel file selected. Exiting.")
        return

    # Check if folder and Excel file exist
    if not os.path.exists(pdf_folder):
        print(f"Error: The folder {pdf_folder} does not exist.")
        return
    if not os.path.exists(excel_path):
        print(f"Error: The file {excel_path} does not exist.")
        return

    # Load Excel file
    try:
        wb = load_workbook(excel_path)
        ws_trade_details = wb['Trade Entry Details']
        ws_trade_outcome = wb['Trade Outcome']
    except Exception as e:
        print(f"Error opening Excel file {excel_path}: {e}")
        return

    # Check existing dates in the Excel file
    existing_dates = check_existing_dates(ws_trade_details)

    # Scan the folder for PDFs
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

    if not pdf_files:
        print(f"No PDF files found in the folder {pdf_folder}.")
        return

    all_trades = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        print(f"Processing {pdf_file}...")

        # Extract text and parse trades from the PDF
        pdf_text = extract_pdf_text(pdf_path)
        trades = parse_pdf_for_trades(pdf_text)

        if trades:
            all_trades.extend(trades)
        else:
            print(f"No valid trade data found in {pdf_file}.")

    # Prompt for confirmation and append/overwrite as necessary
    if all_trades:
        print("Proposed changes:")
        for trade in all_trades:
            print(f"- Date: {trade['Trade Date']}, Symbol: {trade['Symbol']}, Quantity: {trade['Quantity']}, Price: {trade['Price']}, Commission: {trade['Commission']}")
        
        confirm = input("Do you confirm these changes? (Y/N): ").strip().upper()
        if confirm == 'Y':
            append_trades_to_excel(all_trades, ws_trade_details, ws_trade_outcome, existing_dates)
            insert_trades_chronologically(all_trades, ws_trade_details)

            # Save the updated Excel file
            updated_excel_path = excel_path.replace(".xlsx", "_updated.xlsx")
            wb.save(updated_excel_path)
            print(f"Trades have been updated and saved to {updated_excel_path}.")
        else:
            print("Operation canceled.")
    else:
        print("No valid trade data to process.")

if __name__ == "__main__":
    main()