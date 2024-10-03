# Trade Journal Updater

This project automates the process of extracting trade data from PDFs and adding it to an Excel file. It scans a folder for PDFs, checks for existing entries in the Excel file by date, and inserts new rows in chronological order based on trade dates. You are prompted for confirmation to overwrite or append any existing data.

## Features

- **PDF Scanning**: Automatically scans a folder for PDF files and extracts trade data. 
- **Date Verification**: Checks the "Trade Entry Details" sheet in the Excel file for existing dates and prompts whether to overwrite or append data.
- **Chronological Insertion**: Inserts new rows into the Excel file in chronological order based on the trade dates.
- **Error Handling**: Includes basic error handling for file missing errors, invalid PDF format, and date validation.

## Requirements

- Python 3.x
- The following Python libraries (included in `requirements.txt`):
  - `pdfplumber`
  - `pandas`
  - `openpyxl`
  

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/TradeJournal.git
   cd TradeJournal
   pip install -r requirements.txt
   copytrades.py

   
