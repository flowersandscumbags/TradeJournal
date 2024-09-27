import pdfplumber
import logging
import openpyxl
from openpyxl import Workbook

# Setup logging configuration
logging.basicConfig(
    filename='extraction_log.log', 
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def extract_raw_text_from_page(pdf_path, page_number):
    logging.info(f"Starting extraction from page {page_number} of {pdf_path}")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_number >= len(pdf.pages):
                logging.error(f"Page {page_number} does not exist in the PDF.")
                return None

            page = pdf.pages[page_number]
            raw_text = page.extract_text()
            logging.info(f"Successfully extracted text from page {page_number}.")
            return raw_text
    except Exception as e:
        logging.error(f"Failed to extract text from page {page_number}: {e}")
        return None

def write_text_to_excel(text, excel_path):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Data"

        # Split the text by lines and write each line to a row in Excel
        for row_num, line in enumerate(text.split('\n'), 1):
            ws.cell(row=row_num, column=1, value=line)
        
        wb.save(excel_path)
        logging.info(f"Successfully wrote extracted data to {excel_path}")
    except Exception as e:
        logging.error(f"Failed to write data to Excel: {e}")

def main():
    pdf_path = pdf_path = 'c:/users/Mike/Documents/TradeJournal/Trades/2024-09-03.pdf'
    excel_output_path = 'extracted_page_3.xlsx'
    page_number = 2  # Page numbers are zero-indexed, so page 3 is index 2

    logging.info("Starting the extraction process.")
    
    # Extract raw text from page 3 of the PDF
    raw_text = extract_raw_text_from_page(pdf_path, page_number)

    if raw_text:
        logging.info("Writing extracted text to Excel.")
        write_text_to_excel(raw_text, excel_output_path)
    else:
        logging.error("No text extracted, skipping writing to Excel.")

if __name__ == "__main__":
    main()
