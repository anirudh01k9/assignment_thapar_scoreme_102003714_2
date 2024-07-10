import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook

def extract_tables_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    tables = []
    current_table = []
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        blocks = page.get_text("dict")["blocks"]
        
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    row = [span["text"] for span in line["spans"]]
                    row = [cell.strip() for cell in row if cell.strip()]  # Clean and ignore empty cells
                    if row:
                        current_table.append(row)
                    elif current_table:
                        tables.append(current_table)
                        current_table = []
                
    if current_table:  # Append any remaining table
        tables.append(current_table)
                
    return tables

def write_tables_to_excel(tables, excel_path):
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')
    
    for i, table in enumerate(tables):
        df = pd.DataFrame(table)
        df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False, header=False)
    
    writer.close()

def main():
    pdf_paths = ["test3.pdf", "test5.pdf"]
    for pdf_path in pdf_paths:
        tables = extract_tables_from_pdf(pdf_path)
        excel_path = pdf_path.replace('.pdf', '.xlsx')
        write_tables_to_excel(tables, excel_path)
        print(f'Extracted tables from {pdf_path} and saved to {excel_path}')

if __name__ == "__main__":
    main()
