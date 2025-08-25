import PyPDF2
import pdfplumber
import re
import pandas as pd
import os
import glob
from collections import defaultdict
from typing import List, Dict

class BulkPOItemExtractor:
    def __init__(self, folder_path: str):
        self.folder_path = folder_path
        self.all_pos_data = {}
        self.combined_items = defaultdict(lambda: {'total_quantity': 0, 'po_files': [], 'quantities_per_po': {}})

    def get_pdf_files(self) -> List[str]:
        pdf_files = glob.glob(os.path.join(self.folder_path, "*.pdf"))
        print(f"Found {len(pdf_files)} PDF files" if pdf_files else f"No PDF files found in {self.folder_path}")
        return pdf_files

    def extract_text_pdfplumber(self, pdf_path: str) -> str:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                full_text = ""
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for row in table[1:]:
                                if row and len(row) > 2 and re.match(r'^\d{5}$', str(row[0])):
                                    text += "\n".join(str(cell) for cell in row if cell) + "\n"
                    full_text += text
                return full_text if full_text else "\n".join(page.extract_text() or "" for page in pdf.pages)
        except Exception as e:
            print(f"Error extracting from {os.path.basename(pdf_path)} with pdfplumber: {e}")
            return ""

    def extract_text_pypdf2(self, pdf_path: str) -> str:
        try:
            with open(pdf_path, 'rb') as file:
                return "".join(page.extract_text() or "" for page in PyPDF2.PdfReader(file).pages)
        except Exception as e:
            print(f"Error extracting from {os.path.basename(pdf_path)} with PyPDF2: {e}")
            return ""

    def extract_po_metadata(self, text: str) -> Dict[str, str]:
        metadata = {}

        # Extract PO Number (Document Ref)
        po_number_match = re.search(r'Document Ref:\s*(\d+)', text)
        if po_number_match:
            metadata['po_number'] = po_number_match.group(1).strip()

        # Extract Vendor/Location
        vendor_location_match = re.search(r'Vendor.*?\n(.*?)\n', text, re.DOTALL)
        if vendor_location_match:
            full_line = vendor_location_match.group(1).strip()
            # Match first occurrence of location code pattern (e.g., GUJ - MEGA - CQLA)
            location_match = re.search(r'([A-Z]{2,3}\s*-\s*[A-Z]+(?:\s*-\s*[A-Z]+)*)', full_line)
            if location_match:
                metadata['vendor_location'] = location_match.group(1).strip()

        # Extract PO Date
        po_date_match = re.search(r'PO Date:\s*(\d{2}\.\d{2}\.\d{4})', text)
        if po_date_match:
            metadata['po_date'] = po_date_match.group(1)

        # Extract Total Including Sales Tax
        total_amount_match = re.search(r'Total Including Sales Tax\s*([\d,]+\.?\d*)', text)
        if total_amount_match:
            metadata['total_amount'] = total_amount_match.group(1)

        return metadata

    def parse_items_from_text(self, text: str) -> List[Dict]:
        items = []
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        i = 0
        while i < len(lines):
            line = lines[i]
            item_match = re.match(r'^(\d{5})\s+(.+?)\s+(\d+\.?\d*)\s+\d+\.?\d*\s+\d+\.?\d*\s+\d+\.?\d*\s+\d+\.?\d*$', line)
            if item_match:
                item_number = item_match.group(1)
                name = item_match.group(2).strip()
                qty = float(item_match.group(3))
                items.append({'item_number': item_number, 'name': name, 'quantity': qty})
                i += 1
                continue
            i += 1
        return items

    def get_short_po_name(self, po_file: str, metadata: Dict = None) -> str:
        if metadata and 'po_number' in metadata:
            return metadata['po_number']
        return po_file.replace('.pdf', '')[:20].replace(' ', '_')

    def process_single_pdf(self, pdf_path: str) -> Dict:
        filename = os.path.basename(pdf_path)
        text = self.extract_text_pdfplumber(pdf_path)
        if not text:
            text = self.extract_text_pypdf2(pdf_path)
        if not text:
            return {'filename': filename, 'success': False, 'items': [], 'metadata': {}}
        metadata = self.extract_po_metadata(text)
        items = self.parse_items_from_text(text)
        return {'filename': filename, 'success': bool(items), 'items': items, 'metadata': metadata}

    def process_all_pdfs(self):
        for pdf_path in self.get_pdf_files():
            result = self.process_single_pdf(pdf_path)
            self.all_pos_data[result['filename']] = result
            if result['success'] and result['items']:
                for item in result['items']:
                    item_key = item['name'].strip()
                    self.combined_items[item_key]['total_quantity'] += item['quantity']
                    self.combined_items[item_key]['po_files'].append(result['filename'])
                    self.combined_items[item_key]['quantities_per_po'][result['filename']] = item['quantity']

    def save_to_excel(self, output_file: str = 'po_analysis.xlsx'):
        try:
            with pd.ExcelWriter(os.path.join(self.folder_path, output_file), engine='openpyxl') as writer:

                # PO Summary Sheet
                po_summary_data = []
                for f, d in self.all_pos_data.items():
                    po_summary_data.append({
                        'PO Number': d['metadata'].get('po_number', ''),
                        'Vendor/Location': d['metadata'].get('vendor_location', ''),
                        'PO Date': d['metadata'].get('po_date', ''),
                        'Total Sales (PKR)': d['metadata'].get('total_amount', '')
                    })
                pd.DataFrame(po_summary_data).to_excel(writer, sheet_name='PO Summary', index=False)

                # Quantity Summary Sheet
                summary_data = []
                po_columns = [self.get_short_po_name(f, self.all_pos_data[f]['metadata']) for f in self.all_pos_data.keys()]
                for item_name, data in self.combined_items.items():
                    row = {'Row Labels': item_name}
                    for po_file in self.all_pos_data.keys():
                        po_short_name = self.get_short_po_name(po_file, self.all_pos_data[po_file]['metadata'])
                        row[po_short_name] = data['quantities_per_po'].get(po_file, 0)
                    row['Grand Total'] = sum(row[po] for po in po_columns)
                    summary_data.append(row)
                summary_df = pd.DataFrame(summary_data).fillna(0)
                summary_df = summary_df[['Row Labels'] + po_columns + ['Grand Total']]
                summary_df.to_excel(writer, sheet_name='Quantity Summary', index=False)

            print(f"\n💾 Data saved to: {os.path.join(self.folder_path, output_file)}")
        except Exception as e:
            print(f"Error saving to Excel: {e}")

    def run_analysis(self):
        print(f"🚀 Starting Bulk PO Analysis...\n📁 Folder: {self.folder_path}")
        self.process_all_pdfs()
        self.save_to_excel()
        total_pos = sum(1 for data in self.all_pos_data.values() if data['success'])
        total_unique_items = len(self.combined_items)
        total_quantity = sum(data['total_quantity'] for data in self.combined_items.values())
        print(f"\n{'='*60}\n📈 FINAL STATISTICS\n{'='*60}")
        print(f"✅ Successfully processed POs: {total_pos}")
        print(f"📦 Total unique items: {total_unique_items}")
        print(f"🔢 Total quantity across all POs: {total_quantity:.0f}")
        print(f"💾 Results saved to Excel file")

def main():
    folder_path = r"C:\Users\Hassan Shahzad\Downloads\PO"
    if not os.path.exists(folder_path):
        print(f"❌ Folder not found: {folder_path}\nPlease make sure the folder path is correct.")
        return
    BulkPOItemExtractor(folder_path).run_analysis()

if __name__ == "__main__":
    main()
