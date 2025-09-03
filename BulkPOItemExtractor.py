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
        self.combined_items = defaultdict(lambda: {
            'total_quantity': 0.0,
            'po_files': [],
            'quantities_per_po': {},
            'diy_code': ''
        })

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
                            for row in table:
                                if row:
                                    full_text += "\n".join(str(cell) for cell in row if cell) + "\n"
                    full_text += "\n" + text
                return full_text.strip()
        except Exception as e:
            print(f"Error extracting from {os.path.basename(pdf_path)} with pdfplumber: {e}")
            return ""

    def extract_text_pypdf2(self, pdf_path: str) -> str:
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                return "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception as e:
            print(f"Error extracting from {os.path.basename(pdf_path)} with PyPDF2: {e}")
            return ""

    def extract_po_metadata(self, text: str) -> Dict[str, str]:
        metadata = {}
        po_number_match = re.search(r'Document Ref:\s*(\d+)', text)
        if po_number_match:
            metadata['po_number'] = po_number_match.group(1).strip()

        vendor_location_match = re.search(r'Vendor.*?\n(.*?)\n', text, re.DOTALL | re.IGNORECASE)
        if vendor_location_match:
            full_line = vendor_location_match.group(1).strip()
            loc_match = re.search(r'([A-Z]{2,3}\s*-\s*[A-Z]+(?:\s*-\s*[A-Z]+)*)', full_line)
            if loc_match:
                metadata['vendor_location'] = loc_match.group(1).strip()

        po_date_match = re.search(r'PO Date:\s*(\d{2}\.\d{2}\.\d{4})', text)
        if po_date_match:
            metadata['po_date'] = po_date_match.group(1)

        total_amount_match = re.search(r'Total Including Sales Tax\s*([\d,]+\.?\d*)', text)
        if total_amount_match:
            metadata['total_amount'] = total_amount_match.group(1)

        return metadata

    def _to_float(self, s: str) -> float:
        try:
            return float(s.replace(',', '').strip())
        except Exception:
            return 0.0

    def parse_items_from_text(self, text: str) -> List[Dict]:
        """
        Robust parser:
         - Detect item starts like "00010 <name...>" (3-5 digits)
         - From the item line:
             * if there are trailing numeric columns (e.g. "36.000 154.00 ...") we pick the first qty-like number
               (pattern with 3 decimals is common for qty: 36.000) and strip trailing numeric columns from name.
         - Otherwise: look ahead up to 50 lines for DIY code and qty (Pcs lines, standalone nums under threshold).
         - Ignore very large standalone numbers (> 10,000) as quantity (likely IDs).
        """
        items: List[Dict] = []
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        n = len(lines)
        i = 0

        def is_item_start(line: str):
            return re.match(r'^\s*(\d{3,5})\s+(.+)$', line)

        while i < n:
            m = is_item_start(lines[i])
            if not m:
                i += 1
                continue

            item_number = m.group(1)
            raw = m.group(2).strip()
            diy_code = ""
            qty = None

            # Attempt to extract qty that is present on the same line (common in flattened PDFs)
            # Prefer patterns like 36.000 (three decimals) which is often the 'Quantity' column
            same_line_qty = re.search(r'(\d+\.\d{3})', raw)
            if same_line_qty:
                qty = self._to_float(same_line_qty.group(1))
                # remove trailing numeric columns starting from the matched qty
                raw = raw[:same_line_qty.start()].strip()
            else:
                # Another possibility: the line may contain "36.00 Pcs" etc appended
                same_line_pcs = re.search(r'([\d,]+(?:\.\d+)?)[^\S\r\n]*Pcs', raw, flags=re.I)
                if same_line_pcs:
                    qty = self._to_float(same_line_pcs.group(1))
                    raw = raw[:same_line_pcs.start()].strip()
                else:
                    # If raw ends with two or more numeric tokens (e.g. "20092 36.000 154.00"), remove trailing numeric tokens
                    # but only if there are >=2 numeric tokens at the end — this preserves a single trailing item-code like "# 20092"
                    trailing_nums = re.findall(r'([\d,]+(?:\.\d+)?)\s*$', raw)
                    # do nothing here; we'll fallback to look-ahead for qty

                    pass

            # Look ahead for DIY code and quantity if not already found
            j = i + 1
            window_end = min(n, i + 50)
            while j < window_end and not is_item_start(lines[j]):
                linej = lines[j]

                # DIY detection: handle "DIY28000..." or split "DIY" then digits on next line
                if not diy_code:
                    d1 = re.search(r'\bDIY\d+\b', linej)
                    if d1:
                        diy_code = d1.group(0)
                    else:
                        # handle split 'DIY' then digits on next line
                        if 'DIY' in linej and not re.search(r'\d', linej):
                            # check next line for digits
                            if j + 1 < n and re.fullmatch(r'\d{5,}', lines[j + 1]):
                                diy_code = 'DIY' + lines[j + 1].strip()
                                # skip that numeric-only line later by incrementing j
                                j += 1

                # Quantity detection priority:
                if qty is None:
                    # explicit Pcs line
                    m_pcs = re.search(r'([\d,]+(?:\.\d+)?)\s*(?:Pcs|Pieces|Piece|P\.cs)\b', linej, flags=re.I)
                    if m_pcs:
                        cand = self._to_float(m_pcs.group(1))
                        if 0 < cand < 10000:  # sanity threshold
                            qty = cand

                if qty is None:
                    # standalone numeric like '36.000' or '24.000'
                    m_num = re.fullmatch(r'([\d,]+(?:\.\d+)?)', linej)
                    if m_num:
                        cand = self._to_float(m_num.group(1))
                        if 0 < cand < 10000:  # reject huge numbers (likely IDs)
                            qty = cand

                if qty is None:
                    # look for "Qty" tokens
                    m_q = re.search(r'(?:Qty|Quantity)[^\d\n\r]*([\d,]+(?:\.\d+)?)', linej, flags=re.I)
                    if m_q:
                        cand = self._to_float(m_q.group(1))
                        if 0 < cand < 10000:
                            qty = cand

                if diy_code and qty is not None:
                    break
                j += 1

            # fallback: number directly before a line that is exactly "Piece"/"Pieces"
            if qty is None:
                for k in range(i + 1, window_end):
                    if re.fullmatch(r'(?i)piece|pieces', lines[k]):
                        prev_line = lines[k - 1] if k - 1 >= i + 1 else ""
                        m_prev = re.search(r'([\d,]+(?:\.\d+)?)', prev_line)
                        if m_prev:
                            cand = self._to_float(m_prev.group(1))
                            if 0 < cand < 10000:
                                qty = cand
                                break

            if qty is None:
                # warn but keep item with qty 0.0
                print(f"⚠️  Quantity not found for item {item_number} -> '{raw}'. Defaulting to 0.")
                qty = 0.0

            # Clean name: strip trailing numeric groups if there are 2+ numeric tokens at the end (these are likely column dumps)
            # but keep cases like single trailing code "# 20092"
            # if raw ends with two or more numeric tokens, drop them
            trailing_numeric_tokens = re.findall(r'([\d,]+(?:\.\d+)?)', raw)
            # If last two tokens are numeric, remove them from name
            if len(trailing_numeric_tokens) >= 2:
                # Find start of the first of the final numeric tokens and truncate from there
                # Use regex to find where the last-two-numbers block starts
                m_block = re.search(r'(\s+[\d,]+(?:\.\d+)?\s+[\d,]+(?:\.\d+)?\s*)$', raw)
                if m_block:
                    raw = raw[:m_block.start()].strip()

            items.append({
                'item_number': item_number,
                'name': raw,
                'quantity': float(qty),
                'diy_code': diy_code or ""
            })

            # advance to next item start
            i = j

        print(f"🔍 Parsed {len(items)} items from text")
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
                    key = f"{item['name']} ({item['diy_code']})" if item['diy_code'] else item['name']
                    self.combined_items[key]['total_quantity'] += float(item['quantity'])
                    self.combined_items[key]['po_files'].append(result['filename'])
                    self.combined_items[key]['quantities_per_po'][result['filename']] = float(item['quantity'])
                    self.combined_items[key]['diy_code'] = item['diy_code']

    def save_to_excel(self, output_file: str = 'po_analysis.xlsx'):
        try:
            out_path = os.path.join(self.folder_path, output_file)
            with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                # PO Summary
                po_summary_rows = []
                for f, d in self.all_pos_data.items():
                    po_summary_rows.append({
                        'PO Number': d['metadata'].get('po_number', ''),
                        'Vendor/Location': d['metadata'].get('vendor_location', ''),
                        'PO Date': d['metadata'].get('po_date', ''),
                        'Total Sales (PKR)': d['metadata'].get('total_amount', '')
                    })
                pd.DataFrame(po_summary_rows).to_excel(writer, sheet_name='PO Summary', index=False)

                # Quantity Summary
                summary_rows = []
                po_columns = [self.get_short_po_name(f, self.all_pos_data[f]['metadata']) for f in self.all_pos_data.keys()]

                for item_name, data in self.combined_items.items():
                    base_name = item_name.split(' (DIY')[0]
                    row = {'Row Labels': base_name, 'DIY Code': data.get('diy_code', '')}
                    for po_file in self.all_pos_data.keys():
                        po_col = self.get_short_po_name(po_file, self.all_pos_data[po_file]['metadata'])
                        row[po_col] = data['quantities_per_po'].get(po_file, 0.0)
                    row['Grand Total'] = sum(float(row.get(po, 0.0)) for po in po_columns)
                    summary_rows.append(row)

                summary_df = pd.DataFrame(summary_rows).fillna(0)
                cols = ['Row Labels', 'DIY Code'] + po_columns + ['Grand Total']
                cols = [c for c in cols if c in summary_df.columns]
                summary_df = summary_df[cols]
                summary_df.to_excel(writer, sheet_name='Quantity Summary', index=False)

            print(f"\n💾 Data saved to: {out_path}")
        except Exception as e:
            print(f"Error saving to Excel: {e}")

    def run_analysis(self):
        print(f"🚀 Starting Bulk PO Analysis...\n📁 Folder: {self.folder_path}")
        self.process_all_pdfs()
        self.save_to_excel()
        total_pos = sum(1 for d in self.all_pos_data.values() if d['success'])
        total_unique_items = len(self.combined_items)
        total_quantity = sum(d['total_quantity'] for d in self.combined_items.values())
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
