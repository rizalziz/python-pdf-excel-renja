import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pdfplumber
import pandas as pd
import os
import logging
import traceback
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter


logging.basicConfig(
    filename='converter.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class PDFToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF ke Excel Converter - SIPD-RI")
        self.root.geometry("600x500")
        
        self.pdf_files = []
        self.selected_pdf = tk.StringVar()
        
        self.setup_ui()
        self.scan_pdf_files()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        ttk.Label(main_frame, text="PDF ke Excel Converter", 
                 font=("Arial", 16, "bold")).grid(row=0, column=0, pady=(0, 20))
        
        ttk.Label(main_frame, text="Pilih File PDF:", 
                 font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        pdf_combo = ttk.Combobox(main_frame, textvariable=self.selected_pdf, 
                                state="readonly", width=50)
        pdf_combo.grid(row=2, column=0, pady=(0, 20), sticky=(tk.W, tk.E))
        
        ttk.Button(main_frame, text="Scan Folder PDF", 
                  command=self.scan_pdf_files).grid(row=3, column=0, pady=(0, 10), sticky=tk.W)
        
        ttk.Separator(main_frame, orient="horizontal").grid(
            row=4, column=0, sticky=(tk.W, tk.E), pady=20)
        
        ttk.Label(main_frame, text="Opsi Konversi:", 
                 font=("Arial", 10, "bold")).grid(row=5, column=0, sticky=tk.W, pady=(0, 10))
        
        self.remove_footer_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(main_frame, text="Hapus footer 'SIPD-RI***'", 
                       variable=self.remove_footer_var).grid(row=6, column=0, sticky=tk.W, pady=2)
        
        self.auto_width_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(main_frame, text="Auto lebar kolom Excel", 
                       variable=self.auto_width_var).grid(row=7, column=0, sticky=tk.W, pady=2)
        
        ttk.Separator(main_frame, orient="horizontal").grid(
            row=8, column=0, sticky=(tk.W, tk.E), pady=20)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=9, column=0, pady=(10, 0))
        
        ttk.Button(button_frame, text="Konversi ke Excel", 
                  command=self.convert_pdf_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Keluar", 
                  command=self.root.quit).pack(side=tk.LEFT, padx=5)
        
        self.status_label = ttk.Label(main_frame, text="Siap", 
                                     foreground="blue")
        self.status_label.grid(row=10, column=0, pady=(20, 0))
    
    def scan_pdf_files(self):
        folder = os.getcwd()
        self.pdf_files = [f for f in os.listdir(folder) 
                         if f.lower().endswith('.pdf')]
        
        if hasattr(self, '_pdf_combo'):
            self._pdf_combo['values'] = self.pdf_files
        else:
            self._pdf_combo = self.root.nametowidget(
                self.root.children['!frame'].children['!combobox'])
            self._pdf_combo['values'] = self.pdf_files
        
        if self.pdf_files:
            self.selected_pdf.set(self.pdf_files[0])
            self.status_label.config(text=f"Ditemukan {len(self.pdf_files)} file PDF")
        else:
            self.selected_pdf.set("")
            self.status_label.config(text="Tidak ada file PDF di folder ini")
    
    def clean_table_data(self, table):
        if not table:
            return []
        
        cleaned = []
        for row in table:
            cleaned_row = [str(cell).strip() if cell is not None else "" 
                          for cell in row]
            if any(cell for cell in cleaned_row):
                cleaned.append(cleaned_row)
        
        return cleaned
    
    def _clean_numeric_columns(self, df):
        if df.empty:
            return df
        
        target_cols = [9, 10, 11, 12, 19]
        
        for col_idx in target_cols:
            if col_idx < len(df.columns):
                col_name = df.columns[col_idx]
                df[col_name] = df[col_name].apply(self._clean_numeric_value)
                df[col_name] = df[col_name].str.replace(',', '.', regex=False)
                df[col_name] = pd.to_numeric(df[col_name], errors='coerce')
        
        return df
    
    def _clean_numeric_value(self, value):
        if pd.isna(value) or str(value).strip() == '':
            return value
        
        s = str(value).strip()
        s = s.replace('.', '')
        cleaned = ''.join(c for c in s if c.isdigit() or c == ',')
        return cleaned
    
    def remove_footer_rows(self, df):
        if self.remove_footer_var.get():
            mask = ~df.apply(lambda row: any(
                'sipd-ri' in str(cell).lower() for cell in row if pd.notna(cell)
            ), axis=1)
            df = df[mask].reset_index(drop=True)
        return df
    
    def format_excel(self, ws):
        if self.auto_width_var.get():
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width
        
        header_font = Font(bold=True, size=11)
        header_alignment = Alignment(horizontal="center", vertical="center", 
                                     wrap_text=True)
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border
        
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=True)
    
    def _normalize_text(self, text):
        if text is None:
            return ""
        text = str(text).lower().strip()
        text = text.replace('\n', ' ').replace('\r', ' ')
        text = ' '.join(text.split())
        return text
    
    def _is_header_row(self, row, header):
        if len(row) != len(header):
            return False
        
        norm_row = [self._normalize_text(cell) for cell in row]
        norm_header = [self._normalize_text(cell) for cell in header]
        
        matches = sum(1 for a, b in zip(norm_row, norm_header) if a == b and a)
        return matches >= len(header) * 0.7
    
    def _is_numbering_row(self, row):
        cells = [str(cell).strip() for cell in row if cell is not None]
        if not cells:
            return False
        
        first_cell = cells[0]
        if not first_cell.isdigit():
            return False
        
        numeric_count = sum(1 for c in cells if c.replace('.', '').replace('-', '').isdigit())
        return numeric_count == len(cells)
    
    def _is_header_by_keywords(self, row):
        keywords = ['no', 'kode', 'urusan', 'bidang', 'program', 'kegiatan', 'sub kegiatan']
        row_text = ' '.join(self._normalize_text(cell) for cell in row)
        matches = sum(1 for kw in keywords if kw in row_text)
        return matches >= 3
    
    def _row_values_match_header(self, row, header_values):
        row_texts = [self._normalize_text(cell) for cell in row]
        header_texts = [self._normalize_text(cell) for cell in header_values]
        
        if len(row_texts) != len(header_texts):
            return False
        
        matches = sum(1 for a, b in zip(row_texts, header_texts) if a == b and a)
        return matches >= len(header_texts) * 0.7
    
    def _row_values_match_header_rows(self, row, header_rows):
        for header_row in header_rows:
            if self._row_values_match_header(row, header_row):
                return True
        return False

    def _is_potential_parent(self, val):
        val = str(val).strip()
        if not val:
            return False
        if len(val) == 1 and val.isdigit():
            return True
        if '.' in val:
            return True
        return False
    
    def _find_direct_children(self, data_df):
        value_to_idx = {}
        for idx in range(len(data_df)):
            if len(data_df.columns) > 1:
                val = str(data_df.iloc[idx, 1]).strip()
                if val and self._is_potential_parent(val):
                    value_to_idx[val] = idx
        
        groups = []
        
        for idx in range(len(data_df)):
            if len(data_df.columns) > 1:
                val = str(data_df.iloc[idx, 1]).strip()
                if not val or not self._is_potential_parent(val):
                    continue
                
                prefix = val + "."
                parent_segments = len(val.split('.'))
                children = []
                
                for child_val, child_idx in value_to_idx.items():
                    if child_val.startswith(prefix):
                        child_segments = len(child_val.split('.'))
                        if child_segments == parent_segments + 1:
                            children.append(child_idx)
                
                if children:
                    groups.append((idx, children))
        
        return groups
    
    def _apply_sum_formulas(self, ws, data_df, header_row_count):
        groups = self._find_direct_children(data_df)
        if not groups:
            return
        
        target_cols = [9, 10, 11, 12, 19]
        
        for parent_idx, children_indices in groups:
            parent_excel_row = parent_idx + header_row_count + 1
            child_rows = [child_idx + header_row_count + 1 for child_idx in children_indices]
            child_rows.sort()
            
            for col_idx in target_cols:
                if col_idx < len(data_df.columns):
                    col_letter = get_column_letter(col_idx + 1)
                    if len(child_rows) == 1:
                        formula = f"={col_letter}{child_rows[0]}"
                    else:
                        child_refs = ",".join(f"{col_letter}{r}" for r in child_rows)
                        formula = f"=SUM({child_refs})"
                    ws.cell(row=parent_excel_row, column=col_idx + 1, value=formula)
    
    def convert_pdf_to_excel(self):
        if not self.selected_pdf.get():
            messagebox.showwarning("Peringatan", "Pilih file PDF terlebih dahulu!")
            return
        
        pdf_path = self.selected_pdf.get()
        self.status_label.config(text="Sedang memproses...")
        self.root.update()
        
        try:
            logger.info(f"Mulai konversi: {pdf_path}")
            all_tables = []
            master_header_rows = None
            
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                for i, page in enumerate(pdf.pages):
                    self.status_label.config(
                        text=f"Memproses halaman {i+1}/{total_pages}...")
                    self.root.update()
                    
                    tables = page.extract_tables()
                    for table in tables:
                        cleaned_table = self.clean_table_data(table)
                        if not cleaned_table:
                            continue
                        
                        if master_header_rows is None:
                            master_header_rows = cleaned_table[:4]
                            all_tables.append(cleaned_table[4:])
                        else:
                            first_rows = cleaned_table[:4]
                            is_repeated_header = any(
                                self._row_values_match_header(first_rows[j], master_header_rows[j])
                                for j in range(min(4, len(first_rows), len(master_header_rows)))
                            )
                            if is_repeated_header:
                                all_tables.append(cleaned_table[4:])
                            else:
                                all_tables.append(cleaned_table)
            
            if not all_tables:
                messagebox.showinfo("Info", "Tidak ada tabel yang ditemukan di PDF ini.")
                self.status_label.config(text="Siap")
                return
            
            flat_data = []
            for table in all_tables:
                flat_data.extend(table)
            
            if not flat_data and not master_header_rows:
                messagebox.showinfo("Info", "Tidak ada data di PDF ini.")
                self.status_label.config(text="Siap")
                return
            
            logger.info(f"PDF berhasil dibuka, total halaman: {total_pages}")
            
            header_df = pd.DataFrame(master_header_rows)
            data_df = pd.DataFrame(flat_data)
            
            logger.info(f"Data awal: {len(data_df)} baris")
            
            if not data_df.empty:
                data_df = data_df[~data_df.apply(lambda row: self._row_values_match_header_rows(row.tolist(), master_header_rows), axis=1)].reset_index(drop=True)
                logger.info(f"Setelah filter header: {len(data_df)} baris")
            
            data_df = self.remove_footer_rows(data_df)
            data_df = data_df[~data_df.apply(lambda row: self._is_numbering_row(row.tolist()), axis=1)].reset_index(drop=True)
            data_df = data_df.dropna(how='all').reset_index(drop=True)
            data_df.columns = [str(c) for c in data_df.columns]
            data_df = data_df.loc[:, ~data_df.columns.str.contains('^Unnamed')]
            data_df = self._clean_numeric_columns(data_df)
            
            logger.info(f"Setelah filter footer/numbering/NaN: {len(data_df)} baris")
            
            excel_path = pdf_path.replace('.pdf', '.xlsx')
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Data PDF"
            
            for r_idx, row in enumerate(dataframe_to_rows(header_df, index=False, header=False), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)
            
            start_row = len(master_header_rows) + 1
            for r_idx, row in enumerate(dataframe_to_rows(data_df, index=False, header=False), start_row):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)
            
            self._apply_sum_formulas(ws, data_df, len(master_header_rows))
            self.format_excel(ws)
            
            wb.save(excel_path)
            logger.info(f"File Excel berhasil disimpan: {excel_path}")
            
            self.status_label.config(text=f"Selesai! Disimpan: {excel_path}")
            messagebox.showinfo("Sukses", 
                               f"Konversi berhasil!\n\nFile Excel: {excel_path}\n"
                               f"Header: {len(master_header_rows)} baris\n"
                               f"Data: {len(data_df)} baris")
            
        except Exception as e:
            error_msg = f"Terjadi kesalahan:\n{str(e)}"
            logger.error(error_msg)
            logger.error(traceback.format_exc())
            messagebox.showerror("Error", error_msg)
            self.status_label.config(text="Error")


def main():
    root = tk.Tk()
    app = PDFToExcelConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()