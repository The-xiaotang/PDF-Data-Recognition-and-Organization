import re
import pandas as pd
import fitz  # PyMuPDF
import sys
import pdfplumber
from PySide6.QtCore import QThread, Signal
from pdf2docx import Converter
from docx2pdf import convert

class WorkerSignals(QThread):
    finished = Signal(str)  # Message
    error = Signal(str)
    progress = Signal(int)

class PdfToWordWorker(WorkerSignals):
    def __init__(self, pdf_path, docx_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.docx_path = docx_path

    def run(self):
        try:
            cv = Converter(self.pdf_path)
            cv.convert(self.docx_path, start=0, end=None)
            cv.close()
            self.finished.emit(f"Successfully converted to {self.docx_path}")
        except Exception as e:
            self.error.emit(str(e))

class WordToPdfWorker(WorkerSignals):
    def __init__(self, docx_path, pdf_path):
        super().__init__()
        self.docx_path = docx_path
        self.pdf_path = pdf_path

    def run(self):
        try:
            convert(self.docx_path, self.pdf_path)
            self.finished.emit(f"Successfully converted to {self.pdf_path}")
        except Exception as e:
            self.error.emit(str(e))

class PdfToExcelWorker(WorkerSignals):
    def __init__(self, pdf_path, excel_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.excel_path = excel_path

    def run(self):
        try:
            doc = fitz.open(self.pdf_path)
            text = ""
            for page in doc:
                text += page.get_text() + "\n"
            doc.close()
            
            # Safely print text for debugging
            try:
                safe_text = text.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding)
                print("--- Extracted Text Start ---")
                print(safe_text)
                print("--- Extracted Text End ---")
            except Exception as print_error:
                print(f"Could not print text to console: {print_error}")

            data = {}
            
            def extract(pattern, default=""):
                match = re.search(pattern, text)
                return match.group(1).strip() if match else default

            # Basic Info
            data['检测号'] = extract(r"检测号[：:]\s*([A-Za-z0-9]+)")
            data['报告系统版本号'] = extract(r"报告系统版本号\s*([A-Za-z0-9\s\.]+?)(?=\s+生信分析版本号|$)")
            data['生信分析版本号'] = extract(r"生信分析版本号\s*([A-Za-z0-9\s\.]+?)(?=\s+上机号|$)")
            data['上机号'] = extract(r"上机号[：:]\s*(\d+)")
            
            # Simplified Name/Gender/Age extraction
            match_name = re.search(r"姓\s*名[：:]\s*(\S+)", text)
            if not match_name:
                 match_name = re.search(r"姓\s*名\s*[：:]?\s*(\S+)", text)
            data['姓名'] = match_name.group(1).strip() if match_name else ""

            match_gender = re.search(r"性\s*别[：:]\s*([男女])", text)
            if not match_gender:
                 match_gender = re.search(r"性\s*别\s*[：:]?\s*([男女])", text)
            data['性别'] = match_gender.group(1).strip() if match_gender else ""

            match_age = re.search(r"年\s*龄[：:]\s*(\d+)\s*岁?", text)
            if not match_age:
                 match_age = re.search(r"年\s*龄\s*[：:]?\s*(\d+)\s*岁?", text)
            data['年龄'] = (match_age.group(1).strip() + "岁") if match_age else ""
            
            data['采样日期'] = extract(r"采样日期[：:]\s*(\d{4}-\d{2}-\d{2})")
            
            match_specimen = re.search(r"标本类型[：:]\s*(.+?)(?=\s+住院号|\s+病理号|\n|$)", text)
            data['标本类型'] = match_specimen.group(1).strip() if match_specimen else ""
            
            data['住院号'] = extract(r"住院号[：:]\s*(\S+)")
            data['病理号'] = extract(r"病理号[：:]\s*(\S+)")
            
            # ID Card Extraction Logic
            match_id_label = re.search(r"身\s*份\s*证\s*号[：:]\s*(.*?)(?=\s+姓\s*名|\s+送\s*检|\n|$)", text)
            id_card = match_id_label.group(1).strip() if match_id_label else ""
            id_card = id_card.replace(" ", "").replace("-", "") 

            is_valid_id = re.match(r"^(\d{15}|\d{17}[0-9Xx])$", id_card)
            
            if not is_valid_id:
                 match_global = re.search(r"(?<!\d)(\d{18}|\d{17}[0-9Xx])(?!\d)", text)
                 if match_global:
                     id_card = match_global.group(1)
            
            data['身份证号'] = id_card

            data['送检日期'] = extract(r"送检日期[：:]\s*(\d{4}-\d{2}-\d{2})")
            data['病历号'] = extract(r"病历号[：:]\s*(\d+)")
            data['送检医生'] = extract(r"送检医生[：:]\s*(\S+)")
            
            match_unit = re.search(r"送检单位[：:]\s*(.+?)(?=\s+身份证号|\n|$)", text)
            unit_raw = match_unit.group(1).strip() if match_unit else ""
            unit_clean = re.sub(r"\d{18}|\d{17}[Xx]", "", unit_raw).strip()
            data['送检单位'] = unit_clean

            data['检测项目'] = extract(r"检测项目[：:]\s*(.+)")
            data['检测方法'] = extract(r"检测方法[：:]\s*(.+)")
            data['送检材料'] = extract(r"送检材料[：:]\s*(.+)")
            data['临床诊断'] = extract(r"临床诊断[：:]\s*(.+)")

            columns = [
                '检测号', '报告系统版本号', '生信分析版本号', '上机号', 
                '姓名', '性别', '年龄', '采样日期', '标本类型', 
                '住院号', '病理号', '身份证号', '送检日期', '病历号', 
                '送检医生', '送检单位', '检测项目', '检测方法', 
                '送检材料', '临床诊断'
            ]
            
            df = pd.DataFrame([data], columns=columns)
            df.to_excel(self.excel_path, index=False)
            
            self.finished.emit(f"Successfully extracted to {self.excel_path}")
        except Exception as e:
            self.error.emit(str(e))

class PdfToRearrangementWorker(WorkerSignals):
    def __init__(self, pdf_path, excel_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.excel_path = excel_path

    def run(self):
        try:
            # Phase 1: Basic text extraction for Name and Detection No (PyMuPDF is faster/reliable for plain text)
            doc = fitz.open(self.pdf_path)
            text = ""
            for page in doc:
                text += page.get_text() + "\n"
            doc.close()

            data_header = {}
            match_name = re.search(r"姓\s*名[：:]\s*(\S+)", text)
            if not match_name:
                 match_name = re.search(r"姓\s*名\s*[：:]?\s*(\S+)", text)
            data_header['姓名'] = match_name.group(1).strip() if match_name else ""
            
            match_det_no = re.search(r"检测号[：:]\s*([A-Za-z0-9]+)", text)
            data_header['检测号'] = match_det_no.group(1).strip() if match_det_no else ""

            # Phase 2: Table extraction for Rearrangement Data (using pdfplumber)
            # The image shows a table with headers: 重排基因 | 左断裂点位置 | 右断裂点位置
            
            all_rows = []
            found_table = False
            
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table: continue
                        
                        # Check header row
                        header = [str(cell).replace('\n', '').strip() for cell in table[0] if cell]
                        header_str = "".join(header)
                        
                        # Look for key headers from user image
                        if "重排基因" in header_str and "断裂点" in header_str:
                            found_table = True
                            
                            # Identify column indices
                            def find_idx(keywords):
                                for idx, col in enumerate(header):
                                    if any(k in col for k in keywords):
                                        return idx
                                return -1

                            idx_gene = find_idx(["重排基因"])
                            idx_left = find_idx(["左断裂", "Left"])
                            idx_right = find_idx(["右断裂", "Right"])
                            
                            # Extract rows
                            for row in table[1:]:
                                if not row or len(row) < 2: continue
                                
                                def get_cell(idx):
                                    if idx != -1 and idx < len(row):
                                        val = row[idx]
                                        return val.replace('\n', ' ').strip() if val else ""
                                    return ""
                                
                                gene_val = get_cell(idx_gene)
                                # Filter out empty rows or placeholder rows if needed
                                if not gene_val or gene_val == "-": 
                                    gene_val = "无" # Or keep as "-" based on preference, user said "有就写有，没有就写无"
                                    # If it is explicitly "-", let's treat it as "无" for consistency if desired, 
                                    # but if it's a real gene name, keep it.
                                
                                left_val = get_cell(idx_left)
                                right_val = get_cell(idx_right)
                                
                                if gene_val == "无" and not left_val and not right_val:
                                    # If row is essentially empty/dash, we capture it as one entry of "无"
                                    pass

                                row_data = {
                                    '姓名': data_header['姓名'],
                                    '检测号': data_header['检测号'],
                                    '重排基因': gene_val,
                                    '左断裂点位': left_val if left_val else "-",
                                    '右断裂点位': right_val if right_val else "-"
                                }
                                all_rows.append(row_data)

            # If no table found or empty table, create a default "None" row
            if not all_rows:
                # Fallback to regex text search if table extraction failed (legacy logic)
                # But since user asked for "Correction/Efficiency", table extraction is primary.
                # If table is truly missing, we assume "无"
                all_rows.append({
                    '姓名': data_header['姓名'],
                    '检测号': data_header['检测号'],
                    '重排基因': "无",
                    '左断裂点位': "-",
                    '右断裂点位': "-"
                })

            columns = ['姓名', '检测号', '重排基因', '左断裂点位', '右断裂点位']
            df = pd.DataFrame(all_rows, columns=columns)
            df.to_excel(self.excel_path, index=False)
            
            self.finished.emit(f"Successfully extracted {len(all_rows)} rearrangement records to {self.excel_path}")
        except Exception as e:
            self.error.emit(str(e))

class PdfToMutationWorker(WorkerSignals):
    def __init__(self, pdf_path, excel_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.excel_path = excel_path

    def run(self):
        try:
            # First pass: Extract Detection No using PyMuPDF (lighter/faster for text)
            doc = fitz.open(self.pdf_path)
            text = ""
            for page in doc:
                text += page.get_text() + "\n"
            doc.close()
            
            match_det_no = re.search(r"检测号[：:]\s*([A-Za-z0-9]+)", text)
            detection_no = match_det_no.group(1).strip() if match_det_no else ""

            # Second pass: Extract Table Data using pdfplumber
            # We are looking for columns: 突变基因, 转录本 ID, 外显子, 核苷酸改变, 氨基酸改变, 突变频率
            # Note: Column names in PDF might vary slightly (e.g. "Gene", "Transcript", "Exon", "c.Change", "p.Change", "VAF")
            # We'll try to identify the table by its header or content.
            
            all_rows = []
            
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        # Heuristic: Check if table has relevant headers
                        # Flatten table to check headers easily
                        if not table: continue
                        
                        header = [str(cell).replace('\n', '') for cell in table[0] if cell]
                        # Check for keywords like "基因" (Gene), "突变" (Mutation), "频率" (Frequency)
                        header_str = " ".join(header)
                        
                        if "基因" in header_str and "改变" in header_str:
                            # This looks like the mutation table
                            # Map columns based on index
                            # Standard format: Gene | ... | Transcript | Exon | c. | p. | VAF
                            
                            # Let's try to find indices dynamically
                            def find_idx(keywords):
                                for idx, col in enumerate(header):
                                    if any(k in col for k in keywords):
                                        return idx
                                return -1

                            idx_gene = find_idx(["基因", "Gene"])
                            idx_trans = find_idx(["转录本", "Transcript"])
                            idx_exon = find_idx(["外显子", "Exon"])
                            idx_nuc = find_idx(["核苷酸", "c."])
                            idx_aa = find_idx(["氨基酸", "p."])
                            idx_vaf = find_idx(["频率", "VAF", "%"])
                            
                            # Iterate data rows (skip header)
                            for row in table[1:]:
                                if not row or len(row) < 3: continue # Skip empty or short rows
                                
                                # Helper to get cell data safely
                                def get_cell(idx):
                                    if idx != -1 and idx < len(row):
                                        val = row[idx]
                                        return val.replace('\n', ' ').strip() if val else ""
                                    return ""

                                # Check if row is valid (has gene name)
                                gene_val = get_cell(idx_gene)
                                if not gene_val: continue
                                
                                row_data = {
                                    '检测号': detection_no,
                                    '突变基因': gene_val,
                                    '转录本 ID': get_cell(idx_trans),
                                    '外显子': get_cell(idx_exon),
                                    '核苷酸改变': get_cell(idx_nuc),
                                    '氨基酸改变': get_cell(idx_aa),
                                    '突变频率': get_cell(idx_vaf)
                                }
                                all_rows.append(row_data)

            # Create DataFrame
            columns = ['检测号', '突变基因', '转录本 ID', '外显子', '核苷酸改变', '氨基酸改变', '突变频率']
            df = pd.DataFrame(all_rows, columns=columns)
            
            # Save
            df.to_excel(self.excel_path, index=False)
            
            self.finished.emit(f"Successfully extracted {len(all_rows)} mutations to {self.excel_path}")
            
        except Exception as e:
            self.error.emit(str(e))
