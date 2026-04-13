import streamlit as st
import pandas as pd
import re
import math
import io
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

# Word 處理相關
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 網頁配置
st.set_page_config(page_title="題庫轉換神器 V11", page_icon="📝", layout="centered")

st.title("📝 題庫轉換 Web 版")
st.markdown("已移除 PDF 功能。支援：自動偵測選項數、詳解全紫排版、出處標註、Excel 自動列高")

# ==================== UI 介面 ====================
uploaded_file = st.file_uploader("選取原始題庫 (.xlsx)", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    exam_title_input = st.text_input("考卷標題 (Word 大標題)", "113年 專業科目測驗")
with col2:
    excel_filename_input = st.text_input("Excel 輸出檔名", "精修題庫資料")

# 強制確保輸入不為 None
exam_title = str(exam_title_input) if exam_title_input else "測驗題庫"
excel_filename = str(excel_filename_input) if excel_filename_input else "精修題庫"

def sanitize(name):
    # 確保輸入是字串且移除非法字元
    name = str(name)
    return re.sub(r'[\\/:*?"<>|]', '_', name)

# ==================== 處理邏輯 ====================
if uploaded_file is not None:
    if st.button("🚀 開始轉換資料", use_container_width=True):
        try:
            # 1. 智慧讀取 Excel
            df_raw = pd.read_excel(uploaded_file, header=None)
            header_idx = -1
            for i, row in df_raw.iterrows():
                row_str = " ".join([str(x) for x in row if pd.notna(x)])
                if any(k in row_str for k in ["題目", "Question", "內容"]):
                    header_idx = i
                    break
            
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=header_idx if header_idx != -1 else 0)
            df.columns = [str(c).strip() for c in df.columns]
            cols = df.columns.tolist()

            # 2. 欄位自動對應
            q_col = next((c for c in cols if any(k in c for k in ["題目", "內容", "Question"])), None)
            ans_col = next((c for c in cols if any(k in c for k in ["答案", "Answer", "Correct", "判定"])), None)
            expl_col = next((c for c in cols if any(k in c for k in ["詳解", "說明", "Explanation"])), None)
            src_col = next((c for c in cols if any(k in c for k in ["出處", "來源", "Source"])), None)
            
            # 動態偵測選項數
            opt_cols = [c for c in cols if ("選項" in c or "Option" in c or re.match(r'^[A-F]$', c)) 
                        and c not in [q_col, ans_col, expl_col, src_col]]
            opt_cols.sort()
            max_opts = len(opt_cols)
            opt_labels = [chr(65 + i) for i in range(max_opts)]

            processed_rows = []
            def clean(t): 
                if pd.isna(t): return ""
                return str(t).strip().replace('$', '')

            # 3. 資料整理
            for i, row in df.iterrows():
                q_txt = clean(row.get(q_col, ""))
                if not q_txt or q_txt.lower() == "nan": continue # 避免抓到空行或標題
                
                row_dict = {'題號': len(processed_rows) + 1, '題目內容': q_txt}
                for idx, lbl in enumerate(opt_labels):
                    row_dict[f'選項{lbl}'] = clean(row.get(opt_cols[idx], ""))
                
                ans_raw = clean(row.get(ans_col, ""))
                ans_match = re.search(r'[A-F]', ans_raw.upper())
                row_dict['正確答案'] = ans_match.group(0) if ans_match else ""
                row_dict['針對各選項之詳解'] = clean(row.get(expl_col, ""))
                row_dict['出處'] = clean(row.get(src_col, ""))
                processed_rows.append(row_dict)

            if not processed_rows:
                st.error("❌ 找不到有效的題目資料，請確認 Excel 內容。")
                st.stop()

            # ==================== 產出 Excel ====================
            output_df = pd.DataFrame(processed_rows)
            excel_out = io.BytesIO()
            output_df.to_excel(excel_out, index=False)
            excel_out.seek(0)
            
            wb = load_workbook(excel_out)
            ws = wb.active
            
            # 動態設定欄寬
            ans_idx = 3 + max_opts
            col_widths = {'A': 8, 'B': 45}
            for i in range(max_opts): col_widths[get_column_letter(3+i)] = 30
            col_widths[get_column_letter(ans_idx)] = 15
            col_widths[get_column_letter(ans_idx+1)] = 60
            col_widths[get_column_letter(ans_idx+2)] = 40

            for letter, width in col_widths.items(): ws.column_dimensions[letter].width = width
            
            thin = Side(style='thin')
            border = Border(top=thin, bottom=thin, left=thin, right=thin)
            
            for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
                max_h_lines = 1
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                    if r_idx == 1:
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    if cell.column_letter in ['A', get_column_letter(ans_idx)]:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    cw = col_widths.get(cell.column_letter, 20)
                    est = math.ceil((len(str(cell.value)) * 1.8) / cw)
                    if est > max_h_lines: max_h_lines = est
                if r_idx > 1: ws.row_dimensions[r_idx].height = max_h_lines * 18
            
            final_excel = io.BytesIO()
            wb.save(final_excel)
            
            # ==================== 產出 Word ====================
            doc = Document()
            sec = doc.sections[0]
            sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1.27)
            doc.styles['Normal'].font.name = 'Times New Roman'
            doc.styles['Normal'].element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
            
            PURPLE, BLUE = RGBColor(112, 48, 160), RGBColor(0, 50, 150)

            title_p = doc.add_paragraph()
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = title_p.add_run(exam_title)
            run.bold = True
            run.font.size = Pt(16)

            for r in processed_rows:
                doc.add_paragraph(f"{r['題號']}. {r['題目內容']}").paragraph_format.space_after = Pt(6)
                for lbl in opt_labels:
                    txt = r.get(f'選項{lbl}', '')
                    if txt:
                        op = doc.add_paragraph(f"({lbl}) {txt}")
                        op.paragraph_format.left_indent, op.paragraph_format.space_after = Pt(18), Pt(0)
                
                ans_p = doc.add_paragraph()
                ans_p.paragraph_format.space_before = Pt(6)
                ans_p.add_run("Ans : ").bold = True
                ans_p.add_run(f"({r['正確答案']})")

                expl = str(r['針對各選項之詳解'])
                if expl and expl.lower() != "nan" and expl.strip():
                    h = doc.add_paragraph()
                    h.paragraph_format.space_before, h.paragraph_format.space_after = Pt(4), Pt(0)
                    run = h.add_run("詳解 :"); run.bold, run.font.color.rgb = True, PURPLE
                    for line in expl.split('\n'):
                        if not line.strip(): continue
                        lp = doc.add_paragraph(); lp.paragraph_format.left_indent, lp.paragraph_format.space_after = Pt(18), Pt(2)
                        m = re.match(r'^([A-F])\s*([\(（].*?[\)）]|[:：])', line.strip())
                        if m:
                            r1 = lp.add_run(m.group(0)); r1.bold, r1.font.color.rgb = True, PURPLE
                            r2 = lp.add_run(line.strip()[len(m.group(0)):]); r2.font.color.rgb = PURPLE
                        else:
                            lp.add_run(line.strip()).font.color.rgb = PURPLE
                
                src = str(r['出處'])
                if src and src.lower() != "nan" and src.strip():
                    sp = doc.add_paragraph(); sp.paragraph_format.space_before = Pt(2)
                    rl = sp.add_run("出處 : "); rl.bold, rl.font.color.rgb = True, BLUE
                    sp.add_run(src).font.color.rgb = BLUE
                doc.add_paragraph("")

            word_out = io.BytesIO()
            doc.save(word_out)

            # --- 下載按鈕 ---
            st.success("✅ 轉換完成！")
            st.download_button("📊 下載 Excel 精修版", final_excel.getvalue(), f"{sanitize(excel_filename)}.xlsx")
            st.download_button("📄 下載 Word 考卷", word_out.getvalue(), f"{sanitize(exam_title)}.docx")

        except Exception as e:
            st.error(f"轉換過程出錯（建議檢查 Excel 格式）：{e}")
            st.exception(e) # 這會顯示詳細錯誤，方便除錯
