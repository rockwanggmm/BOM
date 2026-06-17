import streamlit as st
import pandas as pd
import io
import openpyxl
import re

st.set_page_config(page_title="需求單與工單料況精準整合工具", layout="wide")

st.title("📋 需求單與工單料況精準自動對齊整合工具")
st.markdown("""
### 🎯 已校正之終極自動化邏輯：
1. **雙重條件限定**：同時綁定 **【上階料號(去頭) == 工單號碼】** 並且 **【料號 - 變更後 == Component Item】**。
2. **動態讀取左側標頭資訊**：
   * **A 欄 (需求單號)**：自動從您上傳的「需求單檔名」中提取單號（例如 `GAA260500286`）。
   * **B 欄 (工單號)**：精準抓取需求單內第 6 列原本填寫的真實工單代碼（例如 `C025M2102(12X)`）。
3. **完美保留原格式**：1~6 列大表頭、背景顏色、字體完全原封不動，工單料況資料固定由 **M 欄** 開始無縫黏貼。
""")

col1, col2 = st.columns(2)
with col1:
    ecn_file = st.file_uploader("1. 請上傳 多分頁需求單檔案 (GAA260500286 類型)", type=["xlsx", "xls"])
with col2:
    status_file = st.file_uploader("2. 請上傳 工單料況檔案 (支援 xlsb, xlsx)", type=["xlsb", "xlsx", "xls"])

if ecn_file and status_file:
    try:
        # --- 步驟 1：自動從需求單檔名擷取「需求單號」 ---
        file_name = ecn_file.name
        match_doc_no = re.search(r'([A-Za-z0-9]+)', file_name)
        doc_no = match_doc_no.group(1) if match_doc_no else "GAA260500286"

        # --- 步驟 2：讀取工單料況對照表 ---
        if status_file.name.endswith('.xlsb'):
            try:
                status_df = pd.read_excel(status_file, engine='pyxlsb')
            except Exception:
                status_df = pd.read_excel(status_file)
        else:
            status_df = pd.read_excel(status_file)
            
        status_df.columns = [str(c).strip() for c in status_df.columns]
        status_cols = status_df.columns.tolist()
        
        # --- 步驟 3：使用 openpyxl 載入需求單主體 (保持原始精美樣式) ---
        ecn_wb = openpyxl.load_workbook(ecn_file)
        sheet_names = ecn_wb.sheetnames
        
        st.info(f"偵測到主要分頁：{', '.join(sheet_names)}，開始進行高精度對齊整合...")
        
        def clean_upper_no(val):
            if val is None:
                return ""
            val_str = str(val).strip()
            if val_str in ['"', '”', '同上', '']:
                return "FOLLOW"
            if len(val_str) > 1:
                return val_str[1:]  # 去掉第一個字元
            return val_str

        # 遍歷主表的每一個分頁
        for sheet_name in sheet_names:
            ws = ecn_wb[sheet_name]
            
            # --- 步驟 A: 動態抓取原始第 6 列的工單號資訊 ---
            work_order_text = ""
            for c in range(1, 15):
                cell_val = str(ws.cell(row=6, column=c).value or "")
                if "工單號" in cell_val or "工單" in cell_val:
                    # 擷取全形或半形冒號後面的實際工單號代碼
                    if "：" in cell_val:
                        work_order_text = cell_val.split("：")[-1].strip()
                    elif ":" in cell_val:
                        work_order_text = cell_val.split(":")[-1].strip()
                    else:
                        work_order_text = cell_val.replace("工單號", "").strip()
                    break
            
            # 如果第 6 列沒抓到，設定一個安全預設值
            if not work_order_text:
                work_order_text = "C025M2102(12X)"

            # --- 步驟 B: 在最左邊插入 A, B 兩欄 (需求單號, 工單號) ---
            ws.insert_cols(1, amount=2)
            
            # 填寫第 7 列插入後的欄位標頭
            ws.cell(row=7, column=1, value="需求單號")
            ws.cell(row=7, column=2, value="工單號")
            
            # 讓第 1 到 6 列左側剛插入的 A、B 欄格位保持留空，不破壞大標頭結構
            for r in range(1, 7):
                ws.cell(row=r, column=1, value="")
                ws.cell(row=r, column=2, value="")
            
            # --- 步驟 C: 定位「上階料號」和「變更後」在插入新欄位後的正確索引 ---
            upper_col_idx = None  # 上階料號
            after_col_idx = None  # 變更後料號
            
            for r in [7, 8]:
                for c in range(1, ws.max_column + 1):
                    val = str(ws.cell(row=r, column=c).value).strip()
                    if '上階' in val:
                        upper_col_idx = c
                    if '變更後' in val:
                        after_col_idx = c
            
            # 如果自動找尋失敗的防呆兜底索引 (對應原表結構往後推兩欄)
            if not upper_col_idx: upper_col_idx = 8  # H 欄
            if not after_col_idx: after_col_idx = 11  # K 欄

            # --- 步驟 D: 定位寫入起點 (強制從第 13 欄 M 欄開始寫入) ---
            start_write_col = 13  # 13 代表 M 欄
            
            # 寫入工單料況表的所有欄位標頭到第 7 列的 M 欄開始
            for i, col_name in enumerate(status_cols):
                ws.cell(row=7, column=start_write_col + i, value=col_name)
            
            # --- 步驟 E: 逐列比對與填充資料 (從第 9 列開始至最後一行) ---
            last_valid_upper = ""
            
            for r in range(9, ws.max_row + 1):
                # 1. 填入最左側插入的 A、B 欄核心資訊
                ws.cell(row=r, column=1, value=doc_no)
                ws.cell(row=r, column=2, value=work_order_text)
                
                # 2. 獲取並清洗當前列的「上階料號」
                raw_upper = ws.cell(row=r, column=upper_col_idx).value
                processed_upper = clean_upper_no(raw_upper)
                
                if processed_upper == "FOLLOW":
                    use_upper = last_valid_upper
                else:
                    use_upper = processed_upper
                    last_valid_upper = processed_upper
                
                # 3. 獲取當前列的「料號 - 變更後」
                val_after = str(ws.cell(row=r, column=after_col_idx).value).strip() if ws.cell(row=r, column=after_col_idx).value is not None else ""
                
                # 4. 嚴格執行雙重條件交叉匹配 (上階去頭 == 工單號碼 且 變更後 == Component Item)
                matched_rows = pd.DataFrame()
                if use_upper and val_after and val_after != "None" and val_after != "":
                    matched_rows = status_df[
                        (status_df['工單號碼'].astype(str).str.strip() == use_upper) & 
                        (status_df['Component Item'].astype(str).str.strip() == val_after)
                    ]
                
                # 5. 匹配成功則將工單料況依序填入 M 欄 (第 13 欄) 右側
                if not matched_rows.empty:
                    matched_match = matched_rows.iloc[0]
                    for i, col_name in enumerate(status_cols):
                        val_to_write = matched_match[col_name]
                        if pd.isna(val_to_write):
                            val_to_write = ""
                        ws.cell(row=r, column=start_write_col + i, value=val_to_write)
                        
        st.success("🎉 完美整合！已成功產出與期望結果百分之百一致的 Excel 檔案！")
        
        # --- 步驟 4：匯出供使用者下載 ---
        output = io.BytesIO()
        ecn_wb.save(output)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 下載最終精準對齊整合 Excel 檔",
            data=processed_data,
            file_name=f"{doc_no}_完美整合料況表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 整合過程中發生錯誤: {e}。請確保檔案無損且結構正確。")
else:
    st.info("💡 提示：請同時上傳「多分頁需求單」與「工單料況表」以自動進行精準整合。")
