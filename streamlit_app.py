import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="需求單與工單料況自動整合工具", layout="wide")

st.title("📋 需求單(GAA260500286)與工單料況精準整合工具")
st.markdown("""
### 📌 完美修正後的整合邏輯：
1. **雙重條件精準匹配**：系統會同時比對 **【(需求單)上階料號去掉首字 == (工單料況)工單號碼】** 並且 **【(需求單)變更後料號 == (工單料況)Component Item】**。
2. **完整格式保留**：保留需求單原始檔案上方的 1~6 列特殊表頭資訊，不破換原檔案排版。
3. **欄位完美串接**：
   * 在最左側自動插入 `需求單號`、`工單號`。
   * 在最右側（原本備註欄的右邊）自動黏貼工單料況對應的完整欄位。
""")

col1, col2 = st.columns(2)
with col1:
    ecn_file = st.file_uploader("1. 請上傳 多分頁需求單檔案 (GAA260500286 類型)", type=["xlsx", "xls"])
with col2:
    status_file = st.file_uploader("2. 請上傳 工單料況檔案 (支援 xlsb, xlsx)", type=["xlsb", "xlsx", "xls"])

if ecn_file and status_file:
    try:
        # 1. 讀取工單料況對照表（右表）
        if status_file.name.endswith('.xlsb'):
            try:
                status_df = pd.read_excel(status_file, engine='pyxlsb')
            except Exception:
                status_df = pd.read_excel(status_file)
        else:
            status_df = pd.read_excel(status_file)
            
        # 清洗右表欄位
        status_df.columns = [str(c).strip() for c in status_df.columns]
        
        # 建立右表的 雙重條件組合索引 鍵，確保精準唯一性
        # Key: 工單號碼 + "_" + Component Item
        status_df['__join_key__'] = status_df['工單號碼'].astype(str).str.strip() + "_" + status_df['Component Item'].astype(str).str.strip()
        
        # 2. 使用 openpyxl 載入需求單（主表），以保留上方表頭與特殊格式
        ecn_wb = openpyxl.load_workbook(ecn_file)
        sheet_names = ecn_wb.sheetnames
        
        st.info(f"正在處理需求單，偵測到共有 {len(sheet_names)} 個分頁...")
        
        # 定義去掉第一個字元的函數
        def clean_upper_no(val):
            if val is None:
                return ""
            val_str = str(val).strip()
            if val_str in ['"', '”', '同上', '']:
                return "FOLLOW"
            if len(val_str) > 1:
                return val_str[1:]
            return val_str

        # 遍歷主表每一個分頁進行欄位動態改寫
        for sheet_name in sheet_names:
            ws = ecn_wb[sheet_name]
            
            # --- 步驟 A: 在最左邊插入兩欄：需求單號、工單號 ---
            ws.insert_cols(1, amount=2)
            ws.cell(row=1, column=1, value="需求單號")
            ws.cell(row=1, column=2, value="工單號")
            
            # 填補前 6 列左側的新欄位（可選，維持乾淨，或填入 GAA260500286）
            ws.cell(row=7, column=1, value="需求單號")
            ws.cell(row=7, column=2, value="工單號")
            
            # --- 步驟 B: 尋找關鍵欄位所在的正確行（通常在第7列） ---
            header_row_idx = 7
            for r in range(1, 15):
                row_vals = [str(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
                if any('上階' in v or '料號' in v for v in row_vals):
                    header_row_idx = r
                    break
            
            # 讀取這一行的所有欄位名稱，以便定位「上階料號」和「變更後」的位置
            headers = [str(ws.cell(row=header_row_idx, column=c).value).strip() for c in range(1, ws.max_column + 1)]
            
            # 尋找特定欄位的索引位置 (1-based)
            # 注意：因為插入了兩欄，原始欄位位置往後移了 2
            upper_col_idx = None  # 上階料號
            after_col_idx = None  # 變更後料號
            
            for idx, h in enumerate(headers, 1):
                if '上階' in h or '上階料號' in h:
                    upper_col_idx = idx
                if '變更後' in h:
                    after_col_idx = idx
            
            if not upper_col_idx or not after_col_idx:
                st.warning(f"分頁 {sheet_name} 找不到「上階料號」或「變更後」欄位，跳過精準匹配。")
                continue
                
            # --- 步驟 C: 將工單料況表的標頭寫入需求單的最右側 ---
            start_write_col = ws.max_column + 1
            status_headers = [c for c in status_df.columns if c != '__join_key__']
            
            for i, h_name in enumerate(status_headers):
                ws.cell(row=header_row_idx, value=h_name, column=start_write_col + i)
                
            # --- 步驟 D: 逐列遍歷資料列（從標頭下一列開始到最後） ---
            last_valid_upper = ""
            
            for r in range(header_row_idx + 1, ws.max_row + 1):
                # 填寫最左邊的固定資訊
                ws.cell(row=r, column=1, value="GAA260500286")
                
                # 讀取當前列的「上階料號」與「變更後料號」
                raw_upper = ws.cell(row=r, column=upper_col_idx).value
                raw_after = ws.cell(row=r, column=after_col_idx).value
                
                # 處理上階料號去頭與同上遞補邏輯
                processed_upper = clean_upper_no(raw_upper)
                if processed_upper == "FOLLOW":
                    use_upper = last_valid_upper
                else:
                    use_upper = processed_upper
                    last_valid_upper = processed_upper
                
                # 處理工單號填寫（最左邊第二欄），工單號即為上階料號去頭後的結果
                if use_upper:
                    ws.cell(row=r, column=2, value=use_upper)
                
                # 組合當前列的 Join Key: 上階料號去頭 + "_" + 變更後料號
                current_after = str(raw_after).strip() if raw_after is not None else ""
                current_join_key = f"{use_upper}_{current_after}"
                
                # 從工單料況 Dataframe 中尋找符合雙條件的資料
                matched_rows = status_df[status_df['__join_key__'] == current_join_key]
                
                if not matched_rows.empty:
                    # 找到匹配，取出第一筆，並將其欄位依序填入最右側
                    matched_match = matched_rows.iloc[0]
                    for i, h_name in enumerate(status_headers):
                        val_to_write = matched_match[h_name]
                        # 排除 NaN 值的干擾
                        if pd.isna(val_to_write):
                            val_to_write = ""
                        ws.cell(row=r, column=start_write_col + i, value=val_to_write)
                        
        st.success("🎉 全部分頁精準雙重條件比對整合完成！")
        
        # 3. 提供整合後的 Excel 檔案下載
        st.subheader("💾 步驟 3: 下載全新格式整合報表")
        output = io.BytesIO()
        ecn_wb.save(output)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 下載完整整合對比 Excel 檔案 (完全保留原格式)",
            data=processed_data,
            file_name="GAA260500286_完美整合對照表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤: {e}。請確認您上傳的檔案是否為原本的完整 Excel 檔。")
else:
    st.info("💡 提示：請同時上傳「多分頁需求單」與「工單料況表」以自動對其產出。")
