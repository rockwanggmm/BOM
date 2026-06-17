import streamlit as st
import pandas as pd
import io
import openpyxl
import re

st.set_page_config(page_title="多分頁跨表整合對齊工具", layout="wide")

st.title("📋 多分頁需求單「全自動跨表整合」與工單料況對齊工具")
st.markdown("""
### 🎯 終極單一表單整合邏輯：
1. **多頁面大融合**：保留第一個分頁的 1~6 列大表頭、格式與背景顏色，自動將其餘所有分頁的資料列（第 9 列以後）**全部垂直向下整合到同一個頁面**，不再留有多個分頁。
2. **B 欄工單號高級判定**：
   * 優先讀取第 6 列工單號。若發現有多個工單號或異常，則自動採用防呆：**【上階料號去頭、且去尾三碼模組號】** 作為工單號。
3. **雙重條件綁定**：匹配 **【上階料號(去頭) == 工單號碼】** 並且 **【料號 - 變更後 == Component Item】**。
4. **位置變更**：左側插入 A、B 欄，右側工單料況固定由 **P 欄 (第 16 欄)** 開始無縫黏貼。
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
        
        # --- 步驟 3：使用 openpyxl 載入需求單主體 ---
        ecn_wb = openpyxl.load_workbook(ecn_file)
        sheet_names = ecn_wb.sheetnames
        
        st.info(f"偵測到專案分頁：{', '.join(sheet_names)}，開始進行跨表垂直整合與高精度對齊...")
        
        def clean_upper_no(val):
            if val is None:
                return ""
            val_str = str(val).strip()
            if val_str in ['"', '”', '同上', '']:
                return "FOLLOW"
            if len(val_str) > 1:
                return val_str[1:]  # 去掉第一個字元 (如 2C025M2102-PB0 -> C025M2102-PB0)
            return val_str

        # 只保留第一個分頁作為「主表格」，我們會把其餘分頁的資料搬進來
        main_sheet_name = sheet_names[0]
        main_ws = ecn_wb[main_sheet_name]
        
        # 先將主要表格最左側插入 A, B 兩欄
        main_ws.insert_cols(1, amount=2)
        main_ws.cell(row=7, column=1, value="需求單號")
        main_ws.cell(row=7, column=2, value="工單號")
        for r in range(1, 7):
            main_ws.cell(row=r, column=1, value="")
            main_ws.cell(row=r, column=2, value="")
            
        # 固定工單料況從 P 欄開始寫入 (P 欄在 Excel 中是第 16 欄)
        start_write_col = 16 
        for i, col_name in enumerate(status_cols):
            main_ws.cell(row=7, column=start_write_col + i, value=col_name)

        # 用來記錄主表格目前寫到哪一列 (從第 9 列開始)
        current_main_row = 9
        
        # 遍歷所有的分頁，收集並整合資料
        for s_idx, sheet_name in enumerate(sheet_names):
            ws = ecn_wb[sheet_name]
            
            # 抓取該分頁第 6 列的工單文字
            work_order_text = ""
            for c in range(1, 15):
                # 如果是第一頁，因為已經插入兩欄，第 6 列的搜尋範圍往後挪
                search_col = c + 2 if s_idx == 0 else c
                cell_val = str(ws.cell(row=6, column=search_col).value or "")
                if "工單號" in cell_val or "工單" in cell_val:
                    if "：" in cell_val:
                        work_order_text = cell_val.split("：")[-1].strip()
                    elif ":" in cell_val:
                        work_order_text = cell_val.split(":")[-1].strip()
                    else:
                        work_order_text = cell_val.replace("工單號", "").strip()
                    break
            
            # 判斷工單號是否有多個（包含逗號、空白、或多組括號）
            is_multi_order = False
            if work_order_text:
                if ',' in work_order_text or '、' in work_order_text or len(re.findall(r'\([^)]*\)', work_order_text)) > 1:
                    is_multi_order = True
            else:
                is_multi_order = True # 沒抓到也啟動防呆
                
            # 確認該頁面「上階料號」和「變更後」的欄位索引 (未插入欄位前的原始索引)
            upper_col_idx = 6  # 預設 F 欄
            after_col_idx = 9  # 預設 I 欄
            
            # 第一頁因為已經提早插入了兩欄，需要特殊對齊
            scan_row = 7
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(row=scan_row, column=c).value).strip()
                if '上階' in val:
                    upper_col_idx = c if s_idx == 0 else c + 2
                if '變更後' in val:
                    after_col_idx = c if s_idx == 0 else c + 2

            # 如果是第一個分頁，直接原地比對與處理；若是後續分頁，則將資料複製過來合併
            start_row = 9
            max_r = ws.max_row
            
            last_valid_upper = ""
            
            for r in range(start_row, max_r + 1):
                # 讀取「上階料號」與「變更後料號」
                # 調整後續分頁的讀取索引 (因為後續分頁還沒被插入兩欄，所以要修正索引)
                read_upper_idx = upper_col_idx if s_idx == 0 else upper_col_idx - 2
                read_after_idx = after_col_idx if s_idx == 0 else after_col_idx - 2
                
                raw_upper = ws.cell(row=r, column=read_upper_idx).value
                raw_after = ws.cell(row=r, column=read_after_idx).value
                
                # 如果整列都是空的，代表到尾端了，跳過
                if raw_upper is None and raw_after is None and ws.cell(row=r, column=1).value is None:
                    continue
                    
                processed_upper = clean_upper_no(raw_upper)
                if processed_upper == "FOLLOW":
                    use_upper = last_valid_upper
                else:
                    use_upper = processed_upper
                    last_valid_upper = processed_upper
                
                # --- 核心邏輯：判定工單號 (B 欄) ---
                if is_multi_order or not work_order_text:
                    # 啟動防呆：上階料號去頭、且去尾三碼模組號 (例如：C025M2102-PB0 -> 去尾3碼變成 C025M2102)
                    if use_upper:
                        final_work_order = use_upper[:-3] if len(use_upper) > 3 else use_upper
                        if final_work_order.endswith('-'): # 去除可能殘留的橫槓
                            final_work_order = final_work_order[:-1]
                    else:
                        final_work_order = ""
                else:
                    final_work_order = work_order_text
                
                # 如果不是第一個分頁，需要把原本的整列資料 (A~N欄) 搬移到第一頁的 current_main_row
                if s_idx > 0:
                    for col_c in range(1, 14): # 原始前 13 欄 (A~M欄)
                        val_to_move = ws.cell(row=r, column=col_c).value
                        # 搬移到主表的第 col_c + 2 欄位 (因為主表前面多了 A, B 兩欄)
                        main_ws.cell(row=current_main_row, column=col_c + 2, value=val_to_move)
                
                # 填入最左側插入的 A、B 欄核心資訊
                main_ws.cell(row=current_main_row, column=1, value=doc_no)
                main_ws.cell(row=current_main_row, column=2, value=final_work_order)
                
                # 執行雙重條件交叉匹配 (上階去頭 == 工單號碼 且 變更後 == Component Item)
                val_after = str(raw_after).strip() if raw_after is not None else ""
                matched_rows = pd.DataFrame()
                if use_upper and val_after and val_after not in ["None", "", '"', '”', '同上']:
                    matched_rows = status_df[
                        (status_df['工單號碼'].astype(str).str.strip() == use_upper) & 
                        (status_df['Component Item'].astype(str).str.strip() == val_after)
                    ]
                
                # 匹配成功則將工單料況依序填入 P 欄 (第 16 欄) 右側
                if not matched_rows.empty:
                    matched_match = matched_rows.iloc[0]
                    for i, col_name in enumerate(status_cols):
                        val_to_write = matched_match[col_name]
                        if pd.isna(val_to_write):
                            val_to_write = ""
                        main_ws.cell(row=current_main_row, column=start_write_col + i, value=val_to_write)
                
                # 往下指引到大表的下一列
                if s_idx == 0:
                    current_main_row = r + 1
                else:
                    current_main_row += 1

        # 刪除其餘不再需要的分頁，只留整合後的第一頁
        for sheet_name in sheet_names[1:]:
            del ecn_wb[sheet_name]
            
        st.success(f"🎉 跨分頁終極大融合完成！已將 {len(sheet_names)} 個分頁的資料全部整合成單一表單！")
        
        # --- 步驟 4：匯出整合後的單一活頁簿 ---
        output = io.BytesIO()
        ecn_wb.save(output)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 下載全自動大融合整合 Excel 檔",
            data=processed_data,
            file_name=f"{doc_no}_全分頁大融合整合表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 整合過程中發生錯誤: {e}。請確保檔案無損且結構正確。")
else:
    st.info("💡 提示：請同時上傳「多分頁需求單」與「工單料況表」以進行大融合對齊。")
