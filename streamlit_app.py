import streamlit as st
import pandas as pd
import io
import openpyxl
import re

st.set_page_config(page_title="跨表大融合整合對齊工具", layout="wide")

st.title("📋 多分頁需求單「單一頁面全整合」與工單料況對齊工具")
st.markdown("""
### 🎯 最終校正版自動化邏輯：
1. **全單一頁面融合**：僅在第一個分頁插入 A、B 欄並從 **P 欄 (第 16 欄)** 黏貼料況。其餘所有分頁的數據將自動垂直向下追加到第一頁尾端，最終融合成單一頁面。
2. **工單號防呆升級**：若各分頁第 6 列檢測到多個工單號或異常，將精準執行：**【當列上階料號去頭、且去尾 3 碼模組號】** 作為 B 欄工單號。
3. **雙重精準匹配**：嚴格限定 **【上階料號(去頭) == 工單號碼】** 並且 **【料號 - 變更後 == Component Item】**。
""")

col1, col2 = st.columns(2)
with col1:
    ecn_file = st.file_uploader("1. 請上傳 多分頁需求單檔案", type=["xlsx", "xls"])
with col2:
    status_file = st.file_uploader("2. 請上傳 工單料況檔案", type=["xlsb", "xlsx", "xls"])

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
        
        st.info(f"偵測到專案分頁：{', '.join(sheet_names)}，正在執行全自動單頁大融合...")
        
        def clean_upper_no(val):
            if val is None:
                return ""
            val_str = str(val).strip()
            if val_str in ['"', '”', '同上', '']:
                return "FOLLOW"
            if len(val_str) > 1:
                return val_str[1:]  # 去掉第一個字元
            return val_str

        # 鎖定第一個分頁作為整合大表
        main_sheet_name = sheet_names[0]
        main_ws = ecn_wb[main_sheet_name]
        
        # 在第一個頁面最左側插入 A, B 兩欄
        main_ws.insert_cols(1, amount=2)
        main_ws.cell(row=7, column=1, value="需求單號")
        main_ws.cell(row=7, column=2, value="工單號")
        for r in range(1, 7):
            main_ws.cell(row=r, column=1, value="")
            main_ws.cell(row=r, column=2, value="")
            
        # 工單料況表欄位標頭，固定由 P 欄 (第 16 欄) 開始寫入
        start_write_col = 16 
        for i, col_name in enumerate(status_cols):
            main_ws.cell(row=7, column=start_write_col + i, value=col_name)

        # 指引主表格目前收集到的最新空行（從第 9 列開始）
        current_main_row = 9
        
        # 遍歷所有的分頁進行垂直整合
        for s_idx, sheet_name in enumerate(sheet_names):
            ws = ecn_wb[sheet_name]
            
            # 讀取第 6 列的原始工單文字
            work_order_text = ""
            for c in range(1, 15):
                # 第一頁因為最前面插了兩欄，搜尋索引需要往後偏移 2 欄
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
            
            # 判斷是否有多個工單代碼（如果字串包含多個括號或長度過長、有分隔符號，判定為多工單）
            is_multi_order = False
            if work_order_text:
                if ',' in work_order_text or '、' in work_order_text or ' ' in work_order_text or len(re.findall(r'\([^)]*\)', work_order_text)) > 1:
                    is_multi_order = True
            else:
                is_multi_order = True
                
            # 鎖定該分頁的「上階料號」與「料號 - 變更後」的原始欄位索引
            upper_col_idx = 6  # 預設為原 F 欄
            after_col_idx = 9  # 預設為原 I 欄
            
            # 動態校正定位
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(row=7, column=c).value).strip()
                if '上階' in val:
                    upper_col_idx = c if s_idx == 0 else c + 2
                if '變更後' in val:
                    after_col_idx = c if s_idx == 0 else c + 2

            # 確定該頁面資料結束範圍
            max_r = ws.max_row
            last_valid_upper = ""
            
            # 開始掃描各分頁的資料列（從第 9 列開始）
            for r in range(9, max_r + 1):
                read_upper_idx = upper_col_idx if s_idx == 0 else upper_col_idx - 2
                read_after_idx = after_col_idx if s_idx == 0 else after_col_idx - 2
                
                raw_upper = ws.cell(row=r, column=read_upper_idx).value
                raw_after = ws.cell(row=r, column=read_after_idx).value
                
                # 遇到完全空白列則跳過
                if raw_upper is None and raw_after is None and ws.cell(row=r, column=1).value is None:
                    continue
                    
                processed_upper = clean_upper_no(raw_upper)
                if processed_upper == "FOLLOW":
                    use_upper = last_valid_upper
                else:
                    use_upper = processed_upper
                    last_valid_upper = processed_upper
                
                # --- 執行多工單判定與上階料號去頭去尾三碼處理 ---
                if is_multi_order or not work_order_text:
                    if use_upper:
                        final_work_order = use_upper[:-3] if len(use_upper) > 3 else use_upper
                        if final_work_order.endswith('-'):
                            final_work_order = final_work_order[:-1]
                    else:
                        final_work_order = ""
                else:
                    final_work_order = work_order_text
                
                # 若不是第一個分頁，需把原始數據（A~M 欄）拷貝搬移到第一頁的 current_main_row
                if s_idx > 0:
                    for col_c in range(1, 14): # 原始前 13 欄
                        val_to_move = ws.cell(row=r, column=col_c).value
                        # 填入主表（需要留出最前方的 A, B 兩欄空間，所以是 col_c + 2）
                        main_ws.cell(row=current_main_row, column=col_c + 2, value=val_to_move)
                
                # 在大整合表填入左側 A, B 欄核心數據
                main_ws.cell(row=current_main_row, column=1, value=doc_no)
                main_ws.cell(row=current_main_row, column=2, value=final_work_order)
                
                # 雙重條件匹配：上階去頭 == 工單號碼 且 變更後 == Component Item
                val_after = str(raw_after).strip() if raw_after is not None else ""
                matched_rows = pd.DataFrame()
                if use_upper and val_after and val_after not in ["None", "", '"', '”', '同上']:
                    matched_rows = status_df[
                        (status_df['工單號碼'].astype(str).str.strip() == use_upper) & 
                        (status_df['Component Item'].astype(str).str.strip() == val_after)
                    ]
                
                # 匹配成功，從 P 欄（第 16 欄）開始依序填入料況
                if not matched_rows.empty:
                    matched_match = matched_rows.iloc[0]
                    for i, col_name in enumerate(status_cols):
                        val_to_write = matched_match[col_name]
                        if pd.isna(val_to_write):
                            val_to_write = ""
                        main_ws.cell(row=current_main_row, column=start_write_col + i, value=val_to_write)
                
                # 往下一列推進
                if s_idx == 0:
                    current_main_row = r + 1
                else:
                    current_main_row += 1

        # 資料搬移完畢後，移除其餘重複的分頁，只留下大融合的第一頁
        for sheet_name in sheet_names[1:]:
            del ecn_wb[sheet_name]
            
        st.success(f"🎉 跨活頁大融合成功！已將所有分頁的資料垂直整合成單一工作表！")
        
        # --- 步驟 4：匯出供使用者下載 ---
        output = io.BytesIO()
        ecn_wb.save(output)
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 下載全自動大融合整合 Excel 檔",
            data=processed_data,
            file_name=f"{doc_no}_全自動融合整合表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 整合過程中發生錯誤: {e}。請確保檔案格式正確。")
else:
    st.info("💡 提示：請同時上傳「多分頁需求單」與「工單料況表」以自動進行單頁面大融合。")
