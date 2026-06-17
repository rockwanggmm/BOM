import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="需求單與工單料況對比工具", layout="wide")

st.title("📋 多分頁需求單與工單料況自動對比工具")
st.markdown("""
### 📌 調整後的使用與對比邏輯：
1. **主要檔案 (需求單 GAA260500286 類型)**：此檔案為包含 **多分頁**（例如 `FDC`, `FDC (2)`, `FDC (3)` 等）的主體。
2. **對照檔案 (工單料況)**：作為 Lookup 對照源（支援 `.xlsb`、`.xlsx` 等格式）。
3. **對比邏輯**：系統會自動遍歷 **需求單中的每一個分頁**，將各分頁中的 **「上階料號」** 去掉第一個字元 (例如 `2C025M2102-PB0` 轉為 `C025M2102-PB0`) 作為 Key，去匹配工單料況表中的 **「工單號碼」或「Seiban」**。
4. **同上防呆**：若「上階料號」欄位出現 `"` 或空白等同上符號，系統會自動向下遞補前一列的料號進行匹配。
5. **資料整合**：匹配成功後，工單料況的完整欄位資料會自動貼在需求單該分頁該列的後方，查無資料則自動留空。
""")

# 檔案上傳區
col1, col2 = st.columns(2)
with col1:
    ecn_file = st.file_uploader("1. 請上傳 多分頁需求單檔案 (GAA260500286 類型)", type=["xlsx", "xls"])
with col2:
    status_file = st.file_uploader("2. 請上傳 工單料況檔案 (支援 xlsb, xlsx)", type=["xlsb", "xlsx", "xls"])

if ecn_file and status_file:
    try:
        # 1. 讀取工單料況對照表
        # 如果是 xlsb 格式，通常需要 pyxlsb 套件，此處做相容性讀取
        if status_file.name.endswith('.xlsb'):
            try:
                status_df = pd.read_excel(status_file, engine='pyxlsb')
            except Exception:
                status_df = pd.read_excel(status_file)
        else:
            status_df = pd.read_excel(status_file)
            
        status_cols = status_df.columns.tolist()
        
        st.subheader("🔍 步驟 1: 工單料況表 (對照來源) 欄位確認")
        work_order_candidates = [c for c in status_cols if '工單' in str(c) or '號碼' in str(c) or 'Seiban' in str(c)]
        selected_status_key = st.selectbox(
            "請選擇工單料況表中的【工單號碼 / Seiban】對比主鍵：",
            options=status_cols,
            index=status_cols.index(work_order_candidates[0]) if work_order_candidates else 0
        )
        # 清洗工單料況的對比 Key
        status_df['__status_key_clean__'] = status_df[selected_status_key].astype(str).str.strip()

        # 2. 讀取需求單 (多分頁主要檔案)
        st.subheader("🔍 步驟 2: 需求單多分頁比對處理")
        ecn_excel = pd.ExcelFile(ecn_file)
        ecn_sheets = ecn_excel.sheet_names
        st.info(f"偵測到需求單內共有 {len(ecn_sheets)} 個分頁。")
        
        # 定義去掉首字元以及識別同上符號的函數
        def process_ecn_key(val):
            if pd.isna(val):
                return ""
            val_str = str(val).strip()
            # 識別常見的同上符號
            if val_str == '"' or val_str == '同上' or val_str == '”':
                return "FOLLOW_PREVIOUS"
            if len(val_str) > 1:
                return val_str[1:] # 去掉第一個字元
            return val_str

        output_sheets = {}
        
        # 遍歷需求單的每一個分頁
        for sheet in ecn_sheets:
            ecn_df = pd.read_excel(ecn_file, sheet_name=sheet)
            ecn_cols = ecn_df.columns.tolist()
            
            # 自動找寻「上階料號」欄位
            key_col_candidates = [c for c in ecn_cols if '上階' in str(c) or '料號' in str(c)]
            selected_key_col = key_col_candidates[0] if key_col_candidates else ecn_cols[0]
            
            # 建立比對 Key 並處理「"」同上符號向下填補
            match_keys = []
            last_valid_key = ""
            
            for val in ecn_df[selected_key_col]:
                processed = process_ecn_key(val)
                if processed == "FOLLOW_PREVIOUS":
                    match_keys.append(last_valid_key)
                elif processed == "":
                    # 如果中間有空白，可選擇跟隨前一個或留空，這裡預設跟隨前一個有效 Key 處理
                    match_keys.append(last_valid_key)
                else:
                    last_valid_key = processed
                    match_keys.append(processed)
                    
            ecn_df['__match_key__'] = match_keys
            ecn_df['__match_key__'] = ecn_df['__match_key__'].astype(str).str.strip()
            
            # 進行 Left Join 串接工單資料
            merged_df = pd.merge(
                ecn_df,
                status_df,
                left_on='__match_key__',
                right_on='__status_key_clean__',
                how='left'
            )
            
            # 移除輔助用隱藏欄位
            if '__match_key__' in merged_df.columns:
                merged_df = merged_df.drop(columns=['__match_key__'])
            if '__status_key_clean__' in merged_df.columns:
                merged_df = merged_df.drop(columns=['__status_key_clean__'])
                
            output_sheets[sheet] = merged_df
            
        st.success("🎉 所有需求單分頁已成功與工單料況完成對比！")
        
        # 3. 預覽與下載
        st.subheader("📊 整合結果預覽 (以第一分頁為例)")
        st.dataframe(output_sheets[ecn_sheets[0]].head(10))
        
        st.subheader("💾 步驟 3: 下載全新多分頁整合報表")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df_result in output_sheets.items():
                # 寫回原本的分頁名稱
                df_result.to_excel(writer, sheet_name=sheet_name, index=False)
                
        processed_data = output.getvalue()
        st.download_button(
            label="📥 下載對比整合後的 Excel 活頁簿",
            data=processed_data,
            file_name="GAA260500286_多分頁料況整合表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤: {e}")
else:
    st.info("💡 提示：請同時上傳「多分頁需求單」與「工單料況表」以開始進行自動對比。")
