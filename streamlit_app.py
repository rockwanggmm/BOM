import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="需求單與工單料況對比工具", layout="wide")

st.title("📋 需求單(GAA260500286)與工單料況自動對比工具")
st.markdown("""
### 📌 系統對比與整合邏輯：
1. **主要檔案 (需求單)**：包含多個分頁 (如 `FDC`, `FDC (2)`...) 的主表。
2. **對照檔案 (工單料況)**：支援 `.xlsb` 或 `.xlsx` 格式的單一資料源。
3. **對比 Key 轉換**：自動抓取需求單「上階料號」，**去掉第一個字元** (如 `2C025M2102-PB0` -> `C025M2102-PB0`)。
4. **同上符號處理**：若遇到 `"`、`”` 或空白，系統會自動向下沿用前一列的料號進行匹配，確保資料不漏對。
""")

# 檔案上傳元件
col1, col2 = st.columns(2)
with col1:
    ecn_file = st.file_uploader("1. 請上傳 多分頁需求單檔案 (GAA260500286 類型)", type=["xlsx", "xls"])
with col2:
    status_file = st.file_uploader("2. 請上傳 工單料況檔案 (支援 xlsb, xlsx)", type=["xlsb", "xlsx", "xls"])

if ecn_file and status_file:
    try:
        # 1. 讀取工單料況對照表
        if status_file.name.endswith('.xlsb'):
            try:
                status_df = pd.read_excel(status_file, engine='pyxlsb')
            except Exception:
                status_df = pd.read_excel(status_file)
        else:
            status_df = pd.read_excel(status_file)
            
        # 清洗工單料況欄位名稱與資料
        status_df.columns = [str(c).strip() for c in status_df.columns]
        status_cols = status_df.columns.tolist()
        
        st.subheader("🔍 步驟 1: 工單料況表欄位確認")
        work_order_candidates = [c for c in status_cols if '工單' in c or '號碼' in c or 'Seiban' in c]
        selected_status_key = st.selectbox(
            "請確認或指定工單料況表的對比主鍵 (如 Seiban 或 工單號碼)：",
            options=status_cols,
            index=status_cols.index(work_order_candidates[0]) if work_order_candidates else 0
        )
        status_df['__status_key_clean__'] = status_df[selected_status_key].astype(str).str.strip()

        # 2. 處理多分頁需求單
        st.subheader("🔍 步驟 2: 需求單多分頁自動匹配中...")
        ecn_excel = pd.ExcelFile(ecn_file)
        ecn_sheets = ecn_excel.sheet_names
        
        output_sheets = {}
        
        def process_ecn_key(val):
            if pd.isna(val):
                return ""
            val_str = str(val).strip()
            if val_str in ['"', '”', '同上']:
                return "FOLLOW_PREVIOUS"
            if len(val_str) > 1:
                return val_str[1:] # 去掉第一個字元
            return val_str

        # 遍歷需求單的每個分頁
        for sheet in ecn_sheets:
            # 為了避免上方有雜亂表頭，先讀取前幾行探測真正包含「上階料號」的 Header 所在行
            temp_df = pd.read_excel(ecn_file, sheet_name=sheet, nrows=15)
            header_row = 0
            for idx, row in temp_df.iterrows():
                row_str = [str(x) for x in row.values]
                if any('上階' in x or '料號' in x for x in row_str):
                    header_row = idx + 1 # pandas 的 header 是從 0 開始算，故加 1
                    break
            
            # 用找到的正確標頭行重新讀取完整資料
            ecn_df = pd.read_excel(ecn_file, sheet_name=sheet, header=header_row)
            ecn_df.columns = [str(c).strip() for c in ecn_df.columns]
            ecn_cols = ecn_df.columns.tolist()
            
            # 定位「上階料號」欄位
            key_col_candidates = [c for c in ecn_cols if '上階' in c or '料號' in c]
            selected_key_col = key_col_candidates[0] if key_col_candidates else ecn_cols[0]
            
            # 處理同上與去第一個數字邏輯
            match_keys = []
            last_valid_key = ""
            for val in ecn_df[selected_key_col]:
                processed = process_ecn_key(val)
                if processed == "FOLLOW_PREVIOUS" or processed == "":
                    match_keys.append(last_valid_key)
                else:
                    last_valid_key = processed
                    match_keys.append(processed)
            
            ecn_df['__match_key__'] = match_keys
            ecn_df['__match_key__'] = ecn_df['__match_key__'].astype(str).str.strip()
            
            # 合併資料 (Left Join)
            merged_df = pd.merge(
                ecn_df,
                status_df,
                left_on='__match_key__',
                right_on='__status_key_clean__',
                how='left'
            )
            
            # 刪除輔助用的隱藏欄位
            if '__match_key__' in merged_df.columns:
                merged_df = merged_df.drop(columns=['__match_key__'])
            if '__status_key_clean__' in merged_df.columns:
                merged_df = merged_df.drop(columns=['__status_key_clean__'])
                
            output_sheets[sheet] = merged_df
            
        st.success(f"🎉 成功完成 {len(ecn_sheets)} 個分頁的自動交叉對比！")
        
        # 結果預覽
        st.subheader(f"📊 結果預覽 (以第一個分頁 [{ecn_sheets[0]}] 為例)")
        st.dataframe(output_sheets[ecn_sheets[0]].head(10))
        
        # 3. 匯出全新多活頁 Excel 檔案
        st.subheader("💾 步驟 3: 下載全新對對齊整合報表")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df_result in output_sheets.items():
                df_result.to_excel(writer, sheet_name=sheet_name, index=False)
                
        processed_data = output.getvalue()
        st.download_button(
            label="📥 下載對比整合後的 Excel 活頁簿",
            data=processed_data,
            file_name="GAA260500286_料況完整對照整合檔.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤: {e}。請檢查檔案欄位結構是否正常。")
else:
    st.info("💡 提示：請同時上傳「多分頁需求單」與「工單料況表」以自動對其產出。")
