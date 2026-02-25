import streamlit as st
import pandas as pd
import io
import zipfile
import os

# 設定網頁標題
st.set_page_config(page_title="鋼構清單轉檔工具", layout="centered")

st.title("鋼構構件清單轉換器")
st.markdown("### 說明")
st.write("支援批量上傳 CSV，自動轉換為標準 Excel 格式。")

# --- 核心轉換邏輯 ---
def convert_csv_to_excel(file):
    try:
        # 修正重點 1：讀取原始內容時，使用 'cp950' (繁體中文) 解碼
        content = file.getvalue().decode("cp950", errors="ignore").splitlines()
        
        meta = {}
        # 解析 Metadata (案號、案名等)
        if len(content) > 4:
            parts = content[4].split()
            for p in parts:
                if "案號" in p and ":" not in p: pass 
                if p.isdigit(): meta["案號"] = p
            if "頁碼" in content[4]:
                try:
                    meta["頁碼"] = content[4].split("頁碼:")[1].strip()
                except:
                    meta["頁碼"] = ""

        if len(content) > 5:
            line5 = content[5]
            if "日期" in line5:
                try:
                    left, right = line5.split("日期")
                    if ":" in left: meta["案名"] = left.split(":", 1)[1].strip()
                    if ":" in right: meta["日期"] = right.split(":", 1)[1].strip()
                except:
                    pass

        # 修正重點 2：Pandas 讀取數據時，也使用 'cp950'
        file.seek(0)
        df = pd.read_csv(file, skiprows=[0,1,2,3,4,5,6,8], encoding="cp950")
        
        # 清理欄位空白
        df.columns = df.columns.str.strip()
        df = df.dropna(how='all')

        # 3. 寫入 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=3)
            ws = writer.sheets['Sheet1']
            
            # 寫入 Metadata
            ws['A1'] = f"案號: {meta.get('案號', '')}"
            ws['D1'] = f"頁碼: {meta.get('頁碼', '')}"
            ws['A2'] = f"案名: {meta.get('案名', '')}"
            ws['D2'] = f"日期: {meta.get('日期', '')}"
            
        output.seek(0)
        return output, df

    except Exception as e:
        st.error(f"檔案 {file.name} 解析失敗: {e}")
        return None, None

# --- 介面邏輯 ---
uploaded_files = st.file_uploader("請拖拉檔案至此 (可多選)", type="csv", accept_multiple_files=True)

if uploaded_files:
    st.divider()
    
    # 單檔模式
    if len(uploaded_files) == 1:
        file = uploaded_files[0]
        excel_data, df = convert_csv_to_excel(file)
        
        if excel_data:
            st.success(f"{file.name} 轉換成功")
            with st.expander("預覽數據"):
                st.dataframe(df.head())
            
            new_filename = os.path.splitext(file.name)[0] + ".xlsx"
            st.download_button(
                label="下載 Excel",
                data=excel_data,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # 批量模式
    else:
        st.info(f"正在處理 {len(uploaded_files)} 個檔案...")
        zip_buffer = io.BytesIO()
        success_count = 0
        progress_bar = st.progress(0)
        
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for i, file in enumerate(uploaded_files):
                excel_data, _ = convert_csv_to_excel(file)
                if excel_data:
                    fname = os.path.splitext(file.name)[0] + ".xlsx"
                    zf.writestr(fname, excel_data.getvalue())
                    success_count += 1
                progress_bar.progress((i + 1) / len(uploaded_files))
        
        zip_buffer.seek(0)
        
        if success_count > 0:
            st.success(f"完成！成功轉換 {success_count} / {len(uploaded_files)} 個檔案")
            st.download_button(
                label="下載壓縮檔 (ZIP)",
                data=zip_buffer,
                file_name="converted_steel_lists.zip",
                mime="application/zip"
            )