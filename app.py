import streamlit as st
import pandas as pd
import io
import zipfile
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# --- 1. 頁面設定與 CSS 優化 ---
st.set_page_config(
    page_title="鋼構清單轉換助手",
    page_icon="🏗️",
    layout="wide"  # 改為寬螢幕模式，讓預覽更完整
)

# 隱藏預設選單，讓介面更像一個獨立 App
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# --- 側邊欄：使用說明 ---
with st.sidebar:
    st.header("📖 使用說明")
    st.markdown("""
    1. **拖拉檔案**：將您的 CSV 檔案拖入右側區域。
    2. **自動偵測**：系統會自動尋找表頭並修正亂碼。
    3. **下載**：
       - **單檔**：直接預覽並下載 Excel。
       - **多檔**：自動打包成 ZIP 下載。
    """)
    st.divider()
    st.info("💡 支援格式：各種鋼構清單、螺栓清單 (CSV)")
    st.caption("Designed for Steel Structure Engineering")

# --- 主畫面 ---
st.title("🏗️ 鋼構構件清單轉換助手")
st.markdown("### 🚀 快速將 CSV 轉換為專業排版的 Excel")

# --- 核心邏輯 ---
def process_excel_styling(ws, df_len):
    """
    Excel 美化函數：加上邊框、標題底色、自動欄寬
    """
    # 定義樣式
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid") # 深藍色
    header_font = Font(bold=True, color="FFFFFF") # 白色粗體字
    
    # 1. 設定表頭樣式 (第 4 列)
    for cell in ws[4]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # 2. 設定數據區邊框與置中 (從第 5 列開始)
    # 針對每一行
    for row in ws.iter_rows(min_row=5, max_row=4 + df_len):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # 3. 自動調整欄寬
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if cell.value:
                    # 中文加權計算
                    val_str = str(cell.value)
                    # 簡單估算：非 ASCII 字元算 2 格寬度
                    len_val = sum(2 if ord(c) > 127 else 1 for c in val_str)
                    if len_val > max_length:
                        max_length = len_val
            except:
                pass
        
        # 設定寬度 (加一點緩衝)
        adjusted_width = min(max_length + 4, 60) # 上限 60
        ws.column_dimensions[column_letter].width = adjusted_width

def convert_csv_to_excel(file):
    try:
        # --- 讀取與解碼 ---
        content = None
        used_encoding = "utf-8"
        for enc in ["cp950", "utf-8", "utf-8-sig"]:
            try:
                raw_content = file.getvalue().decode(enc)
                content = raw_content.splitlines()
                used_encoding = enc
                break
            except UnicodeDecodeError:
                continue
        
        if content is None:
            return None, None, "❌ 無法識別檔案編碼"

        # --- 表頭偵測 ---
        header_idx = -1
        exclude_keywords = ["案號", "案名", "頁碼", "日期", "統計表", "清單", "次加總", "Total"]
        required_keywords = ["編號", "規格", "材質", "長度", "重量", "單重", "單量", "數量"]
        
        for i, line in enumerate(content[:20]):
            if "," not in line: continue
            if any(k in line for k in exclude_keywords): continue
            if any(k in line for k in required_keywords):
                header_idx = i
                break
        
        if header_idx == -1:
            # 回退機制
            for i, line in enumerate(content[:20]):
                if "," in line and any(k in line for k in required_keywords):
                    header_idx = i
                    break
        
        if header_idx == -1:
            return None, None, "❌ 找不到有效表頭"

        # --- 提取 Metadata ---
        meta = {}
        for line in content[:header_idx]:
            clean_line = line.replace(",", " ").replace(":", " ")
            parts = clean_line.split()
            if "案號" in line:
                for p in parts:
                    if p.isdigit() and len(p) > 3: meta["案號"] = p
            if "頁碼" in line:
                try: meta["頁碼"] = line.split("頁碼")[1].strip(":,").strip()
                except: pass
            if "案名" in line:
                try:
                    start = line.find("案名") + 2
                    end = line.find("日期") if "日期" in line else len(line)
                    meta["案名"] = line[start:end].strip(":, ").strip()
                except: pass
            if "日期" in line:
                try: meta["日期"] = line.split("日期")[1].strip(":, ").strip()
                except: pass

        # --- 數據標準化 ---
        data_lines = content[header_idx:]
        max_cols = 0
        for line in data_lines:
            max_cols = max(max_cols, line.count(",") + 1)
        
        header_line = data_lines[0]
        current_cols = header_line.count(",") + 1
        if current_cols < max_cols:
            header_line += "," * (max_cols - current_cols)
            data_lines[0] = header_line
            
        final_csv_str = "\n".join(data_lines)
        
        # --- Pandas 處理 ---
        df = pd.read_csv(io.StringIO(final_csv_str))
        df.columns = df.columns.str.strip()
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df = df.dropna(how='all')

        # 數字轉型
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='ignore')

        # --- 寫入 Excel (含樣式) ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=3)
            ws = writer.sheets['Sheet1']
            
            # Metadata 寫入
            font_meta = Font(bold=True)
            ws['A1'] = f"案號: {meta.get('案號', '')}"; ws['A1'].font = font_meta
            ws['D1'] = f"頁碼: {meta.get('頁碼', '')}"
            ws['A2'] = f"案名: {meta.get('案名', '')}"; ws['A2'].font = font_meta
            ws['D2'] = f"日期: {meta.get('日期', '')}"
            
            # 呼叫美化函數
            process_excel_styling(ws, len(df))

        output.seek(0)
        return output, df, None

    except Exception as e:
        return None, None, f"解析錯誤: {str(e)}"

# --- 介面互動區 ---
uploaded_files = st.file_uploader("📥 請上傳檔案", type="csv", accept_multiple_files=True)

if uploaded_files:
    st.divider()
    
    # 單檔模式
    if len(uploaded_files) == 1:
        file = uploaded_files[0]
        excel_data, df, error = convert_csv_to_excel(file)
        
        if error:
            st.error(error)
        else:
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.success("✅ 轉換成功！")
                st.metric(label="資料筆數", value=len(df))
                
                new_filename = os.path.splitext(file.name)[0] + ".xlsx"
                st.download_button(
                    label="📥 下載 Excel 檔案",
                    data=excel_data,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", # 讓按鈕變顯眼
                    use_container_width=True
                )
            
            with col2:
                st.write("📊 **數據預覽**")
                st.dataframe(df.head(8), use_container_width=True)

    # 批量模式
    else:
        st.info(f"⚡ 正在批次處理 {len(uploaded_files)} 個檔案...")
        
        zip_buffer = io.BytesIO()
        success_count = 0
        failed_log = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for i, file in enumerate(uploaded_files):
                status_text.text(f"正在處理: {file.name}...")
                excel_data, _, error = convert_csv_to_excel(file)
                
                if excel_data:
                    fname = os.path.splitext(file.name)[0] + ".xlsx"
                    zf.writestr(fname, excel_data.getvalue())
                    success_count += 1
                else:
                    failed_log.append(f"{file.name}: {error}")
                
                progress_bar.progress((i + 1) / len(uploaded_files))
        
        zip_buffer.seek(0)
        progress_bar.empty()
        status_text.empty()
        
        # 顯示結果
        if success_count == len(uploaded_files):
            st.success(f"🎉 完美！全部 {success_count} 個檔案轉換成功！")
        else:
            st.warning(f"⚠️ 完成 {success_count} 個，失敗 {len(failed_log)} 個")
            if failed_log:
                with st.expander("查看失敗原因"):
                    for log in failed_log:
                        st.write(log)
        
        if success_count > 0:
            st.download_button(
                label=f"📦 下載壓縮檔 (已包含 {success_count} 個 Excel)",
                data=zip_buffer,
                file_name="converted_steel_lists.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
else:
    # 這是當沒有上傳檔案時的空狀態顯示 (Empty State)
    st.info("👆 請從上方上傳 CSV 檔案以開始使用")