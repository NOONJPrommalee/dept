import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
import gc
import os
import shutil
from datetime import datetime

# --- 1. ตั้งค่าหน้าเว็บ & Path ---
st.set_page_config(page_title="ZWMR019 Upload", layout="wide")
st.title("📝 Data Upload : Debt Flow (ZWMR019)")

BASE_DIR = r"D:\work\บน\dept\project_folder\convert"
ARCHIVE_DIR = os.path.join(BASE_DIR, "Completed_Archive")

# --- 2. Mapping & Logic ---
mapping_dict_activity = {
    'รหัสการไฟฟ้า': 'pea_code',
    'ผู้ปฏิบัติงาน': 'worker_id',
    'การดำเนินการ': 'action_name',
    'ใบแจ้งดำเนินการ': 'notice_doc_no',
    'เอกสารเสนองดจ่ายไฟ': 'disconnect_doc_no',
    'บัญชีแสดงสัญญา': 'ca_no',
    'ชื่อ-สกุล':'customer_name',
    'เลขที่มิเตอร์ที่ดำเนินการ': 'meter_no',
    'หน่วยอ่าน': 'read_unit',
    'Flag': 'flag',
    'วันที่บันทึกจริง': 'actual_record_date',
    'วันที่ดำเนินการ': 'action_date',
    'เวลาที่ดำเนินการ': 'action_time',
    'ผู้บันทึกข้อมูล': 'recorder_id',
}

def smart_read_activity(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            df_check = pd.read_excel(file_path, header=None, skiprows=5, nrows=30)
        else:
            df_check = pd.read_csv(file_path, header=None, skiprows=5, nrows=30)
        
        h_idx_offset = -1
        for i, row in df_check.iterrows():
            if row.astype(str).str.contains('บัญชีแสดงสัญญา').any():
                h_idx_offset = i
                break
        
        if h_idx_offset == -1:
            st.error(f"❌ ไม่พบหัวตาราง 'บัญชีแสดงสัญญา' ในไฟล์ {os.path.basename(file_path)}")
            return None
            
        actual_header_row = 5 + h_idx_offset
        
        if ext == '.xlsx':
            df = pd.read_excel(file_path, engine='openpyxl', header=actual_header_row)
        elif ext == '.xls':
            try:
                df = pd.read_excel(file_path, engine='xlrd', header=actual_header_row)
            except:
                dfs = pd.read_html(file_path)
                df_html = dfs[0]
                for i in range(5, 50):
                    if df_html.iloc[i].astype(str).str.contains('บัญชีแสดงสัญญา').any():
                        df_html.columns = [str(c).strip() for c in df_html.iloc[i]]
                        df = df_html.iloc[i+1:].reset_index(drop=True)
                        return df
                return None
        else:
            df = pd.read_csv(file_path, header=actual_header_row, on_bad_lines='skip')
            
        return df
    except Exception as e:
        st.error(f"❌ Error logic ZWMR019 {os.path.basename(file_path)}: {e}")
        return None

# --- 3. Sidebar ---
st.sidebar.header("🔌 Database Connection")
db_user = "root"
db_pass = "" 
db_host = "localhost"
db_name = st.sidebar.text_input("Database Name", value="debt")
table_name = st.sidebar.text_input("Table Name", value="dept_activity_master")

st.sidebar.divider()
st.sidebar.header("⚙️ ตั้งค่าการอัปโหลด")
upload_mode = st.sidebar.radio(
    "โหมดการอัปโหลด",
    ["ล้างข้อมูลเดิม อัปโหลดใหม่ (Overwrite)", "เพิ่มเติมข้อมูลเดิม (Append)"],
    index=0
)
st.sidebar.warning(f"โหมด: {upload_mode.split(' ')[0]}")

# --- 4. Upload & Process ---
uploaded_files = st.file_uploader("เลือกไฟล์ Excel (xls/xlsx) : ZWMR019", type=["xlsx", "xls"], accept_multiple_files=True)

df_final = pd.DataFrame()

if uploaded_files:
    all_dataframes = []
    session_filenames = []

    with st.spinner('⏳ กำลังประมวลผลไฟล์...'):
        for uploaded_file in uploaded_files:
            if uploaded_file.name.startswith("~$"): continue
            
            temp_path = os.path.join(BASE_DIR, uploaded_file.name)
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            df_temp = smart_read_activity(temp_path)
            
            if df_temp is not None:
                len_raw = len(df_temp)
                
                # 1. Clean columns and rename
                df_temp.columns = [str(c).strip() for c in df_temp.columns]
                df_temp = df_temp.rename(columns=mapping_dict_activity)
                
                # 2. Ensure all expected columns exist (even if empty) to fix "missing columns" issue
                for eng_col in mapping_dict_activity.values():
                    if eng_col not in df_temp.columns:
                        df_temp[eng_col] = np.nan
                
                # 3. Select only mapped columns in order
                ordered_cols = list(dict.fromkeys(mapping_dict_activity.values()))
                df_temp = df_temp[ordered_cols].copy()

                # 4. Strip whitespace and convert empty strings to NaN
                for col in df_temp.select_dtypes(include=['object']).columns:
                    df_temp[col] = df_temp[col].fillna("").astype(str).str.strip().replace(['', 'nan', 'NaN', 'None'], np.nan)

                # 5. Filter out duplicate headers and purely empty rows
                if 'ca_no' in df_temp.columns:
                    # Filter out rows that look like headers
                    header_labels = ['บัญชีแสดงสัญญา', 'ca_no', 'BA']
                    df_temp = df_temp[~df_temp['ca_no'].astype(str).isin(header_labels)]
                    # Ensure ca_no contains at least one digit (filters out random text/empty)
                    df_temp = df_temp[df_temp['ca_no'].astype(str).str.contains(r'\d', na=False)]

                # Drop rows where all columns are NaN
                df_temp = df_temp.dropna(how='all')

                if not df_temp.empty:
                    # Handle Date columns
                    date_cols = ['notice_date', 'due_date', 'actual_record_date', 'action_date']
                    for col in date_cols:
                        if col in df_temp.columns:
                            df_temp[col] = pd.to_datetime(df_temp[col], dayfirst=True, errors='coerce').dt.date
                    
                    if 'action_time' in df_temp.columns:
                        df_temp['action_time'] = df_temp['action_time'].astype(str).str.strip()

                    all_dataframes.append(df_temp)
                    st.write(f"📊 **{uploaded_file.name}**: อ่านได้ {len_raw:,} แถว | Cleaned {len(df_temp):,} แถว")
                else:
                    st.warning(f"⚠️ ไฟล์ {uploaded_file.name}: ไม่มีข้อมูลหลังจากทำความสะอาด")
                
                shutil.move(temp_path, os.path.join(ARCHIVE_DIR, uploaded_file.name))
                session_filenames.append(uploaded_file.name)

    if all_dataframes:
        df_final = pd.concat(all_dataframes, ignore_index=True)
        st.divider()
        st.subheader(f"📊 ตัวอย่างข้อมูลรวม ({len(df_final):,} แถว)")
        st.dataframe(df_final.head(10).fillna("Null"))

# --- 5. Export ---
if not df_final.empty:
    col1, col2 = st.columns(2)
    with col1:
        csv = df_final.to_csv(index=False).encode('utf-8-sig')
        st.download_button(label="📥 ดาวน์โหลด CSV", data=csv, file_name=f"ZWMR019_cleaned.csv", mime="text/csv", use_container_width=True)

    with col2:
        if st.button(f"📤 ส่งข้อมูลเข้า MySQL", type="primary", use_container_width=True):
            try:
                conn_str = f"mysql+pymysql://{db_user}:{db_pass}@{db_host}/{db_name}"
                engine = create_engine(conn_str, pool_pre_ping=True)
                with engine.connect() as conn:
                    if "Overwrite" in upload_mode:
                        st.warning(f"🗑️ กำลังล้างข้อมูลทั้งหมดในตาราง {table_name} (TRUNCATE)...")
                        conn.execute(text(f"TRUNCATE TABLE {table_name}"))
                        conn.commit()
                        st.success(f"✅ ล้างข้อมูลและรีเซ็ต ID เรียบร้อยแล้ว")

                    # Upload
                    st.info(f"⏳ กำลังนำเข้าข้อมูลใหม่ {len(df_final):,} แถว...")
                    total_rows = len(df_final)
                    ui_batch_size = 20000
                    progress_up = st.progress(0)
                    status_up = st.empty()
                    for start_idx in range(0, total_rows, ui_batch_size):
                        end_idx = min(start_idx + ui_batch_size, total_rows)
                        chunk = df_final.iloc[start_idx:end_idx]
                        chunk.to_sql(table_name, con=engine, if_exists='append', index=False, chunksize=1000, method='multi')
                        
                        percent = min(end_idx / total_rows, 1.0)
                        progress_up.progress(percent)
                        status_up.write(f"🚀 อัปโหลดแล้ว: {end_idx:,} / {total_rows:,} แถว ({percent*100:.1f}%)")
                
                st.balloons()
                st.success(f"🚀 อัปโหลดสำเร็จ! ({len(df_final):,} แถว)")
                
                if session_filenames:
                    for fname in session_filenames:
                        fpath = os.path.join(ARCHIVE_DIR, fname)
                        if os.path.exists(fpath): os.remove(fpath)
                
                del df_final
                gc.collect()

            except Exception as e:
                st.error(f"❌ Database Error: {e}")
                if 'engine' in locals(): engine.dispose()
