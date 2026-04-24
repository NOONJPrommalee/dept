import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
import gc
import os
import shutil
from datetime import datetime

# --- 1. ตั้งค่าหน้าเว็บ & Path ---
st.set_page_config(page_title="ZCAKR005 Upload", layout="wide")
st.title("📝 Data Upload : Debt Flow (ZCAKR005)")

BASE_DIR = r"D:\work\บน\dept\project_folder\convert"
ARCHIVE_DIR = os.path.join(BASE_DIR, "Completed_Archive")
os.makedirs(ARCHIVE_DIR, exist_ok=True)

# --- 2. Mapping & Logic ---
mapping_dict_zcakr005 = {
    'วันที่อนุมัติ': 'approve_date',
    'ผลอนุมัติ': 'approve_status',
    'รหัส กฟฟ.': 'pea_code',
    'ชื่อ กฟฟ.': 'pea_name',
    'สายจดหน่วย': 'mru',
    'หมายเลขผู้': 'ca_no',
    'ชื่อผู้ใช้ไฟ': 'customer_name',
    'VIP': 'vip_status',
    'หมายเลขเอกสาร': 'doc_no',
    'รายการ': 'item_type',
    'บิลเดือน': 'bill_month',
    'จำนวนเงิน': 'amount',
    'วันที่ครบกำหนด': 'due_date',
    'DP': 'dp',
    'รายละเอียด': 'details',
    'วันที่เสนอ': 'prop_date',
    'เอกสารเสนอ': 'prop_doc',
    'เลขที่ใบงาน': 'work_order',
    'พนักงานคุม': 'employee',
    'หมายเหตุ': 'remark'
}

def smart_read_zcakr005(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    header_keywords = ['วันที่อนุมัติ', 'หมายเลขผู้', 'CA', 'Contract Account', 'รหัส กฟฟ.', 'บิลเดือน', 'เอกสารเสนอ']
    
    def find_h_idx(df):
        if df is None or df.empty: return -1
        # Scan first 60 rows
        for i, row in df.head(60).iterrows():
            row_text = " ".join([str(x).replace('\xa0', ' ').strip().lower() for x in row.values])
            # A header row should have at least 2 keywords to avoid metadata rows
            matches = [kw.lower() in row_text for kw in header_keywords]
            if sum(matches) >= 2:
                return i
        return -1

    try:
        # Priority 1: Real Excel (XLSX or XLS)
        if ext == '.xlsx':
            df_peak = pd.read_excel(file_path, engine='openpyxl', header=None, nrows=100)
            h = find_h_idx(df_peak)
            if h != -1: return pd.read_excel(file_path, engine='openpyxl', header=h)
        else: # .xls variant
            try:
                # Try as binary XLS (97-2003)
                df_peak = pd.read_excel(file_path, engine='xlrd', header=None, nrows=100)
                h = find_h_idx(df_peak)
                if h != -1: return pd.read_excel(file_path, engine='xlrd', header=h)
            except: pass
            
            # Priority 2: UTF-16 TSV (Very common for SAP exports named .xls)
            try:
                df_peak = pd.read_csv(file_path, sep='\t', encoding='utf-16', header=None, names=range(200), on_bad_lines='skip', nrows=100)
                h = find_h_idx(df_peak)
                if h != -1: 
                    return pd.read_csv(file_path, sep='\t', encoding='utf-16', header=h, on_bad_lines='skip', low_memory=False)
            except: pass
            
            # Priority 3: HTML Fallback
            try:
                html_dfs = pd.read_html(file_path)
                for table in html_dfs:
                    h = find_h_idx(table.head(100))
                    if h != -1:
                        table.columns = [str(c).strip() for c in table.iloc[h]]
                        return table.iloc[h+1:].reset_index(drop=True)
            except: pass

        # Absolute Fallback: Flexible CSV
        for enc in ['utf-8-sig', 'tis-620', 'cp1252']:
            try:
                df_peak = pd.read_csv(file_path, header=None, nrows=100, on_bad_lines='skip', encoding=enc, sep=None, engine='python')
                h = find_h_idx(df_peak)
                if h != -1:
                     return pd.read_csv(file_path, header=h, on_bad_lines='skip', encoding=enc, sep=None, engine='python')
            except: continue

        st.error(f"❌ ไม่สามารถระบุรูปแบบไฟล์ หรือไม่พบหัวตารางในไฟล์ {os.path.basename(file_path)}")
        st.info("💡 ไฟล์นี้ควรมีคอลัมน์อย่างน้อย 2 อย่าง: " + ", ".join(header_keywords))
        return None

    except Exception as e:
        st.error(f"❌ Error logic ZCAKR005: {os.path.basename(file_path)}: {e}")
        return None

# --- 3. Sidebar ---
st.sidebar.header("🔌 Database Connection")
db_user = "root"
db_pass = "" 
db_host = "localhost"
db_name = st.sidebar.text_input("Database Name", value="debt")
table_name = st.sidebar.text_input("Table Name", value="dept_zcakr005_master")

st.sidebar.divider()
st.sidebar.header("⚙️ ตั้งค่าการอัปโหลด")

# --- Month Selection for Filtering ---
st.sidebar.subheader("📅 เลือกเดือนที่อัปโหลด (Approve Date)")
current_year = datetime.now().year
years = list(range(current_year - 5, current_year + 5))
months_th = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]
sel_year = st.sidebar.selectbox("ปี (YYYY)", years, index=years.index(current_year))
sel_month_name = st.sidebar.selectbox("เดือน", months_th, index=datetime.now().month - 1)
sel_month_idx = months_th.index(sel_month_name) + 1
target_period_sql = f"%.{sel_month_idx:02d}.{sel_year}"  # For SQL LIKE: %.03.2026
target_period_df = f"{sel_month_idx:02d}.{sel_year}"   # For DF filter: 03.2026

st.sidebar.info(f"💡 ระบบจะทำการ **ลบข้อมูลเดิม** ของเดือน **{sel_month_name} {sel_year}** ออกก่อน แล้วจึงนำเข้าข้อมูลใหม่จากไฟล์ที่ท่านอัปโหลด")

# --- 4. Upload & Process ---
uploaded_files = st.file_uploader("เลือกไฟล์ Excel (xls/xlsx) : ZCAKR005", type=["xlsx", "xls"], accept_multiple_files=True)

df_final = pd.DataFrame()

if uploaded_files:
    all_dataframes = []
    
    with st.spinner('⏳ กำลังประมวลผลไฟล์...'):
        for uploaded_file in uploaded_files:
            if uploaded_file.name.startswith("~$"): continue
            
            temp_path = os.path.join(BASE_DIR, uploaded_file.name)
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            df_temp = smart_read_zcakr005(temp_path)
            
            if df_temp is not None:
                len_raw = len(df_temp)
                
                # 1. Clean columns and rename
                # Build mapping dictionary
                new_cols_map = {}
                for col in df_temp.columns:
                    c_clean = str(col).strip()
                    if 'วันที่อนุมัติ' in c_clean: new_cols_map[col] = 'approve_date'
                    elif 'ผลอนุ' in c_clean: new_cols_map[col] = 'approve_status'
                    elif 'รหัส กฟฟ' in c_clean: new_cols_map[col] = 'pea_code'
                    elif 'ชื่อ กฟฟ' in c_clean: new_cols_map[col] = 'pea_name'
                    elif 'สายจดหน่วย' in c_clean: new_cols_map[col] = 'mru'
                    elif 'หมายเลขผู้' in c_clean or 'CA' in c_clean: new_cols_map[col] = 'ca_no'
                    elif 'ชื่อผู้ใช้ไฟ' in c_clean: new_cols_map[col] = 'customer_name'
                    elif 'VIP' in c_clean: new_cols_map[col] = 'vip_status'
                    elif 'หมายเลขเอกสาร' in c_clean or 'เลขที่เอกสาร' in c_clean: new_cols_map[col] = 'doc_no'
                    elif 'รายการ' in c_clean: new_cols_map[col] = 'item_type'
                    elif 'บิลเดือน' in c_clean or 'รอบบิล' in c_clean: new_cols_map[col] = 'bill_month'
                    elif 'จำนวนเงิน' in c_clean: new_cols_map[col] = 'amount'
                    elif 'วันที่ครบกำหนด' in c_clean: new_cols_map[col] = 'due_date'
                    elif 'DP' in c_clean: new_cols_map[col] = 'dp'
                    elif 'รายละเอียด' in c_clean: new_cols_map[col] = 'details'
                    elif 'วันที่เสนอ' in c_clean: new_cols_map[col] = 'prop_date'
                    elif 'เอกสารเสนอ' in c_clean: new_cols_map[col] = 'prop_doc'
                    elif 'ใบงาน' in c_clean: new_cols_map[col] = 'work_order'
                    elif 'พนักงาน' in c_clean: new_cols_map[col] = 'employee'
                    elif 'หมายเหตุ' in c_clean: new_cols_map[col] = 'remark'
                
                df_temp = df_temp.rename(columns=new_cols_map)
                
                # 2. Ensure all expected columns exist
                for eng_col in mapping_dict_zcakr005.values():
                    if eng_col not in df_temp.columns:
                        df_temp[eng_col] = np.nan
                
                # 3. Select only mapped columns in order
                ordered_cols = list(dict.fromkeys(mapping_dict_zcakr005.values()))
                df_temp = df_temp[ordered_cols].copy()

                # 4. Strip whitespace and convert empty to NaN
                for col in df_temp.select_dtypes(include=['object']).columns:
                    df_temp[col] = df_temp[col].fillna("").astype(str).str.strip().replace(['', 'nan', 'NaN', 'None'], np.nan)

                # 5. Filter out duplicate headers, garbage rows, and empty essentials
                if 'ca_no' in df_temp.columns:
                    header_labels = ['หมายเลขผู้', 'ca_no', 'วันที่อนุมัติ', 'รหัส กฟฟ.', 'หมายเลขผู้ใช้ไฟ']
                    df_temp = df_temp[~df_temp['ca_no'].astype(str).isin(header_labels)]
                    df_temp = df_temp[df_temp['ca_no'].astype(str).str.contains(r'\d', na=False)]

                    # Handle Date columns
                    def parse_thai_month(s):
                        if pd.isna(s): return s
                        s = str(s).strip()
                        month_map = {
                            'ม.ค.': '01', 'ก.พ.': '02', 'มี.ค.': '03', 'เม.ย.': '04',
                            'พ.ค.': '05', 'มิ.ย.': '06', 'ก.ค.': '07', 'ส.ค.': '08',
                            'ก.ย.': '09', 'ต.ค.': '10', 'พ.ย.': '11', 'ธ.ค.': '12'
                        }
                        try:
                            # Try standard datetime first
                            dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
                            if not pd.isna(dt): return dt.date()
                            
                            # Try Thai format like "ก.พ.-69"
                            if '-' in s:
                                parts = s.split('-')
                                m_part = parts[0].strip()
                                y_part = parts[1].strip()
                                if m_part in month_map:
                                    m = month_map[m_part]
                                    if len(y_part) == 2:
                                        y_val = int(y_part)
                                        # If year is small (e.g. 26), it's likely AD (2026)
                                        # If year is large (e.g. 69), it's likely BE (2569)
                                        if y_val < 60:
                                            y = 2000 + y_val
                                        else:
                                            y = (2500 + y_val) - 543
                                    else:
                                        y = int(y_part)
                                        if y > 2500: y -= 543
                                    return datetime(y, int(m), 1).date()
                        except: pass
                        return s

                    # Special Handle for bill_month (yyyy-mm-dd)
                    if 'bill_month' in df_temp.columns:
                        df_temp['bill_month'] = df_temp['bill_month'].apply(parse_thai_month)
                        df_temp['bill_month'] = pd.to_datetime(df_temp['bill_month'], errors='coerce').dt.strftime('%Y-%m-%d')
                    
                    # Handle other dates (dd.mm.yyyy)
                    dot_date_cols = ['approve_date', 'due_date', 'prop_date']
                    for col in dot_date_cols:
                        if col in df_temp.columns:
                            df_temp[col] = pd.to_datetime(df_temp[col], dayfirst=True, errors='coerce').dt.strftime('%d.%m.%Y')
                    
                    if 'amount' in df_temp.columns:
                        df_temp['amount'] = df_temp['amount'].astype(str).str.replace(',', '').pipe(pd.to_numeric, errors='coerce').fillna(0.00)

                # Final Null Check for Required Columns
                # (Removed strict bill_month filter as requested)
                
                df_temp = df_temp.dropna(how='all')

                if not df_temp.empty:
                    all_dataframes.append(df_temp)
                    st.write(f"📊 **{uploaded_file.name}**: อ่านได้ {len_raw:,} แถว | Cleaned {len(df_temp):,} แถว")
                else:
                    st.warning(f"⚠️ ไฟล์ {uploaded_file.name}: ไม่มีข้อมูลที่ถูกต้องหลังจากทำความสะอาด")
                    if len_raw > 0:
                        st.info("💡 อาจเป็นเพราะระบบหาหัวตารางไม่เจอ หรือข้อมูลในไฟล์ไม่ตรงกับรูปแบบที่กำหนด")
                        with st.expander("ตรวจสอบหัวตารางที่พบ"):
                            st.write(df_temp.columns.tolist())
                
                shutil.move(temp_path, os.path.join(ARCHIVE_DIR, uploaded_file.name))

    if all_dataframes:
        df_final = pd.concat(all_dataframes, ignore_index=True)
        
        # --- Filter by Selected Month/Year (approve_date) ---
        if 'approve_date' in df_final.columns:
            # approve_date format is dd.mm.yyyy
            df_final = df_final[df_final['approve_date'].astype(str).str.endswith(target_period_df, na=False)].copy()
        
        if df_final.empty:
            st.error(f"❌ ไม่พบข้อมูลที่มีวันที่อนุมัติ (Approve Date) ตรงกับเดือน {sel_month_name} {sel_year}")
            st.info("กรุณาตรวจสอบไฟล์ที่อัปโหลด หรือเปลี่ยนการเลือกเดือนใน Sidebar")
        else:
            st.divider()
            st.subheader(f"📊 ตัวอย่างข้อมูลรวมเฉพาะเดือน {sel_month_name} {sel_year} ({len(df_final):,} แถว)")
            st.dataframe(df_final.head(10).fillna("Null"))

# --- 5. Export ---
if not df_final.empty:
    col1, col2 = st.columns(2)
    with col1:
        csv = df_final.to_csv(index=False).encode('utf-8-sig')
        st.download_button(label="📥 ดาวน์โหลด CSV", data=csv, file_name="ZCAKR005_cleaned.csv", mime="text/csv", use_container_width=True)

    with col2:
        if not df_final.empty:
            # --- First Row Validation ---
            first_row_date = "Unknown"
            is_match = True
            if 'approve_date' in df_final.columns and len(df_final) > 0:
                first_row_date = df_final.iloc[0]['approve_date']
                if not str(first_row_date).endswith(target_period_df):
                    is_match = False
            
            if not is_match:
                st.warning(f"⚠️ **คำเตือน**: ข้อมูลแถวแรกมีวันที่อนุมัติเป็น `{first_row_date}` ซึ่งไม่ตรงกับเดือนที่เลือก ({sel_month_name} {sel_year})")
                st.info("กรุณาตรวจสอบให้แน่ใจว่าเลือกเดือนถูกต้องก่อนกดอัปโหลด")

            if st.button("📤 ส่งข้อมูลเข้า MySQL", type="primary", use_container_width=True):
                try:
                    conn_str = f"mysql+pymysql://{db_user}:{db_pass}@{db_host}/{db_name}"
                    engine = create_engine(conn_str, pool_pre_ping=True)
                    with engine.connect() as conn:
                        st.warning(f"🗑️ กำลังล้างข้อมูลเดิมเฉพาะเดือน {sel_month_name} {sel_year} (approve_date LIKE '{target_period_sql}')...")
                        delete_query = text(f"DELETE FROM {table_name} WHERE approve_date LIKE :period")
                        conn.execute(delete_query, {"period": target_period_sql})
                        conn.commit()
                        st.success(f"✅ ล้างข้อมูลเดิมเรียบร้อยแล้ว")

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
                    
                    # ล้างไฟล์ใน Completed_Archive
                    if os.path.exists(ARCHIVE_DIR):
                        for f in os.listdir(ARCHIVE_DIR):
                            f_path = os.path.join(ARCHIVE_DIR, f)
                            try:
                                if os.path.isfile(f_path): os.unlink(f_path)
                                elif os.path.isdir(f_path): shutil.rmtree(f_path)
                            except Exception as e:
                                st.warning(f"⚠️ ไม่สามารถลบไฟล์ {f}: {e}")
                        st.info(f"🧹 ทำความสะอาดโฟลเดอร์ {os.path.basename(ARCHIVE_DIR)} เรียบร้อยแล้ว")
                    
                    del df_final
                    gc.collect()

                except Exception as e:
                    st.error(f"❌ Database Error: {e}")
                    if 'engine' in locals(): engine.dispose()
