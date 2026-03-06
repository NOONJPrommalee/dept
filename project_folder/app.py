import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
import gc
import os
import shutil

# --- 1. ตั้งค่าหน้าเว็บ & Path ---
st.set_page_config(page_title="Smart Multi-Group Uploader", layout="wide")
st.title("🚀 Smart Multi-Group Uploader : Debt Flow")

# กำหนด Path
BASE_DIR = r"D:\work\บน\dept\project_folder\convert"
ARCHIVE_DIR = os.path.join(BASE_DIR, "Completed_Archive")

# สร้างโฟลเดอร์ถ้ายังไม่มี
for folder in [BASE_DIR, ARCHIVE_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# --- 2. ฟังก์ชัน Smart Read (อ่านไฟล์ได้ทุกรูปแบบโดยไม่ต้องมี Excel ติดตั้ง) ---
# Mapping dictionary สำหรับแปลงชื่อคอลัมน์จากภาษาไทยเป็นอังกฤษ
mapping_dict = {
    'ประเภทธุรกิจ': 'bus_type', 'คลาสบัญชี': 'acc_class', 'ชื่อ กฟฟ.(TRSG)': 'pea_name_trsg',
    'กฟฟ.(TRSG)': 'pea_code_main','สาย': 'line_code', 'หมายเลขผู้ใช้ไฟฟ้า': 'ca_no',
    'ชื่อ-สกุล': 'customer_name', 'เลขที่เอกสาร CA': 'ca_doc_no', 'สัญญา': 'contract_no',
    'คู่ค้าทางธุรกิจ': 'bp_no', 'บิลเดือน': 'bill_month', 'เงินที่ค้างชำระ': 'outstanding_amount',
    'ค่าภาษีฯ': 'tax_amount', 'ประเภทการชำระเงิน': 'payment_type', 'บัญชีแยกประเภททั่วไป': 'gl_account',
    'ประเภทอัตรา': 'rate_type', 'วันที่เอกสาร': 'doc_date', 'วันที่ครบกำหนด': 'due_date',
    'ประเภทเอกสาร': 'doc_type', 'รายการหลัก': 'main_item', 'รายการย่อย': 'sub_item',
    'ล๊อคการติดตามหนี้': 'dunning_lock', 'เลขที่เอกสารผ่อนชำระ': 'installment_doc_no',
    'วันครบกำหนดแจ้งเตือน': 'notice_due_date', 'ผลการวางหนังสือแจ้งเตือน': 'notice_result'
}

def smart_read_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.xlsx':
            return pd.read_excel(file_path, engine='openpyxl', header=17)
        else:
            # สำหรับ .xls ลองอ่านแบบมาตรฐานก่อน
            try:
                return pd.read_excel(file_path, engine='xlrd', header=17)
            except Exception:
                # ลองอ่านแบบ UTF-16 TSV (มักเป็นไฟล์ที่ Export จากระบบอื่น)
                try:
                    # ค้นหา Header Programmatically (เพราะตำแหน่งอาจเปลี่ยนไปตามการ Export)
                    df_check = pd.read_csv(file_path, sep='\t', encoding='utf-16', header=None, names=range(100), on_bad_lines='skip', nrows=50)
                    mask = df_check.apply(lambda r: r.astype(str).str.contains('หมายเลขผู้ใช้ไฟฟ้า').any(), axis=1)
                    if mask.any():
                        h_idx = df_check[mask].index[0]
                        df = pd.read_csv(file_path, sep='\t', encoding='utf-16', header=h_idx, on_bad_lines='skip', low_memory=False)
                        df.columns = [str(c).strip() for c in df.columns]
                        return df
                    else:
                        return pd.read_csv(file_path, sep='\t', encoding='utf-16', header=17, on_bad_lines='skip')
                except Exception:
                    # ถ้าพัง แสดงว่าเป็นไฟล์ .xls ปลอม (โครงสร้างภายในเป็น HTML/XML)
                    dfs = pd.read_html(file_path)
                    df = dfs[0]
                    # ตั้งหัวตารางจากแถวที่ 17 (Index 16)
                    df.columns = [str(c).strip() for c in df.iloc[16]]
                    df = df.iloc[17:].reset_index(drop=True)
                    return df
    except Exception as e:
        # ถ้ายังพังอีก ลองอ่านแบบปกติ (CSV/UTF-8)
        try:
            return pd.read_csv(file_path, header=17, on_bad_lines='skip')
        except Exception:
            st.error(f"❌ ไม่สามารถอ่านไฟล์ {os.path.basename(file_path)}: {e}")
            return None

# --- 3. ส่วนการตั้งค่า Database & Group (Sidebar) ---
st.sidebar.header("🔌 Database Connection")
db_user = st.sidebar.text_input("Username", value="root")
db_pass = st.sidebar.text_input("Password", type="password", value="") 
db_host = st.sidebar.text_input("Host", value="localhost")
db_name = "dept_backup"
table_name = "dept_master"

st.sidebar.divider()
st.sidebar.header("📂 เลือกเขตและช่วงเดือน")

# Dropdown เลือกกลุ่ม (D, E, F)
selected_group = st.sidebar.selectbox("เลือกเขตที่ต้องการอัปโหลด", ["D", "E", "F"], index=1)

current_year = 2026
year_list = [str(y) for y in range(2024, 2033)]
selected_year = st.sidebar.selectbox("เลือก ปี (ค.ศ.)", year_list, index=year_list.index("2026"))
selected_month = st.sidebar.selectbox("เลือก เดือน", [f"{m:02d}" for m in range(1, 13)], index=1)

period_param = f"{selected_year}-{selected_month}-01"

st.sidebar.divider()
st.sidebar.header("⚙️ ตั้งค่าการอัปโหลด")
upload_mode = st.sidebar.radio(
    "โหมดการอัปโหลด",
    ["ล้างข้อมูลเดิม อัปโหลดใหม่ (Overwrite)", "เพิ่มเติมข้อมูลเดิม (Append)"],
    index=0,
    help="Overwrite: ลบข้อมูลเก่าของกลุ่มที่เลือกก่อน | Append: เพิ่มข้อมูลต่อท้ายโดยไม่ลบ"
)

st.sidebar.warning(f"โหมด: {upload_mode.split(' ')[0]} เฉพาะข้อมูลที่ขึ้นต้นด้วย '{selected_group}'")

# --- 4. ส่วนการ Upload และประมวลผล ---
uploaded_files = st.file_uploader("เลือกไฟล์ Excel (xls/xlsx) : ZBLR030", type=["xlsx", "xls"], accept_multiple_files=True)

df_final = pd.DataFrame()

if uploaded_files:
    all_dataframes = []
    session_filenames = []  # เก็บรายชื่อไฟล์ที่ประมวลผลสำเร็จในรอบนี้

    with st.spinner('⏳ กำลังประมวลผลไฟล์...'):
        for uploaded_file in uploaded_files:
            # กรองไฟล์ชั่วคราว ~$ ออก
            if uploaded_file.name.startswith("~$"):
                continue

            temp_path = os.path.join(BASE_DIR, uploaded_file.name)
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # อ่านไฟล์ด้วย Smart Read
            df_temp = smart_read_file(temp_path)
            
            if df_temp is not None:
                # --- [Clean Data Logic] ---
                len_raw = len(df_temp)
                
                # 1. ลบคอลัมน์ที่ว่าง (Empty Columns) ออกทั้งหมดก่อน เพื่อปรับจำนวนคอลัมน์ให้เหมาะสม
                df_temp = df_temp.dropna(axis=1, how='all')
                
                # 2. Rename ตาม mapping และ Strip whitespace หัวตาราง
                df_temp.columns = [str(c).strip() for c in df_temp.columns]
                df_temp = df_temp.rename(columns=mapping_dict)
                
                # 🧹 ล้างช่องว่างในทุกคอลัมน์ที่เป็นข้อความ และแปลงค่าว่างเป็น NaN เพื่อให้ MySQL บันทึกเป็น NULL
                for col in df_temp.select_dtypes(include=['object']).columns:
                    df_temp[col] = df_temp[col].fillna("").astype(str).str.strip().replace('', np.nan)

                # 3. กำหนดคอลัมน์หลัก (กรณีมีคอลัมน์ที่ Map ได้)
                final_cols = [v for v in mapping_dict.values() if v in df_temp.columns]
                if final_cols:
                    df_temp = df_temp[list(dict.fromkeys(final_cols))].copy()

                # กรองเฉพาะกลุ่มที่เลือก (D, E, F) - เช็คจากชื่อคอลัมน์ 'pea_code_main' ถ้ามี
                # (แต่ใน mapping_dict ล่าสุด ลบ COL_27_TEMP ที่แมปเป็น pea_code_main ออกแล้ว 
                # หากต้นฉบับไม่มีชื่อคอลัมน์นี้ จะถือว่าไม่กรองกลุ่ม หรือคุณอาจต้องเพิ่มชื่อหัวตารางจริงของ pea_code_main ใน mapping_dict)
                # กรองเฉพาะกลุ่มที่เลือก (D, E, F) - เช็คจากชื่อคอลัมน์ 'pea_code_main' ถ้ามี
                if 'pea_code_main' in df_temp.columns:
                    # Strip และแปลงเป็น string เพื่อความแน่นอน
                    df_temp['pea_code_main'] = df_temp['pea_code_main'].astype(str).str.strip()
                    
                    # ตรวจสอบยอดเงินก่อนกรองกลุ่ม
                    mask_group = df_temp['pea_code_main'].str.startswith(selected_group, na=False)
                    df_excluded = df_temp[~mask_group].copy()
                    df_temp = df_temp[mask_group].copy()
                    
                    if not df_excluded.empty:
                        # คำนวณยอดเงินที่ถูกตัดออกเก็บไว้ในตัวแปร (ไม่แสดงผล)
                        for col in ['outstanding_amount']:
                            if col in df_excluded.columns:
                                excl_sum = df_excluded[col].astype(str).str.replace(',', '').pipe(pd.to_numeric, errors='coerce').sum()
                
                if not df_temp.empty:
                    len_group = len(df_temp)

                    # 4. กรองแถวขยะ (กรองเฉพาะรหัสการไฟฟ้าตามคำขอผู้ใช้)
                    dropna_subset = [c for c in ['pea_code_main'] if c in df_temp.columns]
                    if dropna_subset:
                        df_temp = df_temp.dropna(subset=dropna_subset, how='any')
                    
                    if 'ca_no' in df_temp.columns:
                        df_temp = df_temp[df_temp['ca_no'].astype(str).str.contains(r'\d', na=False)]
                        exclude_headers = ['หมายเลขผู้ใช้ไฟฟ้า', 'ca_no', 'เลขที่เอกสาร CA', 'สัญญา']
                        df_temp = df_temp[~df_temp['ca_no'].astype(str).isin(exclude_headers)]
                    
                    len_clean = len(df_temp)

                    # จัดการตัวเลข (ลบ comma ออกก่อนแปลงเป็นตัวเลข)
                    for col in ['outstanding_amount', 'tax_amount']:
                        if col in df_temp.columns:
                            df_temp[col] = df_temp[col].astype(str).str.replace(',', '').pipe(pd.to_numeric, errors='coerce').fillna(0.00)

                    # จัดการ bill_month (จาก MM/YYYY เป็น YYYY-MM-01)
                    if 'bill_month' in df_temp.columns:
                        df_temp = df_temp[df_temp['bill_month'].notna() & (df_temp['bill_month'].astype(str) != 'nan')]
                        df_temp['bill_month'] = df_temp['bill_month'].astype(str).apply(
                            lambda x: f"{x.split('/')[1]}-{x.split('/')[0].zfill(2)}-01" if '/' in x else x
                        )
                        df_temp = df_temp[df_temp['bill_month'].str.match(r'\d{4}-\d{2}-\d{2}', na=False)]

                    if not df_temp.empty:
                        all_dataframes.append(df_temp)
                        total_out = df_temp['outstanding_amount'].sum() if 'outstanding_amount' in df_temp.columns else 0
                        st.write(f"📊 **{uploaded_file.name}**: อ่านได้ {len_raw:,} แถว | กลุ่ม {selected_group} {len_group:,} แถว | Cleaned {len_clean:,} แถว | ยอดรวม: {total_out:,.2f}")
                    else:
                        st.warning(f"⚠️ ไฟล์ {uploaded_file.name}: ไม่มีข้อมูลกลุ่ม '{selected_group}' (อ่านได้ {len_raw:,} แถว)")
                
                # ย้ายไป Archive และลบไฟล์ชั่วคราว
                shutil.move(temp_path, os.path.join(ARCHIVE_DIR, uploaded_file.name))
                session_filenames.append(uploaded_file.name)

    if all_dataframes:
        df_final = pd.concat(all_dataframes, ignore_index=True)
        st.divider()
        st.subheader(f"📊 ตัวอย่างข้อมูลรวมกลุ่ม {selected_group} ({len(df_final):,} แถว)")
        st.dataframe(df_final.head(10).fillna("Null"))

# --- 5. ส่วนส่งข้อมูลเข้า MySQL & Download ---
if not df_final.empty:
    col1, col2 = st.columns(2)
    
    with col1:
        # ปุ่ม Download CSV
        csv = df_final.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 ดาวน์โหลดข้อมูล Cleaned (CSV)",
            data=csv,
            file_name=f"cleaned_data_group_{selected_group}.csv",
            mime="text/csv",
            use_container_width=True
        )

    with col2:
        # ปุ่มส่งเข้า MySQL
        if st.button(f"📤 ส่งข้อมูลกลุ่ม {selected_group} เข้า MySQL", type="primary", use_container_width=True):
            try:
                conn_str = f"mysql+pymysql://{db_user}:{db_pass}@{db_host}/{db_name}"
                engine = create_engine(conn_str, pool_pre_ping=True)
                
                # ใช้ engine.connect() แทน begin() เพื่อให้จัดการ commit แยกชุดได้
                with engine.connect() as conn:
                    # 🚩 จัดการข้อมูลเก่าตามโหมดที่เลือก
                    if "Overwrite" in upload_mode:
                        st.warning(f"🗑️ กำลังล้างข้อมูลเก่าของกลุ่ม {selected_group}...")
                        
                        # นับจำนวนแถวที่จะลบก่อนเพื่อให้แสดง % ได้
                        count_query = text(f"SELECT COUNT(*) FROM {table_name} WHERE pea_code_main LIKE :pattern")
                        total_to_delete = conn.execute(count_query, {"pattern": f"{selected_group}%"}).scalar()
                        
                        progress_del = st.progress(0)
                        status_del = st.empty()
                        
                        total_deleted = 0
                        if total_to_delete > 0:
                            while True:
                                # ลบบันทึกทีละชุด (50,000 แถว) เพื่อป้องกัน Timeout/Lock
                                delete_query = text(f"DELETE FROM {table_name} WHERE pea_code_main LIKE :pattern LIMIT 50000")
                                result = conn.execute(delete_query, {"pattern": f"{selected_group}%"})
                                conn.commit()  # ทำการ commit ทันทีในแต่ละชุด
                                
                                rows_deleted = result.rowcount
                                total_deleted += rows_deleted
                                
                                percent = min(total_deleted / total_to_delete, 1.0)
                                progress_del.progress(percent)
                                status_del.write(f"✅ ลบข้อมูลแล้ว: {total_deleted:,} / {total_to_delete:,} แถว ({percent*100:.1f}%)")
                                
                                if rows_deleted == 0:
                                    break
                        else:
                            progress_del.progress(1.0)
                            status_del.write("✅ ไม่พบข้อมูลเก่าที่ต้องลบ")
                    else:
                        st.info(f"⏭️ โหมด Append: ข้ามขั้นตอนการลบข้อมูลเก่าของกลุ่ม {selected_group}")

                    # ⏳ นำเข้าข้อมูลใหม่แบบ Append (แบ่งชุดการอัปเดต UI แต่ใช้ chunk 1000 ข้างใน)
                    st.info(f"⏳ กำลังนำเข้าข้อมูลใหม่ {len(df_final):,} แถว...")
                    total_rows = len(df_final)
                    ui_batch_size = 20000 # อัปเดตหน้าจอทุก 20,000 แถว
                    progress_up = st.progress(0)
                    status_up = st.empty()
                    
                    for start_idx in range(0, total_rows, ui_batch_size):
                        end_idx = min(start_idx + ui_batch_size, total_rows)
                        chunk = df_final.iloc[start_idx:end_idx]
                        
                        # ยังคงใช้ chunksize=1000 เพื่อความปลอดภัยของ Server ตามเดิม
                        chunk.to_sql(table_name, con=engine, if_exists='append', index=False, chunksize=1000, method='multi')
                        
                        percent = min(end_idx / total_rows, 1.0)
                        progress_up.progress(percent)
                        status_up.write(f"🚀 อัปโหลดแล้ว: {end_idx:,} / {total_rows:,} แถว ({percent*100:.1f}%)")
                
                # รัน Procedures
                #with engine.begin() as conn:
                    #st.write("⚙️ กำลังประมวลผล Stored Procedures...")
                    #conn.execute(text("CALL sp_refresh_dashboard_master();"))
                    #conn.execute(text("CALL sp_update_kpi_debt_reduction(:period)"), {"period": period_param})
                
                # ตรวจสอบจำนวนแถวใน DB จริงอีกครั้งเพื่อความมั่นใจ
                with engine.connect() as conn:
                    verify_query = text(f"SELECT COUNT(*) FROM {table_name} WHERE pea_code_main LIKE :pattern")
                    db_count = conn.execute(verify_query, {"pattern": f"{selected_group}%"}).scalar()

                st.balloons()
                st.success(f"🚀 อัปโหลดกลุ่ม {selected_group} สำเร็จ! (ข้อมูลชุดนี้: {len(df_final):,} แถว | ยอดรวมใน Database: {db_count:,} แถว)")
                
                # --- 🚩 ลบไฟล์ใน Archive อัตโนมัติหลังอัปโหลดสำเร็จ ---
                if session_filenames:
                    st.info("🗑️ กำลังลบไฟล์ต้นฉบับออกจาก Archive เพื่อประหยัดพื้นที่...")
                    for fname in session_filenames:
                        fpath = os.path.join(ARCHIVE_DIR, fname)
                        if os.path.exists(fpath):
                            os.remove(fpath)
                    st.success("✅ ลบไฟล์ต้นฉบับเรียบร้อยแล้ว")
                
                del df_final
                gc.collect()

            except Exception as e:
                st.error(f"❌ Database Error: {e}")
                if 'engine' in locals(): engine.dispose()
