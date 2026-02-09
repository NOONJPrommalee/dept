import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
import gc
import os
import shutil
import pythoncom
import win32com.client as win32

# --- 1. ตั้งค่าหน้าเว็บ & Path ---
st.set_page_config(page_title="RPA Excel to MySQL", layout="wide")
st.title("🚀 Multi-Excel RPA & MySQL Uploader")

# กำหนด Path (ปรับให้ตรงกับเครื่องของคุณ)
BASE_DIR = r"D:\work\บน\dept\project_folder\convert"
ARCHIVE_DIR = os.path.join(BASE_DIR, "Completed_Archive")

# สร้างโฟลเดอร์ถ้ายังไม่มี
for folder in [BASE_DIR, ARCHIVE_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# --- 2. ฟังก์ชัน RPA สำหรับแปลงไฟล์ (.xls -> .xlsx) ---
def rpa_convert_xls_to_xlsx(folder_path):
    pythoncom.CoInitialize() # แก้ปัญหา Thread ใน Streamlit
    try:
        # ใช้ Dispatch ธรรมดาเพื่อเลี่ยงปัญหา TypeError makepy ในบางเครื่อง
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        for filename in os.listdir(folder_path):
            if filename.lower().endswith(".xls") and not filename.startswith("~$"):
                xls_full_path = os.path.abspath(os.path.join(folder_path, filename))
                xlsx_full_path = xls_full_path + "x"
                
                wb = excel.Workbooks.Open(xls_full_path)
                wb.SaveAs(xlsx_full_path, FileFormat=51) # 51 = .xlsx
                wb.Close()
                
                # ย้าย .xls ต้นฉบับไป Archive
                shutil.move(xls_full_path, os.path.join(ARCHIVE_DIR, filename))
        return True
    except Exception as e:
        st.error(f"RPA Error: {e}")
        return False
    finally:
        try: excel.Quit()
        except: pass
        pythoncom.CoUninitialize()

# --- 3. ส่วนการตั้งค่า Database (Sidebar) ---
st.sidebar.header("🔌 Database Connection")
db_user = st.sidebar.text_input("Username", value="root")
db_pass = st.sidebar.text_input("Password", type="password", value="") 
db_host = st.sidebar.text_input("Host", value="localhost")
db_name = "dept"
table_name = "dept_test"

# --- 4. ส่วนการ Upload และประมวลผล ---
uploaded_files = st.file_uploader("เลือกไฟล์ Excel (xls/xlsx)", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    # ขั้นตอนที่ 1: บันทึกไฟล์ลงเครื่องก่อน
    for uploaded_file in uploaded_files:
        temp_path = os.path.join(BASE_DIR, uploaded_file.name)
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
    
    # ขั้นตอนที่ 2: รัน RPA แปลงไฟล์ .xls
    with st.spinner('🤖 RPA กำลังจัดการไฟล์และแปลง Format...'):
        rpa_convert_xls_to_xlsx(BASE_DIR)

    all_dataframes = []
    mapping_dict = {
        'ประเภทธุรกิจ': 'bus_type', 'คลาสบัญชี': 'acc_class', 'ชื่อ กฟฟ.(TRSG)': 'pea_name_trsg',
        'COL_27_TEMP': 'pea_code_main', 'สาย': 'line_code', 'หมายเลขผู้ใช้ไฟฟ้า': 'ca_no',
        'ชื่อ-สกุล': 'customer_name', 'เลขที่เอกสาร CA': 'ca_doc_no', 'สัญญา': 'contract_no',
        'คู่ค้าทางธุรกิจ': 'bp_no', 'บิลเดือน': 'bill_month', 'เงินที่ค้างชำระ': 'outstanding_amount',
        'ค่าภาษีฯ': 'tax_amount', 'ประเภทการชำระเงิน': 'payment_type', 'บัญชีแยกประเภททั่วไป': 'gl_account',
        'ประเภทอัตรา': 'rate_type', 'วันที่เอกสาร': 'doc_date', 'วันที่ครบกำหนด': 'due_date',
        'ประเภทเอกสาร': 'doc_type', 'รายการหลัก': 'main_item', 'รายการย่อย': 'sub_item',
        'ล๊อคการติดตามหนี้': 'dunning_lock', 'เลขที่เอกสารผ่อนชำระ': 'installment_doc_no',
        'วันครบกำหนดแจ้งเตือน': 'notice_due_date', 'ผลการวางหนังสือแจ้งเตือน': 'notice_result'
    }

    # ขั้นตอนที่ 3: อ่านไฟล์ .xlsx ทั้งหมดมา Clean
    files_to_process = [f for f in os.listdir(BASE_DIR) if f.endswith(".xlsx")]
    
    for filename in files_to_process:
        xlsx_path = os.path.join(BASE_DIR, filename)
        try:
            df_temp = pd.read_excel(xlsx_path, engine='openpyxl', header=17)
            
            # --- [Clean Data Logic อัปเกรดเพื่อแก้ปัญหา Row เกิน] ---
            df_temp.columns = [str(c).strip() for c in df_temp.columns]
            
            if len(df_temp.columns) >= 27:
                cols = list(df_temp.columns)
                cols[26] = 'COL_27_TEMP' 
                df_temp.columns = cols
                df_temp = df_temp.rename(columns=mapping_dict)
                
                # 🚩 กรองเฉพาะคอลัมน์ที่ต้องการ
                final_cols = [v for v in mapping_dict.values() if v in df_temp.columns]
                df_temp = df_temp[final_cols].copy()

                # 🚩 แก้ปัญหา Row เกิน: กรองแถวที่เป็นหัวตารางซ้ำ หรือแถวสรุปยอด
                # 1. ลบแถวที่ ca_no มีค่าว่าง
                df_temp = df_temp.dropna(subset=['ca_no', 'pea_code_main'], how='any')
                
                # 2. กรองเอาเฉพาะแถวที่ ca_no "มีตัวเลข" (ป้องกันหัวตารางภาษาไทยหลุดมา)
                df_temp = df_temp[df_temp['ca_no'].astype(str).str.contains(r'\d', na=False)]
                
                # 3. กำจัดแถวที่มีชื่อคอลัมน์ติดมาในข้อมูล (Exclude Headers)
                exclude_headers = ['หมายเลขผู้ใช้ไฟฟ้า', 'ca_no', 'เลขที่เอกสาร CA', 'สัญญา']
                df_temp = df_temp[~df_temp['ca_no'].astype(str).isin(exclude_headers)]

                # จัดการตัวอักษรและตัวเลขเงิน
                for col in df_temp.columns:
                    if df_temp[col].dtype == 'object':
                        df_temp[col] = df_temp[col].astype(str).str.strip().replace('nan', np.nan)
                
                money_cols = ['outstanding_amount', 'tax_amount']
                for col in money_cols:
                    if col in df_temp.columns:
                        df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.00)

                all_dataframes.append(df_temp)
                st.write(f"✅ ประมวลผล {filename} สำเร็จ: เหลือ {len(df_temp):,} แถว")
            
        except Exception as e:
            st.error(f"❌ Error ในไฟล์ {filename}: {e}")
        
        finally:
            # ลบไฟล์ .xlsx ชั่วคราวทิ้งทันทีหลังอ่านเสร็จ
            if os.path.exists(xlsx_path):
                os.remove(xlsx_path)

# --- ส่วนรวมข้อมูลและทำความสะอาดครั้งสุดท้าย ---
    if all_dataframes:
        df_final = pd.concat(all_dataframes, ignore_index=True)
        
        # 🚩🚩🚩 จุดที่เพิ่มใหม่: ลบไฟล์ใน Completed_Archive ทิ้งทั้งหมด 🚩🚩🚩
        try:
            if os.path.exists(ARCHIVE_DIR):
                # ใช้คำสั่งลบไฟล์ทั้งหมดข้างใน หรือลบตัวโฟลเดอร์แล้วสร้างใหม่
                shutil.rmtree(ARCHIVE_DIR) 
                os.makedirs(ARCHIVE_DIR) # สร้างกลับมาใหม่เพื่อรอรับไฟล์รอบหน้า
                st.toast("🧹 ล้างประวัติไฟล์เก่าใน Archive เรียบร้อยแล้ว")
        except Exception as cleanup_error:
            st.warning(f"⚠️ ไม่สามารถล้าง Archive ได้บางส่วน: {cleanup_error}")
        #st.divider()

    # --- ส่วนแสดงผลและปุ่มนำเข้า Database ---
    if all_dataframes:
        df_final = pd.concat(all_dataframes, ignore_index=True)
        st.divider()
        st.subheader(f"📊 ตัวอย่างข้อมูลรวมที่ clean เรียบร้อยแล้ว ({len(df_final):,} แถว)")
        st.dataframe(df_final.head(10))

        # ปุ่มตรวจสอบข้อมูล
        st.download_button(
            label="ดาวน์โหลดไฟล์ตรวจสอบ (CSV)",
            data=df_final.to_csv(index=False).encode('utf_8_sig'),
            file_name='cleaned_data_check.csv',
            mime='text/csv',
        )

if st.button("📤 ส่งข้อมูลเข้า MySQL และรัน Procedures", type="primary"):
    try:
        # 1. เปลี่ยนมาใช้ pymysql และเพิ่ม pool_pre_ping เพื่อเช็ค connection ก่อนส่ง
        conn_str = f"mysql+pymysql://{db_user}:{db_pass}@{db_host}/{db_name}"
        engine = create_engine(
            conn_str, 
            pool_pre_ping=True,      # ตรวจสอบการเชื่อมต่อก่อนใช้
            pool_recycle=900        # รีไซเคิล connection ทุก 15 นาที
        )
        
        # 2. ขั้นตอน Truncate: เปิดและปิดทันทีเพื่อไม่ให้ค้าง connection
        with engine.begin() as conn:
            conn.execute(text(f"TRUNCATE TABLE {table_name}"))
            # ไม่ต้องใส่ conn.commit() ถ้าใช้ engine.begin() มันจะทำให้เอง
        
        # 3. ขั้นตอนการ Insert: 
        # ลองลด chunksize ลงเหลือ 1000 หรือ 500 หากเน็ตไม่เสถียร
        with st.spinner('⏳ กำลังนำเข้าข้อมูล...'):
            df_final.to_sql(
                table_name, 
                con=engine, 
                if_exists='append', 
                index=False, 
                chunksize=1000,   # ลดขนาดลงมาหน่อยเพื่อความชัวร์
                method='multi'    # ช่วยให้ insert เร็วขึ้น (เฉพาะ pymysql/mysqldb)
            )
        
        # 4. ขั้นตอนรัน Procedure: เปิด connection ใหม่เพื่อกัน timeout
        with st.spinner('⚙️ รัน Stored Procedure...'):
            with engine.begin() as conn:
                # เพิ่มการตั้งค่า session ป้องกัน timeout ขณะรัน procedure นานๆ
                conn.execute(text("SET SESSION wait_timeout=600;")) 
                conn.execute(text("CALL sp_refresh_dashboard_master();"))
        
        st.balloons()
        st.success(f"🚀 นำเข้าสำเร็จรวม {len(df_final):,} แถว!")
        
        # Clear memory
        del df_final
        gc.collect()

    except Exception as e:
        st.error(f"❌ Database Error: {e}")
        # กรณี error ให้ลองเช็คว่า engine ยังอยู่ไหม
        if 'engine' in locals():
            engine.dispose()