import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
import gc

# --- 1. ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="Excel to MySQL Cleaner", layout="wide")
st.title("🚀 Excel Data Cleaner & MySQL Uploader")

# --- 2. ส่วนการตั้งค่า Database (Sidebar) ---
st.sidebar.header("🔌 Database Connection")
db_user = st.sidebar.text_input("Username", value="root")
db_pass = st.sidebar.text_input("Password", type="password", value="") 
db_host = st.sidebar.text_input("Host", value="localhost")

# กำหนดชื่อ Database และ Table ตามความต้องการของคุณ
db_name = "dept"
table_name = "dept_master"

# แสดงสถานะเป้าหมายการอัปโหลดใน Sidebar เพื่อความชัดเจน
st.sidebar.info(f"📍 Target: {db_name}.{table_name}")

st.session_state['db_config'] = {
    'user': db_user, 'pass': db_pass, 'host': db_host, 'name': db_name, 'table': table_name
}

# --- 3. ส่วนการ Upload ไฟล์ ---
uploaded_file = st.file_uploader("เลือกไฟล์ Excel สำหรับ dept_master", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        engine_type = 'openpyxl' if uploaded_file.name.endswith('.xlsx') else 'xlrd'
        df = pd.read_excel(uploaded_file, engine=engine_type, header=17)
        
        # ลบช่องว่างหัวคอลัมน์
        df.columns = [str(c).strip() for c in df.columns]

        # 🚩 ขั้นตอนที่ 1: จัดการรหัส กฟฟ. (ตำแหน่ง AA หรือ index 26)
        if len(df.columns) >= 27:
            cols = list(df.columns)
            cols[26] = 'COL_27_TEMP' 
            df.columns = cols
        else:
            st.error("❌ ไฟล์ Excel มีคอลัมน์ไม่ครบตามโครงสร้าง")
            st.stop()

        mapping_dict = {
            'ประเภทธุรกิจ': 'bus_type',
            'คลาสบัญชี': 'acc_class',
            'ชื่อ กฟฟ.(TRSG)': 'pea_name_trsg',
            'COL_27_TEMP': 'pea_code_main', 
            'สาย': 'line_code',
            'หมายเลขผู้ใช้ไฟฟ้า': 'ca_no',
            'ชื่อ-สกุล': 'customer_name',
            'เลขที่เอกสาร CA': 'ca_doc_no',
            'สัญญา': 'contract_no',
            'คู่ค้าทางธุรกิจ': 'bp_no',
            'บิลเดือน': 'bill_month',
            'เงินที่ค้างชำระ': 'outstanding_amount',
            'ค่าภาษีฯ': 'tax_amount',
            'ประเภทการชำระเงิน': 'payment_type',
            'บัญชีแยกประเภททั่วไป': 'gl_account',
            'ประเภทอัตรา': 'rate_type',
            'วันที่เอกสาร': 'doc_date',
            'วันที่ครบกำหนด': 'due_date',
            'ประเภทเอกสาร': 'doc_type',
            'รายการหลัก': 'main_item',
            'รายการย่อย': 'sub_item',
            'ล๊อคการติดตามหนี้': 'dunning_lock',
            'เลขที่เอกสารผ่อนชำระ': 'installment_doc_no',
            'วันครบกำหนดแจ้งเตือน': 'notice_due_date',
            'ผลการวางหนังสือแจ้งเตือน': 'notice_result'
        }

        # --- [ลำดับการคลีนข้อมูลที่ถูกต้อง] ---
        
        # A. เปลี่ยนชื่อคอลัมน์ก่อน
        df_mapped = df.rename(columns=mapping_dict)

        # B. เลือกเฉพาะคอลัมน์ที่ต้องการ
        final_cols = [v for v in mapping_dict.values() if v in df_mapped.columns]
        df_final = df_mapped[final_cols].copy()

        # C. คลีนแถว: ลบแถวขยะ และ Row 0
        df_final = df_final.dropna(subset=['ca_no', 'pea_code_main'], how='any')

        # D. คลีนแถว: ลบแถวหัวข้อ กฟฟ.
        df_final = df_final[~df_final['pea_code_main'].astype(str).str.contains('กฟฟ.', na=False)]

        # E. จัดการ Data Type
        for col in df_final.columns:
            if df_final[col].dtype == 'object':
                df_final[col] = df_final[col].astype(str).str.strip().replace('nan', np.nan)

        money_cols = ['outstanding_amount', 'tax_amount']
        for col in money_cols:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0.00)

        # F. รีเซ็ต Index
        df_final = df_final.reset_index(drop=True)

        st.success(f"✅ เตรียมข้อมูลเรียบร้อย (เหลือ {len(df_final):,} แถว)")
        st.dataframe(df_final.head(5))

        # --- 4. ส่วนการส่งข้อมูล ---
        if st.button("📤 ส่งข้อมูลเข้า dept_master และรัน Procedures", type="primary"):
            try:
                # สร้าง Connection String
                conn_str = f"mysql+mysqlconnector://{db_user}:{db_pass}@{db_host}/{db_name}"
                engine = create_engine(conn_str)
                
                # 1. ล้างข้อมูลในตารางเป้าหมาย
                with engine.connect() as conn:
                    conn.execute(text(f"TRUNCATE TABLE {table_name}"))
                    conn.commit()
                
                # 2. นำเข้าข้อมูลใหม่
                with st.spinner('⏳ กำลังนำเข้าข้อมูลสู่ตาราง dept_master...'):
                    df_final.to_sql(
                        table_name, 
                        con=engine, 
                        if_exists='append', 
                        index=False,
                        chunksize=5000
                    )
                
                # 3. รัน Stored Procedures
                with st.spinner('⚙️ กำลังประมวลผล Procedures...'):
                    with engine.begin() as conn:
                        # ระบุชื่อ Procedure ที่ต้องการรัน
                        # สามารถเพิ่มกี่ตัวก็ได้โดยการเพิ่มบรรทัด conn.execute
                        procedure_name = "sp_refresh_dashboard_master" 
                        conn.execute(text(f"CALL {procedure_name}();"))
                
                st.balloons()
                st.success(f"🚀 นำเข้าข้อมูล {len(df_final):,} แถว และรัน {procedure_name} สำเร็จ!")
            
            except Exception as e:
                st.error(f"❌ Error during upload/processing: {e}")
            finally:
                # Free large dataframes after upload to reduce memory pressure
                del df_final
                gc.collect()

    except Exception as e:
        st.error(f"❌ Error during processing: {e}")
