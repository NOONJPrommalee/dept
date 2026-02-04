import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import plotly.express as px

# --- 1. ตั้งค่าหน้าเว็บ Dashboard ---
st.set_page_config(page_title="Debt Management Dashboard 2026", layout="wide")

# --- 2. ฟังก์ชันโหลดข้อมูล (เพิ่มการจัดการ Error กรณีไม่มีตาราง) ---
@st.cache_data(show_spinner=False)
def load_data_from_db(conn_str, table_name):
    engine = create_engine(conn_str)
    try:
        return pd.read_sql(f"SELECT * FROM {table_name}", con=engine)
    except Exception as e:
        return pd.DataFrame() # ส่งคืนตารางว่างหากเกิด Error

# --- 3. ตั้งค่าการเชื่อมต่อ (Default เป็น dept.dept_master) ---
if 'db_config' not in st.session_state:
    st.session_state['db_config'] = {
        'user': 'root', 
        'pass': '', 
        'host': 'localhost', 
        'name': 'dept', 
        'table': 'dept_master'
    }

st.title("📊 Debt Dashboard (Master Data)")
st.markdown("### ระบบบริหารจัดการลูกหนี้ค้างชำระ (Table: dept_master)")

# ปุ่มสำหรับกด Refresh ข้อมูลด้วยตัวเอง
if st.sidebar.button("🔄 อัปเดตข้อมูลจาก Database"):
    st.cache_data.clear()
    if 'df_raw' in st.session_state:
        del st.session_state.df_raw
    st.rerun()

# --- 4. เริ่มกระบวนการดึงข้อมูล ---
conf = st.session_state['db_config']
conn_str = f"mysql+mysqlconnector://{conf['user']}:{conf['pass']}@{conf['host']}/{conf['name']}"
table_name = conf['table']

try:
    if 'df_raw' not in st.session_state:
        with st.spinner("⏳ กำลังโหลดข้อมูลล่าสุด..."):
            st.session_state.df_raw = load_data_from_db(conn_str, table_name)
    
    df_dash = st.session_state.df_raw.copy()

    if not df_dash.empty:
        # --- [A. เตรียมข้อมูล & Mapping] ---
        # 1. จัดกลุ่มคลาสบัญชี
        class_mapping = {
            'เอกชน - รายย่อย': 'เอกชน-รายย่อย', 'เอกชน - รายใหญ่': 'เอกชน-รายใหญ่',
            'ราชการ - รายย่อย': 'ราชการ', 'ราชการ - รายใหญ่': 'ราชการ',
            'ราชการ': 'ราชการ', 'รัฐวิสาหกิจ - รายย่อย': 'รัฐวิสาหกิจ',
            'รัฐวิสาหกิจ - รายใหญ่': 'รัฐวิสาหกิจ', 'รัฐวิสาหกิจ': 'รัฐวิสาหกิจ'
        }
        df_dash['คลาสหลัก'] = df_dash['acc_class'].map(class_mapping).fillna('อื่นๆ')

        # 2. สร้างคอลัมน์แสดงผล กฟฟ.
        df_dash['กฟฟ_display'] = df_dash['pea_code_main'].astype(str) + " : " + df_dash['pea_name_trsg'].astype(str)

        # --- [B. ส่วนตัวกรอง (Sidebar หรือด้านบน)] ---
        with st.container():
            c1, c2, c3 = st.columns(3)
            sel_class = c1.selectbox("📂 เลือกคลาสบัญชี", ["ทั้งหมด", "เอกชน-รายย่อย", "เอกชน-รายใหญ่", "ราชการ", "รัฐวิสาหกิจ"])
            sel_pea = c2.selectbox("🏢 เลือกสังกัด กฟฟ.", ["ทั้งหมด"] + sorted(df_dash['กฟฟ_display'].unique().tolist()))
            sel_doc = c3.selectbox("📄 เลือกประเภทเอกสาร", ["ทั้งหมด"] + sorted(df_dash['doc_type'].unique().tolist()))

        # กรองข้อมูลตามที่เลือก
        df_filtered = df_dash.copy()
        if sel_class != "ทั้งหมด": df_filtered = df_filtered[df_filtered['คลาสหลัก'] == sel_class]
        if sel_pea != "ทั้งหมด": df_filtered = df_filtered[df_filtered['กฟฟ_display'] == sel_pea]
        if sel_doc != "ทั้งหมด": df_filtered = df_filtered[df_filtered['doc_type'] == sel_doc]

        # --- [C. Metric Cards] ---
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("👥 รายการค้างรวม (CA)", f"{df_filtered['ca_no'].nunique():,} ราย")
        m2.metric("📖 บิลค้างทั้งหมด", f"{len(df_filtered):,} บิล")
        m3.metric("💰 เงินค้างชำระรวม", f"{df_filtered['outstanding_amount'].sum():,.2f} บาท")

        # --- [D. กราฟวิเคราะห์] ---
        st.write("#### 📉 บทวิเคราะห์ลูกหนี้")
        g1, g2 = st.columns(2)
        
        with g1:
            st.markdown("##### จำนวนบิลค้างชำระ จำแนกรายคลาส")
            c_data = df_filtered.groupby('คลาสหลัก').size().reset_index(name='count')
            fig1 = px.bar(c_data, x='คลาสหลัก', y='count', text='count', color='คลาสหลัก')
            st.plotly_chart(fig1, use_container_width=True)
            
        with g2:
            st.markdown("##### จำนวนรายค้าง (CA) จำแนกตามจำนวนบิล")
            b_counts = df_filtered.groupby('ca_no').size().reset_index(name='n')
            b_counts['group'] = b_counts['n'].apply(lambda n: f"{n} บิล" if n <= 3 else ">3 บิล")
            m_data = b_counts.groupby('group').size().reset_index(name='count_ca')
            fig2 = px.bar(m_data, x='group', y='count_ca', text='count_ca', color_discrete_sequence=['#5bc0de'])
            st.plotly_chart(fig2, use_container_width=True)

        # --- [E. ส่วนสรุปกราฟเส้นลูกหนี้บิลไม่ต่อเนื่อง] ---
        st.divider()
        st.markdown("### 📈 แนวโน้มลูกหนี้บิลไม่ต่อเนื่อง (สะสม)")
        col_filter, col_visuals = st.columns([1, 3])

        with col_filter:
            st.info("กำหนดเงื่อนไขการวิเคราะห์")
            min_amt = st.slider("ยอดค้างรวมไม่น้อยกว่า (บาท)", 0, 5000, 350, 50)
            min_bls = st.slider("จำนวนบิลค้างไม่น้อยกว่า (บิล)", 1, 12, 3, 1)

        # คำนวณ Stats ราย CA
        ca_stats = df_filtered.groupby('ca_no').agg({'outstanding_amount': 'sum', 'acc_class': 'count'}).reset_index()
        target_ids = ca_stats[(ca_stats['outstanding_amount'] >= min_amt) & (ca_stats['acc_class'] >= min_bls)]['ca_no']
        df_dis = df_filtered[df_filtered['ca_no'].isin(target_ids)].copy()

        if not df_dis.empty:
            # 1. สร้าง Dictionary สำหรับแปลงเลขเดือนเป็นชื่อย่อไทย
            thai_months = {
                '01': 'ม.ค.', '02': 'ก.พ.', '03': 'มี.ค.', '04': 'เม.ย.',
                '05': 'พ.ค.', '06': 'มิ.ย.', '07': 'ก.ค.', '08': 'ส.ค.',
                '09': 'ก.ย.', '10': 'ต.ค.', '11': 'พ.ย.', '12': 'ธ.ค.'
            }

            def map_period_thai(val):
                s = str(val).strip()
                # ตรวจสอบว่ามี 2569 หรือ 2026 อยู่ในข้อความหรือไม่
                if '2569' in s or '2026' in s:
                    # พยายามดึงเลขเดือนจาก format เช่น 256901 หรือ 2026-01
                    import re
                    month_match = re.search(r'(?:2569|2026)[-/]?(\d{2})', s)
                    if month_match:
                        m_code = month_match.group(1)
                        return f"{thai_months.get(m_code, m_code)}69"
                return "ก่อนปี 2569 (สะสม)"

            df_dis['period_display'] = df_dis['bill_month'].apply(map_period_thai)
            
            # นับจำนวนราย (Unique CA)
            trend_df = df_dis.groupby('period_display').agg({'ca_no': 'nunique'}).reset_index()

            # 2. สร้างระบบ Sort Key เพื่อให้เรียงลำดับเวลาได้ถูกต้อง
            # ให้ 'ก่อนปี 2569' เป็น 00, ม.ค.69 เป็น 01, ก.พ.69 เป็น 02...
            month_sort = {v+'69': k for k, v in thai_months.items()}
            def get_sort_key(x):
                if "ก่อน" in x: return "00"
                return month_sort.get(x, "99")

            trend_df['sort_key'] = trend_df['period_display'].apply(get_sort_key)
            trend_df = trend_df.sort_values('sort_key')

            with col_visuals:
                fig_line = px.line(
                    trend_df, 
                    x='period_display', 
                    y='ca_no', 
                    markers=True, 
                    text='ca_no', 
                    title=f"จำนวนรายที่ตรงเงื่อนไขแยกตามงวดบิลเดือน"
                )
                fig_line.update_traces(
                    line_color='#FF4B4B', 
                    line_width=3,
                    textposition="top center",
                    texttemplate='%{y:,d}'
                )
                fig_line.update_layout(
                    xaxis_title="งวดเดือน",
                    yaxis_title="จำนวนราย (CA)",
                    height=450,
                    xaxis={'type': 'category'} # บังคับให้เรียงตามที่ sort ไว้ใน dataframe
                )
                st.plotly_chart(fig_line, use_container_width=True)

        # --- [F. ตาราง Pivot] ---
        st.divider()
        st.markdown("##### 📋 ตารางสรุปข้อมูลแยกตามหน่วยงาน")
        summary = df_filtered.groupby(['กฟฟ_display', 'คลาสหลัก']).agg({
            'ca_no': 'nunique', 'acc_class': 'count', 'outstanding_amount': 'sum'
        }).reset_index()
        
        if not summary.empty:
            summary.columns = ['หน่วยงาน', 'คลาสหลัก', 'CA ค้าง', 'จำนวนบิล', 'เงินค้าง']
            pivot_df = summary.pivot(index='หน่วยงาน', columns='คลาสหลัก', values=['CA ค้าง', 'จำนวนบิล', 'เงินค้าง'])
            pivot_df = pivot_df.swaplevel(0, 1, axis=1).sort_index(axis=1)
            st.dataframe(pivot_df.style.format("{:,.2f}"), use_container_width=True)

    else:
        st.warning("⚠️ ไม่พบข้อมูลในตาราง `dept_master` กรุณาอัปโหลดข้อมูลที่หน้า Upload ก่อน")

except Exception as e:
    st.error(f"❌ เกิดข้อผิดพลาดในการแสดงผล: {e}")