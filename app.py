import streamlit as st
import pandas as pd
import plotly.express as px

# ตั้งค่าหน้าจอแดชบอร์ด
st.set_page_config(
    page_title="SAP MB51 Movement Analytics",
    page_icon="📦",
    layout="wide"
)

st.title("📦 ระบบวิเคราะห์ข้อมูลการเคลื่อนไหวสินค้า (SAP MB51)")
st.caption("อัปโหลดไฟล์ Export MB51 (.xlsx) เพื่อตรวจสอบทรานแซกชัน, Movement Type และยอดรับ-จ่าย")

# ส่วนอัปโหลดไฟล์
uploaded_file = st.sidebar.file_uploader("📥 อัปโหลดไฟล์ MB51 (Excel)", type=["xlsx", "xls"])

@st.cache_data
def load_and_clean_data(file):
    # อ่านไฟล์ Excel
    df = pd.read_excel(file)
    
    # ลบช่องว่างหัวตาราง
    df.columns = [str(col).strip() for col in df.columns]
    
    # แมปชื่อคอลัมน์มาตรฐานของ SAP MB51 (รองรับทั้งภาษาอังกฤษและแบบย่อ)
    col_mapping = {
        'Material': 'Material',
        'Material Description': 'Description',
        'Plant': 'Plant',
        'Storage Location': 'Storage_Location',
        'Movement Type': 'Movement_Type',
        'Posting Date': 'Posting_Date',
        'Quantity': 'Quantity',
        'Base Unit of Measure': 'UoM',
        'Amount in LC': 'Amount',
        'User Name': 'User'
    }
    
    # ปรับใช้ชื่อคอลัมน์ที่ตรงกัน
    rename_dict = {col: col_mapping[col] for col in df.columns if col in col_mapping}
    df = df.rename(columns=rename_dict)
    
    # แปลงคอลัมน์วันที่
    if 'Posting_Date' in df.columns:
        df['Posting_Date'] = pd.to_datetime(df['Posting_Date'], errors='coerce')
        
    # จัดการคอลัมน์ Quantity (แปลงเป็นตัวเลข รองรับเครื่องหมายลบ)
    if 'Quantity' in df.columns:
        df['Quantity'] = pd.to_numeric(df['Quantity'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
    # จัดการคอลัมน์ Amount หากมี
    if 'Amount' in df.columns:
        df['Amount'] = pd.to_numeric(df['Amount'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
    return df

if uploaded_file is not None:
    try:
        df = load_and_clean_data(uploaded_file)
        
        # --- ตัวกรองข้อมูล (Sidebar Filters) ---
        st.sidebar.header("🔍 ตัวกรองข้อมูล")
        
        # กรองช่วงวันที่
        if 'Posting_Date' in df.columns and not df['Posting_Date'].dropna().empty:
            min_date = df['Posting_Date'].min().date()
            max_date = df['Posting_Date'].max().date()
            selected_date = st.sidebar.date_input("ช่วงวันที่บันทึก (Posting Date)", [min_date, max_date])
            if len(selected_date) == 2:
                df = df[(df['Posting_Date'].dt.date >= selected_date[0]) & (df['Posting_Date'].dt.date <= selected_date[1])]
        
        # กรอง Movement Type
        if 'Movement_Type' in df.columns:
            mvt_list = sorted(df['Movement_Type'].dropna().astype(str).unique())
            selected_mvt = st.sidebar.multiselect("Movement Type", options=mvt_list, default=mvt_list[:5] if len(mvt_list) > 5 else mvt_list)
            if selected_mvt:
                df = df[df['Movement_Type'].astype(str).isin(selected_mvt)]

        # กรอง Storage Location
        if 'Storage_Location' in df.columns:
            sloc_list = sorted(df['Storage_Location'].dropna().astype(str).unique())
            selected_sloc = st.sidebar.multiselect("Storage Location", options=sloc_list, default=sloc_list)
            if selected_sloc:
                df = df[df['Storage_Location'].astype(str).isin(selected_sloc)]

        # --- KPI Metrics สรุปผล ---
        col1, col2, col3, col4 = st.columns(4)
        
        total_records = len(df)
        unique_materials = df['Material'].nunique() if 'Material' in df.columns else 0
        total_qty = df['Quantity'].sum() if 'Quantity' in df.columns else 0
        total_val = df['Amount'].sum() if 'Amount' in df.columns else 0

        col1.metric("จำนวนรายการทั้งหมด", f"{total_records:,} รายการ")
        col2.metric("จำนวน Material ที่เคลื่อนไหว", f"{unique_materials:,} รหัส")
        col3.metric("ปริมาณการเคลื่อนไหวสุทธิ", f"{total_qty:,.2f}")
        col4.metric("มูลค่ารวม (LC)", f"{total_val:,.2f}")

        st.divider()

        # --- กราฟและแดชบอร์ด ---
        chart_col1, chart_col2 = st.columns(2)

        # กราฟที่ 1: สรุปปริมาณตาม Movement Type
        if 'Movement_Type' in df.columns and 'Quantity' in df.columns:
            with chart_col1:
                st.subheader("📊 ยอดรวมตาม Movement Type")
                mvt_summary = df.groupby('Movement_Type')['Quantity'].sum().abs().reset_index()
                fig_mvt = px.bar(
                    mvt_summary, 
                    x='Movement_Type', 
                    y='Quantity', 
                    color='Movement_Type',
                    text_auto='.2s',
                    title="Volume by Movement Type"
                )
                st.plotly_chart(fig_mvt, use_container_width=True)

        # กราฟที่ 2: Top 10 Material ที่มียอดเคลื่อนไหวสูงสุด
        if 'Material' in df.columns and 'Quantity' in df.columns:
            with chart_col2:
                st.subheader("🏆 Top 10 Materials ที่มีการเคลื่อนไหวสูงสุด")
                top_mat = df.groupby('Material')['Quantity'].apply(lambda x: x.abs().sum()).nlargest(10).reset_index()
                fig_mat = px.bar(
                    top_mat, 
                    x='Quantity', 
                    y='Material', 
                    orientation='h',
                    text_auto='.2s',
                    title="Top 10 Materials (Absolute Quantity)"
                )
                fig_mat.update_layout(yaxis={'categoryorder': 'total ascending'})
                st.plotly_chart(fig_mat, use_container_width=True)

        # --- ตารางแสดงข้อมูลรายละเอียด ---
        st.subheader("📋 รายการข้อมูลรายละเอียด")
        st.dataframe(df, use_container_width=True, height=400)

        # ปุ่มดาวน์โหลดข้อมูลที่ผ่านการกรองแล้ว
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 ดาวน์โหลดข้อมูลที่กรองแล้วเป็น CSV",
            data=csv,
            file_name="filtered_MB51_data.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")
else:
    st.info("กรุณาอัปโหลดไฟล์ MB51 (.xlsx หรือ .xls) ทางแถบด้านซ้ายเพื่อเริ่มต้นวิเคราะห์")
