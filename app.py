import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Material Stock & Requisition Analytics", layout="wide")

st.title("📊 ระบบวิเคราะห์การเบิกและยอดคงเหลือวัสดุ (SAP Analytics)")

# --- Sidebar: อัปโหลดไฟล์ ---
st.sidebar.header("📁 จัดการไฟล์ข้อมูล")
mb51_file = st.sidebar.file_uploader("1. อัปโหลดไฟล์ประวัติการเบิก MB51 (Excel/CSV)", type=["xlsx", "csv"])
mb52_file = st.sidebar.file_uploader("2. อัปโหลดไฟล์ยอดสต็อกปัจจุบัน MB52 (ถ้ามี)", type=["xlsx", "csv"])

def load_data(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        return pd.read_excel(file)

if mb51_file is not None:
    df_mb51 = load_data(mb51_file)
    
    # ทำความสะอาดชื่อคอลัมน์ (ตัดช่องว่างหัว-ท้าย)
    df_mb51.columns = [str(col).strip() for col in df_mb51.columns]

    # ตรวจสอบและระบุชื่อคอลัมน์มาตรฐานจาก SAP MB51
    # รองรับทั้งชื่อมาตรฐานภาษาอังกฤษและแบบประมวลผลแล้ว
    col_mat = next((c for c in df_mb51.columns if c.lower() in ["material", "material number", "เลขวัสดุ", "รหัสวัสดุ"]), None)
    col_desc = next((c for c in df_mb51.columns if c.lower() in ["material description", "item description", "รายละเอียด", "ชื่อวัสดุ"]), None)
    col_qty = next((c for c in df_mb51.columns if c.lower() in ["requisition amount", "quantity", "qty in un. of entry", "จำนวน"]), None)
    col_date = next((c for c in df_mb51.columns if c.lower() in ["material requisition month", "posting date", "entry date", "เดือน"]), None)
    col_grp = next((c for c in df_mb51.columns if "group description" in c.lower() or "group" in c.lower()), None)
    col_type = next((c for c in df_mb51.columns if "material type" in c.lower() or "type" in c.lower()), None)

    if not col_mat or not col_qty:
        st.error("❌ ไม่พบคอลัมน์รหัสวัสดุหรือจำนวนในไฟล์ MB51 กรุณาตรวจสอบหัวตารางของไฟล์")
    else:
        # เตรียมข้อมูลวันที่/เดือน
        df_mb51[col_qty] = pd.to_numeric(df_mb51[col_qty], errors="coerce").fillna(0)
        
        # จัดรูปแบบเดือนให้อยู่ในรูปแบบ YYYY-MM
        try:
            df_mb51["Period_Month"] = pd.to_datetime(df_mb51[col_date]).dt.strftime("%Y-%m")
        except Exception:
            df_mb51["Period_Month"] = df_mb51[col_date].astype(str)

        # คำนวณจำนวนเดือนทั้งหมดในชุดข้อมูลเพื่อหาค่าเฉลี่ย
        total_months_count = df_mb51["Period_Month"].nunique()
        if total_months_count == 0:
            total_months_count = 1

        # จัดกลุ่มข้อมูลหาผลรวมการเบิกและค่าเฉลี่ยต่อเดือน
        group_keys = [col_mat]
        if col_desc:
            group_keys.append(col_desc)
        if col_type:
            group_keys.append(col_type)
        if col_grp:
            group_keys.append(col_grp)

        # สรุปยอดเบิกรวมและค่าเฉลี่ยต่อเดือนตามรหัสวัสดุ
        summary_df = df_mb51.groupby(group_keys).agg(
            Total_Requisition=(col_qty, "sum"),
            Active_Months=("Period_Month", "nunique")
        ).reset_index()

        # คำนวณยอดเบิกเฉลี่ยต่อเดือน (ใช้จำนวนเดือนจริงในไฟล์เพื่อความสม่ำเสมอ)
        summary_df["Avg_Monthly_Requisition"] = (summary_df["Total_Requisition"] / total_months_count).round(2)

        # --- กรณีมีไฟล์ MB52 (ยอดคงเหลือจริง) ---
        if mb52_file is not None:
            df_mb52 = load_data(mb52_file)
            df_mb52.columns = [str(col).strip() for col in df_mb52.columns]
            
            stock_mat_col = next((c for c in df_mb52.columns if c.lower() in ["material", "material number", "เลขวัสดุ"]), None)
            stock_qty_col = next((c for c in df_mb52.columns if c.lower() in ["unrestricted", "unrestricted-use stock", "ยอดคงเหลือ", "stock"]), None)

            if stock_mat_col and stock_qty_col:
                df_mb52[stock_qty_col] = pd.to_numeric(df_mb52[stock_qty_col], errors="coerce").fillna(0)
                current_stock = df_mb52.groupby(stock_mat_col)[stock_qty_col].sum().reset_index()
                current_stock.rename(columns={stock_mat_col: col_mat, stock_qty_col: "Current_Stock"}, inplace=True)
                
                summary_df = pd.merge(summary_df, current_stock, on=col_mat, how="left")
                summary_df["Current_Stock"] = summary_df["Current_Stock"].fillna(0)
            else:
                summary_df["Current_Stock"] = "ข้อมูลใน MB52 ไม่ถูกต้อง"
        else:
            # หากไม่มี MB52 กำหนดเป็น N/A พร้อมคำแนะนำ
            summary_df["Current_Stock"] = "ต้องอัปโหลด MB52"

        # --- ตัวกรองข้อมูล (Filters) ในแถบด้านข้าง ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 ตัวกรองข้อมูล")

        # ค้นหาตามเลขวัสดุหรือคำอธิบาย
        search_query = st.sidebar.text_input("ค้นหารหัสหรือชื่อวัสดุ:")
        if search_query:
            if col_desc:
                summary_df = summary_df[
                    summary_df[col_mat].astype(str).str.contains(search_query, case=False, na=False) |
                    summary_df[col_desc].astype(str).str.contains(search_query, case=False, na=False)
                ]
            else:
                summary_df = summary_df[summary_df[col_mat].astype(str).str.contains(search_query, case=False, na=False)]

        # กรอง Material Group
        if col_grp and col_grp in summary_df.columns:
            groups = ["ทั้งหมด"] + sorted(summary_df[col_grp].dropna().unique().tolist())
            selected_grp = st.sidebar.selectbox("กลุ่มวัสดุ (Material Group):", groups)
            if selected_grp != "ทั้งหมด":
                summary_df = summary_df[summary_df[col_grp] == selected_grp]

        # --- แสดงผลหน้าหลัก ---
        st.markdown("### 📋 สรุปรายการข้อมูลวัสดุ ค่าเฉลี่ยการเบิก และยอดสต็อก")
        
        # จัดลำดับและเปลี่ยนชื่อคอลัมน์ให้อ่านเข้าใจง่าย
        display_columns = {
            col_mat: "Material Number",
            col_desc if col_desc else "": "Material Description",
            col_type if col_type else "": "Material Type",
            col_grp if col_grp else "": "Material Group",
            "Total_Requisition": "Total Requisition",
            "Avg_Monthly_Requisition": "Avg Requisition / Month",
            "Current_Stock": "Current Stock"
        }
        
        # กรองเฉพาะคอลัมน์ที่มีอยู่จริง
        valid_cols = [k for k in display_columns.keys() if k in summary_df.columns]
        output_df = summary_df[valid_cols].rename(columns=display_columns)

        # การแสดงผล Metrics ภาพรวม
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("จำนวนรายการวัสดุ", f"{len(output_df):,} รายการ")
        col_m2.metric("ยอดเบิกรวมทั้งหมด", f"{output_df['Total Requisition'].sum():,.2f}")
        col_m3.metric("ช่วงเวลาที่วิเคราะห์", f"{total_months_count} เดือน")

        st.dataframe(output_df, use_container_width=True)

        # ปุ่มดาวน์โหลดผลลัพธ์
        csv_data = output_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📥 ดาวน์โหลดข้อมูลสรุปเป็น CSV",
            data=csv_data,
            file_name="Material_Summary_Analysis.csv",
            mime="text/csv",
        )

else:
    st.info("👈 กรุณาอัปโหลดไฟล์รายงาน MB51 จากแถบเมนูด้านซ้ายเพื่อเริ่มการวิเคราะห์")
