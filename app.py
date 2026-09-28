import io
import pandas as pd
import plotly.express as px
import streamlit as st

# -------------------------------------------------------------
# 1. Page Configuration
# -------------------------------------------------------------
st.set_page_config(
    page_title="Material Requisition & Stock Analytics",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📦 ระบบวิเคราะห์การเบิกจ่ายและสต็อกวัสดุคงเหลือ")
st.caption(
    "ประมวลผลข้อมูล Transaction จากรายงาน MB51 โดยตรง (คำนวณยอดเฉลี่ยรายเดือน สต็อก และแยกแผนก)"
)


# -------------------------------------------------------------
# 2. Data Processing Function
# -------------------------------------------------------------
@st.cache_data
def process_mb51_file(uploaded_file):
    excel_obj = pd.ExcelFile(uploaded_file)
    sheet_names = excel_obj.sheet_names

    # 1. อ่านชีท MB51 เป็นหลัก
    if "MB51" in sheet_names:
        df_mb51 = pd.read_excel(excel_obj, sheet_name="MB51")
    else:
        df_mb51 = pd.read_excel(excel_obj, sheet_name=0)

    df_mb51.columns = [str(c).strip() for c in df_mb51.columns]

    # 2. ดึงรายละเอียดวัสดุ (ไทย/อังกฤษ) จากชีท MM60 และ ZRMM0015
    desc_dict = {}
    if "MM60" in sheet_names:
        df_mm60 = pd.read_excel(excel_obj, sheet_name="MM60")
        df_mm60.columns = [str(c).strip() for c in df_mm60.columns]
        m_col = (
            "Material" if "Material" in df_mm60.columns else df_mm60.columns[0]
        )
        desc_en = (
            "Material description"
            if "Material description" in df_mm60.columns
            else None
        )
        desc_th = "รายละเอียด" if "รายละเอียด" in df_mm60.columns else None

        for _, r in df_mm60.iterrows():
            mat = str(r[m_col]).strip()
            en = (
                str(r[desc_en]).strip()
                if desc_en and pd.notna(r[desc_en])
                else ""
            )
            th = (
                str(r[desc_th]).strip()
                if desc_th and pd.notna(r[desc_th])
                else ""
            )
            desc_dict[mat] = {"desc_en": en, "desc_th": th}

    if "ZRMM0015" in sheet_names:
        df_zr = pd.read_excel(excel_obj, sheet_name="ZRMM0015")
        df_zr.columns = [str(c).strip() for c in df_zr.columns]
        m_col = "物料" if "物料" in df_zr.columns else df_zr.columns[0]
        desc_en = "物料說明" if "物料說明" in df_zr.columns else None
        desc_th = "物控物料說明" if "物控物料說明" in df_zr.columns else None

        for _, r in df_zr.iterrows():
            mat = str(r[m_col]).strip()
            en = (
                str(r[desc_en]).strip()
                if desc_en and pd.notna(r[desc_en])
                else ""
            )
            th = (
                str(r[desc_th]).strip()
                if desc_th and pd.notna(r[desc_th])
                else ""
            )
            if mat not in desc_dict:
                desc_dict[mat] = {"desc_en": en, "desc_th": th}
            else:
                if not desc_dict[mat]["desc_en"] and en:
                    desc_dict[mat]["desc_en"] = en
                if not desc_dict[mat]["desc_th"] and th:
                    desc_dict[mat]["desc_th"] = th

    # 3. จัดการแผนก (คอลัมน์ X: 領料站 หรือ คอลัมน์ S: 收貨儲存地點)
    if (
        "Data" in sheet_names
        and "物料文件" in df_mb51.columns
        and "物料" in df_mb51.columns
    ):
        df_data = pd.read_excel(excel_obj, sheet_name="Data")
        df_data.columns = [str(c).strip() for c in df_data.columns]
        if "儲存地點" in df_data.columns:
            rec_map = df_data[df_data["儲存地點"] != "2R10"].drop_duplicates(
                subset=["物料文件", "物料"]
            )[["物料文件", "物料", "儲存地點"]].rename(
                columns={"儲存地點": "儲存地點_data"}
            )

            df_mb51 = df_mb51.merge(rec_map, on=["物料文件", "物料"], how="left")
            if "領料站" in df_mb51.columns:
                df_mb51["領料站"] = df_mb51["領料站"].fillna(
                    df_mb51["儲存地點_data"].astype(str).str[1:3]
                )
            elif "收貨儲存地點" in df_mb51.columns:
                df_mb51["收貨儲存地點"] = df_mb51["收貨儲存地點"].fillna(
                    df_mb51["儲存地點_data"]
                )

    # 4. รหัสวัสดุ (Material Code)
    mat_col = next(
        (
            c
            for c in df_mb51.columns
            if c in ["物料編號", "物料", "Material", "Material Number"]
        ),
        df_mb51.columns[1],
    )
    df_mb51["Material_Code"] = df_mb51[mat_col].astype(str).str.strip()

    # 5. จำนวนการเบิกแต่ละรอบ (คอลัมน์ AA: Quantity)
    qty_col = next(
        (
            c
            for c in df_mb51.columns
            if c in ["Quantity", "數量", "以輸入單位表示的數量"]
        ),
        None,
    )
    df_mb51["Requisition_Qty"] = (
        pd.to_numeric(
            df_mb51[qty_col].astype(str).str.replace(",", ""), errors="coerce"
        )
        .abs()
        .fillna(0)
    )

    # 6. ยอดคงเหลือ Stock ปัจจุบัน (คอลัมน์ AC: Stock)
    stock_col = next(
        (c for c in df_mb51.columns if c in ["Stock", "庫存"]), None
    )
    if stock_col:
        df_mb51["Stock_Clean"] = pd.to_numeric(
            df_mb51[stock_col], errors="coerce"
        ).fillna(0)
    else:
        df_mb51["Stock_Clean"] = 0.0

    # 7. หน่วยนับ (Unit)
    unit_col = next(
        (c for c in df_mb51.columns if c in ["Unit", "輸入單位", "單位"]),
        None,
    )
    df_mb51["Unit_Clean"] = (
        df_mb51[unit_col].astype(str).str.strip() if unit_col else "UNIT"
    )

    # 8. แผนกที่เบิก (Department)
    if "領料站" in df_mb51.columns:
        dept_val = df_mb51["領料站"].astype(str).str.strip()
    elif "收貨儲存地點" in df_mb51.columns:
        dept_val = df_mb51["收貨儲存地點"].astype(str).str[1:3]
    else:
        dept_val = "Unspecified"

    df_mb51["Department"] = (
        dept_val.replace({"nan": "Other", "": "Other", "an": "Other"})
        .fillna("Other")
        .str.upper()
    )

    # 9. วันที่/เดือนที่เบิก (คอลัมน์ L: 過帳日期 หรือ คอลัมน์ U: 領料月)
    if "領料月" in df_mb51.columns and df_mb51["領料月"].dropna().count() > 0:
        df_mb51["Year_Month"] = df_mb51["領料月"].astype(str).str.strip()
    elif "過帳日期" in df_mb51.columns:

        def parse_date_serial(v):
            if pd.isna(v):
                return "Unknown"
            if isinstance(v, (int, float)):
                try:
                    return pd.to_datetime(
                        v, unit="D", origin="1899-12-30"
                    ).strftime("%Y-%m")
                except:
                    return "Unknown"
            try:
                return pd.to_datetime(v).strftime("%Y-%m")
            except:
                return "Unknown"

        df_mb51["Year_Month"] = df_mb51["過帳日期"].apply(parse_date_serial)
    else:
        df_mb51["Year_Month"] = "2026-Period"

    # กรองแถวว่าง
    df_mb51 = df_mb51[
        ~df_mb51["Material_Code"].isin(["nan", "", "None", "Total"])
    ]

    # คำนวณช่วงเดือนทั้งหมด (เช่น 2026-01 ถึง 2026-09)
    all_months = sorted(
        [m for m in df_mb51["Year_Month"].unique() if m not in ["Unknown", "nan"]]
    )
    num_months = len(all_months) if len(all_months) > 0 else 1

    # 10. สรุปข้อมูลรายวัสดุ
    mat_summary = df_mb51.groupby("Material_Code").agg(
        Unit=("Unit_Clean", "first"),
        Current_Stock=("Stock_Clean", "first"),
        Total_Requisition=("Requisition_Qty", "sum"),
    )

    # คำนวณยอดเบิกเฉลี่ยต่อเดือน = ยอดเบิกรวม / จำนวนเดือนทั้งหมด (9 เดือน)
    mat_summary["Monthly_Average"] = (
        mat_summary["Total_Requisition"] / num_months
    ).round(2)

    # คำนวณระยะเวลาสต็อกพอใช้อีกกี่เดือน (MOS)
    mat_summary["Stock_Coverage_Months"] = mat_summary.apply(
        lambda r: (
            round(r["Current_Stock"] / r["Monthly_Average"], 1)
            if r["Monthly_Average"] > 0
            else 0
        ),
        axis=1,
    )

    mat_summary["Description_TH"] = mat_summary.index.map(
        lambda m: desc_dict.get(m, {}).get("desc_th", "-")
    )
    mat_summary["Description_EN"] = mat_summary.index.map(
        lambda m: desc_dict.get(m, {}).get("desc_en", "-")
    )

    # 11. ทำ Pivot แยกตามแต่ละเดือน (Monthly Columns)
    month_pivot = df_mb51.pivot_table(
        index="Material_Code",
        columns="Year_Month",
        values="Requisition_Qty",
        aggfunc="sum",
        fill_value=0,
    )
    month_cols = [f"M_{c}" for c in month_pivot.columns]
    month_pivot.columns = month_cols

    # 12. ทำ Pivot แยกตามแต่ละแผนก (Department Columns)
    dept_pivot = df_mb51.pivot_table(
        index="Material_Code",
        columns="Department",
        values="Requisition_Qty",
        aggfunc="sum",
        fill_value=0,
    )
    dept_cols = [f"Dept_{c}" for c in dept_pivot.columns]
    dept_pivot.columns = dept_cols

    # รวมทุกตารางเข้าด้วยกัน
    final_df = (
        mat_summary.join(month_pivot)
        .join(dept_pivot)
        .reset_index()
        .sort_values(by="Total_Requisition", ascending=False)
    )

    return final_df, df_mb51, all_months, month_cols, dept_cols


# -------------------------------------------------------------
# 3. Sidebar
# -------------------------------------------------------------
st.sidebar.header("📁 อัปโหลดไฟล์รายงาน")
uploaded_file = st.sidebar.file_uploader(
    "เลือกไฟล์ MB51.XLSX", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    final_df, raw_df, all_months, month_cols, dept_cols = process_mb51_file(
        uploaded_file
    )

    st.sidebar.markdown("---")
    st.sidebar.header("🔍 ค้นหาและกรองข้อมูล")

    search_text = st.sidebar.text_input(
        "🔎 ค้นหารหัส หรือ รายละเอียดวัสดุ:", ""
    ).strip()

    filtered_df = final_df.copy()
    filtered_raw = raw_df.copy()

    if search_text:
        match_condition = (
            filtered_df["Material_Code"]
            .astype(str)
            .str.contains(search_text, case=False, na=False)
            | filtered_df["Description_TH"]
            .astype(str)
            .str.contains(search_text, case=False, na=False)
            | filtered_df["Description_EN"]
            .astype(str)
            .str.contains(search_text, case=False, na=False)
        )
        filtered_df = filtered_df[match_condition]
        filtered_raw = filtered_raw[
            filtered_raw["Material_Code"].isin(filtered_df["Material_Code"])
        ]

    # -------------------------------------------------------------
    # 4. KPI Summary
    # -------------------------------------------------------------
    total_requisition = filtered_df["Total_Requisition"].sum()
    monthly_avg_total = total_requisition / len(all_months) if all_months else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("จำนวนรายการวัสดุ", f"{len(filtered_df):,} รายการ")
    m2.metric("ยอดเบิกรวมทั้งหมด", f"{total_requisition:,.2f}")
    m3.metric(
        f"ยอดเบิกเฉลี่ยต่อเดือน ({len(all_months)} เดือน)",
        f"{monthly_avg_total:,.2f}",
    )
    m4.metric(
        "สต็อกคงเหลือปัจจุบัน", f"{filtered_df['Current_Stock'].sum():,.2f}"
    )

    st.markdown("---")

    # -------------------------------------------------------------
    # 5. Main Dashboard Tabs
    # -------------------------------------------------------------
    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📋 ตารางสรุปวัสดุ (รายเดือน & แผนก)",
            "📅 สรุปยอดเบิกรายเดือน (Monthly Summary)",
            "🏢 วิเคราะห์รายแผนก",
            "📈 กราฟแนวโน้มการเบิกรายเดือน",
        ]
    )

    # -------------------------------------------------------------
    # TAB 1: Material Master Table
    # -------------------------------------------------------------
    with tab1:
        st.subheader("📋 ตารางข้อมูลการเบิก ยอดเฉลี่ยต่อเดือน และยอดสต็อกคงเหลือ")

        # สวิตช์เลือกแสดงคอลัมน์รายเดือน / แผนก
        col_view1, col_view2 = st.columns(2)
        with col_view1:
            show_monthly_cols = st.checkbox("แสดงคอลัมน์ยอดเบิกแต่ละเดือน", value=True)
        with col_view2:
            show_dept_cols = st.checkbox("แสดงคอลัมน์ยอดเบิกแต่ละแผนก", value=True)

        base_cols = [
            "Material_Code",
            "Description_TH",
            "Description_EN",
            "Unit",
            "Current_Stock",
            "Total_Requisition",
            "Monthly_Average",
            "Stock_Coverage_Months",
        ]

        active_cols = base_cols.copy()
        if show_monthly_cols:
            active_cols += month_cols
        if show_dept_cols:
            active_cols += dept_cols

        display_table = filtered_df[active_cols].copy()

        # เปลี่ยนชื่อหัวตารางสำหรับแสดงบนหน้าเว็บ (ภาษาไทย)
        rename_headers_display = {
            "Material_Code": "เลขวัสดุ",
            "Description_TH": "รายละเอียด (ไทย)",
            "Description_EN": "รายละเอียด (EN)",
            "Unit": "หน่วยนับ",
            "Current_Stock": "Stock ปัจจุบัน",
            "Total_Requisition": "ยอดเบิกรวมสะสม",
            "Monthly_Average": "ยอดเบิกเฉลี่ย/เดือน",
            "Stock_Coverage_Months": "สต็อกพอใช้อีก (เดือน)",
        }
        for m in month_cols:
            rename_headers_display[m] = f"เดือน {m.replace('M_', '')}"
        for d in dept_cols:
            rename_headers_display[d] = f"แผนก {d.replace('Dept_', '')}"

        table_for_web = display_table.rename(columns=rename_headers_display)

        # กำหนด Format ตัวเลข
        fmt_dict = {
            "Stock ปัจจุบัน": "{:,.2f}",
            "ยอดเบิกรวมสะสม": "{:,.2f}",
            "ยอดเบิกเฉลี่ย/เดือน": "{:,.2f}",
            "สต็อกพอใช้อีก (เดือน)": "{:,.1f}",
        }
        if show_monthly_cols:
            for m in month_cols:
                fmt_dict[rename_headers_display[m]] = "{:,.2f}"
        if show_dept_cols:
            for d in dept_cols:
                fmt_dict[rename_headers_display[d]] = "{:,.2f}"

        st.dataframe(
            table_for_web.style.format(fmt_dict),
            use_container_width=True,
            height=450,
        )

        # ---------------------------------------------------------
        # ตารางสำหรับดาวน์โหลดเป็นภาษาอังกฤษทั้งหมด (English Only)
        # ---------------------------------------------------------
        export_base_cols = [
            "Material_Code",
            "Description_EN",
            "Description_TH",
            "Unit",
            "Current_Stock",
            "Total_Requisition",
            "Monthly_Average",
            "Stock_Coverage_Months",
        ]
        export_table = filtered_df[export_base_cols + month_cols + dept_cols].copy()

        rename_headers_en = {
            "Material_Code": "Material Number",
            "Description_EN": "Material Description (EN)",
            "Description_TH": "Material Description (TH)",
            "Unit": "Base Unit",
            "Current_Stock": "Current Stock",
            "Total_Requisition": "Total Requisition",
            "Monthly_Average": "Monthly Average Issue",
            "Stock_Coverage_Months": "Stock Coverage (Months / MOS)",
        }
        for m in month_cols:
            clean_m = m.replace("M_", "")
            rename_headers_en[m] = f"Issue {clean_m}"
        for d in dept_cols:
            clean_dept = d.replace("Dept_", "")
            rename_headers_en[d] = f"Dept {clean_dept}"

        export_table = export_table.rename(columns=rename_headers_en)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export_table.to_excel(
                writer, sheet_name="Material_Summary", index=False
            )

        st.download_button(
            label="📥 Download Summary Report (Excel .xlsx)",
            data=buffer.getvalue(),
            file_name="Material_Requisition_Summary_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -------------------------------------------------------------
    # TAB 2: Monthly Summary Table
    # -------------------------------------------------------------
    with tab2:
        st.subheader("📅 ตารางสรุปยอดเบิกรายเดือนและค่าเฉลี่ย (ตั้งแต่ 1/1/2026 ถึงปัจจุบัน)")

        # คำนวณสรุปรายเดือน
        monthly_stat = (
            filtered_raw.groupby("Year_Month")
            .agg(
                Total_Requisition=("Requisition_Qty", "sum"),
                Transaction_Count=("Requisition_Qty", "count"),
                Unique_Materials=("Material_Code", "nunique"),
            )
            .reset_index()
        )
        monthly_stat["Avg_Per_Transaction"] = (
            monthly_stat["Total_Requisition"] / monthly_stat["Transaction_Count"]
        )
        monthly_stat["Avg_Per_Material"] = (
            monthly_stat["Total_Requisition"] / monthly_stat["Unique_Materials"]
        )

        display_monthly = monthly_stat.rename(
            columns={
                "Year_Month": "ประจำเดือน (Period)",
                "Total_Requisition": "ยอดเบิกรวม (Total Issue)",
                "Transaction_Count": "จำนวนครั้งที่เบิก (Transactions)",
                "Unique_Materials": "จำนวนชนิดวัสดุ (Active Items)",
                "Avg_Per_Transaction": "ค่าเฉลี่ยต่อการเบิก 1 ครั้ง",
                "Avg_Per_Material": "ค่าเฉลี่ยต่อชนิดวัสดุ",
            }
        )

        st.dataframe(
            display_monthly.style.format(
                {
                    "ยอดเบิกรวม (Total Issue)": "{:,.2f}",
                    "จำนวนครั้งที่เบิก (Transactions)": "{:,} ครั้ง",
                    "จำนวนชนิดวัสดุ (Active Items)": "{:,} รหัส",
                    "ค่าเฉลี่ยต่อการเบิก 1 ครั้ง": "{:,.2f}",
                    "ค่าเฉลี่ยต่อชนิดวัสดุ": "{:,.2f}",
                }
            ),
            use_container_width=True,
        )

        st.info(
            f"💡 **สรุปภาพรวม:** ตั้งแต่วันที่ 1/1/2026 จนถึงปัจจุบัน (รวม {len(all_months)} เดือน) "
            f"มียอดการเบิกรวมทั้งสิ้น **{total_requisition:,.2f}** หรือคิดเป็น **ยอดเบิกเฉลี่ย {monthly_avg_total:,.2f} ต่อเดือน**"
        )

    # -------------------------------------------------------------
    # TAB 3: Department Breakdown
    # -------------------------------------------------------------
    with tab3:
        st.subheader("🏢 ยอดการเบิกแยกตามแต่ละแผนกสำหรับวัสดุที่เลือก")
        selected_material = st.selectbox(
            "เลือกเลขวัสดุที่ต้องการดูรายละเอียดแผนก:",
            options=filtered_df["Material_Code"].unique(),
        )

        if selected_material:
            mat_data = filtered_raw[filtered_raw["Material_Code"] == selected_material]
            dept_chart_data = (
                mat_data.groupby("Department")["Requisition_Qty"]
                .sum()
                .reset_index()
            )

            col_c1, col_c2 = st.columns([1, 1])
            with col_c1:
                fig_bar = px.bar(
                    dept_chart_data,
                    x="Department",
                    y="Requisition_Qty",
                    color="Department",
                    text_auto=".2s",
                    title=f"ยอดเบิกแยกตามแผนก: {selected_material}",
                    labels={
                        "Department": "แผนกที่เบิก",
                        "Requisition_Qty": "จำนวนที่เบิก",
                    },
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            with col_c2:
                fig_pie = px.pie(
                    dept_chart_data,
                    names="Department",
                    values="Requisition_Qty",
                    hole=0.4,
                    title="สัดส่วนการเบิกของแต่ละแผนก (%)",
                )
                st.plotly_chart(fig_pie, use_container_width=True)

    # -------------------------------------------------------------
    # TAB 4: Trend Over Time
    # -------------------------------------------------------------
    with tab4:
        st.subheader("📅 แนวโน้มการเบิกจ่ายในแต่ละเดือน (Jan 2026 - ปัจจุบัน)")
        monthly_trend = (
            filtered_raw.groupby(["Year_Month", "Department"])["Requisition_Qty"]
            .sum()
            .reset_index()
        )

        fig_trend = px.bar(
            monthly_trend,
            x="Year_Month",
            y="Requisition_Qty",
            color="Department",
            barmode="stack",
            text_auto=".2s",
            title="ยอดการเบิกจ่ายรวมรายเดือน แยกตามแผนก",
            labels={
                "Year_Month": "ประจำเดือน (YYYY-MM)",
                "Requisition_Qty": "จำนวนที่เบิก",
            },
        )
        st.plotly_chart(fig_trend, use_container_width=True)

else:
    st.info(
        "👈 กรุณาอัปโหลดไฟล์ Excel (`MB51.XLSX`) ที่แถบด้านซ้ายมือเพื่อเริ่มการประมวลผล"
    )
