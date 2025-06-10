import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO

st.title("VOC / SVOC Analysis App")  # ชื่อแอปหลัก

# --------------------------------------------------------------------
# ส่วนเลือกโหมดการวิเคราะห์ VOC หรือ SVOC
# --------------------------------------------------------------------
analysis_mode = st.selectbox(
    "เลือกโหมดการวิเคราะห์",
    ["VOC", "SVOC"],
    key="analysis_mode"
)

# --------------------------------------------------------------------
# ส่วนที่ 1: รับค่า Peak Area Standard (1),(2),(3) แล้วคำนวณหา Average
# --------------------------------------------------------------------
st.subheader("Input Peak Area Standard")

peak_area_std1 = st.number_input(
    "Peak Area Standard (1)",
    min_value=0.0, 
    value=0.0, 
    step=1.0, 
    key="peak_area_std1"
)
peak_area_std2 = st.number_input(
    "Peak Area Standard (2)",
    min_value=0.0, 
    value=0.0, 
    step=1.0, 
    key="peak_area_std2"
)
peak_area_std3 = st.number_input(
    "Peak Area Standard (3)",
    min_value=0.0, 
    value=0.0, 
    step=1.0, 
    key="peak_area_std3"
)

valid_stds = [v for v in [peak_area_std1, peak_area_std2, peak_area_std3] if v > 0]
if len(valid_stds) > 0:
    avg_peak_area_std = sum(valid_stds) / len(valid_stds)
else:
    avg_peak_area_std = 0.0

st.text_input(
    "Average Peak Area Standard", 
    value=str(avg_peak_area_std), 
    disabled=True, 
    key="avg_peak_area_std"
)

# --------------------------------------------------------------------
# ส่วนที่ 2: รับค่าปริมาณ Standard (µg) และ ปริมาณตัวอย่าง (mg)
# --------------------------------------------------------------------
st.subheader("Input Standard and Sample Amount")

standard_vol_ug = st.number_input(
    "ระบุปริมาณ Standard (µg)",
    min_value=0.0, 
    value=0.0, 
    step=1.0, 
    key="standard_vol_ug"
)
sample_vol_mg = st.number_input(
    "ระบุปริมาณตัวอย่าง (mg)",
    min_value=0.0, 
    value=0.0, 
    step=1.0, 
    key="sample_vol_mg"
)

# --------------------------------------------------------------------
# ส่วนที่ 3: อัปโหลดไฟล์ Excel/CSV แล้วอ่านเป็น DataFrame
# --------------------------------------------------------------------
st.subheader("Upload File")

uploaded_file = st.file_uploader(
    "Upload Excel (.xlsx) or CSV file", 
    type=["xlsx", "csv"], 
    key="upload_file"
)

df = pd.DataFrame()

if uploaded_file is not None:
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        st.success("อัปโหลดไฟล์สำเร็จ!")
    except Exception as e:
        st.error(f"ไม่สามารถอ่านไฟล์ได้: {e}")

# ถ้าไม่อัปโหลดไฟล์หรือว่าง ให้สร้างโครงเปล่า
if df.empty:
    df = pd.DataFrame({
        "Component RT": [],
        "Compound Name": [],
        "CAS#": [],
        "Component Area": [],
        "Area %": [],
        "Emission (ug/g)": []
    })

# --------------------------------------------------------------------
# ส่วนที่ 3.1: กรองข้อมูลตามโหมดการวิเคราะห์
#   - SVOC -> ตัดข้อมูลที่ Component RT < 12.1
#   - VOC -> ใช้ข้อมูลทั้งหมด
# --------------------------------------------------------------------
if not df.empty:
    if analysis_mode == "SVOC":
        if "Component RT" in df.columns:
            df = df[df["Component RT"] >= 12.1]
        else:
            st.warning("ไม่พบคอลัมน์ 'Component RT' จึงไม่สามารถกรองข้อมูลสำหรับ SVOC ได้")

# --------------------------------------------------------------------
# ส่วนที่ 3.2: ถ้าเป็น SVOC -> คำนวณ Area% ใหม่จาก Component Area
# --------------------------------------------------------------------
if not df.empty and analysis_mode == "SVOC":
    if "Component Area" in df.columns:
        total_area = df["Component Area"].fillna(0).sum()
        if total_area > 0:
            df["Area %"] = (df["Component Area"] / total_area) * 100
        else:
            st.warning("Total Component Area = 0 ไม่สามารถคำนวณ Area% ใหม่ได้")
    else:
        st.warning("ไม่พบคอลัมน์ 'Component Area' จึงไม่สามารถคำนวณ Area% ได้")

# --------------------------------------------------------------------
# ส่วนที่ 4: คำนวณ Emission (ug/g)
#     - emission_decimal_series เป็นทศนิยม
#     - df["Emission (ug/g)"] เก็บเป็น int (ปัดลง)
# --------------------------------------------------------------------
emission_decimal_series = pd.Series(dtype=float)

if avg_peak_area_std > 0 and standard_vol_ug > 0 and sample_vol_mg > 0 and not df.empty:
    if "Component Area" in df.columns:
        emission_decimal_series = (
            (standard_vol_ug / avg_peak_area_std)
            * (df["Component Area"] / sample_vol_mg)
            * 1000
        ).fillna(0)
        
        df["Emission (ug/g)"] = emission_decimal_series.astype(int)
    else:
        st.warning("ไม่พบคอลัมน์ 'Component Area' ในไฟล์ที่อัปโหลด จึงไม่สามารถคำนวณ Emission ได้")

# --------------------------------------------------------------------
# ส่วนที่ 5: สรุปผล
# --------------------------------------------------------------------
# (A) Sum All Area%
sum_area_percent = 0.0
if "Area %" in df.columns:
    sum_area_percent = df["Area %"].fillna(0).sum()
sum_area_percent_int = int(round(sum_area_percent))

# (B) Sum <1 ppm (Emission) -> ยังคงใช้ค่าจาก "emission_decimal_series" ก่อนปัด
sum_less_1ppm_val = 0
if not emission_decimal_series.empty:
    sum_less_1ppm_val = emission_decimal_series[emission_decimal_series < 1].sum()
sum_less_1ppm_val_int = int(round(sum_less_1ppm_val))

# (C) Sum ≥ 1 ppm (Emission) -> ใช้จาก df["Emission (ug/g)"] (int)
sum_greater_1ppm_val = 0
if not df.empty and "Emission (ug/g)" in df.columns:
    sum_greater_1ppm_val = df.loc[df["Emission (ug/g)"] >= 1, "Emission (ug/g)"].sum()

# (D) Sum All Emission (ug/g) = sum_less_1ppm_val_int + sum_greater_1ppm_val
sum_all_emission_val = sum_less_1ppm_val_int + sum_greater_1ppm_val

# (E) Sum <1 ppm Area% -> จาก df
sum_area_less_1ppm = 0.0
if not df.empty and "Emission (ug/g)" in df.columns and "Area %" in df.columns:
    sum_area_less_1ppm = df.loc[df["Emission (ug/g)"] < 1, "Area %"].fillna(0).sum()
sum_area_less_1ppm_int = int(round(sum_area_less_1ppm))

# --------------------------------------------------------------------
# ส่วนที่ 6: ปรับ format ใน DataFrame
# --------------------------------------------------------------------
if not df.empty:
    if "Component RT" in df.columns:
        df["Component RT"] = df["Component RT"].round(2)
    if "Component Area" in df.columns:
        df["Component Area"] = df["Component Area"].fillna(0).astype(int)
    if "Area %" in df.columns:
        df["Area %"] = df["Area %"].round(2)

# --------------------------------------------------------------------
# ส่วนที่ 7: แสดง Data Table
# --------------------------------------------------------------------
st.subheader("Data Table")
st.dataframe(df, use_container_width=True)

# --------------------------------------------------------------------
# ส่วนที่ 8: แสดงผลสรุป
# --------------------------------------------------------------------
st.subheader("Summary")

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.text_input(
        "Sum All Area%",
        value=str(sum_area_percent_int),
        disabled=True,
        key="sum_all_area_percent"
    )

with col2:
    st.text_input(
        "Sum <1 ppm Area%",
        value=str(sum_area_less_1ppm_int),
        disabled=True,
        key="sum_area_less_1ppm"
    )

with col3:
    st.text_input(
        "Sum <1 ppm",
        value=str(sum_less_1ppm_val_int),
        disabled=True,
        key="sum_less_1ppm"
    )

with col4:
    st.text_input(
        "Sum ≥ 1 ppm",
        value=str(sum_greater_1ppm_val),
        disabled=True,
        key="sum_greater_1ppm"
    )

with col5:
    st.text_input(
        "Sum All Emission (ug/g)",
        value=str(sum_all_emission_val),
        disabled=True,
        key="sum_all_emission"
    )

# --------------------------------------------------------------------
# สร้าง Summary DataFrame สำหรับ Export
# --------------------------------------------------------------------
summary_df = pd.DataFrame({
    "Sum All Area%": [sum_area_percent_int],
    "Sum <1 ppm Area%": [sum_area_less_1ppm_int],
    "Sum <1 ppm (decimal)": [sum_less_1ppm_val_int],
    "Sum ≥ 1 ppm (int)": [sum_greater_1ppm_val],
    "Sum All Emission (ug/g)": [sum_all_emission_val]
})

# --------------------------------------------------------------------
# ส่วนที่ 9: Export to Excel
#   - ตัด Emission <1 ppm ออก (เหลือ Emission >= 1)
#   - ตัดคอลัมน์ Component Area
# --------------------------------------------------------------------
st.subheader("Export Data")

if st.button("Export to Excel"):
    # (1) Clone df => df_export
    df_export = df.copy()

    # (2) ตัดแถว Emission <1 ppm
    if "Emission (ug/g)" in df_export.columns:
        df_export = df_export[df_export["Emission (ug/g)"] >= 1]

    # (3) ตัดคอลัมน์ Component Area (หากมี)
    columns_to_drop = ["Component Area"]
    df_export.drop(columns=columns_to_drop, inplace=True, errors="ignore")

    # เลือกคอลัมน์ที่จะเหลือ
    desired_cols = ["Component RT", "Compound Name", "CAS#", "Area %", "Emission (ug/g)"]
    existing_cols = [c for c in desired_cols if c in df_export.columns]
    df_export = df_export[existing_cols]

    # (4) กำหนดชื่อไฟล์ตามชื่อไฟล์ที่อัปโหลด + "_export.xlsx"
    export_filename = "calculated_emission_filtered.xlsx"  # default
    if uploaded_file is not None:
        import os
        base_name = os.path.splitext(uploaded_file.name)[0]  # ตัดนามสกุล
        export_filename = f"{base_name}_export.xlsx"         # ex: "mydata_export.xlsx"

    # (5) เขียนลง Excel โดยมี 2 Sheet: "Data" และ "Summary"
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name="Data")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        writer.close()

    excel_data = output.getvalue()

    st.download_button(
        label="Download Excel",
        data=excel_data,
        file_name=export_filename,  # ใช้ชื่อไฟล์อิงจาก uploaded_file
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
