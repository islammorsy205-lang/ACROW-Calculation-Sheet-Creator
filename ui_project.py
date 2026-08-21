import os
import fitz
import json
import streamlit as st
from datetime import date
from helpers import get_best_image_match
from config import SHORING_OPTIONS_SLAB, SECTIONS_DB

def render_project_details():
    # 1. زرار الـ Upload بقى صغير وشيك ومدمج فوق تفاصيل المشروع
    c_up, c_space = st.columns([1, 4])
    with c_up:
        uploaded_proj = st.file_uploader("📂 Load Project (.acrow)", type=['acrow', 'json'], key="ui_proj_upload")
        if uploaded_proj is not None:
            try:
                saved_data = json.load(uploaded_proj)
                for k, v in saved_data.items():
                    if "download" not in k.lower() and "upload" not in k.lower():
                        st.session_state[k] = v
                st.success("✅ Loaded!")
                st.rerun()
            except:
                st.error("❌ Invalid File!")

    st.subheader("1. Project Details & References")
    col1, col2, col3 = st.columns(3)

    with col1: 
        project_name = st.text_input("Main Project Name", "Acrow Mega Project", key="ui_proj_name")
        contractor = st.text_input("Client / Contractor", "Main Contractor", key="ui_contractor")
        # 2. إضافة خانة إدخال اسم المهندس يدوياً
        calc_by_input = st.text_input("Calculated By (Initials)", "I.M", max_chars=5, key="ui_calc_by")
        calc_by = f"Eng. {calc_by_input}" if calc_by_input else "Eng."

    with col2: 
        calc_subject = st.text_input("Structure Element", "CALCULATION SHEET FOR SOLID SLAB", key="ui_calc_subject")
        system_name = st.selectbox(
            "Formwork System Name", 
            SHORING_OPTIONS_SLAB + [
                "Timber H20 & Soldier System", 
                "Acrow Beam S12 & Soldier System", 
                "Eco-form Panel System", 
                "Tech-form Panel System", 
                "Curved Steel Panel System", 
                "Circular Steel Panel System"
            ],
            key="ui_system_name"
        )

    with col3: 
        # 3. تثبيت حرف الـ S- في رقم المشروع
        proj_no = st.text_input("Project No.", "S-2026", key="ui_proj_no")
        date_val = date.today().strftime("%d/%m/%Y")
        chk_by = "Eng. M.F."
        st.text_input("Checked By", chk_by, disabled=True)

    st.markdown("**Select Design Code, Cover Image & Method Statements:**")
    col_r1, col_r2, col_r3 = st.columns(3)

    with col_r1: 
        ref_code = st.selectbox("Design Codes & References:", ["British Standard (BS)", "American Code (ACI)", "None"], key="ui_ref_code")

    with col_r2:
        # 4. اختيار الصورة الذكي بناءً على العنصر والسيستم
        av_img = [f for f in os.listdir('.') if f.lower().endswith(('.jpg', '.jpeg', '.png')) and not f.startswith('ref_')]
        best_img_idx = get_best_image_match(calc_subject, system_name, av_img) if av_img else 0
        cover_img = st.selectbox("Cover Page Image:", av_img if av_img else ["No images found."], index=best_img_idx, key="ui_cover_img")

    with col_r3:
        av_ds = [f for f in os.listdir('.') if f.lower().endswith('.pdf') and f.lower().startswith('data sheet')]
        c1, c2 = st.columns([10, 1])
        with c1: 
            data_sheets = st.multiselect("Select Data Sheets:", av_ds, key="ui_data_sheets")
        with c2:
            st.markdown("<div style='margin-top: 29px;'></div>", unsafe_allow_html=True)
            if data_sheets:
                merged_ds = fitz.open()
                for pdf_file in data_sheets:
                    with fitz.open(pdf_file) as doc_pdf: 
                        merged_ds.insert_pdf(doc_pdf)
                st.download_button("⬇️", data=merged_ds.write(), file_name="Data_Sheets.pdf", mime="application/pdf", key="ui_download_ds")
                merged_ds.close()

    def_sec, def_main = 0, 2 
    for f in data_sheets:
        if "h20" in f.lower(): 
            def_sec = list(SECTIONS_DB.keys()).index("Timber H20")
        if "soldier" in f.lower(): 
            def_main = list(SECTIONS_DB.keys()).index("Soldier U100")
            
    return project_name, contractor, calc_subject, system_name, proj_no, calc_by, date_val, chk_by, ref_code, cover_img, data_sheets, def_sec, def_main