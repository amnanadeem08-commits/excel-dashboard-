"""Streamlit UI for Excel Dashboard Studio."""

import streamlit as st

from config.palettes import PALETTE_NAMES
from config.template_options import CURRENCY_OPTIONS, STYLE_OPTIONS, TEMPLATE_TYPES
from core.workbook_builder import build_workbook

st.set_page_config(page_title="Excel Dashboard Studio", layout="centered")

st.title("Excel Dashboard Studio")


template_type = st.selectbox("Template type", TEMPLATE_TYPES)

st.subheader("Business details")
with st.form("business_details_form"):
    business_name = st.text_input("Business name")
    phone = st.text_input("Phone")
    email = st.text_input("Email")
    address = st.text_area("Address", height=80)
    website = st.text_input("Website")

    col_currency, col_tax = st.columns(2)
    with col_currency:
        currency = st.selectbox("Currency", CURRENCY_OPTIONS)
    with col_tax:
        tax_rate = st.number_input("Tax/VAT/GST percentage", min_value=0.0, max_value=100.0, value=0.0, step=0.5)

    logo = st.file_uploader("Logo upload", type=["png", "jpg", "jpeg"])
    palette_name = st.selectbox("Color palette", PALETTE_NAMES)
    style_name = st.selectbox("Style", STYLE_OPTIONS)

    generate = st.form_submit_button("Generate Workbook", type="primary")

if generate:
    business_details = {
        "business_name": business_name.strip(),
        "phone": phone.strip(),
        "email": email.strip(),
        "address": address.strip(),
        "website": website.strip(),
        "currency": currency,
        "tax_rate": tax_rate,
    }

    try:
        filename, workbook_bytes, file_path = build_workbook(
            template_type=template_type,
            business_details=business_details,
            palette_name=palette_name,
            style_name=style_name,
            uploaded_logo=logo,
        )
        st.session_state["generated_workbook"] = workbook_bytes
        st.session_state["generated_filename"] = filename
        st.session_state["generated_path"] = str(file_path)
        st.success(f"Workbook generated: {filename}")
    except Exception as exc:
        st.error(f"Could not generate workbook: {exc}")

if "generated_workbook" in st.session_state:
    st.download_button(
        label="Download .xlsx",
        data=st.session_state["generated_workbook"],
        file_name=st.session_state["generated_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.info(f"Saved locally to: {st.session_state['generated_path']}")

