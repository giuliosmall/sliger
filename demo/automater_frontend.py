"""Streamlit demo used at PyCon Italy 2023.

Expects a service-account JSON key named ``service_account.json`` in the
working directory, and a template presentation shared with that account.
"""

from __future__ import annotations

import os

import redirect as rd
import streamlit as st

from sliger import Sliger

TEMPLATE_PRESENTATION_ID = "1sROK5h0qjeyk0TLGCcCsq2MtCJCYDQO7O5f0EoaiDHY"

st.set_page_config(page_title="Sliger", page_icon="🐯")
st.markdown(
    "<h1 style='text-align: center;'>Slide of the Tiger - Sliger 🐯</h1>",
    unsafe_allow_html=True,
)

service_account_json = os.path.join(os.getcwd(), "service_account.json")

account_uuid = st.text_input("Enter the account UUID (1/4)")
company_name = st.text_input("Enter the company name (2/4)") if account_uuid else None
presentation_name = (
    st.text_input("Enter the presentation name (3/4)") if company_name else None
)
currency = (
    st.selectbox("Select the currency (4/4):", ["EUR €", "USD $", "GBP £"])
    if presentation_name
    else None
)

inputs = {
    "account_uuid": account_uuid,
    "company_name": company_name,
    "presentation_name": presentation_name,
    "currency": currency,
}

if st.button("Launch") and presentation_name:
    with st.spinner("Running..."), st.expander("Output", expanded=False):
        with rd.stdout:
            client = Sliger(
                service_account_json,
                TEMPLATE_PRESENTATION_ID,
                config_path="config.toml",
            )
            st.write("Duplicating presentation")
            new_presentation_id = client.duplicate_presentation(
                f"Slido @ {inputs['company_name']}",
                anyone_can_view=True,
            )
            client.presentation_id = new_presentation_id

            st.write("Jinjifying presentation")
            client.jinjify(inputs)

            st.write("Imagifying presentation")
            client.imagify(inputs)

    st.markdown(
        f"Done! [View presentation](https://docs.google.com/presentation/d"
        f"/{new_presentation_id}/edit)"
    )
