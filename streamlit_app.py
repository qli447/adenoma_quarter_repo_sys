import streamlit as st
from report_generator import generate_report

# 页面配置
st.set_page_config(page_title="Gopath Report Generator", layout="centered")
st.title("📊 Gopath Quarterly Adenoma Report")

# 选择facility
facility_list = st.secrets["FACILITY_LIST"]
facility = st.selectbox("Select a facility:", facility_list)

# Generate Button
if st.button("Generate Report"):
    with st.spinner("Generating report..."):
        ppt_path = generate_report(facility)
        if ppt_path is None:
            st.warning("No data found for this facility in the previous quarter.")
        else:
            with open(ppt_path, "rb") as f:
                st.success("Report generated successfully!")
                st.download_button(
                    label="📥 Download PowerPoint Report",
                    data=f,
                    file_name=ppt_path.split("/")[-1],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

