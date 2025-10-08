# app.py
import streamlit as st
import TimeScheduleGenerator as tsg

st.set_page_config(page_title="Train Schedule Generator", layout="wide")
st.title("Train Schedule Generator")

uploaded_file = st.file_uploader("Upload Tankenliste.xlsm", type=["xlsm"])
if uploaded_file is not None:
    with open("Tankenliste.xlsm", "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success("File uploaded!")

    if st.button("Generate Schedule"):
        try:
            html_content = tsg.generate_schedule_html()
            html_file_name = "time_schedule_gantt_colored1.html"
            with open(html_file_name, "w", encoding="utf-8") as f:
                f.write(html_content)
            st.components.v1.html(html_content, height=800, scrolling=True)
            with open(html_file_name, "rb") as f:
                st.download_button("Download Schedule HTML", f, file_name=html_file_name, mime="text/html")
            st.success("Schedule generated successfully!")
        except Exception as e:
            st.error(f"Error generating schedule: {e}")
