# app.py
import streamlit as st
from TimeScheduleGenerator import generate_schedule_html

st.title("Train Schedule Generator")

st.write("Click the button to generate the train schedule from the latest Excel on OneDrive.")

if st.button("Generate Schedule"):
    try:
        html_file = generate_schedule_html()
        st.success("Schedule generated successfully!")
        with open(html_file, "rb") as f:
            st.download_button(
                label="Download Schedule HTML",
                data=f,
                file_name=html_file,
                mime="text/html"
            )
    except Exception as e:
        st.error(f"Error generating schedule: {e}")
