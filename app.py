# app.py
import streamlit as st
import TimeScheduleGenerator as tsg  # make sure this file is in the same folder

st.set_page_config(page_title="Train Schedule Generator", layout="wide")
st.title("Train Schedule Generator")

st.write("Upload your `Tankenliste.xlsm` Excel file, then click 'Generate Schedule'.")

uploaded_file = st.file_uploader(
    "Upload Tankenliste.xlsm", 
    type=["xlsm"]
)

if uploaded_file is not None:
    if st.button("Generate Schedule"):
        try:
            # Generate HTML from uploaded Excel file
            html_content = tsg.generate_schedule_html(uploaded_file)  # update your function to accept file-like object
            
            # Provide download button
            st.download_button(
                label="Download Schedule HTML",
                data=html_content,
                file_name="time_schedule_gantt_colored1.html",
                mime="text/html"
            )
            
            # Optionally, display the schedule in the app
            st.components.v1.html(html_content, height=800, scrolling=True)
            
            st.success("Schedule generated successfully!")
        except Exception as e:
            st.error(f"Error generating schedule: {e}")
else:
    st.info("Please upload your Excel file to proceed.")
