import streamlit as st
import subprocess
import os

st.title("Train Schedule App")

if st.button("Generate Schedule"):
    # Run your TimeScheduleGenerator.py
    result = subprocess.run(
        ["python", "TimeScheduleGenerator.py"],
        capture_output=True,
        text=True
    )
    st.text("Output:")
    st.text(result.stdout)
    st.text(result.stderr)

    # Display the HTML if generated
    html_file = "time_schedule_gantt_colored1.html"
    if os.path.exists(html_file):
        with open(html_file, "r", encoding="utf-8") as f:
            html_content = f.read()
        st.components.v1.html(html_content, height=800, scrolling=True)
