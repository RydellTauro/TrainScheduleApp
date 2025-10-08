# app.py
import streamlit as st
import TimeScheduleGenerator as tsg

st.set_page_config(page_title="Train Schedule Generator", layout="wide")
st.title("Train Schedule Generator")

st.write(
    """
    This app generates the train schedule from the Excel file and creates a Gantt chart
    with the list of trains in each composition.
    """
)

# Button to generate schedule
if st.button("Generate Schedule"):
    try:
        # Generate the HTML content
        html_content = tsg.generate_schedule_html()

        # Save the HTML file in the same folder
        html_file = "time_schedule_gantt_colored1.html"
        with open(html_file, "w", encoding="utf-8") as f:
            f.write(html_content)

        # Display the HTML in Streamlit
        st.components.v1.html(html_content, height=800, scrolling=True)

        # Provide a download button
        with open(html_file, "rb") as f:
            st.download_button(
                label="Download Schedule HTML",
                data=f,
                file_name=html_file,
                mime="text/html"
            )

        st.success("Schedule generated successfully!")

    except Exception as e:
        st.error(f"Error generating schedule: {e}")
