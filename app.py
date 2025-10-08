# app.py
import streamlit as st
import TimeScheduleGenerator as tsg
import io

st.set_page_config(page_title="Train Schedule Generator", layout="wide")
st.title("Train Schedule Generator")

st.markdown("Upload your Excel file (`Tankenliste.xlsm`) to generate the schedule.")

# File uploader
uploaded_file = st.file_uploader("Choose Excel file", type=["xlsm"])
if uploaded_file is not None:
    # Save uploaded file temporarily
    with open("Tankenliste.xlsm", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    st.success("File uploaded successfully!")

    if st.button("Generate Schedule"):
        try:
            # Generate HTML string
            html_content = tsg.generate_schedule_html()
            
            # Save HTML file as time_schedule_gantt_colored1.html
            html_file_name = "time_schedule_gantt_colored1.html"
            with open(html_file_name, "w", encoding="utf-8") as f:
                f.write(html_content)
            
            # Show HTML in Streamlit
            st.components.v1.html(html_content, height=800, scrolling=True)
            
            # Provide download button
            with open(html_file_name, "rb") as f:
                st.download_button(
                    label="Download Schedule HTML",
                    data=f,
                    file_name=html_file_name,
                    mime="text/html"
                )
            
            st.success("Schedule generated successfully!")

        except Exception as e:
            st.error(f"Error generating schedule: {e}")

        except Exception as e:
            st.error(f"Error generating schedule: {e}")
