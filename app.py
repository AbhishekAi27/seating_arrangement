import streamlit as st
import os
import shutil
import zipfile
from seating_arrangement import ExamSeatingSystem

def clean_directory_contents(dir_path):
    """
    Safely removes all contents of a directory without deleting the directory itself.
    This prevents 'Device or resource busy' errors with Docker volumes.
    """
    if not os.path.exists(dir_path):
        os.makedirs(dir_path, exist_ok=True)
        return

    for filename in os.listdir(dir_path):
        file_path = os.path.join(dir_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            st.error(f"Failed to delete {file_path}. Reason: {e}")

# --- Main App Configuration ---

# Define paths
DATA_DIR = "data"
PHOTOS_DIR = os.path.join(DATA_DIR, "photos")
INPUT_FILE = "input_data.xlsx"
OUTPUT_DIR = "output"
MAPPING_FILE = "roll-names-mapping.csv"

# Ensure directories exist
os.makedirs(PHOTOS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

st.title("Examination Seating & Attendance Management System")

# 1. File Uploads
uploaded_excel = st.file_uploader("1. Upload Input Excel (input_data.xlsx)", type=["xlsx"])
uploaded_zip = st.file_uploader("2. Upload Photos Zip (photos.zip)", type=["zip"])
uploaded_map = st.file_uploader("3. Upload Name Map (roll-names-mapping.csv)", type=["csv"])

# 2. Configuration
col1, col2 = st.columns(2)
with col1:
    buffer = st.number_input("Buffer (Empty seats per room)", min_value=0, value=5)
with col2:
    mode = st.selectbox("Allocation Mode", ["Dense", "Sparse"])

# 3. Execution Button
if st.button("Generate Seating Plan"):
    if uploaded_excel and uploaded_zip:
        
        # --- FIX: Use helper function instead of rmtree ---
        clean_directory_contents(OUTPUT_DIR)
        # -------------------------------------------------

        # Save Files
        with open(INPUT_FILE, "wb") as f:
            f.write(uploaded_excel.getbuffer())
            
        if uploaded_map:
            with open(MAPPING_FILE, "wb") as f:
                f.write(uploaded_map.getbuffer())
        
        # Clean Data dir before extracting (Optional but recommended)
        clean_directory_contents(PHOTOS_DIR)

        with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
            zip_ref.extractall(DATA_DIR)
            
        st.success("Files uploaded successfully. Processing...")

        # Run System
        try:
            sys_obj = ExamSeatingSystem()
            
            if sys_obj.load_data():
                # Process Schedule
                progress_bar = st.progress(0)
                total_slots = len(sys_obj.schedule)
                
                if total_slots == 0:
                    st.warning("No exams found in schedule.")
                
                for i, slot in enumerate(sys_obj.schedule):
                    sys_obj.allocate_session(slot, buffer, mode.lower())
                    if total_slots > 0:
                        progress_bar.progress((i + 1) / total_slots)
                
                # Generate All Reports (Excel + PDFs)
                sys_obj.generate_excel_reports()
                sys_obj.generate_attendance_sheets()
                
                st.success("Allocation Complete! PDFs generated.")

                # Zip Output for Download
                shutil.make_archive("exam_output", 'zip', OUTPUT_DIR)
                
                with open("exam_output.zip", "rb") as f:
                    st.download_button(
                        label="Download All Files (Zip)",
                        data=f,
                        file_name="exam_output.zip",
                        mime="application/zip"
                    )
            else:
                st.error("Failed to load data. Check logs/format.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
            import traceback
            st.text(traceback.format_exc())
    else:
        st.warning("Please upload at least the Excel file and Photos Zip.")