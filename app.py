import streamlit as st
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from io import BytesIO

# Predefined locations
LOCATIONS = [
    "Chalmers", "Bay View Dental", "Portsoy Medical",
    "Fraserburgh Hospital", "Ugie Hospital", "Peterhead Hospital", "Macduff Vaccination"
]

# Initialize session state
if "entries" not in st.session_state:
    st.session_state.entries = []
if "start_time" not in st.session_state:
    st.session_state.start_time = None
if "air_temp" not in st.session_state:
    st.session_state.air_temp = 0
if "snow" not in st.session_state:
    st.session_state.snow = "No"
if "plough" not in st.session_state:
    st.session_state.plough = "No"
if "grit" not in st.session_state:
    st.session_state.grit = "No"
if "entry_added" not in st.session_state:
    st.session_state.entry_added = False

# If an entry was just added, reset the form
if st.session_state.entry_added:
    st.session_state.start_time = None
    st.session_state.air_temp = 0
    st.session_state.snow = "No"
    st.session_state.plough = "No"
    st.session_state.grit = "No"
    st.session_state.entry_added = False

st.title("Winter Maintenance Record")

# Location selector is outside the form
location = st.selectbox("Select Location", LOCATIONS)

# Manual start button
if st.button("Start"):
    st.session_state.start_time = datetime.now().strftime("%H:%M")
    st.rerun()

# Data entry section
st.write(f"Entering data for: **{location}**")

start_time_display = st.session_state.start_time if st.session_state.start_time else "Not started"
st.write(f"Entry started at: **{start_time_display}**")

# Use session state to store widget values immediately
snow_options = ["Yes", "No"]
plough_options = ["Yes", "No"]
grit_options = ["Yes", "No"]

st.number_input("Air Temperature (°C)", key="air_temp", step=1)
st.radio("Snow Present?", snow_options, key="snow", index=snow_options.index(st.session_state.snow))
st.radio("Plough Used?", plough_options, key="plough", index=plough_options.index(st.session_state.plough))
st.radio("Grit Spread?", grit_options, key="grit", index=grit_options.index(st.session_state.grit))

# "Add Entry" button now reads from session state
if st.button("Add Entry"):
    if st.session_state.start_time is None:
        st.error("Please click 'Start' before adding an entry.")
    else:
        finish_time = datetime.now().strftime("%H:%M")
        date_today = datetime.now().strftime("%d-%m-%Y")

        st.session_state.entries.append({
            "Date": date_today,
            "Start Time": st.session_state.start_time,
            "Finish Time": finish_time,
            "Location": location,
            "Air Temp": st.session_state.air_temp,
            "Snow": st.session_state.snow,
            "Plough": st.session_state.plough,
            "Grit": st.session_state.grit
        })

        st.session_state.entry_added = True
        st.success("Entry added successfully!")
        st.rerun()

# Display entries and download option
if st.session_state.entries:
    st.subheader("Current Entries")
    st.table(st.session_state.entries)

    # Function to create Word document
    def create_word(entries):
        doc = Document()
        doc.add_heading("Winter Maintenance Record", 0)

        # Create table with headers
        table = doc.add_table(rows=1, cols=8)
        # Set width for the 'Location' column (index 3)
        table.columns[3].width = Inches(1.2)
        hdr_cells = table.rows[0].cells
        headers = [
            'Date', 'Start Time', 'Finish Time', 'Location',
            'Air Temp', 'Snow', 'Plough Used', 'Grit Spread'
        ]
        for i, header in enumerate(headers):
            hdr_cells[i].text = header

        # Populate table rows
        for entry in entries:
            row_cells = table.add_row().cells
            row_cells[0].text = entry["Date"]
            row_cells[1].text = entry["Start Time"]
            row_cells[2].text = entry["Finish Time"]
            
            # Set font size for location to prevent wrapping
            location_cell = row_cells[3]
            p = location_cell.paragraphs[0]
            run = p.add_run(entry["Location"])
            run.font.size = Pt(8)

            row_cells[4].text = str(entry["Air Temp"])
            row_cells[5].text = entry["Snow"]
            row_cells[6].text = entry["Plough"]
            row_cells[7].text = entry["Grit"]

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    # Download button
    word_file = create_word(st.session_state.entries)
    if st.download_button(
        label="Download Word File",
        data=word_file,
        file_name=f"grit_report_{datetime.now().strftime('%Y-%m-%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ):
        st.session_state.entries = []
        st.success("Entries have been cleared after download!")
