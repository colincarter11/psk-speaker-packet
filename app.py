import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO

st.title("Pullspark Speaker Packet Generator")

st.markdown("""
Upload your completed **Onsite Packet (.xlsx)** and get a fully filled-in **Speaker Packet (.docx)** based on the official template.
""")

# Client uploads the Excel file
excel_file = st.file_uploader("Upload Onsite Packet", type=["xlsx"])

# Fixed Word template file (always stored locally)
TEMPLATE_PATH = "APP TEST Speaker Packet Template.docx"

def extract_context_from_excel(file, selected_speaker=None):
    # Read Event Details sheet (first tab)
    sheet1 = pd.read_excel(file, sheet_name="Event Details", header=None)
    sheet2 = pd.read_excel(file, sheet_name="Onsite Schedule")

    # Turn Column A+B into key-value dictionary
    kv = sheet1.set_index(0)[1].to_dict()

    # Build context from mapped fields
    context = {
        "event_name": kv.get("Event Name", ""),
        "dates": kv.get("Dates", ""),
        "time": kv.get("Time", ""),
        "location_name": kv.get("Location Name", ""),
        "location_address": kv.get("Location Address", ""),
        "event_audience_details": kv.get("Event Audience Details", kv.get("Evenet Audience Details", "")),
        "expected_attendance": kv.get("Expected Attendance", ""),
        "host_name_1": kv.get("Host Name 1", ""),
        "cell_phone_1": kv.get("Cell Phone 1", ""),
        "host_name_2": kv.get("Host Name 2", ""),
        "cell_phone_2": kv.get("Cell Phone 2", ""),
        "parking_details": kv.get("Parking Details", ""),
        "event_producer_email": kv.get("Event Producer Email", ""),
        "deadline": kv.get("Deadline", ""),
        "stage_layout": kv.get("Stage Layout", ""),
        "design": kv.get("Design", "")
    }

    # Schedule
    schedule_df = sheet2[sheet2["Time"].notna() & (sheet2["Time"] != "Time")]
    schedule_df = schedule_df.dropna(how='all', subset=["Time", "What", "Who"])
    schedule_df = schedule_df.rename(columns={"Time": "time", "What": "what", "Who": "who"})

    # Add fallback Speaker column
    if "Speaker" not in schedule_df.columns:
        schedule_df["Speaker"] = "All"

    # Filter for selected speaker or include all
    if selected_speaker and selected_speaker != "All Speakers":
        schedule_df = schedule_df[
            schedule_df["Speaker"].fillna("").str.lower().str.contains(selected_speaker.lower()) |
            schedule_df["Speaker"].fillna("").str.lower().str.contains("all")
        ]

    # Format time values
    schedule = schedule_df[["time", "what", "who"]].to_dict(orient="records")
    for item in schedule:
        try:
            if isinstance(item["time"], pd.Timestamp):
                item["time"] = item["time"].strftime("%-I:%M %p").lstrip("0")
            elif isinstance(item["time"], str) and item["time"].endswith(":00"):
                from datetime import datetime
                parsed_time = datetime.strptime(item["time"], "%H:%M:%S")
                item["time"] = parsed_time.strftime("%-I:%M %p").lstrip("0")
        except Exception:
            pass

    context["schedule"] = schedule
    return context

# Main UI flow
if excel_file:
    try:
        # Load speakers
        full_df = pd.read_excel(excel_file, sheet_name="Onsite Schedule")
        speakers_raw = full_df.get("Speaker", pd.Series(["All"])).dropna().unique()
        speakers = sorted(set([s.strip() for s in speakers_raw if str(s).strip() != ""]))
        speakers.insert(0, "All Speakers")

        # Dropdown
        selected_speaker = st.selectbox("Select a speaker to generate their packet:", speakers)

        # Generate based on selection
        if selected_speaker:
            context = extract_context_from_excel(excel_file, selected_speaker)

            doc = DocxTemplate(TEMPLATE_PATH)
            doc.render(context)

            output = BytesIO()
            doc.save(output)
            output.seek(0)

            st.success(f"✅ Speaker Packet generated for {selected_speaker}!")
            st.download_button(
                label="📥 Download Speaker Packet (.docx)",
                data=output,
                file_name=f"{selected_speaker.replace(' ', '_')}_Speaker_Packet.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
