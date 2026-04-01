import streamlit as st
import pandas as pd
from datetime import datetime, time

#
st.set_page_config(
    page_title="Rehearsal Helper", page_icon="https://www.cdandfs.com/favicon.ico"
)

# Read data from XLSX file
classes_data = pd.read_excel("input.xlsx", sheet_name="classes")
rehearsals_data = pd.read_excel("input.xlsx", sheet_name="rehearsals")

# Main app
st.title("CD&FS May Rehearsals and Performances")

# Get user input for class name
class_name_list = classes_data["class_name"].unique()
class_name = st.selectbox("Enter class name", [""] + class_name_list.tolist())

# Filter data based on user input
if class_name:
    classes_match = classes_data[
        classes_data["class_name"].str.fullmatch(class_name, case=False)
    ]
    classes_match["time_of_day"] = classes_match["time_of_day"].apply(
        lambda x: x.strftime("%I:%M %p") if pd.notnull(x) else ""
    )
        classes_match["teacher"] = classes_match["teacher"].if pd.notnull(x) else ""
    )
    rehearsals_match = rehearsals_data[
        rehearsals_data["class_name"].str.fullmatch(class_name, case=False)
    ]

    # Display classes data
    if not classes_match.empty:
        st.write("### Class Information:")

        def df_to_html(df):
            return (
                "<style>"
                "table.cdfs { width:100%; border-collapse:collapse; font-size:0.88rem; }"
                "table.cdfs th { background:#1c1c1c; color:#fff; padding:6px 10px;"
                "  text-align:left; font-weight:600; white-space:nowrap; }"
                "table.cdfs td { padding:6px 10px; border-bottom:1px solid #e0e0e0;"
                "  vertical-align:top; white-space:normal; word-wrap:break-word; }"
                "table.cdfs tr:nth-child(even) td { background:#f7f5f2; }"
                "</style>"
                + df.to_html(index=False, classes="cdfs", border=0)
            )

        st.markdown(
            df_to_html(
                classes_match.rename(columns={
                    "class_name":    "Class Name",
                    "teacher":       "Teacher",
                    "assistant":     "Assistant",
                    "day_of_week":   "Class Day of Week",
                    "time_of_day":   "Class Time",
                    "cdfs_costume": "Costume Information",
                    "student_costume": "Pieces provided by student",
                    "dressing_room": "Dressing Room Assignment",
                })
            ),
            unsafe_allow_html=True,
        )

    # Display rehearsals data with download buttons
    if not rehearsals_match.empty:
        st.write("### Rehearsal and Performance Information:")
        st.write("##### Links to CD&FS Website Rehearsal Information:")
        st.write(
            "\t [Information for Technical Rehearsal](https://www.cdandfs.com/tech-rehearsals.html)"
        )
        st.write(
            "\t [Information for Dress Rehearsal](https://www.cdandfs.com/dress-rehearsals.html)"
        )
        st.write(
            "\t [Information for the Performances](https://www.cdandfs.com/performances.html)"
        )

        rehearsals_table = rehearsals_match.copy()
        rehearsals_table.sort_values(by=["date", "start_time"], inplace=True)
        rehearsals_table["Date"] = rehearsals_table["date"].apply(
            lambda x: x.strftime("%a, %b %d")
        )
        rehearsals_table["Start Time"] = rehearsals_table["start_time"].apply(
            lambda x: x.strftime("%I:%M %p")
        )
        rehearsals_table["End Time"] = rehearsals_table["end_time"].apply(
            lambda x: x.strftime("%I:%M %p")
        )
        rehearsals_table["Arrival Time"] = rehearsals_table["arrival_time"].apply(
            lambda x: x.strftime("%I:%M %p") if isinstance(x, time) else ""
        )
        # rehearsals_table["name"] = rehearsals_table.apply(
        #     lambda row: f"[{row['name']}]({row['url']})" if row["url"] else row["name"],
        #     axis=1,
        # )
        rehearsals_table = rehearsals_table[
            [
                "name",
                "Date",
                "class_name",
                "dance_name",
                "location",
                "Start Time",
                "End Time",
                "Arrival Time",
                "information",
            ]
        ]
        rehearsals_table = rehearsals_table.rename(
            columns={
                "name": "Rehearsal/Performance",
                "location": "Location",
                "class_name": "Class",
                "dance_name": "Dance Name",
                "information": "Information",
            }
        )
        rehearsals_table["Information"] = rehearsals_table["Information"].fillna("")
        rehearsals_table["Rehearsal/Performance"] = rehearsals_table["Rehearsal/Performance"].fillna("")

        st.write("##### Rehearsal & Performance Schedule as a Table:")
        st.markdown(df_to_html(rehearsals_table), unsafe_allow_html=True)

        st.write("##### Rehearsal & Performance Schedule as a List:")
        for index, row in rehearsals_table.iterrows():
            st.write(f"**{row['Rehearsal/Performance']}**")
            st.write(f"* *_Date_*: {row['Date']}")
            st.write(f"* *_Class_*: {row['Class']}")
            st.write(f"* *_Dance Name_*: {row['Dance Name']}")
            st.write(f"* *_Location_*: {row['Location']}")
            st.write(f"* *_Start Time_*: {row['Start Time']}")
            st.write(f"* *_End Time_*: {row['End Time']}")
            st.write(f"* *_Arrival Time_*: {row['Arrival Time']}")
            st.write(f"* *_Information_*: {row['Information']}")
            st.write("\n")
