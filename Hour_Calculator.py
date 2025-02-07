import streamlit as st
import pandas as pd


st.title("DSC Hour Calculator")


uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    
    df = pd.read_excel(uploaded_file)


    if "Datum" not in df.columns:
        st.error("Error: The uploaded file must contain a 'Datum' column.")
    else:
        
        df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True, errors='coerce')

        df = df.dropna(subset=["Datum"])

        min_date = df["Datum"].min()
        max_date = df["Datum"].max()

        start_date = st.date_input("Start Date", min_date)
        end_date = st.date_input("End Date", max_date)

        # Time input section
        st.subheader("Set Time Values (in minutes)")
        t1 = st.number_input("T1 (gestanzt)", min_value=1, value=5)
        t2 = st.number_input("T2 (sample cutter)", min_value=1, value=10)
        t3 = st.number_input("T3 (Für die Auswertung)", min_value=1, value=15)

        if start_date > end_date:
            st.error("Error: Start Date cannot be after End Date.")
        else:
            filtered_df = df[(df["Datum"] >= pd.to_datetime(start_date)) & 
                             (df["Datum"] <= pd.to_datetime(end_date))]


            selected_columns = ["Fortlaufende Nummer", "Projekt", "Datum", 
                                "Probenform", "Messung Durchgeführt", "Auswertung Durchgeführt"]
            

            existing_columns = [col for col in selected_columns if col in filtered_df.columns]
            filtered_df = filtered_df[existing_columns]

 
            st.subheader("Filtered Data")
            st.dataframe(filtered_df)

            # Step 1: Count variations of "gestanzt" in "Probenform"
            gestanzt_pattern = r"gestanzt|gestanzz|gestanztz|gestantzt|gestanzr"
            gestanzt_count = filtered_df["Probenform"].astype(str).str.lower().str.contains(gestanzt_pattern, regex=True).sum()

            # Step 2: Count occurrences of "sample cutter" in "Probenform"
            sample_cutter_count = filtered_df["Probenform"].astype(str).str.lower().str.contains(r"sample cutter", regex=True).sum()

            # Step 3: Count occurrences of MH, AK, HD in "Messung Durchgeführt" and calculate time dynamically
            mh_time_messung = 0
            ak_time_messung = 0
            hd_time_messung = 0

            for _, row in filtered_df.iterrows():
                person = str(row["Messung Durchgeführt"]).strip()
                probenform = str(row["Probenform"]).lower()

                if pd.notna(person):
                    if pd.notna(probenform):
                        if any(variation in probenform for variation in ["gestanzt", "gestanzz", "gestanztz", "gestantzt", "gestanzr"]):
                            time_to_add = t1  # User-inputted T1
                        elif "sample cutter" in probenform:
                            time_to_add = t2  # User-inputted T2
                        else:
                            time_to_add = 0  # No time added for other cases

                        if person == "MH":
                            mh_time_messung += time_to_add
                        elif person == "AK":
                            ak_time_messung += time_to_add
                        elif person == "HD":
                            hd_time_messung += time_to_add

            # Step 4: Count occurrences of MH, AK, HD in "Auswertung Durchgeführt"
            auswertung_counts = filtered_df["Auswertung Durchgeführt"].astype(str).value_counts()
            mh_auswertung = auswertung_counts.get("MH", 0)
            ak_auswertung = auswertung_counts.get("AK", 0)
            hd_auswertung = auswertung_counts.get("HD", 0)

            # Step 5: Calculate total time (in minutes) for each person in "Auswertung Durchgeführt"
            mh_total_time_auswertung = mh_auswertung * t3  # User-inputted T3
            ak_total_time_auswertung = ak_auswertung * t3
            hd_total_time_auswertung = hd_auswertung * t3

            # Step 6: Convert total time to Hour:Minute format
            def format_time(minutes):
                hours = minutes // 60
                remaining_minutes = minutes % 60
                return f"{hours}:{remaining_minutes:02d}"

            mh_time_messung_formatted = format_time(mh_time_messung)
            ak_time_messung_formatted = format_time(ak_time_messung)
            hd_time_messung_formatted = format_time(hd_time_messung)

            mh_total_time_auswertung_formatted = format_time(mh_total_time_auswertung)
            ak_total_time_auswertung_formatted = format_time(ak_total_time_auswertung)
            hd_total_time_auswertung_formatted = format_time(hd_total_time_auswertung)


            summary_data = {
                "Category": [
                    "Total 'gestanzt' samples",
                    "Total 'sample cutter' samples",
                    "MH (Messung Durchgeführt)", "AK (Messung Durchgeführt)", "HD (Messung Durchgeführt)",
                    "MH (Auswertung Durchgeführt)", "AK (Auswertung Durchgeführt)", "HD (Auswertung Durchgeführt)"
                ],
                "Count": [
                    gestanzt_count,
                    sample_cutter_count,
                    filtered_df["Messung Durchgeführt"].astype(str).value_counts().get("MH", 0),
                    filtered_df["Messung Durchgeführt"].astype(str).value_counts().get("AK", 0),
                    filtered_df["Messung Durchgeführt"].astype(str).value_counts().get("HD", 0),
                    mh_auswertung, ak_auswertung, hd_auswertung
                ]
            }
            summary_df = pd.DataFrame(summary_data)

         
            st.subheader("Summary Table")
            st.dataframe(summary_df)

      
            st.subheader("Total Time each person spent at the DSC machine (Hour:Min)")
            st.write(f"**MH Total Time (Messung):** {mh_time_messung_formatted}")
            st.write(f"**AK Total Time (Messung):** {ak_time_messung_formatted}")
            st.write(f"**HD Total Time (Messung):** {hd_time_messung_formatted}")
            st.write(f"**MH Total Time (Auswertung):** {mh_total_time_auswertung_formatted}")
            st.write(f"**AK Total Time (Auswertung):** {ak_total_time_auswertung_formatted}")
            st.write(f"**HD Total Time (Auswertung):** {hd_total_time_auswertung_formatted}")
