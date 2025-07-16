import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Title
st.title("🛏️ Room Occupancy Report Generator")

# Upload
uploaded_file = st.file_uploader("Upload the DAILY SALES Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.info("📄 Processing Excel sheets...")

    rates = pd.ExcelFile(uploaded_file)
    all_data = []

    for sheet_name in rates.sheet_names:
        df_raw = rates.parse(sheet_name, header=None)

        # Extract the date (3rd row, A column)
        sheet_date = pd.to_datetime(df_raw.iloc[2, 0], errors='coerce')

        # Combine row 4 and 5 headers
        row_main = df_raw.iloc[3].fillna('').astype(str)
        row_sub = df_raw.iloc[4].fillna('').astype(str)

        combined_columns = []
        for main, sub in zip(row_main, row_sub):
            main = main.strip()
            sub = sub.strip()
            if not main and not sub:
                combined_columns.append(None)
            else:
                combined_columns.append(f"{main} ({sub})" if sub else main)

        df_raw.columns = combined_columns
        df_raw = df_raw.loc[:, df_raw.columns.notna()]
        df_cleaned = df_raw.iloc[6:].reset_index(drop=True)

        # Remove rows after 'Function Room'
        particulars_col = next((col for col in df_cleaned.columns if 'Particulars' in str(col)), None)
        if particulars_col:
            stop_index = df_cleaned[df_cleaned[particulars_col] == 'Function Room'].index
            if not stop_index.empty:
                df_cleaned = df_cleaned.loc[:stop_index[0]]

        df_cleaned["Date"] = sheet_date

        cols = df_cleaned.columns.tolist()
        cols = ['Date'] + [col for col in cols if col != 'Date']
        df_cleaned = df_cleaned[cols]

        all_data.append(df_cleaned)

    df = pd.concat(all_data, ignore_index=True)

    df['Particulars'] = df['Particulars'].astype(str).str.strip()
    
    # Room mapping
    room_mapping_data = [
        ("Room 201A", "Deluxe Studio Room", "Pamana"),
        ("Room 203A", "Deluxe Studio Room", "Pamana"),
        ("Room 204A", "Deluxe Studio Room - Seaview", "Pamana"),
        ("Room 205A", "Deluxe Studio Room - Seaview", "Pamana"),
        ("Room 202A", "Deluxe Triple Room", "Pamana"),
        ("Room 301A", "Deluxe Triple Room", "Pamana"),
        ("Room 302A", "Deluxe Triple Room", "Pamana"),
        ("Room 303A", "Deluxe Double Room", "Pamana"),
        ("Room 304A", "Deluxe Double Room", "Pamana"),
        ("Room 305A", "Deluxe Double Room - Seaview", "Pamana"),
        ("Room 306A", "Deluxe Double Room - Seaview", "Pamana"),
        ("Room 307A", "Deluxe Double Room - Seaview", "Pamana"),
        ("Room 101B", "Dormitory - 10pax", "Pamana"),
        ("Room 102B", "Dormitory - 6pax", "Pamana"),
        ("Room 201B", "Studio Room - Seaview", "Pamana"),
        ("Room 202B", "Studio Room - Seaview", "Pamana"),
        ("Room 203B", "Studio Room - Seaview", "Pamana"),
        ("Room 204B", "Studio Room - Seaview", "Pamana"),
        ("Room 201C", "Double Room", "Annex"),
        ("Room 202C", "Double Room", "Annex"),
        ("Room 203C", "Double Room", "Annex"),
        ("Room 204C", "Double Room", "Annex"),
        ("Room 101D", "Standard Double Room", "Annex"),
        ("Room 102D", "Standard Double Room", "Annex"),
        ("Bahay Kubo", "Bahay Kubo", "Annex"),
    ]
    room_map_df = pd.DataFrame(room_mapping_data, columns=['Particulars', 'Room Type', 'Room Group'])

    # Preprocess
    df = df.merge(room_map_df, on='Particulars', how='left')
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
    df['Rooms Rates'] = pd.to_numeric(df['Rooms Rates'], errors='coerce').fillna(0)

    # Output Excel buffer
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    border_fmt = workbook.add_format({'border': 1})
    center_bold_fmt = workbook.add_format({'bold': True, 'align': 'center'})
    header_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'font_color': 'red'})
    money_fmt = workbook.add_format({'num_format': '"₱"#,##0.00', 'border': 1})

    columns_per_room = ['DATE', 'TOTAL ROOMS', 'NO. OF OCCUPIED ROOMS', 'OCCUPANCY PERCENTAGE', 'AMOUNT']
    col_widths = [15, 15, 20, 22, 15]

    for group_name in ["Pamana", "Annex"]:
        room_map_group = room_map_df[room_map_df['Room Group'] == group_name]
        room_types = room_map_group['Room Type'].unique().tolist()
        worksheet = workbook.add_worksheet(group_name)
        writer.sheets[group_name] = worksheet
        row = 0
        for i in range(0, len(room_types), 2):
            for offset in range(2):
                if i + offset >= len(room_types):
                    continue
                room_type = room_types[i + offset]
                group = room_map_group[room_map_group['Room Type'] == room_type]
                total_rooms = len(group)
                room_df = df[(df['Room Type'] == room_type) & (df['Room Group'] == group_name) & (df['Date'].notna())]
                if room_df.empty:
                    continue
                start_date = room_df['Date'].min().normalize()
                end_date = room_df['Date'].max().normalize()
                date_range = pd.date_range(start=start_date, end=end_date)
                data_rows = []
                for date in date_range:
                    day_data = room_df[room_df['Date'].dt.date == date.date()]
                    occupied = int(day_data['No. of Rooms'].sum())
                    occupancy = f"{(occupied / total_rooms) * 100:.0f}%" if total_rooms else "0%"
                    amount = day_data['Rooms Rates'].sum()
                    data_rows.append([date.strftime("%b %d,%Y"), total_rooms, occupied, occupancy, amount if amount > 0 else ""])
                total_occupied = sum(row[2] for row in data_rows)
                total_possible = total_rooms * len(date_range)
                overall_percent = f"{(total_occupied / total_possible) * 100:.2f}%" if total_possible else "0%"
                total_amount = sum(row[4] if isinstance(row[4], (int, float)) else 0 for row in data_rows)
                data_rows.append(["TOTAL", total_rooms * len(date_range), total_occupied, overall_percent, total_amount])
                data_rows.append(["OCC. PERCENT.", "", "", overall_percent, ""])
                base_col = 1 + offset * (len(columns_per_room) + 1)
                row += 1
                title = f"{room_type.upper()} ({', '.join(group['Particulars'])})"
                worksheet.merge_range(row, base_col, row, base_col + len(columns_per_room) - 1, title, header_fmt)
                for j, col_name in enumerate(columns_per_room):
                    worksheet.write(row + 1, base_col + j, col_name, center_bold_fmt)
                    worksheet.set_column(base_col + j, base_col + j, col_widths[j])
                for k, row_data in enumerate(data_rows):
                    for j, val in enumerate(row_data):
                        fmt = money_fmt if j == 4 and isinstance(val, (float, int)) else border_fmt
                        worksheet.write(row + 2 + k, base_col + j, val, fmt)
            row += len(date_range) + 4

    writer.close()
    output.seek(0)

    st.success("✅ Report generated!")
    st.download_button("📥 Download Excel Report", output, file_name="room_occupancy_report.xlsx")
