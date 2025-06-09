import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import tempfile

st.title("Abwesenheitsstatistik Verarbeiter")

uploaded_csv = st.file_uploader("Lade die CSV-Datei hoch", type=["csv"])
uploaded_xlsx = st.file_uploader("Lade die XLSX-Vorlage hoch", type=["xlsx"])
sheet_name = st.text_input("Tabellenblattname", value="März-Mai 24")

if uploaded_csv and uploaded_xlsx:
    df = pd.read_csv(uploaded_csv, header=1, sep=';')
    result_dict = {}

    for _, row in df.iterrows():
        if pd.isna(row['Lehrperson']):
            continue
        key = row['Lehrperson']
        vom = pd.to_datetime(row['vom'], dayfirst=True)
        bis = pd.to_datetime(row['bis'], dayfirst=True)
        if vom.year == 2024: vom = vom.replace(year=2023)
        if bis.year == 2024: bis = bis.replace(year=2023)
        result_dict.setdefault(key, []).append([vom, bis])

    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(uploaded_xlsx.read())
        tmp_path = tmp.name

    wb = load_workbook(tmp_path, data_only=True)
    if sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
    else:
        sheet = wb.create_sheet(sheet_name)

    dates_in_sheet = []
    initial_date = None
    for col in range(7, sheet.max_column + 1):
        date_value = sheet.cell(row=2, column=col).value
        day_name = sheet.cell(row=5, column=col).value
        try:
            if col == 7:
                initial_date = pd.to_datetime(date_value, format="%d.%m.%Y")
                if day_name not in ['Sa', 'So']:
                    dates_in_sheet.append((col, initial_date))
            elif isinstance(date_value, str) and date_value.startswith('=') and initial_date:
                offset = int(date_value.split('+')[-1])
                d = initial_date + pd.Timedelta(days=offset)
                if day_name not in ['Sa', 'So']:
                    dates_in_sheet.append((col, d))
            else:
                d = pd.to_datetime(date_value, format="%d.%m.%Y")
                if day_name not in ['Sa', 'So']:
                    dates_in_sheet.append((col, d))
        except:
            continue

    for row_idx, (key, periods) in enumerate(result_dict.items(), start=6):
        sheet.cell(row=row_idx, column=3, value=key)
        for vom, bis in periods:
            for col, date in dates_in_sheet:
                if vom <= date <= bis:
                    sheet.cell(row=row_idx, column=col, value='x')

    output = io.BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)

    st.success("Datei verarbeitet!")
    st.download_button("Download der bearbeiteten Excel-Datei", data=output,
                       file_name="Personalkennzahlen.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
