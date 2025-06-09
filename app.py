import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import tempfile

st.title("Abwesenheits-Eintrager")

uploaded_csv = st.file_uploader("CSV-Datei hochladen", type=["csv"])
uploaded_xlsx = st.file_uploader("XLSX-Vorlage mit Makros hochladen", type=["xlsx"])
sheet_name = st.text_input("Name des Tabellenblatts", value="März-Mai 24")

# Sobald beide Dateien da sind, geht’s los
if uploaded_csv and uploaded_xlsx:

    with st.spinner("Verarbeite..."):

        # CSV einlesen
        df = pd.read_csv(uploaded_csv, header=1, sep=';')
        result_dict = {}

        for _, row in df.iterrows():
            if pd.isna(row['Lehrperson']):
                continue

            key = row['Lehrperson']
            vom = pd.to_datetime(row['vom'], dayfirst=True)
            bis = pd.to_datetime(row['bis'], dayfirst=True)

            if vom.year == 2024:
                vom = vom.replace(year=2023)
            if bis.year == 2024:
                bis = bis.replace(year=2023)

            result_dict.setdefault(key, []).append([vom, bis])

        # XLSX-Datei temporär schreiben
        xlsx_bytes = uploaded_xlsx.read()
        if len(xlsx_bytes) == 0:
            st.error("Die XLSX-Datei ist leer oder beschädigt.")
            st.stop()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(xlsx_bytes)
            tmp_path = tmp.name

        wb = load_workbook(tmp_path, data_only=True)

        # Blatt auswählen oder erstellen
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            sheet = wb.create_sheet(sheet_name)

        # Datumsspalten erkennen
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

        # Daten eintragen
        for row_index, (key, periods) in enumerate(result_dict.items(), start=6):
            sheet.cell(row=row_index, column=3, value=key)
            for vom, bis in periods:
                if vom and bis:
                    for col, date in dates_in_sheet:
                        if date and vom <= date <= bis:
                            sheet.cell(row=row_index, column=col, value='x')

        # Workbook speichern
        output = io.BytesIO()
        wb.save(output)
        wb.close()
        output.seek(0)

        st.success("Verarbeitung abgeschlossen.")
        st.download_button(
            label="📥 Download fertige Excel-Datei",
            data=output,
            file_name="Personalkennzahlen.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
