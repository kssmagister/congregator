import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Titel der App
st.title("Gruppenbildungs-App für den Lateintag 2023")

# Information Bereich auf der linken Seite
st.sidebar.text("Mit dieser App kann man")
st.sidebar.text("mit den Daten aus einer")
st.sidebar.text("Excel-Datei Gruppeneinteilungen")
st.sidebar.text("erstellen.")
st.sidebar.text("1. Lade eine Excel-Datei hoch")
st.sidebar.text("2. Wähle die Zahl der")
st.sidebar.text("   Gruppenmitglieder aus.")
st.sidebar.text("3. Lade die neu erstellte")
st.sidebar.text("   Excel-Datei herunter.")

# Schritt 1: Hochladen der Excel-Datei
uploaded_file = st.file_uploader("Bitte laden Sie eine Excel-Datei hoch", type=['xlsx'])

if uploaded_file is not None:
    data = pd.read_excel(uploaded_file, header=None, skiprows=1)
    data.columns = ['Name', 'Ort']
    
    # Zeitstempel und Dateiname erstellen
    timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    uploaded_file_name = uploaded_file.name.split('.')[0]
    output_file_name = f"{uploaded_file_name}-{timestamp}.xlsx"
    output_file = f"./Daten/{output_file_name}"

    # Möglichkeit zur Auswahl der Gruppengröße
    group_size = st.number_input('Wählen Sie die Gruppengröße', min_value=2, value=4)

    # Schritt 2: Zufällig die Reihenfolge der Daten ändern
    data_shuffled = data.sample(frac=1).reset_index(drop=True)

    # Schritt 3: Gruppieren Sie die Daten
    groups = []
    group = []
    unique_orts = set()

    for i, row in data_shuffled.iterrows():
        if len(group) < group_size:
            if row['Ort'] not in unique_orts:
                group.append({'Name': row['Name'], 'Ort': row['Ort']})
                unique_orts.add(row['Ort'])
            else:
                # Suche nach einer Zeile mit einem anderen Ort, um Duplikate zu vermeiden
                for j, next_row in data_shuffled.loc[i:].iterrows():
                    if next_row['Ort'] not in unique_orts:
                        group.append({'Name': next_row['Name'], 'Ort': next_row['Ort']})
                        unique_orts.add(next_row['Ort'])
                        data_shuffled.drop(j, inplace=True)
                        break
                else:
                    # Wenn kein anderer Ort gefunden wird, fügen Sie den aktuellen Ort hinzu
                    group.append({'Name': row['Name'], 'Ort': row['Ort']})
                    unique_orts.add(row['Ort'])
            if len(group) == group_size:
                groups.append(group)
                group = []
                unique_orts = set()
        else:
            groups.append(group)
            group = [{'Name': row['Name'], 'Ort': row['Ort']}]
            unique_orts = {row['Ort']}

    if group:
        groups.append(group)

    # Schritt 4: Schreiben Sie die erstellten Gruppen in eine neue Excel-Datei
    with pd.ExcelWriter(output_file) as writer:
        for i, group in enumerate(groups):
            group_df = pd.DataFrame(group)
            group_df.to_excel(writer, sheet_name=f'Gruppe {i+1}', index=False)

    st.success("Die Gruppen wurden erfolgreich erstellt.")
    
    # Option zum Herunterladen der neuen Excel-Datei
    if st.button('Herunterladen'):
        with open(output_file, 'rb') as file:
            st.download_button(
                label="Herunterladen",
                data=file,
                file_name=output_file_name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
    
    # Löschen der temporären Datei nach dem Herunterladen
    if os.path.exists(output_file):
        os.remove(output_file)
else:
    st.warning("Bitte laden Sie eine Datei hoch.")
