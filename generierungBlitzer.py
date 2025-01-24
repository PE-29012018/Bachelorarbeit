import pandas as pd
from math import radians, sin, cos, sqrt, atan2
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Pfad zur Excel-Datei
excel_datei = 'Testdatei.xlsx'

# Name des neuen Tabellenblatts
neues_tabellenblatt = 'Kamera auf Hauptfahrstreifen'

# Lese die Excel-Datei
daten = pd.read_excel(excel_datei, sheet_name='Anschlussstellen & Knoten 2022')

# Filtere die Datensätze für Autobahnen und Hauptfahrstreifen
autobahnen = daten[daten['Typ'] == 'Autobahn']
hauptfahrstreifen = daten[daten['Typ'] == 'Hauptfahrstreifen']

# Erstelle ein neues DataFrame für die Kameras auf dem Hauptfahrstreifen
kameras = pd.DataFrame(columns=daten.columns)

# Iteriere über die Autobahnen
for autobahn in autobahnen['Autobahn'].unique():
    # Filtere die Datensätze für die aktuelle Autobahn
    autobahn_daten = autobahnen[autobahnen['Autobahn'] == autobahn]

    # Ermittle den maximalen Kilometerwert der Autobahn
    max_kilometer = autobahn_daten['Kilometer'].max()

    # Bestimme die Anzahl der Kameras, die gleichmäßig platziert werden sollen
    anzahl_kameras = int(max_kilometer / 5000)  # Annahme: 1 Kamera pro 5 km

    # Berechne den Abstand zwischen den Kameras
    abstand = max_kilometer / anzahl_kameras

    # Iteriere über die Kilometerwerte und platziere die Kameras
    for i in range(1, anzahl_kameras + 1):
        kilometer = i * abstand

        # Finde den nächstgelegenen Kilometerwert in den Datensätzen
        naechster_kilometer = min(autobahn_daten['Kilometer'], key=lambda x: abs(x - kilometer))

        # Filtere die Datensätze für den nächstgelegenen Kilometerwert
        naechster_daten = autobahn_daten[autobahn_daten['Kilometer'] == naechster_kilometer]

        # Füge die Koordinaten zum neuen DataFrame hinzu
        kameras = kameras.append(naechster_daten)

# Setze den Wert des Typs auf 'Hauptfahrstreifen'
kameras['Typ'] = 'Hauptfahrstreifen'

# Generiere die Koordinaten für die Kameras
for index, row in kameras.iterrows():
    latitude = row['Latitude']
    longitude = row['Longitude']

    # Konvertiere die Koordinaten von Grad in Radian
    lat_rad = radians(latitude)
    lon_rad = radians(longitude)

    # Berechne die Abweichungen in der Breiten- und Längengraden für eine Entfernung von 2 bis 5 km
    delta_lat = radians(2) * 0.009  # 1 km entspricht ca. 0.009 Breitengraden
    delta_lon = radians(2) * 0.009 / cos(lat_rad)  # 1 km entspricht ca. 0.009 Längengraden (abhängig von der Breite)

    # Berechne die Koordinaten der Kamera auf dem Hauptfahrstreifen
    new_latitude = latitude + delta_lat
    new_longitude = longitude + delta_lon

    # Füge die Koordinaten zum DataFrame hinzu
    kameras.at[index, 'Latitude'] = new_latitude
    kameras.at[index, 'Longitude'] = new_longitude

    # Koordinaten in der Konsole ausgeben
    print(f"Kamera {index + 1}: Latitude: {new_latitude}, Longitude: {new_longitude}")

# Erstelle eine neue Excel-Datei mit dem aktualisierten DataFrame
with pd.ExcelWriter('Testdatei_neu.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    writer.book = load_workbook(excel_datei, read_only=False, keep_vba=True)
    kameras.to_excel(writer, sheet_name=neues_tabellenblatt, index=False)

print("Die Kameras auf dem Hauptfahrstreifen wurden erfolgreich generiert.")
