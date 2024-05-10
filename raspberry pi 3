import serial
import time
import csv
import openpyxl
import os
from datetime import datetime

# Configureer de seriÃ«le verbinding
ser = serial.Serial('/dev/ttyUSB0', 115200)  # Gebruik de juiste poort van je slimme meter

# Initialisatie
buffer = []
start_time = time.time()  # Definieer start_time hier
save_interval = 3600  # 60 minuten in seconden
daily_total_energy = 0  # Totale opbrengst gedurende de dag in kilowattuur
last_date = datetime.now().day  # Houd de laatste datum bij

while True:
    try:
        current_time = time.time()
        current_datetime = datetime.fromtimestamp(current_time)

        # Huidige datum en tijd in het Amerikaanse notatieformaat
        current_date = current_datetime.strftime('%d-%m-%Y')
        current_time_str = current_datetime.strftime('%H:%M:%S')  # Gebruik 24-uursnotatie

        # Als de datum veranderd is, zet daily_total_energy op nul
        if current_datetime.day != last_date:
            daily_total_energy = 0
            last_date = current_datetime.day  # Update last_date

        # Aanmaken van het Excel-werkboek met de datum in de bestandsnaam
        excel_file = f'gemiddelde_stroom_{current_date}.xlsx'

        # Aanmaken van het CSV-back-upbestand met de datum in de bestandsnaam
        csv_file = f'gemiddelde_stroom_{current_date}.csv'

        # Lees de ruwe gegevens van de slimme meter
        raw_data = ser.readline().decode().strip()

        # Controleer of de ruwe gegevens de actieve teruggeleverde stroom bevatten
        if "1-0:2.7.0" in raw_data:
            # Splits de gegevens om de waarde te isoleren
            parts = raw_data.split("(")
            value = parts[1].split("*")[0]  # Haal de waarde uit de ruwe gegevens
            power_watts = float(value) * 1000  # Zet kilowatts om naar watts

            # Voeg de meting toe aan de buffer
            buffer.append(power_watts)

            # Controleer of de 30 minuten voorbij zijn om een gemiddelde te berekenen en op te slaan
            if current_time - start_time >= save_interval:
                # Bereken het gemiddelde van de metingen
                average_power = sum(buffer) / len(buffer)

                # Formatteren met 2 of 3 cijfers voor watts of kilowatts
                if average_power >= 1000:
                    formatted_average_power = f'{average_power/1000:.3f} kW'
                else:
                    formatted_average_power = f'{average_power:.2f} W'

                # Voeg de dagelijkse opbrengst toe aan de totale opbrengst
                daily_total_energy += (average_power * save_interval) / (1000 * 3600)  # kWh

                # Formatteren van de totale opbrengst met 2 cijfers achter de komma
                formatted_total_energy = round(daily_total_energy, 2)

                # Aanmaken of openen van het Excel-werkboek
                if not os.path.isfile(excel_file):
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    sheet['A1'] = 'Datum'
                    sheet['B1'] = 'Tijd'
                    sheet['C1'] = 'Gemiddelde Vermogen'
                    sheet['D1'] = 'Totale Opbrengst (kWh)'

                else:
                    workbook = openpyxl.load_workbook(excel_file)
                    sheet = workbook.active

                # Voeg het gemiddelde en de dagelijkse totaalopbrengst toe aan het Excel-werkboek
                next_row = sheet.max_row + 1
                sheet.cell(row=next_row, column=1, value=current_date)
                sheet.cell(row=next_row, column=2, value=current_time_str)
                sheet.cell(row=next_row, column=3, value=formatted_average_power)
                sheet.cell(row=next_row, column=4, value=formatted_total_energy)

                # Opslaan van het Excel-werkboek
                workbook.save(excel_file)

                # Aanmaken of openen van het CSV-back-upbestand
                if not os.path.isfile(csv_file):
                    with open(csv_file, mode='w', newline='') as file:
                        writer = csv.writer(file, delimiter=',', quoting=csv.QUOTE_MINIMAL)
                        writer.writerow(['Datum', 'Tijd', 'Gemiddelde Vermogen', 'Totale Opbrengst (kWh)'])

                # Voeg het gemiddelde en de dagelijkse totaalopbrengst toe aan het CSV-back-upbestand
                with open(csv_file, mode='a', newline='') as file:
                    writer = csv.writer(file, delimiter=',', quoting=csv.QUOTE_MINIMAL)
                    writer.writerow([current_date, current_time_str, formatted_average_power, formatted_total_energy])

                # Reset de buffer en starttijd voor de volgende 30 minuten
                buffer = []
                start_time = current_time  # Update start_time

    except KeyboardInterrupt:
        break

# Sluit de seriÃ«le verbinding bij het afsluiten van het script
ser.close()

