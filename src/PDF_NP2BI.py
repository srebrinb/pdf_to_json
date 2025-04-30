import os
import csv
import argparse
from pdf_processor import PdfProcessor

def export_to_csv(output_directory, pdf_name, current_month, object_code, object_name, current_object_address, rows):
    # Създаване на директория за CSV файловете, ако не съществува
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Име на CSV файла
    csv_file_path = os.path.join(output_directory, f"{object_code}.csv")

    # Проверка дали файлът вече съществува
    file_exists = os.path.isfile(csv_file_path)

    # Записване на данните в CSV файла
    with open(csv_file_path, mode="a", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)

        # Добавяне на заглавията, ако файлът е нов
        if not file_exists:
            writer.writerow([
                "PDF Path", "Month", "Object Code", "Object Name", "Object Address",
                "Active Power", "Reactive Power", "Total Power", "Power Factor (cosφ)",
                "CO₂ Emissions (kg)", "Voltage", "Current", "Phase Angle"
            ])

        # Добавяне на редовете
        for row in rows:
            active_power = row[2]
            reactive_power = row[3]
            total_power = (active_power**2 + reactive_power**2)**0.5
            power_factor = active_power / total_power if total_power != 0 else 0
            co2_emissions = total_power * 0.3
            voltage = 230
            current = total_power / voltage if voltage != 0 else 0
            phase_angle = 0 if total_power == 0 else f"=DEGREES(ACOS({active_power}/{total_power}))"

            writer.writerow([
                pdf_name, current_month, object_code, object_name, current_object_address,
                active_power, reactive_power, total_power, power_factor,
                co2_emissions, voltage, current, phase_angle
            ])

def process_pdfs_in_directory(directory, output_directory):
    pdf_processor = PdfProcessor()
    for pdf_file in os.listdir(directory):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(directory, pdf_file)
            extracted_text = pdf_processor.extract_text(pdf_path)

            # Decode the text from windows-1251 and re-encode to UTF-8
            try:
                extracted_text = extracted_text.encode('latin1').decode('windows-1251')
            except UnicodeError as e:
                print(f"Error: Unable to decode text for {pdf_file}. Details: {e}")
                continue

            if extracted_text:
                # Парсиране на текста
                rows_by_object_code = {}
                current_object_name = None
                current_object_address = None
                current_month = None
                current_energy_sum = 0.0
                current_energy_sum2 = 0.0
                current_energy_sum3 = 0.0
                find_date_doc = "Основание: Електрическа енергия за месец"
                find_str_con1 = "Достъп високо напрежение"
                find_str_con2 = "Общо сума"
                find_str_con3 = "Надбавка за използвана реактивна енергия"

                for line in extracted_text.splitlines():
                    line = line.strip()
                    if line.startswith("Наименование на обекта:"):
                        current_object_name = line.replace("Наименование на обекта:", "").strip()
                        if current_object_name:
                            parts = current_object_name.rsplit(" ", 1)
                            if len(parts) == 2:
                                object_code = parts[1]
                                object_name = parts[0]
                            else:
                                object_code = current_object_name
                                object_name = ""
                    elif line.startswith("Адрес на обекта: "):
                        current_object_address = line.replace("Адрес на обекта: ", "").strip()
                        current_object_address = current_object_address.replace("Кодов номер:", "").strip()
                    elif line.startswith(find_date_doc):
                        current_month = line.replace(find_date_doc, "").strip()

                    if line.startswith(find_str_con1):
                        line = line.replace(find_str_con1, "").strip()
                        if "кВтч" in line:
                            energy_part = line.split("кВтч")[0].strip()
                            energy_part = energy_part.replace(" ", "")
                            energy_part = energy_part[::-1]
                            energy_part = energy_part.lstrip("0")
                            try:
                                current_energy_sum += float(energy_part)
                            except ValueError:
                                pass
                    if line.startswith(find_str_con2):
                        line = line.replace(find_str_con2, "").strip()
                        line = line.replace(",", "").strip()
                        energy_part = line
                        energy_part = energy_part.replace(" ", "")
                        energy_part = energy_part[::-1]
                        energy_part = energy_part.lstrip("0")
                        try:
                            current_energy_sum2 += float(energy_part)
                        except ValueError:
                            pass
                    if line.startswith(find_str_con3):
                        line = line.replace(find_str_con3, "").strip()
                        if "кВАрч" in line:
                            energy_part = line.split("кВАрч")[0].strip()
                            energy_part = energy_part.replace(" ", "")
                            energy_part = energy_part[::-1]
                            energy_part = energy_part.lstrip("0")
                        try:
                            current_energy_sum3 += float(energy_part)
                        except ValueError:
                            pass

                    if current_energy_sum2:
                        if object_code not in rows_by_object_code:
                            rows_by_object_code[object_code] = []
                        rows_by_object_code[object_code].append([
                            pdf_file, current_month, current_energy_sum, current_energy_sum3, current_energy_sum2
                        ])
                        current_energy_sum = 0.0
                        current_energy_sum3 = 0.0
                        current_energy_sum2 = 0.0

                # Експорт на данните в CSV
                for object_code, rows in rows_by_object_code.items():
                    export_to_csv(output_directory, pdf_file, current_month, object_code, current_object_name, current_object_address, rows)

def main():
    parser = argparse.ArgumentParser(description="Process multiple PDF files and export data to CSV files for Power BI.")
    parser.add_argument("output_directory", help="Path to the directory where CSV files will be saved.")
    parser.add_argument("pdf_directory", help="Path to the directory containing PDF files.")
    args = parser.parse_args()

    if not os.path.isdir(args.pdf_directory):
        print("The specified directory does not exist.")
        return

    process_pdfs_in_directory(args.pdf_directory, args.output_directory)

if __name__ == "__main__":
    main()