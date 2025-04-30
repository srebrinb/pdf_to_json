import os
import argparse
from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from datetime import datetime

def process_pdfs(directory):
    """Обработва всички PDF файлове и съхранява данните в речник."""
    data_by_object_code = {}

    for pdf_file in os.listdir(directory):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(directory, pdf_file)

            # Симулиране на извличане на текст от PDF (заменете с реална логика)
            extracted_text = simulate_pdf_extraction(pdf_path)

            if extracted_text:
                current_object_name = None
                current_object_address = None
                current_month = None
                current_energy_sum = 0.0
                current_energy_sum3 = 0.0

                for line in extracted_text.splitlines():
                    line = line.strip()
                    if line.startswith("Наименование на обекта:"):
                        current_object_name = line.replace("Наименование на обекта:", "").strip()
                        object_code = current_object_name.split()[-1]  # Примерен код
                    elif line.startswith("Адрес на обекта:"):
                        current_object_address = line.replace("Адрес на обекта:", "").strip()
                    elif line.startswith("Основание: Електрическа енергия за месец"):
                        current_month = line.replace("Основание: Електрическа енергия за месец", "").strip()
                        current_month = datetime.strptime(current_month, "%B %Y")  # Преобразуване в дата
                    elif "кВтч" in line:
                        current_energy_sum = float(line.split()[0])  # Примерно извличане
                    elif "кВАрч" in line:
                        current_energy_sum3 = float(line.split()[0])  # Примерно извличане

                if object_code not in data_by_object_code:
                    data_by_object_code[object_code] = {
                        "object_name": current_object_name,
                        "object_address": current_object_address,
                        "rows": []
                    }
                data_by_object_code[object_code]["rows"].append([
                    pdf_file, current_month, current_energy_sum, current_energy_sum3
                ])

    return data_by_object_code

def generate_excel(data_by_object_code, excel_path):
    """Генерира Excel файл с таблици и графики за всеки обект."""
    workbook = Workbook()

    for object_code, data in data_by_object_code.items():
        sheet = workbook.create_sheet(title=object_code[:31])  # Ограничение на имената на sheet-овете до 31 символа

        # Добавяне на антетка
        sheet.append(["Код на обекта:", object_code])
        sheet.append(["Име на обекта:", data["object_name"]])
        sheet.append(["Адрес на обекта:", data["object_address"]])
        sheet.append([])  # Празен ред за разделяне

        # Добавяне на заглавия на таблицата
        sheet.append(["PDF Path", "За месец (Дата)", "Активна мощност (W)", "Реактивна мощност (W)"])

        # Добавяне на редовете
        for row in data["rows"]:
            sheet.append([row[0], row[1].strftime("%Y-%m"), row[2], row[3]])

        # Създаване на графика
        chart = LineChart()
        chart.title = "Активна и Реактивна мощност по месеци"
        chart.style = 13
        chart.x_axis.title = "Месец"
        chart.y_axis.title = "Мощност (W)"

        # Данни за графиката
        data = Reference(sheet, min_col=3, max_col=4, min_row=6, max_row=sheet.max_row)
        categories = Reference(sheet, min_col=2, min_row=6, max_row=sheet.max_row)  # Колона B за "Месец"
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)

        # Поставяне на графиката под таблицата
        sheet.add_chart(chart, f"A{sheet.max_row + 2}")

    # Премахване на празния sheet, ако съществува
    if "Sheet" in workbook.sheetnames:
        del workbook["Sheet"]

    workbook.save(excel_path)

def simulate_pdf_extraction(pdf_path):
    """Симулира извличане на текст от PDF (заменете с реална логика)."""
    return f"""
    Наименование на обекта: Обект 1
    Адрес на обекта: ул. Примерна 1
    Основание: Електрическа енергия за месец Януари 2025
    100 кВтч
    50 кВАрч
    """

def main():
    parser = argparse.ArgumentParser(description="Process multiple PDF files and save data to an Excel file.")
    parser.add_argument("excel_path", help="Path to the output Excel file.")
    parser.add_argument("pdf_directory", help="Path to the directory containing PDF files.")
    args = parser.parse_args()

    if not os.path.isdir(args.pdf_directory):
        print("The specified directory does not exist.")
        return

    # Обработка на PDF файловете
    data_by_object_code = process_pdfs(args.pdf_directory)

    # Генериране на Excel файла
    generate_excel(data_by_object_code, args.excel_path)

if __name__ == "__main__":
    main()