import os
import argparse
from openpyxl import Workbook, load_workbook
from pdf_processor import PdfProcessor

def parse_text_to_excel(extracted_text, excel_path):
    rows = []
    current_object_name = None
    current_object_address = None
    current_month = None
    current_energy_sum = 0.0
    current_energy_sum2 = 0.0
    find_str_con1 = "Достъп високо напрежение"
    find_str_con2 = "Достъп средно/ниско напрежение (предоставена мощност по брой дни)"

    for line in extracted_text.splitlines():
        line = line.strip()

        # Проверка за "Наименование на обекта:"
        if line.startswith("Наименование на обекта:"):
            current_object_name = line.replace("Наименование на обекта:", "").strip()
            # Разделяне на current_object_name в две части
            if current_object_name:
                parts = current_object_name.rsplit(" ", 1)  # Разделяне по последния интервал
                if len(parts) == 2:
                    object_code = parts[1]  # Последният стринг
                    object_name = parts[0]  # Началото
                else:
                    object_code = current_object_name  # Ако няма интервал, цялото име е код
                    object_name = ""  # Няма начало
        # Проверка за "Адрес на обекта: "
        elif line.startswith("Адрес на обекта: "):
            current_object_address = line.replace("Адрес на обекта: ", "").strip()
            current_object_address = current_object_address.replace("Кодов номер:", "").strip()
        # Проверка за "За месец:"
        elif line.startswith("За месец:"):
            current_month = line.replace("За месец:", "").strip()

        # Проверка за "Електрическа енергия"
        elif line.startswith(find_str_con1):
            line = line.replace(find_str_con1, "").strip()
            if "кВтч" in line:
                energy_part = line.split("кВтч")[0].strip()
                energy_part = energy_part.replace(" ", "")  # Премахване на интервалите
                energy_part = energy_part[::-1]  # Обръщане на числото
                energy_part = energy_part.lstrip("0")
                try:
                    current_energy_sum += float(energy_part)  # Добавяне към сумата
                except ValueError:
                    pass  # Игнориране на грешки при преобразуване
        # Проверка за "Достъп средно/ниско напрежение"
        elif line.startswith(find_str_con2):
            line = line.replace(find_str_con2, "").strip()
            if "кВтч" in line:
                energy_part = line.split("кВтч")[0].strip()
                energy_part = energy_part.replace(" ", "")  # Премахване на интервалите
                energy_part = energy_part[::-1]  # Обръщане на числото
                energy_part = energy_part.lstrip("0")
                try:
                    current_energy_sum2 += float(energy_part)  # Добавяне към сумата
                except ValueError:
                    pass  # Игнориране на грешки при преобразуване
        # Ако има текущо "Наименование на обекта", "За месец" и "Електрическа енергия", добавяме ред
        if current_object_name and current_month and current_energy_sum:
            rows.append([object_code, object_name, current_object_address, current_month, current_energy_sum, current_energy_sum2])
            current_energy_sum = 0.0  # Нулираме сумата, за да избегнем дублиране
            current_energy_sum2 = 0.0  # Нулираме сумата, за да избегнем дублиране
    # Проверка дали файлът съществува
    file_exists = os.path.exists(excel_path)

    if file_exists:
        workbook = load_workbook(excel_path)
        sheet = workbook.active
    else:
        workbook = Workbook()
        sheet = workbook.active
        # Добавяне на заглавия, ако файлът не съществува
        sheet.append(["Код на обекта", "Име на обекта", "Адрес на обекта", "За месец", "Количество (Електрическа енергия)", find_str_con2])

    # Добавяне на редовете
    for row in rows:
        sheet.append(row)

    # Запис на Excel файла
    workbook.save(excel_path)

def main():
    # Добавяне на argparse за подаване на параметри от командния ред
    parser = argparse.ArgumentParser(description="Extract data from a PDF and save it to an Excel file.")
    parser.add_argument("pdf_path", nargs="?", default="C:\\work\\PDF_Extr\\1.pdf", help="Path to the input PDF file (default: C:\\work\\PDF_Extr\\1.pdf)")
    parser.add_argument("excel_path", nargs="?", default="C:\\work\\PDF_Extr\\1.xlsx", help="Path to the output Excel file (default: C:\\work\\PDF_Extr\\1.xlsx)")

    try:
        args = parser.parse_args()
        pdf_path = args.pdf_path
        excel_path = args.excel_path
    except Exception as e:
        print(f"Error: Unable to parse arguments. Details: {e}")
        return

    if not os.path.exists(pdf_path):
        print("The specified PDF file does not exist.")
        return

    pdf_processor = PdfProcessor()
    extracted_text = pdf_processor.extract_text(pdf_path)

    # Decode the text from windows-1251 and re-encode to UTF-8
    try:
        extracted_text = extracted_text.encode('latin1').decode('windows-1251')
    except UnicodeError as e:
        print(f"Error: Unable to decode text. Details: {e}")
        return

    if extracted_text:
        # Парсиране на текста и запис в Excel
        parse_text_to_excel(extracted_text, excel_path)
        print(f"Data has been written to {excel_path}")
    else:
        print("No text extracted from the PDF.")

if __name__ == "__main__":
    main()