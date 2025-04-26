import os
import csv
from pdf_processor import PdfProcessor

def parse_text_to_csv(extracted_text, csv_path):
    rows = []
    current_object_name = None
    current_month = None

    for line in extracted_text.splitlines():
        line = line.strip()

        # Проверка за "Наименование на обекта:"
        if line.startswith("Наименование на обекта:"):
            current_object_name = line.replace("Наименование на обекта:", "").strip()

        # Проверка за "За месец:"
        elif line.startswith("За месец:"):
            current_month = line.replace("За месец:", "").strip()

        # Ако има текущо "Наименование на обекта" и "За месец", добавяме ред
        elif current_object_name and current_month and line:
            rows.append([current_object_name, current_month, line])

    # Запис в CSV файл с TAB разделител
    with open(csv_path, "w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file, delimiter="\t")
        writer.writerow(["Наименование на обекта", "За месец", "Данни"])  # Заглавия на колоните
        writer.writerows(rows)

def main():
    pdf_path = "C:\\work\\PDF_Extr\\1.pdf"  # input("Enter the path to the PDF file: ")
    csv_path = "C:\\work\\PDF_Extr\\1.csv"  # input("Enter the path to save the CSV file: ")

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
        # Парсиране на текста и запис в CSV
        parse_text_to_csv(extracted_text, csv_path)
        print(f"Data has been written to {csv_path}")
    else:
        print("No text extracted from the PDF.")

if __name__ == "__main__":
    main()