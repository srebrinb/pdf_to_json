import argparse
import json
from datetime import datetime
import os   
import InvToJson
from pdf_processor import PdfProcessor


def main():
    parser = argparse.ArgumentParser(description="Process multiple PDF files and save data to a JSON file.")
    parser.add_argument("json_path", nargs="?", default="test.json", help="Path to the output JSON file.")
    parser.add_argument("pdf_directory", nargs="?", default="StefanP", help="Path to the directory containing PDF files.")
    args = parser.parse_args()

    if not os.path.isdir(args.pdf_directory):
        print("The specified directory does not exist.")
        return
    directory=args.pdf_directory
    output_file = args.json_path
    factura_data = {"invoices": []}
    for pdf_file in os.listdir(directory):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(directory, pdf_file)

            # Симулиране на извличане на текст от PDF (заменете с реална логика)
            #extracted_text = simulate_pdf_extraction(pdf_path)
            pdf_processor = PdfProcessor()
            extracted_text = pdf_processor.extract_text(pdf_path)      
            extracted_text = extracted_text.encode('latin1').decode('windows-1251')
            # Извикване на функцията
            factura_data = InvToJson.parse_factura(extracted_text, factura_data, filename=pdf_file)
    InvToJson.save_to_json(factura_data, output_file)
if __name__ == "__main__":
    main()