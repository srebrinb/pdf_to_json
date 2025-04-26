import pdfplumber
class PdfProcessor:
    def extract_text(self, pdf_path):
        import PyPDF2
        
        text = ""
        with open(pdf_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() + "\n"
        
        return text.strip()
    def extract_text_to_file(self, pdf_path, output_txt_path):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                with open(output_txt_path, "w", encoding="utf-8") as output_file:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:  # Проверка дали страницата съдържа текст
                            # Опит за декодиране на текста
                            try:
                                text = text.encode('latin1').decode('windows-1251')
                            except (UnicodeEncodeError, UnicodeDecodeError):
                                pass  # Ако декодирането не успее, използваме оригиналния текст
                            # Запис на текста във файла
                            output_file.write(text + "\n")
            print(f"Text successfully extracted and saved to {output_txt_path}")
        except Exception as e:
            print(f"An error occurred: {e}")