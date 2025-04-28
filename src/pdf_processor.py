
class PdfProcessor:
    def extract_text(self, pdf_path):
        import PyPDF2
        
        text = ""
        with open(pdf_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() + "\n"
        
        return text.strip()