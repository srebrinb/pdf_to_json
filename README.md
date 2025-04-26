# PDF to Excel Converter

This project provides a simple way to extract text from PDF files and format it into an Excel spreadsheet. It consists of a Python application that utilizes the `PyPDF2` library for PDF processing and the `openpyxl` library for writing to Excel files.

## Project Structure

```
pdf-to-excel
├── src
│   ├── main.py          # Entry point of the application
│   ├── pdf_processor.py  # Contains PdfProcessor class for PDF text extraction
│   ├── excel_writer.py    # Contains ExcelWriter class for writing to Excel
│   └── utils
│       └── __init__.py   # Utility functions or constants
├── requirements.txt      # Project dependencies
└── README.md             # Project documentation
```

## Installation

To set up the project, you need to install the required dependencies. You can do this by running:

```
pip install -r requirements.txt
```

## Usage

1. Place your PDF file in a known directory.
2. Update the `main.py` file with the path to your PDF file and the desired output Excel file path.
3. Run the application:

```
python src/main.py
```

This will extract the text from the specified PDF file and write it to the specified Excel file.

## Dependencies

- PyPDF2
- openpyxl

Make sure to check the `requirements.txt` file for the exact versions of the libraries used in this project.