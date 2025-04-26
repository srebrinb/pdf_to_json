class ExcelWriter:
    def write_to_excel(self, data, excel_path):
        import pandas as pd
        # Ensure data is in the correct format
        if isinstance(data, str):
            # Convert plain text into a list of dictionaries
            data = [{"Line": line} for line in data.split("\n") if line.strip()]

        # Create a DataFrame from the extracted data
        df = pd.DataFrame(data)

        # Write the DataFrame to an Excel file
        df.to_excel(excel_path, index=False)