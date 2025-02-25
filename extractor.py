import pdfplumber
import pandas as pd
import re

class PDFExtractor:
    def __init__(self, file_path):
        self.file_path = file_path

    def extract_data(self):
        transactions = []
        with pdfplumber.open(self.file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    transactions.extend(self.process_text(text))
        return transactions

    def process_text(self, text):
        lines = text.split("\n")
        extracted_data = []

        for line in lines:
            # Expresión regular para detectar líneas con movimientos (Fecha + Descripción + Monto)
            match = re.match(r"(\d{1,2}[\/\-]\w{3,}|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\s+(.*?)\s+([\d,]+\.\d{2})", line)
            if match:
                date, description, amount = match.groups()
                extracted_data.append([date, description.strip(), amount.replace(",", "")])

        return extracted_data

    def save_to_excel(self, data, output_path="output.xlsx"):
        df = pd.DataFrame(data, columns=["Fecha", "Descripción", "Monto"])
        df.to_excel(output_path, index=False)
        print(f"Archivo guardado en: {output_path}")

# Prueba rápida
if __name__ == "__main__":
    extractor = PDFExtractor("estado_de_cuenta.pdf")
    data = extractor.extract_data()
    extractor.save_to_excel(data)
