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
        current_transaction = []

        for line in lines:
            parts = line.split()
            
            # Detectamos si una línea tiene una fecha válida al inicio
            if re.match(r"\d{1,2}[/-]\w{3,}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", parts[0]):
                # Si ya hay datos acumulados en `current_transaction`, los guardamos
                if current_transaction:
                    extracted_data.append(current_transaction)
                # Empezamos una nueva transacción
                current_transaction = [parts[0]]  # Fecha
                description = " ".join(parts[1:-3])  # Descripción en medio
                deposits = parts[-3] if parts[-3].replace(",", "").replace(".", "").isdigit() else "0"
                withdrawals = parts[-2] if parts[-2].replace(",", "").replace(".", "").isdigit() else "0"
                balance = parts[-1] if parts[-1].replace(",", "").replace(".", "").isdigit() else "0"
                current_transaction.extend([description, deposits, withdrawals, balance])
            else:
                # Si la línea no tiene fecha, puede ser parte de la descripción de la transacción anterior
                if current_transaction:
                    current_transaction[1] += " " + " ".join(parts)

        # Guardamos la última transacción detectada
        if current_transaction:
            extracted_data.append(current_transaction)

        return extracted_data

    def save_to_excel(self, data, output_path="output.xlsx"):
        df = pd.DataFrame(data, columns=["Fecha", "Descripción", "Depósitos", "Retiros", "Saldo"])
        df.to_excel(output_path, index=False)
        print(f"Archivo guardado en: {output_path}")

if __name__ == "__main__":
    extractor = PDFExtractor("estado_de_cuenta.pdf")
    data = extractor.extract_data()
    extractor.save_to_excel(data)
