from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget
from extractor import PDFExtractor

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extractor de Estados de Cuenta")
        self.setGeometry(200, 200, 400, 300)

        self.label = QLabel("Carga un estado de cuenta en PDF", self)
        self.label.setStyleSheet("font-size: 14px;")
        
        self.btn_select_file = QPushButton("Seleccionar PDF", self)
        self.btn_select_file.clicked.connect(self.load_pdf)

        self.btn_export_excel = QPushButton("Exportar a Excel", self)
        self.btn_export_excel.setEnabled(False)  # Se activa después de cargar un archivo
        self.btn_export_excel.clicked.connect(self.export_to_excel)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.btn_export_excel)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.file_path = None

    def load_pdf(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar PDF", "", "Archivos PDF (*.pdf);;Todos los archivos (*)", options=options)
        if file_path:
            self.file_path = file_path
            self.label.setText(f"Archivo seleccionado:\n{file_path}")
            self.btn_export_excel.setEnabled(True)

    def export_to_excel(self):
        if self.file_path:
            extractor = PDFExtractor(self.file_path)
            data = extractor.extract_data()

            # Pedir al usuario una ubicación para guardar el archivo
            options = QFileDialog.Options()
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Guardar como",
                "estado_de_cuenta.xlsx",
                "Archivos de Excel (*.xlsx);;Todos los archivos (*)",
                options=options
            )

            if save_path:  # Si el usuario elige una ubicación
                extractor.save_to_excel(data, save_path)
                self.label.setText(f"Datos guardados en:\n{save_path}")


# Prueba rápida de la GUI
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
