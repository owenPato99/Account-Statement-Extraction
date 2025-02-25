from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extractor de Estados de Cuenta")
        self.setGeometry(200, 200, 450, 350)
        self.setFixedSize(450, 350)  
        self.setWindowIcon(QIcon("heza_logo.jpg"))

        # Estilos CSS para la interfaz
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QLabel {
                font-size: 16px;
                font-weight: bold;
                color: #333333;
                margin-bottom: 20px;
            }
            QPushButton {
                background-color: #01b8a5;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 5px;
                border: none;
                margin: 5px;
            }
            QPushButton:hover {
                background-color: #01b8a5;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #01b8a5;
            }
        """)

        # Fuente personalizada
        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(12)

        # Etiqueta principal
        self.label = QLabel("Carga un estado de cuenta en PDF", self)
        self.label.setFont(font)
        self.label.setAlignment(Qt.AlignCenter)

        # Botón para seleccionar archivo
        self.btn_select_file = QPushButton("Seleccionar PDF", self)
        self.btn_select_file.setFont(font)
        self.btn_select_file.clicked.connect(self.load_pdf)

        # Botón para exportar a Excel
        self.btn_export_excel = QPushButton("Exportar a Excel", self)
        self.btn_export_excel.setFont(font)
        self.btn_export_excel.setEnabled(False)
        self.btn_export_excel.clicked.connect(self.export_to_excel)

        # Diseño de la ventana
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.btn_export_excel)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.file_path = None

        # Habilitar arrastrar y soltar
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith('.pdf'):
                self.file_path = file_path
                self.label.setText(f"Archivo seleccionado:\n{file_path}")
                self.btn_export_excel.setEnabled(True)
                break

    def load_pdf(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar PDF", "", "Archivos PDF (*.pdf);;Todos los archivos (*)", options=options)
        if file_path:
            self.file_path = file_path
            self.label.setText(f"Archivo seleccionado:\n{file_path}")
            self.btn_export_excel.setEnabled(True)

    def export_to_excel(self):
        if self.file_path:
            from extractor import PDFExtractor
            extractor = PDFExtractor(self.file_path)
            data = extractor.extract_data()

            # Nombre por defecto del archivo
            default_file_name = "estado_de_cuenta.xlsx"
            
            # Configurar el diálogo de guardado con el nombre por defecto
            save_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Guardar como", 
                default_file_name,  # Nombre por defecto
                "Archivos de Excel (*.xlsx);;Todos los archivos (*)"
            )
            
            if save_path:
                try:
                    extractor.save_to_excel(data, save_path)
                    self.label.setText(f"Datos guardados en:\n{save_path}")
                except PermissionError:
                    self.label.setText("Error: No se pudo guardar el archivo. Verifica los permisos o cierra el archivo si está abierto.")
                except Exception as e:
                    self.label.setText(f"Error inesperado: {str(e)}")

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())