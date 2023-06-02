import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QFileDialog, QMessageBox

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Format Converter")
        self.setGeometry(100, 100, 300, 300)

        self.input_files = []
        self.output_folder = ""

        self.label = QLabel(self)
        self.label.setGeometry(50, 20, 300, 30)
        self.label.setText("Eingabedateien: Nicht ausgewählt")

        self.input_button = QPushButton("Eingabedateien auswählen", self)
        self.input_button.setGeometry(50, 70, 200, 30)
        self.input_button.clicked.connect(self.select_input_files)

        self.label2 = QLabel(self)
        self.label2.setGeometry(50, 120, 300, 30)
        self.label2.setText("Ausgabeordner: Nicht ausgewählt")

        self.output_button = QPushButton("Ausgabeordner auswählen", self)
        self.output_button.setGeometry(50, 170, 200, 30)
        self.output_button.clicked.connect(self.select_output_folder)

        self.convert_button = QPushButton("Konvertieren", self)
        self.convert_button.setGeometry(50, 220, 200, 30)
        self.convert_button.clicked.connect(self.convert_files)

        self.status_label = QLabel(self)
        self.status_label.setGeometry(50, 270, 300, 30)
        self.status_label.setText("")

    def select_input_files(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        files, _ = QFileDialog.getOpenFileNames(self, "Eingabedateien auswählen", "", "CSV-Dateien (*.csv);;XLSX-Dateien (*.xlsx)", options=options)
        self.input_files = files
        self.label.setText(f"Eingabedateien: {len(self.input_files)} ausgewählt")

    def select_output_folder(self):
        self.output_folder = QFileDialog.getExistingDirectory(self, "Ausgabeordner auswählen")
        self.label2.setText(f"Ausgabeordner: {self.output_folder}")

    def convert_files(self):
        if not self.input_files and not self.output_folder:
            QMessageBox.warning(self, "Fehler", "Bitte wählen Sie sowohl Eingabedateien als auch einen Ausgabeordner aus.", QMessageBox.Ok)
            return
        if not self.input_files:
            QMessageBox.warning(self, "Fehler", "Es wurden keine Eingabedateien ausgewählt.", QMessageBox.Ok)
            return
        if not self.output_folder:
            QMessageBox.warning(self, "Fehler", "Es wurde kein Ausgabeordner ausgewählt.", QMessageBox.Ok)
            return

        for file_path in self.input_files:
            file_name, file_ext = os.path.splitext(file_path)
            output_file = os.path.join(self.output_folder, os.path.basename(file_name))

            if file_ext.lower() == '.csv':
                output_file += '.xlsx'
                df = pd.read_csv(file_path, delimiter=';', quotechar='"', decimal=',', encoding='utf-8-sig')
                df.to_excel(output_file, index=False)
            elif file_ext.lower() == '.xlsx':
                output_file += '.csv'
                df = pd.read_excel(file_path)
                df.to_csv(output_file, index=False, sep=';', quotechar='"', decimal=',', float_format='%.15g', encoding='utf-8-sig')

        self.label.setText("Eingabedateien: Nicht ausgewählt")
        self.label2.setText("Ausgabeordner: Nicht ausgewählt")

        self.status_label.setText("Umwandlung erfolgreich abgeschlossen.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
