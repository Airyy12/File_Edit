import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QLabel, QLineEdit, QMessageBox
)

class ExcelMerger(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📊 Gabung File Excel")
        self.setGeometry(200, 200, 400, 200)

        layout = QVBoxLayout()

        self.label = QLabel("Masukkan Nama Bulan:")
        layout.addWidget(self.label)

        self.bulan_input = QLineEdit()
        layout.addWidget(self.bulan_input)

        self.btn_upload = QPushButton("📁 Pilih File Excel (.xlsx/.xls)")
        self.btn_upload.clicked.connect(self.upload_files)
        layout.addWidget(self.btn_upload)

        self.btn_proses = QPushButton("✅ Gabungkan File")
        self.btn_proses.clicked.connect(self.gabungkan)
        layout.addWidget(self.btn_proses)

        self.setLayout(layout)
        self.files = []

    def upload_files(self):
        self.files, _ = QFileDialog.getOpenFileNames(self, "Pilih File Excel", "", "Excel Files (*.xlsx *.xls)")
        if self.files:
            QMessageBox.information(self, "Info", f"{len(self.files)} file dipilih.")

    def gabungkan(self):
        bulan = self.bulan_input.text().strip()
        if not self.files:
            QMessageBox.warning(self, "Warning", "Silakan pilih file terlebih dahulu.")
            return
        if not bulan:
            QMessageBox.warning(self, "Warning", "Masukkan nama bulan terlebih dahulu.")
            return

        try:
            all_data = []
            for file in self.files:
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    all_data.append(df)

            combined = pd.concat(all_data, ignore_index=True)

            output_filename = f"TotalGabungan_{bulan}.xlsx"
            combined.to_excel(output_filename, index=False)

            QMessageBox.information(self, "Sukses", f"File berhasil disimpan: {output_filename}")
            os.startfile(output_filename)

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelMerger()
    window.show()
    sys.exit(app.exec_())
