import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QSpinBox, QMessageBox


class ExcelProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Excel Processor")
        self.setGeometry(200, 200, 600, 600)

        layout = QVBoxLayout()

        # region Вибір файлу
        self.label = QLabel("Choose an Excel file", self)
        layout.addWidget(self.label)

        self.btn_select = QPushButton("Select File", self)
        self.btn_select.clicked.connect(self.load_file)
        layout.addWidget(self.btn_select)
        # endregion

        # region Характеристики
        self.label_col = QLabel("Select Start Column for Characteristics:", self)
        layout.addWidget(self.label_col)

        self.spin_col = QSpinBox(self)
        self.spin_col.setMinimum(1)
        self.spin_col.setValue(43)
        layout.addWidget(self.spin_col)
        # endregion

        # region Ім'я продукту
        self.label_name_col = QLabel("Select Product Name Column:", self)
        layout.addWidget(self.label_name_col)

        self.spin_name_col = QSpinBox(self)
        self.spin_name_col.setMinimum(1)
        self.spin_name_col.setValue(4)
        layout.addWidget(self.spin_name_col)
        # endregion

        # region Артикул
        self.label_sku_col = QLabel("Select Product SKU Column:", self)
        layout.addWidget(self.label_sku_col)

        self.spin_sku_col = QSpinBox(self)
        self.spin_sku_col.setMinimum(1)
        self.spin_sku_col.setValue(2)  # За замовчуванням колонка 2
        layout.addWidget(self.spin_sku_col)
        # endregion

        # region Старий опис
        self.label_old_desc_col = QLabel("Select Old Description Column:", self)
        layout.addWidget(self.label_old_desc_col)

        self.spin_old_desc_col = QSpinBox(self)
        self.spin_old_desc_col.setMinimum(1)
        self.spin_old_desc_col.setValue(11)  # За замовчуванням колонка 11
        layout.addWidget(self.spin_old_desc_col)
        # endregion

        # region Кнопка збереження
        self.btn_process = QPushButton("Process & Save", self)
        self.btn_process.clicked.connect(self.process_excel)
        self.btn_process.setEnabled(False)
        layout.addWidget(self.btn_process)

        self.setLayout(layout)
        self.file_path = ""
        # endregion

    def load_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "",
                                                   "Excel Files (*.xls *.xlsx *.xlsm);;All Files (*)", options=options)

        if file_name:
            self.file_path = file_name
            self.label.setText(f"Selected: {file_name}")
            self.btn_process.setEnabled(True)

    def process_excel(self):
        if not self.file_path:
            return

        try:
            df = pd.read_excel(self.file_path, sheet_name=0)
            char_col_start = self.spin_col.value() - 1
            name_col = self.spin_name_col.value() - 1
            sku_col = self.spin_sku_col.value() - 1
            old_desc_col = self.spin_old_desc_col.value() - 1

            if any(col >= len(df.columns) for col in [name_col, char_col_start, sku_col, old_desc_col]):
                QMessageBox.critical(self, "Column Error", "Selected columns exceed file range.")
                return

            df["Characteristics"] = df.iloc[:, char_col_start:].apply(
                lambda row: "; ".join(f"{col.replace('@', ' ')} — {val}" for col, val in row.items() if pd.notna(val)), axis=1
            )

            result_df = df.iloc[:, [name_col, sku_col, old_desc_col]].copy()
            result_df.columns = ["Name", "Articule", "Old_description"]
            result_df["Characteristics"] = df["Characteristics"]

            script_dir = os.path.dirname(os.path.abspath(__file__))
            old_file_name = os.path.basename(self.file_path)
            clean_name = os.path.splitext(old_file_name)[0]
            save_path = os.path.join(script_dir, f"характеристики_{clean_name}.xlsx")

            result_df.to_excel(save_path, index=False)
            self.label.setText(f"Saved: {save_path}")
        except PermissionError:
            QMessageBox.warning(self, "File Error", "Close the Excel file before processing.")
        except Exception as e:
            QMessageBox.critical(self, "Processing Error", f"An error occurred: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessor()
    window.show()
    sys.exit(app.exec_())
