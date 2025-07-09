
import sys
import os
import traceback
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QLineEdit, QLabel, QTextEdit, QSpinBox
)
from PyQt5.QtCore import Qt
from refactor_expense_sorter import main_processing_function

class ExpenseSorterGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Expense Wizard")
        self.setGeometry(100, 100, 700, 350)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 📂 Input folder
        folder_layout = QHBoxLayout()
        self.folder_input = QLineEdit()
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self.browse_input_folder)
        folder_layout.addWidget(QLabel("Input Folder:"))
        folder_layout.addWidget(self.folder_input)
        folder_layout.addWidget(browse_button)
        layout.addLayout(folder_layout)

        # 💾 Output folder
        output_layout = QHBoxLayout()
        self.output_input = QLineEdit()
        output_browse_button = QPushButton("Browse")
        output_browse_button.clicked.connect(self.browse_output_file)
        output_layout.addWidget(QLabel("Save File As:"))
        output_layout.addWidget(self.output_input)
        output_layout.addWidget(output_browse_button)
        layout.addLayout(output_layout)

        # 📅 Date selectors
        date_layout = QHBoxLayout()
        self.year_input = QSpinBox()
        self.year_input.setRange(2000, 2100)
        self.year_input.setValue(2025)

        self.month_input = QSpinBox()
        self.month_input.setRange(1, 12)
        self.month_input.setValue(5)

        date_layout.addWidget(QLabel("Year:"))
        date_layout.addWidget(self.year_input)
        date_layout.addWidget(QLabel("Month:"))
        date_layout.addWidget(self.month_input)
        layout.addLayout(date_layout)

        # ▶️ Run button
        run_button = QPushButton("Generate Report")
        run_button.clicked.connect(self.run_script)
        layout.addWidget(run_button)

        # 🧾 Output log
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)

        self.setLayout(layout)

    def browse_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Input Folder")
        if folder:
            self.folder_input.setText(folder)

    def browse_output_file(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Report As",
            "",
            "Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)"
     )
        if file_path:
            self.output_input.setText(file_path)




    def run_script(self):
        folder = self.folder_input.text()
        out_folder = self.output_input.text()
        year = self.year_input.value()
        month = self.month_input.value()

        self.log_output.clear()

        if not folder:
            self.log_output.append("❌ Input folder is required.")
            return

        if not out_folder:
          self.log_output.append("❌ Output file path is required.")
          return


        self.log_output.append(f"📂 Input: {folder}")
        self.log_output.append(f"💾 Output: {out_folder}")
        self.log_output.append(f"📅 Year: {year}, Month: {month}")
        self.log_output.append("⚙️ Running script...")

        try:
            path = main_processing_function(folder, year, month, out_folder)
            self.log_output.append(f"✅ Report saved to: {path}")
        except Exception as e:
            self.log_output.append(f"❌ Error: {str(e)}")
            self.log_output.append(traceback.format_exc())

def main():
    app = QApplication(sys.argv)
    window = ExpenseSorterGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
