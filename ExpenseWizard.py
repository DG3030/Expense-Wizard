import sys
import os
import json
import traceback
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QLineEdit, QLabel, QTextEdit, QSpinBox, QDialog, QComboBox, QDateEdit, QMainWindow, QAction, QWidgetAction
)

from PyQt5.QtCore import Qt, QTimer, QPoint, QDate
from PyQt5.QtGui import QFont, QIcon

from expense_sorter import main_processing_function

class GuidedStepOverlay(QDialog):
    def __init__(self, target_widget, message, parent=None):
        super().__init__(parent)
        self.setFixedWidth(300)
        self.setFixedHeight(150)
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("background-color: rgba(30, 30, 30, 250); color: white; border-radius: 8px;")
        self.message = message
        self.target_widget = target_widget

        self.index = 0
        self.current_text = ""

        self.label = QLabel("")
        self.label.setWordWrap(True)
        self.label.setFont(QFont("Courier New", 10))
        self.label.setStyleSheet("padding: 8px;")

        self.next_button = QPushButton("Next")
        self.next_button.setEnabled(False)
        self.next_button.clicked.connect(self.close)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.next_button)
        self.setLayout(layout)

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_text)
        self.timer.start(25)

    def update_text(self):
        if self.index < len(self.message):
            self.current_text += self.message[self.index]
            self.label.setText(self.current_text)
            self.index += 1
        else:
            self.timer.stop()
            self.next_button.setEnabled(True)

    def showEvent(self, event):
        if self.target_widget:
            pos = self.target_widget.mapToGlobal(QPoint(0, self.target_widget.height()))
            self.move(pos + QPoint(10, 10))
        super().showEvent(event)

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("How to Use Expense Wizard")
        self.setMinimumWidth(400)

        help_text = (
    "<b>🧙‍♂️ Welcome to Expense Wizard</b><br><br>"
    "This tool helps you organize and summarize your credit card transactions across any custom date range.<br><br>"

    "<b>1. 📁 Select your Input Folder:</b><br>"
    "Choose a folder containing one or more Discover card <code>.xlsx</code> statement files.<br><br>"

    "<b>2. 💾 Choose your Save File name:</b><br>"
    "This is where the summary will be saved.<br>"
    "If a file already exists, Expense Wizard will auto-rename it (e.g., add <code>_copy1</code>, <code>_copy2</code>, etc.).<br><br>"

    "<b>3. 📅 Select a Date Range:</b><br>"
    "Use the calendar to pick the start and end dates for your report.<br>"
    "Only transactions within this range will be included.<br><br>"

    "<b>4. 📦 Choose Grouping:</b><br>"
    "<ul>"
    "<li><b>Weekly:</b> Breaks the date range into 7-day periods</li>"
    "<li><b>Biweekly:</b> Splits each month into two halves (1st–15th, 16th–end)</li>"
    "<li><b>Monthly:</b> One report per full calendar month in the range</li>"
    "</ul><br>"

    "<b>5. 🧾 Click “Generate Report”:</b><br>"
    "A full Excel report will be created, including:<br>"
    "<ul>"
    "<li>One summary tab per period</li>"
    "<li>One tab per spending category</li>"
    "<li>Pie charts and totals included</li>"
    "<li>Custom labels based on grouping and date range</li>"
    "</ul><br>"

    "That’s it — review your spending the smart way! 💼"
)


        label = QLabel(help_text)
        label.setWordWrap(True)
        label.setFont(QFont("Courier New", 10))

        layout = QVBoxLayout()
        layout.addWidget(label)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)

        self.setLayout(layout)

class ExpenseSorterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Expense Wizard")
        self.setWindowIcon(QIcon("app_icon.ico"))
        self.resize(700,400)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout()
        central_widget.setLayout(self.main_layout)
        

        self.init_ui()
        



    def init_ui(self):
        layout = self.main_layout

        


        menubar = self.menuBar()
        settings_menu = menubar.addMenu("Settings")

        self.export_format_action = QComboBox()
        self.export_format_action.addItems(["Excel (.xlsx)", "CSV (.csv)"])
        export_widget = QWidget()
        export_layout = QHBoxLayout()
        export_layout.setContentsMargins(5, 5, 5, 5)
        export_layout.addWidget(QLabel("Export Format:"))
        export_layout.addWidget(self.export_format_action)
        export_widget.setLayout(export_layout)

        help_menu = menubar.addMenu("Help")

        help_action = QAction("View Instructions", self)
        help_action.triggered.connect(self.show_help)
        help_menu.addAction(help_action)

        self.chart_type_action = QComboBox()
        self.chart_type_action.addItems(["Pie", "Bar", "Column", "Doughnut", "Radar"])
        chart_widget = QWidget()
        chart_layout = QHBoxLayout()
        chart_layout.setContentsMargins(5, 5, 5, 5)
        chart_layout.addWidget(QLabel("Chart Type:"))
        chart_layout.addWidget(self.chart_type_action)
        chart_widget.setLayout(chart_layout)

        chart_action = QWidgetAction(self)
        chart_action.setDefaultWidget(chart_widget)
        settings_menu.addAction(chart_action)





# Add as QWidgetAction to embed in menu
        export_action = QWidgetAction(self)
        export_action.setDefaultWidget(export_widget)
        settings_menu.addAction(export_action)


        folder_layout = QHBoxLayout()
        self.folder_input = QLineEdit()
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self.browse_input_folder)
        folder_layout.addWidget(QLabel("Input Folder:"))
        folder_layout.addWidget(self.folder_input)
        folder_layout.addWidget(browse_button)
        layout.addLayout(folder_layout)

        output_layout = QHBoxLayout()
        self.output_input = QLineEdit()
        output_browse_button = QPushButton("Browse")
        output_browse_button.clicked.connect(self.browse_output_file)
        output_layout.addWidget(QLabel("Save File As:"))
        output_layout.addWidget(self.output_input)
        output_layout.addWidget(output_browse_button)
        layout.addLayout(output_layout)

        range_layout = QHBoxLayout()
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("MM-dd-yyyy")
        self.start_date.setDate(QDate.currentDate().addMonths(-1))

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("MM-dd-yyyy")
        self.end_date.setDate(QDate.currentDate())

        range_layout.addWidget(QLabel("Start Date:"))
        range_layout.addWidget(self.start_date)
        range_layout.addWidget(QLabel("End Date:"))
        range_layout.addWidget(self.end_date)
        layout.addLayout(range_layout)

        grouping_layout = QHBoxLayout()
        self.grouping = QComboBox()
        self.grouping.addItems(["Weekly", "Biweekly", "Monthly"])
        grouping_layout.addWidget(QLabel("Group By:"))
        grouping_layout.addWidget(self.grouping)
        layout.addLayout(grouping_layout)

        

        button_layout = QHBoxLayout()
        self.run_button = QPushButton("Generate Report")
        self.run_button.clicked.connect(self.run_script)
        button_layout.addWidget(self.run_button)
        layout.addLayout(button_layout)

        


        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)
        #layout = self.main_layout 

        self.load_preferences()
        self.run_interactive_tutorial()


    def load_preferences(self):
        config_path = os.path.join(os.path.expanduser("~"), ".expense_sorter_config.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, "r") as f:
                    config = json.load(f)
                    if "last_input" in config:
                        self.folder_input.setText(config["last_input"])
                    if "last_output" in config:
                        self.output_input.setText(config["last_output"])
                    if "start_date" in config:
                        self.start_date.setDate(QDate.fromString(config["start_date"], "MM-dd-yyyy"))
                    if "end_date" in config:
                        self.end_date.setDate(QDate.fromString(config["end_date"], "MM-dd-yyyy"))
                    if "group_by" in config:
                        index = self.grouping.findText(config["group_by"])
                        if index != -1:
                            self.grouping.setCurrentIndex(index)
                    if "export_format" in config:
                        idx = self.export_format_action.findText(config["export_format"])
                        if idx != -1:
                            self.export_format_action.setCurrentIndex(idx)
                    if "chart_type" in config:
                        idx = self.chart_type_action.findText(config["chart_type"])
                        if idx != -1:
                            self.chart_type_action.setCurrentIndex(idx)

                    else:
                        self.export_format_action.setCurrentIndex(
                            self.export_format_action.findText("Excel(.xlsx)")
                        )

            except Exception:
                pass

    def save_preferences(self):
        config_path = os.path.join(os.path.expanduser("~"), ".expense_sorter_config.json")
        config = {
            "last_input": self.folder_input.text(),
            "last_output": self.output_input.text(),
            "start_date": self.start_date.date().toString("MM-dd-yyyy"),
            "end_date": self.end_date.date().toString("MM-dd-yyyy"),
            "group_by": self.grouping.currentText(),
            "export_format": self.export_format_action.currentText(),
            "tutorial_shown": True

        }
        try:
            with open(config_path, "w") as f:
                json.dump(config, f)
        except Exception as e:
            print("Failed to load preferences:",e)
            pass

    def show_help(self):
        help_dialog = HelpDialog(self)
        help_dialog.exec_()

    def run_interactive_tutorial(self):
        config_path = os.path.join(os.path.expanduser("~"), ".expense_sorter_config.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, "r") as f:
                    config = json.load(f)
                    if config.get("tutorial_shown"):
                        return
            except Exception:
                pass

        steps = [
            (None, "📊 Welcome to Expense Wizard!\n\nThis tool helps you organize your credit card statements to get a clear view of where your money is spent."),
            (self.folder_input, "\U0001F4C2 Input Folder: Select the folder containing your expense files."),
            (self.output_input, "\U0001F4BE Save File As: Choose where to save the final report."),
            (self.start_date, "\U0001F4C5 Start Date: Select when to begin filtering your transactions."),
            (self.end_date, "📅 End Date: Select when to stop."),
            (self.run_button, "\u25B6\uFE0F Generate Report: Click to create your organized report.")
        ]

        def show_next(index=0):
            if index >= len(steps):
                with open(config_path, "w") as f:
                    json.dump({"tutorial_shown": True}, f)
                return

            widget, message = steps[index]
            popup = GuidedStepOverlay(widget, message, self)
            popup.finished.connect(lambda _: show_next(index + 1))
            popup.show()

        show_next()

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
        start_date = self.start_date.date().toPyDate()
        end_date = self.end_date.date().toPyDate()
        group_mode = self.grouping.currentText()
        file_format = self.export_format_action.currentText() or ""
        use_csv = "csv" in file_format.lower()
        chart_type = self.chart_type_action.currentText()



        self.log_output.clear()

        if not folder:
            self.log_output.append("\u274C Input folder is required.")
            return

        if not out_folder:
            self.log_output.append("\u274C Output file path is required.")
            return

        self.log_output.append(f"\U0001F4C2 Input: {folder}")
        self.log_output.append(f"\U0001F4BE Output: {out_folder}")
        self.log_output.append(
    f"\U0001F4C5 Range: {self.start_date.date().toString('MM-dd-yyyy')} to {self.end_date.date().toString('MM-dd-yyyy')}, Grouped: {self.grouping.currentText()}"
)

        self.log_output.append("\u2699\uFE0F Running script...")

        try:

            path = main_processing_function(folder, start_date, end_date, out_folder, group_mode,use_csv, chart_type=chart_type)

            self.log_output.append(f"\u2705 Report saved to: {path}")
            self.save_preferences()
        except Exception as e:
            self.log_output.append(f"\u274C Error: {str(e)}")
            self.log_output.append(traceback.format_exc())

def main():
    app = QApplication(sys.argv)
    window = ExpenseSorterGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
