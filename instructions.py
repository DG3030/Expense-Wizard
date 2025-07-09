from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout, QPushButton
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QFont
import sys

class TypewriterDialog(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Instructions")
        self.setGeometry(200, 200, 500, 150)

        self.full_text = "Welcome, traveler. Use the arrow keys to navigate. Press 'I' to view your inventory."
        self.displayed_text = ""
        self.text_index = 0

        self.label = QLabel("")
        self.label.setFont(QFont("Courier New", 12))
        self.label.setWordWrap(True)

        self.skip_button = QPushButton("Skip")
        self.skip_button.clicked.connect(self.skip_animation)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.skip_button)
        self.setLayout(layout)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_text)
        self.timer.start(40)  # milliseconds between letters

    def update_text(self):
        if self.text_index < len(self.full_text):
            self.displayed_text += self.full_text[self.text_index]
            self.label.setText(self.displayed_text)
            self.text_index += 1
        else:
            self.timer.stop()

    def skip_animation(self):
        self.timer.stop()
        self.label.setText(self.full_text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TypewriterDialog()
    window.show()
    sys.exit(app.exec_())
