from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QFileDialog
from PyQt5 import uic
import sys


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi("WOIDS.ui", self)

        # connect widgets to python
        self.browse_button = self.findChild(QPushButton, "browse_button")
        self.folder_path_line = self.findChild(QLineEdit, "folder_path")
        self.check_button = self.findChild(QPushButton, "check_button")
        self.generate_button = self.findChild(QPushButton, "generate_button")
        self.template_button = self.findChild(QPushButton, "template_button")

        # connect button presses to actions
        self.browse_button.clicked.connect(self.browse_folders)
        self.check_button.clicked.connect(self.check_folders)
        self.generate_button.clicked.connect(self.generate_report)
        self.template_button.clicked.connect(self.generate_template)

        self.show()

    def browse_folders(self):
        path = QFileDialog.getExistingDirectory(self, "Select the folder containing the spreadsheet")
        self.folder_path_line.setText(path)

    def check_folders(self):
        pass

    def generate_report(self):
        self.check_folders()
        pass

    def generate_template(self):
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
