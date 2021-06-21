from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QFileDialog, QLabel
from PyQt5 import uic
import sys
from spreadsheet import Sheet

error_text = '<html><head/><body><p><span style=" color:#ef0000;">ERROR: Spreadsheet not found</span></p></body></html>'


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi("WOIDS.ui", self)

        # connect widgets to python
        self.browse_button = self.findChild(QPushButton, "browse_button")
        self.sheet_error_label = self.findChild(QLabel, "sheet_error_label")
        self.file_path_line = self.findChild(QLineEdit, "folder_path")
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
        path = QFileDialog.getOpenFileName(self, "Select the spreadsheet", "E:/WSP/5.17/RED LINE/CABOT (R-13)")
        self.file_path_line.setText(path[0])

        try:
            Sheet(path[0])
            self.sheet_error_label.setText("")
        except ValueError:  # what error
            self.sheet_error_label.setText(error_text)
            print("sheet not found")

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
