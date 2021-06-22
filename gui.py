from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QFileDialog, QLabel, QTextBrowser, \
    QErrorMessage
from PyQt5 import uic
import sys
from spreadsheet import Sheet, LocationInfo
from document import build_document


error_text = '<html><head/><body><p><span style=" color:#ef0000;">ERROR: Spreadsheet not found</span></p></body></html>'


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        # self.setWindowIcon(QIcon("./images/wsp square.png"))
        uic.loadUi("WOIDS.ui", self)

        self.sheet = None

        # connect widgets to python
        self.browse_button = self.findChild(QPushButton, "browse_button")
        self.error_label = self.findChild(QLabel, "error_label")
        self.file_path_line = self.findChild(QLineEdit, "folder_path")
        self.check_button = self.findChild(QPushButton, "check_button")
        self.generate_button = self.findChild(QPushButton, "generate_button")
        self.template_button = self.findChild(QPushButton, "template_button")
        self.info_box = self.findChild(QTextBrowser, "info_box")

        # connect button presses to actions
        self.browse_button.clicked.connect(self.browse_folders)
        self.check_button.clicked.connect(self.check_folders)
        self.generate_button.clicked.connect(self.generate_report)
        self.template_button.clicked.connect(self.generate_template)

        self.show()

    def browse_folders(self):
        # TODO: change default file location
        # TODO: crash on canceling file
        path = QFileDialog.getOpenFileName(self, "Select the spreadsheet", "E:/WSP/5.17/RED LINE/CABOT (R-13)")
        if path[0] != "":
            self.file_path_line.setText(path[0])

            try:
                self.sheet = Sheet(path[0])
                self.error_label.setText("")
            except ValueError:
                self.error_label.setText(error_text)
                self.sheet = None

    def check_folders(self):
        self.info_box.setText("")

        if self.sheet is not None:  # if sheet is loaded
            missing_pics = self.sheet.check_pictures()

            if len(missing_pics) != 0:  # if pictures are missing
                text = ""
                for pic in missing_pics:
                    text += "Missing " + pic + " picture\n"
                self.info_box.setText(text)
            else:  # if all pictures exist
                self.info_box.setText("All entries have a matching picture :)\n")
                return 1
        else:
            error_dialog = QErrorMessage()
            error_dialog.setWindowTitle("Error")
            error_dialog.showMessage("Please select an appropriate spreadsheet\n")

            error_dialog.exec()

        return 0

    def generate_report(self):
        if self.check_folders():
            build_document(self.sheet)
            self.info_box.setText(self.info_box.toPlainText() + "Report generated!\n")
        # TODO: overwrite confirmation

    def generate_template(self):
        pass
        # TODO: make template sheet
        # TODO: copy paste on fcn call
        # TODO: overwrite confirmation


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
