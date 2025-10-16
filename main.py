# author: Adam Friedman

from convertor import convertToJTSK

from excel import process_sheet

import sys
import os

from PySide6.QtWidgets import (
    QApplication,
    QLabel,
    QLineEdit,
    QMainWindow,
    QVBoxLayout,
    QWidget,
    QFileDialog,
    QPushButton,
    QCheckBox,
)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.fileloc = ""
        self.savedir = ""

        self.setWindowTitle("Překladátor")

        self.label = QLabel() # Label for input file name

        self.button = QPushButton("Open", self) # Button that opens file dialog
        self.button.clicked.connect(self.buttonclick)

        self.savebutton = QPushButton("Pick Save Directory", self) # Button that opens file(folder) dialog
        self.savebutton.clicked.connect(self.savebuttonclick)

        self.savelabel = QLabel() # Label for save directory

        self.translatebutton = QPushButton("Translate", self) # Button to trigger conversion
        self.translatebutton.clicked.connect(self.run_translate)

        self.done_label = QLabel("") # Label to display status

        # checkbox to toggle negative XY
        self.checkbox = QCheckBox("Negative XY", self)

        layout = QVBoxLayout()
        layout.addWidget(self.button)
        layout.addWidget(self.label)
        layout.addWidget(self.savebutton)
        layout.addWidget(self.savelabel)
        layout.addWidget(self.translatebutton)
        layout.addWidget(self.checkbox)
        layout.addWidget(self.done_label)

        container = QWidget()
        container.setLayout(layout)

        self.setCentralWidget(container)

    def buttonclick(self):
        self.done_label.setText("")
        fileName = QFileDialog.getOpenFileName(self, "Open Excel Sheet", os.path.expanduser("~"), "Excel Sheet (*.xlsx)")
        print(fileName[0])
        name = str(fileName[0])
        if (name.endswith(".xlsx") or name.endswith(".XLSX")):
            self.label.setText("File to translate: " + str(fileName[0]))
            self.fileloc = str(fileName[0])
        else:
            self.label.setText("Selected file is not an Excel sheet")

    def savebuttonclick(self):
        self.done_label.setText("")
        fileName = QFileDialog.getExistingDirectory(self, "Pick save directory", os.path.expanduser("~"))
        print(fileName)
        self.savelabel.setText("Save directory: " + str(fileName))
        self.savedir = str(fileName)

    def run_translate(self):
        ok = process_sheet(self.fileloc, self.savedir, self.checkbox.isChecked())
        if ok:
            self.done_label.setText("Done!")
        else:
            self.done_label.setText("Error!")

app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()
