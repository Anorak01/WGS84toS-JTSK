# author: Adam Friedman

from convertor import convertToJTSK

from excel import process_sheet

from translations import t

import sys
import os

from PySide6.QtWidgets import (
    QApplication,
    QLabel,
    QMainWindow,
    QVBoxLayout,
    QWidget,
    QFileDialog,
    QPushButton,
    QCheckBox,
    QMenu,
)
from PySide6.QtGui import QAction

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.fileloc = ""
        self.savedir = ""
        self.lang = "cs"
        
        # Menu bar
        menubar = self.menuBar()
        language_menu = menubar.addMenu(t[self.lang]["language_menu"])
        
        # Language options
        english_action = QAction(t[self.lang]["language_en"], self)
        english_action.setCheckable(True)
        english_action.setChecked(self.lang == "en")
        english_action.triggered.connect(lambda: self.switch_language("en"))
        
        czech_action = QAction(t[self.lang]["language_cs"], self)
        czech_action.setCheckable(True)
        czech_action.setChecked(self.lang == "cs")
        czech_action.triggered.connect(lambda: self.switch_language("cs"))
        
        # Actions for language menu
        language_menu.addAction(czech_action)  # Czech first since it's default
        language_menu.addAction(english_action)
        
        self.language_actions = {"en": english_action, "cs": czech_action}

        self.setWindowTitle(t[self.lang]["window_title"])

        self.label = QLabel() # Label for input file name

        self.button = QPushButton(t[self.lang]["pick_input"], self) # Button that opens file dialog
        self.button.clicked.connect(self.buttonclick)

        self.savebutton = QPushButton(t[self.lang]["pick_save_dir"], self) # Button that opens file(folder) dialog
        self.savebutton.clicked.connect(self.savebuttonclick)

        self.savelabel = QLabel() # Label for save directory

        self.translatebutton = QPushButton(t[self.lang]["translate"], self) # Button to trigger conversion
        self.translatebutton.clicked.connect(self.run_translate)

        self.done_label = QLabel("") # Label to display status

        # Checkbox to toggle negative XY
        self.checkbox = QCheckBox(t[self.lang]["negative_xy"], self)

        self.convert_height_checkbox = QCheckBox(t[self.lang]["convert_height"], self)
        self.convert_height_checkbox.setChecked(True)

        layout = QVBoxLayout()
        layout.addWidget(self.button)
        layout.addWidget(self.label)
        layout.addWidget(self.savebutton)
        layout.addWidget(self.savelabel)
        layout.addWidget(self.checkbox)
        layout.addWidget(self.convert_height_checkbox)
        layout.addWidget(self.translatebutton)
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
            self.label.setText(t[self.lang]["file_to_translate"] + str(fileName[0]))
            self.fileloc = str(fileName[0])
        else:
            self.label.setText(t[self.lang]["not_excel_error"])

    def savebuttonclick(self):
        self.done_label.setText("")
        fileName = QFileDialog.getExistingDirectory(self, t[self.lang]["pick_save_dir"], os.path.expanduser("~"))
        print(fileName)
        self.savelabel.setText(t[self.lang]["save_directory"] + str(fileName))
        self.savedir = str(fileName)

    def run_translate(self):
        # Validate input file
        if not self.fileloc or not (self.fileloc.endswith('.xlsx') or self.fileloc.endswith('.XLSX')):
            self.done_label.setText(t[self.lang]["no_file_error"])
            return
            
        # Validate save directory
        if not self.savedir:
            self.done_label.setText(t[self.lang]["no_dir_error"])
            return
            
        self.done_label.setText(t[self.lang]["processing"])
        # Force the UI to update immediately
        QApplication.processEvents()
        
        ok = process_sheet(self.fileloc, self.savedir, self.checkbox.isChecked(), self.convert_height_checkbox.isChecked())
        if ok:
            self.done_label.setText(t[self.lang]["done"])
        else:
            self.done_label.setText(t[self.lang]["error"])
            
    def switch_language(self, new_lang):
        # Update checkmarks in menu
        for lang, action in self.language_actions.items():
            action.setChecked(lang == new_lang)
            
        self.lang = new_lang
        # Update menu text
        self.menuBar().findChild(QMenu).setTitle(t[self.lang]["language_menu"])
        self.language_actions["en"].setText(t[self.lang]["language_en"])
        self.language_actions["cs"].setText(t[self.lang]["language_cs"])
        
        # Update all UI elements with new language
        self.setWindowTitle(t[self.lang]["window_title"])
        self.button.setText(t[self.lang]["pick_input"])
        self.savebutton.setText(t[self.lang]["pick_save_dir"])
        self.translatebutton.setText(t[self.lang]["translate"])
        self.checkbox.setText(t[self.lang]["negative_xy"])
        self.convert_height_checkbox.setText(t[self.lang]["convert_height"])
        
        # Update labels if they have content
        if self.fileloc:
            self.label.setText(t[self.lang]["file_to_translate"] + self.fileloc)
        if self.savedir:
            self.savelabel.setText(t[self.lang]["save_directory"] + self.savedir)
        # Update any status label if it has text
        if self.done_label.text():
            # Find which status this label represents by checking all languages
            current_text = self.done_label.text()
            for status in ["done", "error", "processing", "no_file_error", "no_dir_error"]:
                if any(current_text == t[lang][status] for lang in t.keys()):
                    self.done_label.setText(t[self.lang][status])
                    break

app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()
