# author: Adam Friedman

import datetime
import os
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QFont, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMenuBar,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from convertor import convertToJTSK
from excel import process_sheet
from translations import t


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
        language_menu.addAction(czech_action)
        language_menu.addAction(english_action)

        self.language_actions = {"en": english_action, "cs": czech_action}

        self.setWindowTitle(t[self.lang]["window_title"])

        # Set window size
        self.setMinimumWidth(500)
        self.setFixedHeight(450)

        # Set application style
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QGroupBox {
                background-color: white;
                border: 1px solid #cccccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
                margin-bottom: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
                background-color: white;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton#translateButton {
                background-color: #4CAF50;
                font-weight: bold;
            }
            QPushButton#translateButton:hover {
                background-color: #388E3C;
            }
            QLabel#statusLabel {
                font-weight: bold;
            }
        """)

        # File Selection Group
        file_group = QGroupBox(self)
        file_group.setTitle("File Selection")
        file_group.setFixedHeight(150)
        file_layout = QVBoxLayout()
        file_layout.setSpacing(15)
        file_layout.setContentsMargins(10, 20, 10, 10)

        # Input file selection
        input_layout = QHBoxLayout()
        input_layout.setSpacing(10)
        self.button = QPushButton(t[self.lang]["pick_input"], self)
        self.button.clicked.connect(self.buttonclick)
        self.button.setMinimumWidth(150)
        self.button.setFixedHeight(32)
        size_policy = QSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        self.button.setSizePolicy(size_policy)
        self.label = QLabel()
        self.label.setWordWrap(True)
        input_layout.addWidget(self.button)
        input_layout.addWidget(self.label, 1)
        file_layout.addLayout(input_layout)

        # Output directory selection
        output_layout = QHBoxLayout()
        output_layout.setSpacing(10)
        self.savebutton = QPushButton(t[self.lang]["pick_save_dir"], self)
        self.savebutton.clicked.connect(self.savebuttonclick)
        self.savebutton.setMinimumWidth(150)
        self.savebutton.setFixedHeight(32)
        self.savebutton.setSizePolicy(size_policy)
        self.savelabel = QLabel()
        self.savelabel.setWordWrap(True)
        output_layout.addWidget(self.savebutton)
        output_layout.addWidget(self.savelabel, 1)
        file_layout.addLayout(output_layout)

        # Set minimum height for the group to prevent overlap
        file_group.setMinimumHeight(120)

        file_group.setLayout(file_layout)

        # Options Group
        options_group = QGroupBox(self)
        options_group.setTitle("Options")
        options_group.setFixedHeight(80)
        options_layout = QHBoxLayout()
        options_layout.setSpacing(20)
        options_layout.setContentsMargins(10, 15, 10, 10)

        self.checkbox = QCheckBox(t[self.lang]["negative_xy"], self)
        self.convert_height_checkbox = QCheckBox(t[self.lang]["convert_height"], self)
        self.convert_height_checkbox.setChecked(True)

        # Set fixed height for checkboxes to prevent compression
        self.checkbox.setFixedHeight(25)
        self.convert_height_checkbox.setFixedHeight(25)

        options_layout.addWidget(self.checkbox)
        options_layout.addWidget(self.convert_height_checkbox)
        options_layout.addStretch(1)
        options_group.setLayout(options_layout)

        # Action Section
        action_layout = QVBoxLayout()
        self.translatebutton = QPushButton(t[self.lang]["translate"], self)
        self.translatebutton.setObjectName("translateButton")  # Set ID for styling
        self.translatebutton.clicked.connect(self.run_translate)
        self.translatebutton.setMinimumHeight(40)
        self.translatebutton.setFixedWidth(200)

        self.done_label = QLabel("")
        self.done_label.setObjectName("statusLabel")
        self.done_label.setAlignment(Qt.AlignCenter)
        self.done_label.setMinimumHeight(30)

        # Main Layout
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # Create a widget for the file group
        file_widget = QWidget()
        file_widget_layout = QVBoxLayout(file_widget)
        file_widget_layout.setContentsMargins(0, 0, 0, 0)
        file_widget_layout.addWidget(file_group)

        # Create a widget for the options group
        options_widget = QWidget()
        options_widget_layout = QVBoxLayout(options_widget)
        options_widget_layout.setContentsMargins(0, 0, 0, 0)
        options_widget_layout.addWidget(options_group)

        # Add widgets to main layout
        layout.addWidget(file_widget)
        layout.addWidget(options_widget)

        # Center the translate button
        translate_container = QHBoxLayout()
        translate_container.addStretch(1)
        translate_container.addWidget(self.translatebutton)
        translate_container.addStretch(1)

        layout.addLayout(translate_container)
        layout.addSpacing(10)
        layout.addWidget(self.done_label)

        container = QWidget()
        container.setLayout(layout)

        self.setCentralWidget(container)

    def buttonclick(self):
        self.done_label.setText("")
        fileName = QFileDialog.getOpenFileName(
            self, "Open Excel Sheet", os.path.expanduser("~"), "Excel Sheet (*.xlsx)"
        )
        print(fileName[0])
        name = str(fileName[0])
        if name.endswith(".xlsx") or name.endswith(".XLSX"):
            self.label.setText(t[self.lang]["file_to_translate"] + str(fileName[0]))
            self.fileloc = str(fileName[0])
        else:
            self.label.setText(t[self.lang]["not_excel_error"])

    def savebuttonclick(self):
        self.done_label.setText("")
        fileName = QFileDialog.getExistingDirectory(
            self, t[self.lang]["pick_save_dir"], os.path.expanduser("~")
        )
        print(fileName)
        self.savelabel.setText(t[self.lang]["save_directory"] + str(fileName))
        self.savedir = str(fileName)

    def run_translate(self):
        # Validate input file
        if not self.fileloc or not (
            self.fileloc.endswith(".xlsx") or self.fileloc.endswith(".XLSX")
        ):
            self.done_label.setText(t[self.lang]["no_file_error"])
            self.done_label.setStyleSheet("color: #f44336;")
            return

        # Validate save directory
        if not self.savedir:
            self.done_label.setText(t[self.lang]["no_dir_error"])
            self.done_label.setStyleSheet("color: #f44336;")
            return

        self.done_label.setText(t[self.lang]["processing"])
        self.done_label.setStyleSheet("color: #2196F3;")
        # Force the UI to update immediately
        QApplication.processEvents()

        def save_error(error_msg):
            try:
                error_path = os.path.join(self.savedir, "error.txt")
                with open(error_path, "w", encoding="utf-8") as f:
                    f.write(f"Error occurred at: {datetime.datetime.now()}\n")
                    f.write(f"Input file: {self.fileloc}\n")
                    f.write(f"Output directory: {self.savedir}\n")
                    f.write(f"Settings:\n")
                    f.write(f"- Negative XY: {self.checkbox.isChecked()}\n")
                    f.write(
                        f"- Convert Height: {self.convert_height_checkbox.isChecked()}\n"
                    )
                    f.write(f"\nError details:\n{error_msg}")
            except Exception as write_err:
                print(f"Failed to write error file: {str(write_err)}")

        try:
            ok = process_sheet(
                self.fileloc,
                self.savedir,
                self.checkbox.isChecked(),
                self.convert_height_checkbox.isChecked(),
            )
            if ok:
                self.done_label.setText(t[self.lang]["done"])
                self.done_label.setStyleSheet("color: #4CAF50;")
            else:
                save_error("Process failed with no specific error message")
                self.done_label.setText(t[self.lang]["error"])
                self.done_label.setStyleSheet("color: #f44336;")
        except Exception as e:
            error_msg = f"Error during translation: {str(e)}"
            print(error_msg)
            save_error(error_msg)
            self.done_label.setText(t[self.lang]["error"])
            self.done_label.setStyleSheet("color: #f44336;")

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
            for status in [
                "done",
                "error",
                "processing",
                "no_file_error",
                "no_dir_error",
            ]:
                if any(current_text == t[lang][status] for lang in t.keys()):
                    self.done_label.setText(t[self.lang][status])
                    break


app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()
