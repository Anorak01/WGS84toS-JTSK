# author: Adam Friedman

import datetime
import os
import sys

from PySide6.QtCore import Qt, QSettings, Signal, QTimer, QPointF, QRectF
from PySide6.QtGui import QAction, QColor, QFont, QIcon, QPalette, QPainter, QPen, QPixmap, QPolygonF
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMenuBar,
    QPushButton,
    QSizePolicy,
    QStyle,
    QStyleOptionGroupBox,
    QVBoxLayout,
    QWidget,
)

from convertor import convertToJTSK
from excel import (
    process_sheet,
    missing_column,
    DEFAULT_RECORD_INDEX_COLUMN,
    DEFAULT_LATITUDE_COLUMN,
    DEFAULT_LONGITUDE_COLUMN,
    DEFAULT_DEPTH_COLUMN,
    DEFAULT_HEIGHT_COLUMN,
    DEFAULT_DESCRIPTION_COLUMN,
)
from translations import t

import load_splash

# QGroupBox that shows/hides its content when the title is clicked
class CollapsibleGroupBox(QGroupBox):
    titleClicked = Signal()

    def __init__(self, title, parent=None):
        super().__init__(title, parent)
        self.setMouseTracking(True)

    def titleRect(self):
        option = QStyleOptionGroupBox()
        self.initStyleOption(option)
        return self.style().subControlRect(QStyle.CC_GroupBox, option, QStyle.SC_GroupBoxLabel, self)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.titleRect().contains(event.position().toPoint()):
            self.titleClicked.emit()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.titleRect().contains(event.position().toPoint()):
            self.setCursor(Qt.PointingHandCursor)
        else:
            self.unsetCursor()
        super().mouseMoveEvent(event)

def make_reset_icon() -> QIcon: # simple monochrome circular arrow
    pixmap = QPixmap(16, 16)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setPen(QPen(QColor("#555555"), 2))
    painter.drawArc(QRectF(3, 3, 10, 10), 45 * 16, 270 * 16)
    painter.setPen(Qt.NoPen)
    painter.setBrush(QColor("#555555"))
    painter.drawPolygon(QPolygonF([QPointF(14, 7), QPointF(13, 3), QPointF(9.5, 6)]))
    painter.end()
    return QIcon(pixmap)

# column name settings: (settings key, translation key, default column name)
COLUMN_SETTINGS = [
    ("record_index_column", "record_index_column", DEFAULT_RECORD_INDEX_COLUMN),
    ("latitude_column", "latitude_column", DEFAULT_LATITUDE_COLUMN),
    ("longitude_column", "longitude_column", DEFAULT_LONGITUDE_COLUMN),
    ("depth_column", "depth_column", DEFAULT_DEPTH_COLUMN),
    ("height_column", "height_column", DEFAULT_HEIGHT_COLUMN),
    ("description_column", "description_column", DEFAULT_DESCRIPTION_COLUMN),
]

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.settings = QSettings("Anorak01", "WGS84toS-JTSK")

        self.fileloc = ""
        self.savedir = ""
        self.file_check = None
        # re-checks the input file shortly after a column name is changed
        self.check_timer = QTimer(self)
        self.check_timer.setSingleShot(True)
        self.check_timer.setInterval(500)
        self.check_timer.timeout.connect(self.validate_input_file)
        self.lang = self.settings.value("lang", "cs", type=str)

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

        # Set window size (columns section is collapsed by default; heights must
        # stay at or above the layout's sizeHint, or widgets compress and shift)
        self.setMinimumWidth(500)
        self.setFixedHeight(500)

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
            QGroupBox#columnsGroup[collapsed="true"] {
                background-color: transparent;
                border: none;
            }
            QGroupBox#columnsGroup[collapsed="true"]::title {
                background-color: transparent;
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
            QPushButton#columnResetButton {
                background-color: #e0e0e0;
                color: #555555;
                font-weight: bold;
                padding: 2px;
            }
            QPushButton#columnResetButton:hover {
                background-color: #bdbdbd;
            }
            QLabel#statusLabel {
                font-weight: bold;
            }
        """)

        # File Selection Group
        file_group = QGroupBox(self)
        file_group.setTitle(t[self.lang]["file_picker_box"])
        file_group.setFixedHeight(180)
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

        # Input file check status
        check_layout = QHBoxLayout()
        check_layout.setSpacing(10)
        self.file_check_label = QLabel()
        self.file_check_label.setWordWrap(True)
        check_layout.addWidget(self.file_check_label, 1)
        file_layout.addLayout(check_layout)

        # Set minimum height for the group to prevent overlap
        file_group.setMinimumHeight(120)

        file_group.setLayout(file_layout)

        # Options Group
        self.options_group = QGroupBox(self)
        self.options_group.setTitle(t[self.lang]["options_box"])
        self.options_group.setFixedHeight(80)
        options_layout = QHBoxLayout()
        options_layout.setSpacing(20)
        options_layout.setContentsMargins(10, 15, 10, 10)

        self.checkbox = QCheckBox(t[self.lang]["negative_xy"], self)
        self.checkbox.setChecked(self.settings.value("negative_xy", False, type=bool))
        self.checkbox.checkStateChanged.connect(self.check_changed)
        self.convert_height_checkbox = QCheckBox(t[self.lang]["convert_height"], self)
        self.convert_height_checkbox.setChecked(self.settings.value("convert_height", False, type=bool))
        self.convert_height_checkbox.checkStateChanged.connect(self.check_changed)
        self.subtract_depth_checkbox = QCheckBox(t[self.lang]["subtract_depth"], self)
        self.subtract_depth_checkbox.setChecked(self.settings.value("subtract_depth", True, type=bool))
        self.subtract_depth_checkbox.checkStateChanged.connect(self.check_changed)

        # Set fixed height for checkboxes to prevent compression
        self.checkbox.setFixedHeight(25)
        self.convert_height_checkbox.setFixedHeight(25)
        self.subtract_depth_checkbox.setFixedHeight(25)

        options_layout.addWidget(self.checkbox)
        options_layout.addWidget(self.convert_height_checkbox)
        options_layout.addWidget(self.subtract_depth_checkbox)
        options_layout.addStretch(1)
        self.options_group.setLayout(options_layout)

        # Column Names Group (collapsible, hidden by default)
        self.columns_group = CollapsibleGroupBox("▷ " + t[self.lang]["columns_box"], self)
        self.columns_group.setObjectName("columnsGroup")
        self.columns_group.setProperty("collapsed", True)
        self.columns_group.titleClicked.connect(self.toggle_columns)

        group_layout = QVBoxLayout()
        group_layout.setContentsMargins(10, 15, 10, 10)

        self.columns_content = QWidget()
        columns_layout = QGridLayout(self.columns_content)
        columns_layout.setSpacing(10)
        columns_layout.setContentsMargins(0, 0, 0, 0)

        self.column_labels = {}
        self.column_edits = {}
        self.column_reset_buttons = {}

        for i, (key, label_key, default) in enumerate(COLUMN_SETTINGS):
            row, col = divmod(i, 2)

            entry_layout = QHBoxLayout()
            entry_layout.setSpacing(5)

            label = QLabel(t[self.lang][label_key])
            edit = QLineEdit(self.settings.value(key, default, type=str))
            edit.textChanged.connect(lambda text, k=key: self.column_changed(k, text))

            reset = QPushButton()
            reset.setObjectName("columnResetButton")
            reset.setIcon(make_reset_icon())
            reset.setFixedSize(24, 24)
            reset.setToolTip(t[self.lang]["reset_column"])
            reset.clicked.connect(lambda checked=False, e=edit, d=default: e.setText(d))

            entry_layout.addWidget(label)
            entry_layout.addWidget(edit, 1)
            entry_layout.addWidget(reset)

            columns_layout.addLayout(entry_layout, row, col)

            self.column_labels[key] = label
            self.column_edits[key] = edit
            self.column_reset_buttons[key] = reset

        self.columns_content.setVisible(False)

        group_layout.addWidget(self.columns_content)
        self.columns_group.setLayout(group_layout)

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
        options_widget_layout.addWidget(self.options_group)

        # Create a widget for the columns group
        columns_widget = QWidget()
        columns_widget_layout = QVBoxLayout(columns_widget)
        columns_widget_layout.setContentsMargins(0, 0, 0, 0)
        columns_widget_layout.addWidget(self.columns_group)

        # Add widgets to main layout
        layout.addWidget(file_widget)
        layout.addWidget(options_widget)
        layout.addWidget(columns_widget)

        # Center the translate button
        translate_container = QHBoxLayout()
        translate_container.addStretch(1)
        translate_container.addWidget(self.translatebutton)
        translate_container.addStretch(1)

        layout.addLayout(translate_container)
        layout.addSpacing(10)
        layout.addWidget(self.done_label)
        layout.addStretch(1)

        container = QWidget()
        container.setLayout(layout)

        self.setCentralWidget(container)

    def buttonclick(self):
        self.done_label.setText("")
        self.file_check = None
        self.update_file_check()
        fileName = QFileDialog.getOpenFileName(
            self, "Open Excel Sheet", os.path.expanduser("~"), "Excel Sheet (*.xlsx)"
        )
        print(fileName[0])
        name = str(fileName[0])
        if name.endswith(".xlsx") or name.endswith(".XLSX"):
            self.label.setText(t[self.lang]["file_to_translate"] + str(fileName[0]))
            self.fileloc = str(fileName[0])
            self.validate_input_file()
        else:
            self.label.setText(t[self.lang]["not_excel_error"])

    def validate_input_file(self):
        columns = [(key, self.column_edits[key].text()) for key, label_key, default in COLUMN_SETTINGS]
        try:
            missing = missing_column(self.fileloc, [name for _, name in columns])
        except Exception:
            self.file_check = ("error", "", "")
            self.update_file_check()
            return
        if missing is None:
            self.file_check = ("valid", "", "")
        else:
            self.file_check = ("missing", next(key for key, name in columns if name == missing), missing)
        self.update_file_check()

    def update_file_check(self):
        if self.file_check is None:
            self.file_check_label.setText("")
            return
        state, key, column_name = self.file_check
        if state == "valid":
            self.file_check_label.setText("✓ " + t[self.lang]["file_valid"])
            self.file_check_label.setStyleSheet("color: #4CAF50;")
        elif state == "missing":
            self.file_check_label.setText("✗ " + t[self.lang]["column_not_found"].format(t[self.lang][key], column_name))
            self.file_check_label.setStyleSheet("color: #f44336;")
        else:
            self.file_check_label.setText("✗ " + t[self.lang]["file_read_error"])
            self.file_check_label.setStyleSheet("color: #f44336;")

    def savebuttonclick(self):
        self.done_label.setText("")
        fileName = QFileDialog.getExistingDirectory(
            self, t[self.lang]["pick_save_dir"], os.path.expanduser("~")
        )
        print(fileName)
        self.savelabel.setText(t[self.lang]["save_directory"] + str(fileName))
        self.savedir = str(fileName)

    def check_changed(self):
        self.settings.setValue("negative_xy", self.checkbox.isChecked())
        self.settings.setValue("convert_height", self.convert_height_checkbox.isChecked())
        self.settings.setValue("subtract_depth", self.subtract_depth_checkbox.isChecked())

    def column_changed(self, key, text):
        self.settings.setValue(key, text)
        if self.fileloc:
            self.check_timer.start()

    def toggle_columns(self):
        expanded = not self.columns_content.isVisible()
        self.columns_content.setVisible(expanded)
        self.columns_group.setTitle(("▽ " if expanded else "▷ ") + t[self.lang]["columns_box"])
        self.columns_group.setProperty("collapsed", not expanded)
        self.columns_group.style().unpolish(self.columns_group)
        self.columns_group.style().polish(self.columns_group)
        self.setFixedHeight(595 if expanded else 500)

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
                self.subtract_depth_checkbox.isChecked(),
                self.column_edits["record_index_column"].text(),
                self.column_edits["latitude_column"].text(),
                self.column_edits["longitude_column"].text(),
                self.column_edits["depth_column"].text(),
                self.column_edits["height_column"].text(),
                self.column_edits["description_column"].text()
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

        self.settings.setValue("lang", new_lang)
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
        self.subtract_depth_checkbox.setText(t[self.lang]["subtract_depth"])

        self.options_group.setTitle(t[self.lang]["options_box"])
        arrow = "▽ " if self.columns_content.isVisible() else "▷ "
        self.columns_group.setTitle(arrow + t[self.lang]["columns_box"])
        for key, label_key, default in COLUMN_SETTINGS:
            self.column_labels[key].setText(t[self.lang][label_key])
            self.column_reset_buttons[key].setToolTip(t[self.lang]["reset_column"])

        # Update labels if they have content
        if self.fileloc:
            self.label.setText(t[self.lang]["file_to_translate"] + self.fileloc)
        if self.savedir:
            self.savelabel.setText(t[self.lang]["save_directory"] + self.savedir)
        self.update_file_check()
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
