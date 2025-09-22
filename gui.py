import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QLabel, QLineEdit, QGroupBox, QSpinBox, QCheckBox, QComboBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QTextEdit,
    QTabWidget, QStyleFactory, QSizePolicy, QFormLayout, QFrame
)
from PyQt6.QtCore import Qt, QPoint
from PyQt6.QtGui import QFont, QIcon, QPixmap
from pathlib import Path
from datetime import datetime, timedelta

__version__ = "20250922"
__author__ = "Robert Simon Uy"
"""
pyinstaller --onefile --windowed --name "TimeSheet" --add-data "logo.png;." --icon="logo.png" gui.py
"""

try:
    from timesheet import process_timesheet, str_to_delta
except ImportError:
    def process_timesheet(*args, **kwargs):
        """Dummy function for when import fails."""
        pass


    def str_to_delta(*args, **kwargs):
        """Dummy function for when import fails."""
        pass


    app = QApplication(sys.argv)
    QMessageBox.critical(None, "Import Error",
                         "Could not find 'timesheet.py'.\nPlease ensure it is in the same directory as the application.")
    sys.exit(1)


class ConstrainedComboBox(QComboBox):
    """
    A robust QComboBox subclass that manually controls its popup's geometry.
    This is necessary to fix sizing and positioning bugs when the combobox
    is placed inside a QTableWidget's viewport.
    """

    def showPopup(self):
        super().showPopup()
        view = self.view()
        popup_container = view.parentWidget()
        if popup_container:
            bottom_left = QPoint(0, self.height())
            global_pos = self.mapToGlobal(bottom_left)
            popup_container.move(global_pos)
            popup_container.setFixedWidth(self.width())
            popup_container.setMaximumHeight(200)


class TimesheetApp(QWidget):
    def __init__(self):
        super().__init__()
        script_dir = Path(__file__).parent
        self.icon_path = script_dir / 'logo.png'
        self.init_ui()

    def init_ui(self):
        if self.icon_path.exists():
            self.setWindowIcon(QIcon(str(self.icon_path)))
        else:
            print(f"Warning: Icon file not found at '{self.icon_path}'.")

        # Versioning
        self.setWindowTitle(f'Timesheet Processor v{__version__}')
        self.setGeometry(100, 100, 750, 700)

        self.setWindowTitle('Timesheet Processor')
        self.setGeometry(100, 100, 750, 700)

        main_layout = QVBoxLayout(self)

        tabs = QTabWidget()
        tabs.addTab(self._create_main_tab(), "Home")
        tabs.addTab(self._create_settings_tab(), "Settings")
        tabs.addTab(self._create_breaks_tab(), "Break Times")
        tabs.addTab(self._create_about_tab(), "About")

        main_layout.addWidget(tabs)
        main_layout.addWidget(self._create_log_box())
        self.show()

    def _create_about_tab(self):
        about_tab = QWidget()
        main_layout = QVBoxLayout(about_tab)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        main_layout.setContentsMargins(25, 25, 25, 25)
        main_layout.setSpacing(15)

        # 1. Logo
        if self.icon_path.exists():
            logo_label = QLabel()
            pixmap = QPixmap(str(self.icon_path)).scaled(96, 96, Qt.AspectRatioMode.KeepAspectRatio,
                                                         Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            main_layout.addWidget(logo_label)

        # 2. Application Title and Description
        title_label = QLabel("Timesheet Processor")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)

        description_label = QLabel("A utility to clean, process, and summarize timesheet data from raw punch logs.")
        description_label.setWordWrap(True)
        description_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(description_label)

        # 3. Separator Line
        main_layout.addWidget(self._create_separator())

        # 4. Details Section using QFormLayout
        details_layout = QFormLayout()
        # THIS LINE IS NOW REMOVED to keep labels and values on the same line.
        # details_layout.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        details_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        details_layout.addRow("Version:", QLabel(f"{__version__}"))
        details_layout.addRow("Author:", QLabel(f"<b>{__author__}</b>"))
        details_layout.addRow("License:", QLabel("MIT License"))

        source_code_label = QLabel('<a href="https://github.com/RSv618/timesheet-summarizer">View on GitHub</a>')
        source_code_label.setOpenExternalLinks(True)
        details_layout.addRow("Source Code:", source_code_label)

        main_layout.addLayout(details_layout)

        # 5. Separator Line
        main_layout.addWidget(self._create_separator())

        # 6. "Built With" Section
        built_with_label = QLabel("Built with Python & PyQt6")
        built_with_label.setStyleSheet("color: #888;")
        built_with_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(built_with_label)

        main_layout.addStretch()

        return about_tab

    @staticmethod
    def _create_separator():
        """Helper function to create a horizontal line."""
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        return line

    def _create_settings_tab(self):
        settings_tab = QWidget()
        layout = QVBoxLayout(settings_tab)
        layout.addWidget(self._create_workday_rules_group())
        layout.addWidget(self._create_rounding_group())
        layout.addStretch()

        restore_btn = QPushButton("Restore Default Settings")
        restore_btn.setToolTip("Resets all settings across all tabs to their original values.")
        # noinspection PyUnresolvedReferences
        restore_btn.clicked.connect(self._restore_default_settings_clicked)

        restore_layout = QHBoxLayout()
        restore_layout.addStretch()
        restore_layout.addWidget(restore_btn)
        layout.addLayout(restore_layout)

        return settings_tab

    def _restore_default_settings_clicked(self):
        reply = QMessageBox.question(self, 'Confirm Restore',
                                     "Are you sure you want to restore all settings to their original defaults?\nAll your custom changes will be lost.",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.start_hour_check.setChecked(True)
            self.start_hour_combo.setCurrentText("07:00 AM")
            self.end_hour_check.setChecked(True)
            self.end_hour_combo.setCurrentText("10:00 PM")
            self.buffer_spinbox.setValue(15)
            self.first_in_combo.setCurrentText("09:30 AM")
            self.last_out_combo.setCurrentText("02:30 PM")
            self.breaks_table.setRowCount(0)
            self._add_break_row("Lunch", "12:00 PM", "01:00 PM", False)
            self._add_break_row("Dinner", "06:00 PM", "06:30 PM", True)
            self.rounding_table.setRowCount(0)
            self._add_rounding_row("04:00 PM")
            self._add_rounding_row("05:00 PM")
            self._add_rounding_row("06:00 PM")
            self.log("All settings have been restored to their default values.")
            QMessageBox.information(self, "Settings Restored",
                                    "All settings have been restored to their default values.")

    def _create_workday_rules_group(self):
        group_box = QGroupBox("Workday Rules")
        layout = QVBoxLayout()

        self.start_hour_check = QCheckBox("Paid Hours Start Time")
        self.start_hour_check.setToolTip(
            "Define the earliest time paid work can begin. Any punches before this time\n(e.g., a 5:30 AM punch for a 6:00 AM start) will be adjusted to this time.\nThis prevents payment for time clocked in before the official shift start.")
        self.start_hour_combo = self._create_time_combobox()
        self.start_hour_combo.setCurrentText("07:00 AM")
        layout.addLayout(self._create_two_widget_layout(self.start_hour_check, self.start_hour_combo))

        self.end_hour_check = QCheckBox("Paid Hours End Time")
        self.end_hour_check.setToolTip(
            "Define the latest time paid work can end. Any punches after this time\n(e.g., a 10:30 PM punch for a 10:00 PM end) will be adjusted to this time.\nThis prevents payment for time clocked out after the official shift end.")
        self.end_hour_combo = self._create_time_combobox()
        self.end_hour_combo.setCurrentText("10:00 PM")
        layout.addLayout(self._create_two_widget_layout(self.end_hour_check, self.end_hour_combo))

        self.start_hour_check.setChecked(True)
        self.end_hour_check.setChecked(True)
        # noinspection PyUnresolvedReferences
        self.start_hour_check.stateChanged.connect(
            lambda state: self.start_hour_combo.setEnabled(state == Qt.CheckState.Checked.value))
        # noinspection PyUnresolvedReferences
        self.end_hour_check.stateChanged.connect(
            lambda state: self.end_hour_combo.setEnabled(state == Qt.CheckState.Checked.value))
        self.start_hour_combo.setEnabled(self.start_hour_check.isChecked())
        self.end_hour_combo.setEnabled(self.end_hour_check.isChecked())

        buffer_label = QLabel("Grace Period (minutes):")
        buffer_label.setToolTip(
            "A buffer to correct early/late punches. A 15-minute grace period means a 6:46 AM punch for a 7:00 AM shift start is counted as 7:00 AM.")
        self.buffer_spinbox = QSpinBox()
        self.buffer_spinbox.setRange(1, 30)
        self.buffer_spinbox.setValue(15)
        layout.addLayout(self._create_two_widget_layout(buffer_label, self.buffer_spinbox))

        first_in_label = QLabel("Verify First Punch-In By:")
        first_in_label.setToolTip(
            "Helps validate the first punch of the day. If an employee's very first punch is 'OUT' but occurs before this time,\nthe system will assume it was a mistake and correct it to 'IN'. This rule only applies to the first punch.")
        self.first_in_combo = self._create_time_combobox()
        self.first_in_combo.setCurrentText("10:30 AM")
        layout.addLayout(self._create_two_widget_layout(first_in_label, self.first_in_combo))

        last_out_label = QLabel("Verify Last Punch-Out After:")
        last_out_label.setToolTip(
            "Helps validate the last punch of the day. If an employee's very last punch is 'IN' but occurs after this time,\nthe system will assume it was a mistake and correct it to 'OUT'. This rule only applies to the last punch.")
        self.last_out_combo = self._create_time_combobox()
        self.last_out_combo.setCurrentText("02:30 PM")
        layout.addLayout(self._create_two_widget_layout(last_out_label, self.last_out_combo))

        group_box.setLayout(layout)
        return group_box

    def _create_breaks_tab(self):
        breaks_tab = QWidget()
        layout = QVBoxLayout(breaks_tab)
        layout.addWidget(self._create_breaks_group())
        return breaks_tab

    def _create_breaks_group(self):
        group_box = QGroupBox("Define Scheduled Breaks")
        group_box.setToolTip(
            "- Unpaid Breaks: The program assumes no work is done. The program will ensure this period is not counted in the total hours (e.g., a mandatory 1-hour lunch).\n"
            "- Paid Breaks: For short, paid rest periods. The program will clean up any punches during this window, ensuring the employee is paid for the full duration (e.g., a 15-minute coffee break).")
        layout = QVBoxLayout(group_box)
        self.breaks_table = QTableWidget()
        self.breaks_table.setColumnCount(5)
        self.breaks_table.setHorizontalHeaderLabels(["Name", "Start Time", "End Time", "Paid Break", "Action"])
        self.breaks_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.breaks_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.breaks_table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self._add_break_row("Lunch", "12:00 PM", "01:00 PM", False)
        self._add_break_row("Dinner", "06:00 PM", "06:30 PM", True)
        layout.addWidget(self.breaks_table)
        add_btn = QPushButton("Add Break")
        # noinspection PyUnresolvedReferences
        add_btn.clicked.connect(lambda: self._add_break_row())
        layout.addWidget(add_btn, alignment=Qt.AlignmentFlag.AlignRight)
        return group_box

    def _add_break_row(self, name="", start_time="12:00 PM", end_time="01:00 PM", is_paid=False):
        row_position = self.breaks_table.rowCount()
        self.breaks_table.insertRow(row_position)
        self.breaks_table.setItem(row_position, 0, QTableWidgetItem(name))
        start_combo = self._create_time_combobox()
        start_combo.setCurrentText(start_time)
        self.breaks_table.setCellWidget(row_position, 1, start_combo)
        end_combo = self._create_time_combobox()
        end_combo.setCurrentText(end_time)
        self.breaks_table.setCellWidget(row_position, 2, end_combo)
        paid_check = QCheckBox()
        paid_check.setChecked(is_paid)
        cell_widget = QWidget()
        cell_layout = QHBoxLayout(cell_widget)
        cell_layout.addWidget(paid_check)
        cell_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cell_layout.setContentsMargins(0, 0, 0, 0)
        self.breaks_table.setCellWidget(row_position, 3, cell_widget)
        remove_btn = QPushButton("Remove")
        # noinspection PyUnresolvedReferences
        remove_btn.clicked.connect(self._remove_break_row_clicked)
        self.breaks_table.setCellWidget(row_position, 4, remove_btn)

    def _remove_break_row_clicked(self):
        button = self.sender()
        if button:
            index = self.breaks_table.indexAt(button.pos())
            if index.isValid():
                self.breaks_table.removeRow(index.row())

    def _create_main_tab(self):
        main_tab = QWidget()
        layout = QVBoxLayout(main_tab)

        input_layout = QHBoxLayout()
        input_label = QLabel("Select Input File:")
        self.input_path_edit = QLineEdit()
        self.input_path_edit.setPlaceholderText("Click 'Browse' to select a CSV or XLSX file...")
        browse_btn = QPushButton("Browse...")
        # noinspection PyUnresolvedReferences
        browse_btn.clicked.connect(self._select_input_file)
        input_layout.addWidget(input_label)
        input_layout.addWidget(self.input_path_edit)
        input_layout.addWidget(browse_btn)

        layout.addLayout(input_layout)
        layout.addWidget(self._create_process_button())
        layout.addWidget(self._create_preview_group())
        return main_tab

    def _create_preview_group(self):
        group_box = QGroupBox("File Preview")
        layout = QVBoxLayout(group_box)
        self.preview_label = QLabel("No file selected.")
        self.preview_label.setStyleSheet("font-style: italic; color: #9c9a9a;")
        layout.addWidget(self.preview_label)
        self.preview_table = QTableWidget()
        self.preview_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.preview_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(self.preview_table)
        return group_box

    def _create_rounding_group(self):
        group_box = QGroupBox("Snap Clock-Out Times To")
        group_box.setToolTip(
            "If a punch is close to one of these defined times (e.g., 4:59 PM near 5:00 PM), it will be automatically snapped to it.\nThis is useful for standardizing end-of-shift times.")
        layout = QVBoxLayout(group_box)
        self.rounding_table = QTableWidget()
        self.rounding_table.setColumnCount(2)
        self.rounding_table.setHorizontalHeaderLabels(["Time", "Action"])
        self.rounding_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.rounding_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.rounding_table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self._add_rounding_row("04:00 PM")
        self._add_rounding_row("05:00 PM")
        self._add_rounding_row("06:00 PM")
        layout.addWidget(self.rounding_table)
        add_btn = QPushButton("Add Time")
        # noinspection PyUnresolvedReferences
        add_btn.clicked.connect(lambda: self._add_rounding_row())
        layout.addWidget(add_btn, alignment=Qt.AlignmentFlag.AlignRight)
        group_box.setLayout(layout)
        return group_box

    def _create_process_button(self):
        self.process_btn = QPushButton("Generate Timesheet Report")
        self.process_btn.setFont(QFont('Arial', 14, QFont.Weight.Bold))
        self.process_btn.setMinimumHeight(50)
        self.process_btn.setStyleSheet("""
            QPushButton { background-color: #dc3545; color: white; border: none; border-radius: 5px; padding: 5px; }
            QPushButton:hover { background-color: #c82333; }
            QPushButton:pressed { background-color: #bd2130; }
            QPushButton:disabled { background-color: #cccccc; color: #666666; }
        """)
        # noinspection PyUnresolvedReferences
        self.process_btn.clicked.connect(self._run_processing)
        return self.process_btn

    def _validate_parameters(self):
        errors = []
        start_hour = str_to_delta(self.start_hour_combo.currentText()) if self.start_hour_check.isChecked() else None
        end_hour = str_to_delta(self.end_hour_combo.currentText()) if self.end_hour_check.isChecked() else None
        first_in_thresh = str_to_delta(self.first_in_combo.currentText())
        last_out_thresh = str_to_delta(self.last_out_combo.currentText())
        if start_hour and end_hour and start_hour >= end_hour:
            errors.append("Paid Hours Start Time must be earlier than Paid Hours End Time.")
        if first_in_thresh >= last_out_thresh:
            errors.append("The 'Verify First Punch-In' time must be earlier than the 'Verify Last Punch-Out' time.")
        if start_hour:
            if first_in_thresh < start_hour: errors.append(
                "'Verify First Punch-In' time cannot be earlier than Paid Hours Start Time.")
            if last_out_thresh < start_hour: errors.append(
                "'Verify Last Punch-Out' time cannot be earlier than Paid Hours Start Time.")
        if end_hour:
            if first_in_thresh > end_hour: errors.append(
                "'Verify First Punch-In' time cannot be later than Paid Hours End Time.")
            if last_out_thresh > end_hour: errors.append(
                "'Verify Last Punch-Out' time cannot be later than Paid Hours End Time.")
        breaks_data, break_names = [], set()
        for row in range(self.breaks_table.rowCount()):
            name_item = self.breaks_table.item(row, 0)
            name = name_item.text().strip() if name_item else ""
            if not name:
                errors.append(f"Break at row {row + 1} must have a name.")
                continue
            if name in break_names:
                errors.append(f"Duplicate break name found: '{name}'. Names must be unique.")
            break_names.add(name)

            start_combo = self.breaks_table.cellWidget(row, 1)
            end_combo = self.breaks_table.cellWidget(row, 2)
            if isinstance(start_combo, QComboBox) and isinstance(end_combo, QComboBox):
                start = str_to_delta(start_combo.currentText())
                end = str_to_delta(end_combo.currentText())
            else:
                errors.append(f"Invalid widget in break table at row {row + 1}.")
                continue

            if start >= end:
                errors.append(f"For break '{name}', start time must be before end time.")
            if start_hour and start < start_hour:
                errors.append(f"Break '{name}' cannot start before the Paid Hours Start Time.")
            if end_hour and end > end_hour:
                errors.append(f"Break '{name}' cannot end after the Paid Hours End Time.")
            breaks_data.append({'name': name, 'start': start, 'end': end})
        for i in range(len(breaks_data)):
            for j in range(i + 1, len(breaks_data)):
                b1, b2 = breaks_data[i], breaks_data[j]
                if b1['start'] < b2['end'] and b2['start'] < b1['end']:
                    errors.append(f"Breaks '{b1['name']}' and '{b2['name']}' are overlapping.")
        return errors

    def _run_processing(self):
        self.log_edit.clear()
        input_file = self.input_path_edit.text()
        if not input_file:
            error_msg = "Please select an input file first."
            QMessageBox.critical(self, "Error", error_msg)
            self.log(f"Error: {error_msg}")
            return

        validation_errors = self._validate_parameters()
        if validation_errors:
            error_msg = "Please fix the following configuration errors:\n\n" + "\n".join(
                f"• {e}" for e in validation_errors)
            QMessageBox.critical(self, "Invalid Settings", error_msg)
            self.log("Validation failed. Please check settings.")
            return

        output_file, _ = QFileDialog.getSaveFileName(self, "Save Timesheet Report As", "", "Excel Files (*.xlsx)")
        if not output_file:
            self.log("Process cancelled by user.")
            return

        self.log("Settings validated. Starting process...")

        try:
            buffer = timedelta(minutes=self.buffer_spinbox.value())
            start_hour = str_to_delta(
                self.start_hour_combo.currentText()) if self.start_hour_check.isChecked() else None
            end_hour = str_to_delta(self.end_hour_combo.currentText()) if self.end_hour_check.isChecked() else None
            first_in_thresh = str_to_delta(self.first_in_combo.currentText())
            last_out_thresh = str_to_delta(self.last_out_combo.currentText())
            break_time = {}
            for row in range(self.breaks_table.rowCount()):
                name = self.breaks_table.item(row, 0).text().strip()
                start_combo = self.breaks_table.cellWidget(row, 1)
                end_combo = self.breaks_table.cellWidget(row, 2)
                paid_cell = self.breaks_table.cellWidget(row, 3)
                start_str = start_combo.currentText() if isinstance(start_combo, QComboBox) else ""
                end_str = end_combo.currentText() if isinstance(end_combo, QComboBox) else ""
                is_paid = paid_cell.findChild(QCheckBox).isChecked() if paid_cell else False
                break_time[name] = {'start': start_str, 'end': end_str, 'paid': is_paid}
            round_to_str = []
            for row in range(self.rounding_table.rowCount()):
                combo = self.rounding_table.cellWidget(row, 0)
                if isinstance(combo, QComboBox): round_to_str.append(combo.currentText())
            round_to_delta = [str_to_delta(t) for t in set(round_to_str)]
        except Exception as e:
            error_msg = f"Could not gather parameters: {e}"
            QMessageBox.critical(self, "Parameter Error", error_msg)
            self.log(f"Error: {error_msg}")
            return

        self.process_btn.setEnabled(False)
        self.process_btn.setText("Processing...")
        QApplication.processEvents()

        try:
            from timesheet import read_input_file
            self.log("Reading and validating input file...")
            QApplication.processEvents()
            df_full, logs = read_input_file(input_file)
            for msg in logs:
                self.log(msg)

            self.log("Input file loaded successfully. Starting main process...")
            final_filename = process_timesheet(
                df=df_full, buffer=buffer, start_hour=start_hour, end_hour=end_hour,
                break_time=break_time, first_in_thresh=first_in_thresh,
                last_out_thresh=last_out_thresh, round_to=round_to_delta,
                output_filename=output_file
            )

            if final_filename:
                success_msg = f"Processing complete! Report saved to:\n{final_filename}"
                QMessageBox.information(self, "Success", success_msg)
                self.log(success_msg.replace('\n', ' '))
            else:
                error_msg = "Could not write to the output file. Please ensure it's not open elsewhere."
                QMessageBox.critical(self, "File Error", error_msg)
                self.log(f"Error: {error_msg}")

        except (FileNotFoundError, ValueError) as e:
            error_msg = f"Failed to read or process the input file:\n\n{e}"
            QMessageBox.critical(self, "File or Data Error", error_msg)
            self.log(f"ERROR: {str(e).replace('\n', ' ')}")
        except Exception as e:
            error_msg = f"An unexpected error occurred during processing:\n\n{type(e).__name__}: {e}"
            QMessageBox.critical(self, "Processing Error", error_msg)
            self.log(f"FATAL ERROR: {error_msg.replace('\n', ' ')}")
        finally:
            self.process_btn.setEnabled(True)
            self.process_btn.setText("Generate Timesheet Report")

    @staticmethod
    def _create_time_combobox():
        combo = ConstrainedComboBox()
        combo.setMaxVisibleItems(12)
        for i in range(48):
            time = datetime.strptime("12:00 AM", "%I:%M %p") + timedelta(minutes=30 * i)
            combo.addItem(time.strftime("%I:%M %p"))
        return combo

    @staticmethod
    def _create_two_widget_layout(widget1, widget2):
        layout = QHBoxLayout()
        layout.addWidget(widget1)
        layout.addWidget(widget2)
        return layout

    def _add_rounding_row(self, time="05:00 PM"):
        row_position = self.rounding_table.rowCount()
        self.rounding_table.insertRow(row_position)
        time_combo = self._create_time_combobox()
        time_combo.setCurrentText(time)
        self.rounding_table.setCellWidget(row_position, 0, time_combo)
        remove_btn = QPushButton("Remove")
        # noinspection PyUnresolvedReferences
        remove_btn.clicked.connect(self._remove_rounding_row_clicked)
        self.rounding_table.setCellWidget(row_position, 1, remove_btn)

    def _remove_rounding_row_clicked(self):
        button = self.sender()
        if button:
            index = self.rounding_table.indexAt(button.pos())
            if index.isValid():
                self.rounding_table.removeRow(index.row())

    def _select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Timesheet File", "",
            "Timesheet Files (*.csv *.xlsx *.xls);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            self.input_path_edit.setText(file_path)
            self.log(f"Selected input file: {file_path}")
            self._update_preview(file_path)

    def _update_preview(self, file_path):
        try:
            path = Path(file_path)
            file_suffix = path.suffix.lower()

            if file_suffix == '.csv':
                df = pd.read_csv(file_path, nrows=100)
            elif file_suffix in ['.xlsx', '.xls']:
                df = pd.read_excel(file_path, nrows=100, engine='openpyxl')
            else:
                raise ValueError("Unsupported file type for preview.")

            self.preview_label.setText(f"Previewing first 100 rows from: <strong>{path.name}</strong>")
            self.preview_table.clear()
            self.preview_table.setRowCount(df.shape[0])
            self.preview_table.setColumnCount(df.shape[1])
            self.preview_table.setHorizontalHeaderLabels(df.columns)
            for r_idx, row in enumerate(df.itertuples(index=False)):
                for c_idx, val in enumerate(row):
                    self.preview_table.setItem(r_idx, c_idx, QTableWidgetItem(str(val)))
            self.preview_table.resizeColumnsToContents()
        except Exception as e:
            path = Path(file_path)
            self.preview_label.setText(f"Could not load preview for: <strong>{path.name}</strong>")
            self.preview_table.clear()
            self.preview_table.setRowCount(1)
            self.preview_table.setColumnCount(1)
            self.preview_table.setHorizontalHeaderLabels(["Error"])
            self.preview_table.setItem(0, 0, QTableWidgetItem(f"Could not preview file: {e}"))

    def _create_log_box(self):
        group_box = QGroupBox("Logs")
        layout = QVBoxLayout()
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        log_font = QFont()
        log_font.setPointSize(9)
        self.log_edit.setFont(log_font)
        self.log_edit.setMaximumHeight(80)
        layout.addWidget(self.log_edit)
        group_box.setLayout(layout)
        return group_box

    def log(self, message):
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_edit.append(f"{timestamp} - {message}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    if "Fusion" in QStyleFactory.keys():
        app.setStyle("Fusion")
    ex = TimesheetApp()
    sys.exit(app.exec())