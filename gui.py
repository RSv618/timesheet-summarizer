import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QLabel, QLineEdit, QGroupBox, QSpinBox, QCheckBox, QComboBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QTextEdit,
    QTabWidget, QStyleFactory, QSizePolicy
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon
from pathlib import Path
from datetime import datetime, timedelta

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


class TimesheetApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        script_dir = Path(__file__).parent
        icon_path = script_dir / 'logo.ico'
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        else:
            print(f"Warning: Icon file not found at '{icon_path}'.")

        self.setWindowTitle('Timesheet Processor')
        self.setGeometry(100, 100, 750, 700)

        main_layout = QVBoxLayout(self)

        tabs = QTabWidget()
        tabs.addTab(self._create_main_tab(), "Start Here")
        tabs.addTab(self._create_general_settings_tab(), "General Settings")
        tabs.addTab(self._create_advanced_settings_tab(), "Advanced Settings")

        main_layout.addWidget(tabs)
        main_layout.addWidget(self._create_log_box())
        self.show()

    def _create_general_settings_tab(self):
        settings_tab = QWidget()
        layout = QVBoxLayout(settings_tab)
        layout.addWidget(self._create_general_settings_group())
        restore_btn = QPushButton("Restore Default Settings")
        restore_btn.setToolTip("Resets all settings across all tabs to their original values.")
        # noinspection PyUnresolvedReferences
        restore_btn.clicked.connect(self._restore_default_settings)
        layout.addStretch()
        layout.addWidget(restore_btn, alignment=Qt.AlignmentFlag.AlignRight)
        return settings_tab

    def _restore_default_settings(self):
        self.start_hour_check.setChecked(True)
        self.start_hour_combo.setCurrentText("07:00 AM")
        self.end_hour_check.setChecked(True)
        self.end_hour_combo.setCurrentText("10:00 PM")
        self.buffer_spinbox.setValue(15)
        self.first_in_combo.setCurrentText("10:30 AM")
        self.last_out_combo.setCurrentText("02:30 PM")
        self.breaks_table.setRowCount(0)
        self._add_break_row("Lunch", "12:00 PM", "01:00 PM", False)
        self._add_break_row("Dinner", "06:00 PM", "06:30 PM", True)
        self.rounding_table.setRowCount(0)
        self._add_rounding_row("04:00 PM")
        self._add_rounding_row("05:00 PM")
        self._add_rounding_row("06:00 PM")
        self.log("All settings have been restored to their default values.")
        QMessageBox.information(self, "Settings Restored", "All settings have been restored to their default values.")

    def _create_general_settings_group(self):
        group_box = QGroupBox("Time & Threshold Settings")
        layout = QVBoxLayout()
        self.start_hour_check = QCheckBox("Set Official Start Hour")
        self.start_hour_check.setToolTip(
            "If checked, any punch before this time (within the buffer) will be snapped to this time.\nUseful for standardizing early clock-ins.")
        self.start_hour_combo = self._create_time_combobox()
        self.start_hour_combo.setCurrentText("07:00 AM")
        layout.addLayout(self._create_two_widget_layout(self.start_hour_check, self.start_hour_combo))
        self.end_hour_check = QCheckBox("Set Official End Hour")
        self.end_hour_check.setToolTip(
            "If checked, any punch after this time (within the buffer) will be snapped to this time.\nUseful for standardizing late clock-outs.")
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
        buffer_label = QLabel("Buffer Period (minutes):")
        buffer_label.setToolTip(
            "A grace period around event times.\nFor example, a buffer of 15 minutes means a 6:50 AM punch is considered 'at' 7:00 AM.")
        self.buffer_spinbox = QSpinBox()
        self.buffer_spinbox.setRange(1, 30)
        self.buffer_spinbox.setValue(15)
        layout.addLayout(self._create_two_widget_layout(buffer_label, self.buffer_spinbox))
        first_in_label = QLabel("First-In Threshold:")
        first_in_label.setToolTip(
            "Helps the program guess correctly.\nAny 'out' punch before this time is likely a mistaken first 'in' punch and will be corrected.")
        self.first_in_combo = self._create_time_combobox()
        self.first_in_combo.setCurrentText("10:30 AM")
        layout.addLayout(self._create_two_widget_layout(first_in_label, self.first_in_combo))
        last_out_label = QLabel("Last-Out Threshold:")
        last_out_label.setToolTip(
            "Helps the program guess correctly.\nAny 'in' punch after this time is likely a mistaken last 'out' punch and will be corrected.")
        self.last_out_combo = self._create_time_combobox()
        self.last_out_combo.setCurrentText("02:30 PM")
        layout.addLayout(self._create_two_widget_layout(last_out_label, self.last_out_combo))
        group_box.setLayout(layout)
        return group_box

    def _create_breaks_group(self):
        group_box = QGroupBox("Break Times")
        group_box.setToolTip(
            "Define unpaid or paid breaks.\nThe script will auto-insert missing punches for unpaid breaks\nand clean up punches during paid breaks.")
        layout = QVBoxLayout(group_box)
        self.breaks_table = QTableWidget()
        self.breaks_table.setColumnCount(5)
        self.breaks_table.setHorizontalHeaderLabels(["Name", "Start Time", "End Time", "Is Paid?", "Action"])
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
        layout.addWidget(self._create_file_io_group())
        layout.addWidget(self._create_preview_group())
        return main_tab

    def _create_file_io_group(self):
        group_box = QGroupBox("Workflow")
        layout = QVBoxLayout(group_box)
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
        return group_box

    def _create_preview_group(self):
        group_box = QGroupBox("File Preview")
        layout = QVBoxLayout(group_box)
        self.preview_table = QTableWidget()
        self.preview_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.preview_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(self.preview_table)
        return group_box

    def _create_advanced_settings_tab(self):
        advanced_tab = QWidget()
        layout = QVBoxLayout(advanced_tab)
        layout.addWidget(self._create_breaks_group())
        layout.addWidget(self._create_rounding_group())
        return advanced_tab

    def _create_rounding_group(self):
        group_box = QGroupBox("Round Punch Times To")
        group_box.setToolTip(
            "Define specific end-of-day times.\nPunches near these times will be 'snapped' to them.\nUseful for standardizing shifts that end at 4, 5, or 6 PM.")
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
        self.process_btn = QPushButton("Create Summary")
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
            errors.append("Start Hour must be earlier than End Hour.")
        if first_in_thresh >= last_out_thresh:
            errors.append("First-In Threshold must be earlier than Last-Out Threshold.")
        if start_hour:
            if first_in_thresh < start_hour: errors.append("First-In Threshold cannot be earlier than Start Hour.")
            if last_out_thresh < start_hour: errors.append("Last-Out Threshold cannot be earlier than Start Hour.")
        if end_hour:
            if first_in_thresh > end_hour: errors.append("First-In Threshold cannot be later than End Hour.")
            if last_out_thresh > end_hour: errors.append("Last-Out Threshold cannot be later than End Hour.")
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

            # Safely get widget and its text
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
                errors.append(f"Break '{name}' cannot start before the official Start Hour.")
            if end_hour and end > end_hour:
                errors.append(f"Break '{name}' cannot end after the official End Hour.")
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

        output_file, _ = QFileDialog.getSaveFileName(self, "Save Summary As", "", "Excel Files (*.xlsx)")
        if not output_file:
            self.log("Process cancelled by user.")
            return

        self.log("Settings validated. Starting process...")

        try:
            # --- Parameter Gathering (this part is unchanged) ---
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
                start_str, end_str = "", ""
                if isinstance(start_combo, QComboBox): start_str = start_combo.currentText()
                if isinstance(end_combo, QComboBox): end_str = end_combo.currentText()
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
                success_msg = f"Processing complete! Summary saved to:\n{final_filename}"
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
            self.process_btn.setText("Create Summary")

    @staticmethod
    def _create_time_combobox():
        combo = QComboBox()
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
            self._update_preview(file_path)

    def _update_preview(self, file_path):
        try:
            path = Path(file_path)
            file_suffix = path.suffix.lower()

            if file_suffix == '.csv':
                df = pd.read_csv(file_path, nrows=100)
            elif file_suffix in ['.xlsx', '.xls']:
                # For preview, just read the first sheet to be fast.
                df = pd.read_excel(file_path, nrows=100, engine='openpyxl')
            else:
                raise ValueError("Unsupported file type for preview.")

            self.preview_table.clear()
            self.preview_table.setRowCount(df.shape[0])
            self.preview_table.setColumnCount(df.shape[1])
            self.preview_table.setHorizontalHeaderLabels(df.columns)
            for r_idx, row in enumerate(df.itertuples(index=False)):
                for c_idx, val in enumerate(row):
                    self.preview_table.setItem(r_idx, c_idx, QTableWidgetItem(str(val)))
            self.preview_table.resizeColumnsToContents()
        except Exception as e:
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
        self.log_edit.append(message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    if "Fusion" in QStyleFactory.keys():
        app.setStyle("Fusion")
    ex = TimesheetApp()
    sys.exit(app.exec())