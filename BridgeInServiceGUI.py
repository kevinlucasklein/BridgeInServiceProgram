import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLineEdit,
                             QPushButton, QLabel, QFormLayout, QDateEdit, QTextEdit)
from PyQt5.QtCore import QDate, QRegExp, Qt
from PyQt5.QtGui import QRegExpValidator, QPixmap, QIcon
from PyQt5.QtGui import QFontDatabase, QFont
from bridge_in_service_WIP_3 import Employee, Calculation, DateOperations, Verification, Calculation, ExcelExport, Email
import datetime
from dateutil.relativedelta import relativedelta
import os
from decimal import Decimal, getcontext

class EmployeeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.employment_periods = []
        self.fte_changes = []
        self.employee = None
        self.initUI()
        self.applyStyle()

    def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(base_path, relative_path)
        return path.replace('\\', '/')  # Replace backslashes with forward slashes for URL compatibility


    def applyStyle(self):
        down_arrow_path = EmployeeApp.resource_path('downarrow.png')
        stylesheet = """
            QMainWindow {{
                background-color: #ffffff;
            }}
            QLabel {{
                font-size: 14px;
                color: #333333;
                padding: 2px;
            }}
            QLineEdit, QDateEdit {{
                border: 1px solid #cccccc;
                border-radius: 10px;
                padding: 5px;
                background: #ffffff;
                selection-background-color: #e5e5e5;
                font-size: 14px;  /* Increased font size */
            }}
            QDateEdit::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: #cccccc;
                border-left-style: solid;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }}
            QDateEdit::down-arrow {{
                image: url('{0}');
                width: 16px;
                height: 16px;
            }}
            QPushButton {{
                color: #ffffff;
                background-color: #0078d4;
                border-style: none;
                padding: 6px 12px;
                border-radius: 10px;
                font-size: 14px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #005fa3;
            }}
            QPushButton:pressed {{
                background-color: #003970;
            }}
            QTextEdit, QScrollBar:vertical {{
                border: 1px solid #cccccc;
                border-radius: 10px;
                background: #ffffff;
            }}
            QScrollBar::handle:vertical {{
                background: #cccccc;
                min-height: 20px;
                border-radius: 5px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: #aaaaaa;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                background: none;
            }}
        """.format(down_arrow_path.replace('\\', '/'))  # Ensure path is appropriate for URL
        self.setStyleSheet(stylesheet)




    def initUI(self):
        icon_path = EmployeeApp.resource_path('icon.jpg')
        self.setWindowTitle('Bridge In Service Calculator')
        self.setWindowIcon(QIcon(icon_path))
        self.setGeometry(100, 100, 800, 600)
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)


        self.form_layout = QFormLayout()
        self.employee_id_input = QLineEdit()
        employee_id_regex = QRegExp("^[0-9]{8}$")  # Regular expression for exactly 8 digits
        employee_id_validator = QRegExpValidator(employee_id_regex, self.employee_id_input)
        self.employee_id_input.setValidator(employee_id_validator)
        self.employee_id_input.textChanged.connect(self.validate_employee_id)
        self.first_name_input = QLineEdit()
        self.last_name_input = QLineEdit()
        self.date_and_fte_layout = QHBoxLayout()
        self.most_recent_start_date_input = QDateEdit()
        self.most_recent_start_date_input.setCalendarPopup(True)
        self.most_recent_start_date_input.setDisplayFormat("MM/dd/yyyy")
        self.most_recent_start_date_input.setMinimumSize(120, 32) # Set a familiar format, adjust as needed
        six_months_ago = datetime.datetime.today() - relativedelta(months=6)  # Accurately move six months back
        self.most_recent_start_date_input.setDate(QDate(six_months_ago.year, six_months_ago.month, six_months_ago.day))
        self.most_recent_start_date_input.dateChanged.connect(self.update_fte_change_dates)
        self.fte_input = QLineEdit()
        fte_regex = QRegExp("^1(.0)?|\.75|\.76|\.77|\.78|\.79|\.8[0-9]|\.9[0-9]|0?\.[7-9][0-9]?$")
        fte_validator = QRegExpValidator(fte_regex, self.fte_input)
        self.fte_input.setValidator(fte_validator)
        self.fte_input.textChanged.connect(self.validate_fte)
        self.date_and_fte_layout.addWidget(self.most_recent_start_date_input)
        self.date_and_fte_layout.addWidget(self.fte_input)
        self.form_layout.addRow('Employee ID:', self.employee_id_input)
        self.form_layout.addRow('First Name:', self.first_name_input)
        self.form_layout.addRow('Last Name:', self.last_name_input)
        self.form_layout.addRow('Most Recent Start Date and FTE:', self.date_and_fte_layout)
        
        # Add the form layout to the main layout
        self.layout.addLayout(self.form_layout)

        # Employment periods and FTE changes sections should be added next
        self.periods_layout = QVBoxLayout()
        self.employment_periods_title = QLabel("Employment Periods")
        self.employment_periods_title.setStyleSheet("font-weight: bold; font-size: 14px")
        self.employment_periods_title.setVisible(False)  # Initially hidden
        self.employment_periods_title.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.employment_periods_title)
        
        
        self.add_period_button = QPushButton("Add Employment Period")
        self.add_period_button.clicked.connect(self.add_employment_period)
        self.periods_layout.addWidget(self.add_period_button)
        self.layout.addLayout(self.periods_layout)

        self.fte_changes_layout = QVBoxLayout()
        self.fte_changes_title = QLabel("FTE Changes")
        self.fte_changes_title.setStyleSheet("font-weight: bold; font-size: 14px")  # Styling the title
        self.fte_changes_title.setVisible(False)  # Hide the title until the first change is added
        self.fte_changes_title.setAlignment(Qt.AlignCenter)
        self.fte_changes_layout.addWidget(self.fte_changes_title)
        self.add_fte_change_button = QPushButton("Add FTE Change")
        self.add_fte_change_button.clicked.connect(self.add_fte_change)
        self.fte_changes_layout.addWidget(self.add_fte_change_button)
        self.layout.addLayout(self.fte_changes_layout)

        self.export_to_excel_button = QPushButton("Export to Excel")
        self.export_to_excel_button.clicked.connect(self.export_to_excel)
        self.layout.addWidget(self.export_to_excel_button)

        self.send_email_button = QPushButton("Send Email")
        self.send_email_button.clicked.connect(self.send_email)
        self.layout.addWidget(self.send_email_button)

        # Submit button
        self.submit_button = QPushButton('Submit')
        self.submit_button.clicked.connect(self.submit_data)
        self.layout.addWidget(self.submit_button)

        # Create a text edit field for displaying results
        self.result_display = QTextEdit()
        self.result_display.setReadOnly(True)
        self.layout.addWidget(self.result_display)

    def add_employment_period(self):
        if not self.employment_periods_title.isVisible():
            self.employment_periods_title.setVisible(True)
        start_date = QDateEdit()
        start_date.setCalendarPopup(True)
        end_date = QDateEdit()
        end_date.setCalendarPopup(True)
        remove_button = QPushButton('Delete')
        period_layout = QHBoxLayout()
        period_layout.addWidget(start_date)
        period_layout.addWidget(end_date)
        period_layout.addWidget(remove_button)
        period_entry = {'start': start_date, 'end': end_date, 'layout': period_layout, 'button': remove_button}
        self.employment_periods.append(period_entry)
        self.periods_layout.insertLayout(self.periods_layout.count() - 1, period_layout)
        remove_button.clicked.connect(lambda: self.remove_layout(period_entry, self.employment_periods, self.periods_layout))

    def add_fte_change(self):
        if not self.fte_changes_title.isVisible():
            self.fte_changes_title.setVisible(True)

        change_date = QDateEdit()
        change_date.setCalendarPopup(True)
        most_recent_date = self.most_recent_start_date_input.date()
        change_date.setDate(most_recent_date)
        change_date.setMinimumDate(most_recent_date)

        new_fte = QLineEdit()
        fte_regex = QRegExp("^1(.0)?|\.75|\.76|\.77|\.78|\.79|\.8[0-9]|\.9[0-9]|0?\.[7-9][0-9]?$")
        fte_validator = QRegExpValidator(fte_regex, new_fte)
        new_fte.setValidator(fte_validator)
        new_fte.textChanged.connect(self.validate_fte)

        remove_button = QPushButton('Delete')
        fte_layout = QHBoxLayout()
        fte_layout.addWidget(change_date)
        fte_layout.addWidget(new_fte)
        fte_layout.addWidget(remove_button)

        fte_change_entry = {'date': change_date, 'fte': new_fte, 'layout': fte_layout, 'button': remove_button}
        self.fte_changes.append(fte_change_entry)
        self.fte_changes_layout.insertLayout(self.fte_changes_layout.count() - 1, fte_layout)
        remove_button.clicked.connect(lambda: self.remove_layout(fte_change_entry, self.fte_changes, self.fte_changes_layout))


    def get_latest_fte_change_date(self):
        """Return the maximum date set in FTE changes or the most recent start date if no FTE changes exist."""
        if self.fte_changes:
            return max(entry['date'].date() for entry in self.fte_changes)
        else:
            return self.most_recent_start_date_input.date()
        
    def set_fte_change_date(self, change_date):
        most_recent_date = self.most_recent_start_date_input.date()
        change_date.setDate(most_recent_date)
        change_date.setMinimumDate(most_recent_date)

    def update_fte_change_dates(self):
        """Update all FTE change dates to reflect the most recent start date or to ensure correct order."""
        most_recent_date = self.most_recent_start_date_input.date()
        if self.fte_changes:
            for entry in self.fte_changes:
                change_date = entry['date']
                current_date = change_date.date()
                change_date.setMinimumDate(most_recent_date)
                if current_date < most_recent_date:
                    change_date.setDate(most_recent_date)  # Update the date if it's earlier than the most recent start date



    def remove_layout(self, entry, list_reference, layout_reference):
        # Deleting the widgets from the layout and the list
        entry.get('start', entry.get('date')).deleteLater()
        entry.get('end', entry.get('fte')).deleteLater()
        entry['button'].deleteLater()
        layout_reference.removeItem(entry['layout'])
        list_reference.remove(entry)

        # Check if the list is empty after removal and adjust the title visibility specifically
        if not list_reference:
            if list_reference is self.employment_periods:
                self.employment_periods_title.setVisible(False)  # Hide the employment periods title if no periods are left
            elif list_reference is self.fte_changes:
                self.fte_changes_title.setVisible(False)  # Hide the FTE changes title if no changes are left

        # Recompute layout after removal to ensure updates are visible
        layout_reference.update()

    def submit_data(self):
        getcontext().prec = 28
        employee_id = self.employee_id_input.text()
        first_name = self.first_name_input.text()
        last_name = self.last_name_input.text()
        most_recent_start_date = self.most_recent_start_date_input.date().toString("MM/dd/yyyy")
        fte = self.fte_input.text()

        # Validate data
        valid_id, id_msg = Verification.verify_employee_id(employee_id)
        valid_date, date_msg = Verification.verify_most_recent_start_date(most_recent_start_date)
        valid_fte, fte_msg = Verification.verify_employee_fte(fte)
        
        if not valid_id or not valid_date or not valid_fte:
            validation_message = "\n".join([id_msg, date_msg, fte_msg])
            self.result_display.setText("Validation Failed:\n" + validation_message)
            return

        # Create Employee instance if validations pass
        dt_most_recent_start_date = DateOperations.convert_to_datetime(most_recent_start_date)
        self.employee = Employee(employee_id, first_name, last_name, dt_most_recent_start_date, float(fte))

        # Handle employment periods from the GUI
        for period in self.employment_periods:
            start = period['start'].date().toString("MM/dd/yyyy")
            end = period['end'].date().toString("MM/dd/yyyy")
            self.employee.add_employment_period(DateOperations.convert_to_datetime(start), DateOperations.convert_to_datetime(end))

        # Handle FTE changes from the GUI, ensuring no empty or invalid entries are used
        for change in self.fte_changes:
            change_date = change['date'].date().toString("MM/dd/yyyy")
            new_fte = change['fte'].text().strip()
            if new_fte:  # Check if the FTE field is not empty
                is_valid, _ = Verification.verify_employee_fte(new_fte)  # Optionally check for validity
                if is_valid:
                    self.employee.add_fte_change(DateOperations.convert_to_datetime(change_date), float(new_fte))

        # Perform calculations
        bridge_in_service_date = Calculation.calculate_bridge_in_service_date(self.employee)
        original_pto_accrued, original_monthly_accruals = Calculation.calculate_pto_accrual_rate(self.employee)
        bridge_pto_accrued, bridge_monthly_accruals = Calculation.calculate_bridge_pto_accrual_rate(self.employee)
        accrual_differences = Calculation.calculate_accrual_differences(original_monthly_accruals, bridge_monthly_accruals)
        # Store accrual details in the employee object
        self.employee.original_monthly_accruals = original_monthly_accruals
        self.employee.bridge_monthly_accruals = bridge_monthly_accruals
        self.employee.accrual_differences = accrual_differences

        # Accumulate totals
        # Convert strings to Decimals and sum them up
        total_original = sum(Decimal(details.split()[0]) for details in original_monthly_accruals.values())
        total_bridge = sum(Decimal(details.split()[0]) for details in bridge_monthly_accruals.values())
        total_difference = (total_bridge - total_original).quantize(Decimal('0.00'))
        self.employee.update_pto_accrual_difference(total_difference)

        # Prepare results display
        result_text = f"Processed data for {first_name} {last_name}:<br>"
        result_text += f"Bridge in Service Date: {bridge_in_service_date.strftime('%m/%d/%Y')}<br>"
        result_text += f"PTO to Add: {total_difference:.2f} hours<br><br>"
        result_text += "<table border='1'><tr><th>Month Year</th><th>Original</th><th>Bridge</th><th>Difference</th></tr>"

        for month in accrual_differences:
            result_text += f"<tr><td>{month}</td><td>{original_monthly_accruals.get(month, '0.00 Hours')}</td><td>{bridge_monthly_accruals.get(month, '0.00 Hours')}</td><td>{accrual_differences[month]}</td></tr>"

        result_text += f"<tr style='font-weight:bold;'><td>Total</td><td>{total_original:.2f} Hours</td><td>{total_bridge:.2f} Hours</td><td>{total_difference:.2f} Hours</td></tr>"
        result_text += "</table>"
        self.result_display.setHtml(result_text)

    def export_to_excel(self):
        if self.employee:
            # Assume directory and employee setup already provided
            success = ExcelExport.try_export_employee_data(
                self.employee,
                "C:\\Hospital HR\\Operations\\VOE (Letters, Completed VOEs, etc)\\Bridge in Service",
                self.employee.original_monthly_accruals,
                self.employee.bridge_monthly_accruals,
                self.employee.accrual_differences
            )
            result_msg = 'Excel file exported successfully.' if success else 'Failed to export to Excel.'
            current_text = self.result_display.toPlainText()
            self.result_display.setText(current_text + "\n" + result_msg)

    def send_email(self):
        if self.employee:
            success = Email.try_send_email(self.employee)
            result_msg = 'Email opened successfully.' if success else 'Failed to open email.'
            current_text = self.result_display.toPlainText()
            self.result_display.setText(current_text + "\n" + result_msg)

    def validate_employee_id(self, text):
        if not text:  # Check if the text field is empty
            self.employee_id_input.setStyleSheet("")  # Reset style to default
            self.result_display.setText("")  # Clear any previous error messages
        else:
            is_valid, message = Verification.verify_employee_id(text)  # Use the backend validation logic
            if is_valid:
                self.employee_id_input.setStyleSheet("color: black;")  # Valid input
                self.result_display.setText("")  # Clear any previous error messages
            else:
                self.employee_id_input.setStyleSheet("color: red;")  # Invalid input, show user with red text
                self.result_display.setText(message)  # Display the validation message from the backend

    def validate_fte(self, text):
        if not text:  # Check if the text field is empty
            self.fte_input.setStyleSheet("")
            self.result_display.setText("")
        else:
            is_valid, message = Verification.verify_employee_fte(text)
            if is_valid:
                self.fte_input.setStyleSheet("color: black;")
                self.result_display.setText("")
            else:
                self.fte_input.setStyleSheet("color: red;")
                self.result_display.setText(message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    id = QFontDatabase.addApplicationFont(EmployeeApp.resource_path('Roboto-Regular.ttf'))
    if id != -1:
        font_family = QFontDatabase.applicationFontFamilies(id).at(0)
        app_font = QFont(font_family)
        app.setFont(app_font)
    ex = EmployeeApp()
    ex.show()
    sys.exit(app.exec_())
