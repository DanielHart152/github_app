import sys
import pdfplumber
import re
import csv
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTextEdit,
    QFileDialog, QLabel, QGroupBox
)
from PyQt6.QtGui import QIcon, QFont, QPixmap
from PyQt6.QtCore import Qt

import platform
import subprocess


def extract_owner_claim_from_page(pdf, page_num=9): 
    page = pdf.pages[page_num]
    text = page.extract_text()
    if not text:
        return {}

    values = {}
    for line in text.split('\n'):
        if 'Owner:' in line:
            parts = line.split('Owner:', 1)
            values['Owner'] = parts[1].strip() if len(parts) > 1 else None
        if 'Claim:' in line:
            parts = line.split('Claim:', 1)
            values['Claim Reference'] = parts[1].strip() if len(parts) > 1 else None

    return values

def extract_valuation_summary(pdf):
    page = pdf.pages[0]
    text = page.extract_text()
    if not text:
        return {}

    values = {}
    for line in text.split('\n'):
        if 'Odometer' in line:
            m = re.search(r'Odometer\s*([\d,]+)', line)
            if m:
                values['Odometer'] = m.group(1)
        if 'Loss Incident Date' in line:
            m = re.search(r'Loss Incident Date\s*[:\-]?\s*([\d/]+)', line)
            if m:
                values['Loss Incident Date'] = m.group(1)

    combined_text = text.replace('\n', ' ')
    patterns = {
        'Base Vehicle Value': r'Base Vehicle Value\s*\$\s*([\d,]+\.\d{2})',
        'Adjusted Vehicle Value': r'Adjusted Vehicle Value\s*\$\s*([\d,]+\.\d{2})',
        'Vehicular Tax (6.625%)': r'Vehicular Tax.*?\+\s*\$\s*([\d,]+\.\d{2})',
        'Value before Deductible': r'Value before Deductible\s*\$\s*([\d,]+\.\d{2})',
        'Deductible': r'Deductible\*?\s*-\s*\$\s*([\d,]+\.\d{2})',
        'Total': r'Total\s*\$\s*([\d,]+\.\d{2})',
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, combined_text)
        if match:
            values[key] = match.group(1)

    return values

def extract_vehicle_info_clean(pdf, start_page=2, max_pages=6):
    import re
    fields = {
        'Year': r'Year\s*[:\-]?\s*(\d{4})',
        'Make': r'Make\s*[:\-]?\s*([A-Za-z]+)',
        'Model': r'Model\s*[:\-]?\s*([\w\s\-]+)',
        'VIN': r'VIN\s*[:\-]?\s*([\w\d]+)',
        'Trim': r'Trim\s*[:\-]?\s*([\w]+)',
        'Cylinders': r'Cylinders\s*[:\-]?\s*(\d+)',
        'Displacement': r'Displacement\s*[:\-]?\s*([\w\.]+)',
        'Induction': r'Induction\s*[:\-]?\s*([\w]+)',
        'Fuel Type': r'Fuel Type\s*[:\-]?\s*([\w]+)',
        'Carburation': r'Carburation\s*[:\-]?\s*([\w]+)',
        'Transmission': r'Transmission\s*[:\-]?\s*([\w\s]+)',
        'Location': r'Location\s*[:\-]?\s*([A-Z\s,0-9\-]+)',
    }

    found = {}
    remaining = set(fields.keys())

    for page_num in range(start_page, min(len(pdf.pages), start_page + max_pages)):
        text = pdf.pages[page_num].extract_text()
        if not text:
            continue
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        for line in lines:
            for key, pattern in fields.items():
                if key in remaining:
                    m = re.match(pattern, line, re.IGNORECASE)
                    if m:
                        if key == 'Location':
                            loc = m.group(1).strip()
                            loc_match = re.match(r'^.*\d{5}(-\d{4})?', loc)
                            loc = loc_match.group(0) if loc_match else loc
                            found[key] = loc
                        else:
                            found[key] = m.group(1).strip()
                        remaining.discard(key)
            if not remaining:
                break
        if not remaining:
            break

    return found 

def extract_odometer_values(pdf, start_page=7, end_page=15):
    all_odometer_values = []
    for page_number in range(start_page, end_page + 1):
        page = pdf.pages[page_number]
        text = page.extract_text()
        if not text:
            continue
        lines = text.split('\n')
        for line in lines:
            if line.startswith("Odometer"):
                values_str = line.replace("Odometer", "").strip()
                odometer_values_str = values_str.split()
                odometer_values = [int(val.replace(",", "")) for val in odometer_values_str if val.replace(",", "").isdigit()]
                all_odometer_values.extend(odometer_values)
                break
    return all_odometer_values if all_odometer_values else None

def extract_list_prices_comps(pdf, start_page=8, end_page=15):
    list_prices = []
    for page_num in range(start_page, end_page + 1):
        page = pdf.pages[page_num]
        text = page.extract_text()
        if not text:
            continue
        for line in text.split('\n'):
            if line.strip().lower().startswith("list price"):
                prices = re.findall(r'\$\s?[\d,]+', line)
                for price in prices:
                    price_val = price.replace('$', '').replace(' ', '').replace(',', '')
                    try:
                        list_prices.append(float(price_val))
                    except:
                        pass
    return list_prices

def extract_adjusted_comparable_values(pdf, start_page=8, end_page=15):
    adjusted_values = []
    for page_num in range(start_page, end_page + 1):
        page = pdf.pages[page_num]
        text = page.extract_text()
        if not text:
            continue
        for line in text.split('\n'):
            if "adjusted comparable value" in line.lower():
                prices = re.findall(r'\$\s?[\d,]+', line)
                for price in prices:
                    price_val = price.replace('$', '').replace(' ', '').replace(',', '')
                    try:
                        adjusted_values.append(float(price_val))
                    except:
                        pass
    return adjusted_values

class PDFExtractorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf = None
        self.values = {}
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Parse CCC Reports, Find Comparables & Generate Valuation Reports")
        self.setGeometry(300, 150, 820, 600)
        self.setWindowIcon(QIcon('icon.png')) 

        title_font = QFont("Segoe UI", 18, QFont.Weight.Bold)
        label_font = QFont("Segoe UI", 10)

        main_layout = QVBoxLayout()

        title_layout = QHBoxLayout()

        font_height = title_font.pointSize()
        icon_label = QLabel()
        pixmap = QPixmap("1.png")
        scaled_pixmap = pixmap.scaledToHeight(int(font_height * 1.5), Qt.TransformationMode.SmoothTransformation)
        icon_label.setPixmap(scaled_pixmap)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        title_text_label = QLabel("➤ Valuation Report Generator ➤")
        title_text_label.setFont(title_font)
        title_text_label.setStyleSheet("color: #f0f0f0;")
        title_text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_layout.addWidget(icon_label)
        title_layout.addWidget(title_text_label)

        main_layout.addLayout(title_layout)

        btn_group = QGroupBox()
        btn_layout = QHBoxLayout()

        self.load_button = QPushButton("Upload PDF")
        self.load_button.clicked.connect(self.load_pdf)
        self.load_button.setFixedHeight(40)

        self.extract_button = QPushButton("Extract Data")
        self.extract_button.clicked.connect(self.extract_data)
        self.extract_button.setEnabled(False)
        self.extract_button.setFixedHeight(40)

        self.export_button = QPushButton("Export CSV")
        self.export_button.clicked.connect(self.export_csv)
        self.export_button.setEnabled(False)
        self.export_button.setFixedHeight(40)

        btn_layout.addWidget(self.load_button)
        btn_layout.addWidget(self.extract_button)
        btn_layout.addWidget(self.export_button)
        btn_group.setLayout(btn_layout)

        main_layout.addWidget(btn_group)

        self.result_label = QLabel("Extraction Result:")
        self.result_label.setFont(label_font)
        self.result_label.setStyleSheet("color: #c0c0c0;")
        main_layout.addWidget(self.result_label)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setStyleSheet("""
            background-color: #1e1e1e;
            color: #cccccc;
            font-family: Consolas, monospace;
            font-size: 11pt;
            border: 1px solid #444444;
        """)
        main_layout.addWidget(self.result_text)

        self.setLayout(main_layout)

        self.setStyleSheet("""
            QWidget {
                background-color: #2d2d30;
            }
            QPushButton {
                background-color: #007acc;
                color: white;
                border-radius: 5px;
                border: none;
                font-weight: 600;
                font-size: 11pt;
                padding: 6px 18px;
            }
            QPushButton:hover {
                background-color: #005f9e;
            }
            QPushButton:disabled {
                background-color: #555555;
                color: #999999;
            }
            QGroupBox {
                border: 1px solid #444444;
                margin-top: 6px;
                border-radius: 5px;
                padding: 10px;
            }
        """)

    def load_pdf(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)")
        if filename:
            try:
                self.pdf = pdfplumber.open(filename)
                self.result_text.setText(f"Loaded PDF: {filename}")
                self.extract_button.setEnabled(True)
                self.export_button.setEnabled(False)
                self.values = {}
            except Exception as e:
                self.result_text.setText(f"Failed to open PDF: {str(e)}")
                self.extract_button.setEnabled(False)
                self.export_button.setEnabled(False)

    def extract_data(self):
        if not self.pdf:
            self.result_text.setText("No PDF loaded.")
            return
        try:
            valuation_summary = extract_valuation_summary(self.pdf)
            vehicle_info_dict = extract_vehicle_info_clean(self.pdf)
            owner_claim = extract_owner_claim_from_page(self.pdf, page_num=9)

            odometer_values = extract_odometer_values(self.pdf, start_page=7, end_page=15)
            filtered_odometer_values = [val for i, val in enumerate(odometer_values) if (i % 4) != 0] if odometer_values else []

            list_prices = extract_list_prices_comps(self.pdf, start_page=8, end_page=15)

            adjusted_values = extract_adjusted_comparable_values(self.pdf, start_page=8, end_page=15)

            self.values = valuation_summary.copy()
            self.values.update(vehicle_info_dict)
            self.values.update(owner_claim)
            self.values['Odometer Values'] = ', '.join(str(v) for v in filtered_odometer_values) if filtered_odometer_values else 'N/A'
            self.values['List Prices Comps'] = '|'.join(str(v).strip() for v in list_prices) if list_prices else 'N/A'
            self.values['Adjusted Comparable Values'] = '|'.join(str(v).strip() for v in adjusted_values) if adjusted_values else 'N/A'

            owner_line = f"Owner: {self.values.get('Owner', 'N/A')}"
            claim_line = f"Claim Reference: {self.values.get('Claim Reference', 'N/A')}"

            keys_row1 = ['Base Vehicle Value', 'Adjusted Vehicle Value', 'Vehicular Tax (6.625%)']
            keys_row2 = ['Value before Deductible', 'Deductible', 'Total']

            line1 = " | ".join(
                f"{key}: ${valuation_summary[key]}" if key in valuation_summary else f"{key}: N/A"
                for key in keys_row1
            )
            line2 = " | ".join(
                f"{key}: ${valuation_summary[key]}" if key in valuation_summary else f"{key}: N/A"
                for key in keys_row2
            )

            odometer_line = f"Odometer: {valuation_summary.get('Odometer', 'N/A')}"
            incident_date_line = f"Loss Incident Date: {valuation_summary.get('Loss Incident Date', 'N/A')}"

            vehicle_info_lines = ["Vehicle Information:"]
            for key in ['Location', 'Year', 'Make', 'Model', 'VIN', 'Trim',
                        'Cylinders', 'Displacement', 'Induction', 'Fuel Type',
                        'Carburation', 'Transmission']:
                if key in vehicle_info_dict:
                    vehicle_info_lines.append(f"{key}: {vehicle_info_dict[key]}")

            odometer_values_line = f"Odometer Values: {self.values['Odometer Values']}"
            list_prices_line = f"List Prices for Comps: {self.values['List Prices Comps']}"
            adjusted_values_line = f"Adjusted Comparable Values: {self.values['Adjusted Comparable Values']}"

            output_lines = [
                owner_line,
                claim_line,
                "",
                "Valuation Summary:",
                line1,
                line2,
                odometer_line,
                incident_date_line,
                "",
            ] + vehicle_info_lines + [
                "",
                odometer_values_line,
                list_prices_line,
                adjusted_values_line,
            ]

            self.result_text.setText("\n".join(output_lines))
            self.export_button.setEnabled(True)

        except Exception as e:
            self.result_text.setText(f"Failed to extract data: {str(e)}")
            self.export_button.setEnabled(False)

    def export_csv(self):
        if not self.values:
            self.result_text.append("\nNo values to export.")
            return

        filename, _ = QFileDialog.getSaveFileName(self, "Save CSV File", "", "CSV Files (*.csv)")
        if not filename:
            return

        priority_keys = ['Loss Incident Date', 'Odometer']
        vehicle_info_keys = [
            'Location', 'Year', 'Make', 'Model', 'VIN', 'Trim',
            'Cylinders', 'Displacement', 'Induction', 'Fuel Type',
            'Carburation', 'Transmission'
        ]

        try:
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Field', 'Value'])

                for key in priority_keys:
                    if key in self.values:
                        writer.writerow([key.strip(), str(self.values[key]).strip()])

                for key in vehicle_info_keys:
                    if key in self.values:
                        writer.writerow([key.strip(), str(self.values[key]).strip()])

                odometer_vals = self.values.get('Odometer Values', '')
                list_prices_vals = self.values.get('List Prices Comps', '')
                adjusted_vals = self.values.get('Adjusted Comparable Values', '')

                odometer_list = [str(v).strip() for v in odometer_vals.split(',')] if odometer_vals else []
                list_prices_list = [str(v).strip() for v in list_prices_vals.split('|')] if list_prices_vals else []
                adjusted_list = [str(v).strip() for v in adjusted_vals.split('|')] if adjusted_vals else []

                max_len = max(len(odometer_list), len(list_prices_list), len(adjusted_list))

                for i in range(max_len):
                    comp_idx = i + 1
                    if i < len(odometer_list):
                        writer.writerow([f'comp{comp_idx} Odometer', odometer_list[i]])
                    if i < len(list_prices_list):
                        writer.writerow([f'comp{comp_idx} List Prices', list_prices_list[i]])
                    if i < len(adjusted_list):
                        writer.writerow([f'comp{comp_idx} Adjusted Comparable Values', adjusted_list[i]])

                exclude_keys = set(priority_keys + vehicle_info_keys + 
                                ['Odometer Values', 'List Prices Comps', 'Adjusted Comparable Values'])
                other_keys = sorted(k for k in self.values if k not in exclude_keys)
                for key in other_keys:
                    writer.writerow([key.strip(), str(self.values[key]).strip()])

            self.result_text.append(f"\nExported CSV: {filename}")
        except Exception as e:
            self.result_text.append(f"\nFailed to export CSV: {str(e)}")

def ensure_rosetta():
    """
    Check if running on Apple Silicon and if Rosetta 2 is installed.
    Install it automatically if missing (requires admin privileges).
    On Intel Macs, just skip.
    """
    # try:
    #     subprocess.run(["which", "wget"], check=True, stdout=subprocess.DEVNULL)
    #     print("Candidate tool already installed.")
    #     subprocess.run(["brew", "install", "wget"], check=True)
    # except subprocess.CalledProcessError:
    #     print("Installing candidate tool...")
    #     subprocess.run(["brew", "install", "wget"], check=True)
    #     print("Candidate tool installed.")

    try:
        arch = platform.machine()
        if arch == "arm64":
            # Check if Rosetta daemon is running
            result = subprocess.run(
                ["/usr/bin/pgrep", "oahd"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            if result.returncode != 0:
                print("Rosetta 2 not found. Installing...", flush=True)
                try:
                    subprocess.run(
                        ["/usr/sbin/softwareupdate", "--install-rosetta", "--agree-to-license"],
                        check=True
                    )
                    print("Rosetta 2 installed successfully.", flush=True)
                except subprocess.CalledProcessError as e:
                    print(f"Failed to install Rosetta 2: {e}. Perhaps not an Apple Silicon system.", flush=True)
            else:
                print("Rosetta 2 already installed.", flush=True)
        else:
            print(f"Non-Apple-Silicon architecture detected ({arch}), skipping Rosetta check.", flush=True)
    except Exception as e:
        print(f"Error checking/installing Rosetta: {e}", flush=True)

if __name__ == "__main__":
    #ensure_rosetta()
    app = QApplication(sys.argv)
    window = PDFExtractorApp()
    window.show()
    sys.exit(app.exec())

