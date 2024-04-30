import os
import csv
import shutil
import zipfile
import logging
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton, QLabel, QWidget,
    QFileDialog, QMessageBox, QCheckBox, QProgressBar, QLineEdit, QDialog,
    QHBoxLayout  # Add this import
)
from PyQt5.QtGui import QFont, QIcon, QPixmap
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import qdarkstyle
import re

# Set up logging
LOG_FILENAME = "MASTER.log"
logger = logging.getLogger("MASTER")

# Create formatter
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

# Create rotating file handler
handler = logging.FileHandler(LOG_FILENAME)
handler.setFormatter(formatter)
logger.addHandler(handler)


class UpdateProcessThread(QThread):
    update_progress = pyqtSignal(int)
    update_site_number = pyqtSignal(str)

    def __init__(self):
        super().__init__()

    def run(self):
        global start_button_enabled
        try:
            check_PRONTO_ACES()
            data = parse_dat_files()
            if data:
                write_to_csv(data, 'MASTER.CSV')
                entries_with_files = parse_master_csv()
                unique_ids = list(set(entry[-1] for entry in entries_with_files))
                total_unique_ids = len(unique_ids)
                for idx, unique_id in enumerate(unique_ids, start=1):
                    create_update_folders(unique_id)
                    zip_update_folder(unique_id)
                    self.update_progress.emit(int((idx / total_unique_ids) * 100))
                    self.update_site_number.emit(f"Processing Site Number: {unique_id}")
                    QApplication.processEvents()
                logger.info("Process completed successfully.")
                mark_unused_csv(entries_with_files)                
                # Move files to TAKE5UPDATE directory if Take5.CSV is included
                if self.take5_file_path:
                    with open(self.take5_file_path, 'r') as take5_csv:
                        take5_sites = [line.strip() for line in take5_csv if line.strip()]
                    move_to_take5_update(take5_sites)

                start_button_enabled = True
            else:
                logger.error("No .DAT files found or PARTSBOX directory is empty or doesn't exist.")
                QMessageBox.critical(None, "Error", "No .DAT files found or PARTSBOX directory is empty or doesn't exist.")
        except Exception as e:
            logger.exception(f"An error occurred: {str(e)}")
            QMessageBox.critical(None, "Error", f"An error occurred: {str(e)}")


class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Login")
        self.setWindowIcon(QIcon('icon.png'))
        self.setGeometry(300, 300, 300, 150)

        layout = QVBoxLayout()

        self.password_label = QLabel("Password:")
        layout.addWidget(self.password_label)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.password_input)

        self.login_button = QPushButton("Login")
        self.login_button.clicked.connect(self.check_password)
        layout.addWidget(self.login_button)

        self.setLayout(layout)

    def check_password(self):
        password = self.password_input.text()
        if password == "KYSON":
            self.accept()
        else:
            QMessageBox.warning(self, "Incorrect Password", "The password you entered is incorrect.")


class UpdateProcessApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.login_dialog = LoginDialog()
        if not self.login_dialog.exec_():
            sys.exit()

        self.setWindowTitle("AutoData Parts Catalog Broadcast")
        self.setGeometry(100, 100, 400, 250)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setAlignment(Qt.AlignCenter)

        self.progress_label = QLabel("Processing Site Number: -")
        self.progress_label.setFont(QFont("Arial", 12))
        self.layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.layout.addWidget(self.progress_bar)

        self.include_take5_checkbox = QCheckBox("Include CSV of Take 5 Sites")
        self.include_take5_checkbox.setFont(QFont("Arial", 12))
        self.layout.addWidget(self.include_take5_checkbox)

        # Calculate the width of the buttons
        max_button_width = max(
            len("Start Update Build"),
            len("Select PRONTO_ACES Folder"),
            len("Select PARTSBOX Folder"),
            len("Select CSV of Take 5 Sites")
        ) * 10

        # Start button
        self.start_button = QPushButton("Start Update Build")
        self.start_button.setFont(QFont("Arial", 12))
        self.start_button.setStyleSheet("QPushButton {background-color: #FF5733; color: white; border-radius: 5px;}"
                                         "QPushButton:hover {background-color: #FF5733;}")
        self.start_button.setFixedWidth(max_button_width)  # Set fixed width
        self.start_button.clicked.connect(self.start_update_process)
        self.add_help_button(
            self.start_button, 
            "Click for help", 
            """
            <html>
            <body>
            <p style='font-size:12pt; color:white'><b>Start Update Build Button</b></p>
            <p style='font-size:10pt'>This button initiates the update process, building the UP####.ZIP files. Depending on the amount of Sites it needs to build it for, this process could take a while. This button will not be active until the PRONTO_ACES and PARTSBOX folders are selected. For more information on those, please seek the Help button next to their respective buttons.</p>
            </body>
            </html>
            """
        )

        # Master files button
        self.masterfiles_button = QPushButton("Select PRONTO_ACES Folder")
        self.masterfiles_button.setFont(QFont("Arial", 12))
        self.masterfiles_button.setFixedWidth(max_button_width)  # Set fixed width
        self.masterfiles_button.clicked.connect(self.select_masterfiles_folder)
        self.add_help_button(
            self.masterfiles_button, 
            "Click for help", 
            """
            <html>
            <body>
            <p style='font-size:12pt; color:white'><b>Select PRONTO_ACES Folder Button</b></p>
            <p style='font-size:10pt'>This button is used to select the PRONTO_ACES folder. The folder should include Sub-folders with the 'A_' format. Inside the respective 'A_' folders, there should be the .dbf and .ndx files of the parts catalog update. This is usually found in C:\\Auto Data\\AUTOUPDATE on the Auto Data FTP server witht he folder name PRONTO_ACES.</p>
            </body>
            </html>
            """
        ) 
            
        # Part files button
        self.partfiles_button = QPushButton("Select PARTSBOX Folder")
        self.partfiles_button.setFont(QFont("Arial", 12))
        self.partfiles_button.setFixedWidth(max_button_width)  # Set fixed width
        self.partfiles_button.clicked.connect(self.select_partfiles_folder)
        self.add_help_button(
            self.partfiles_button, 
            "Click for help", 
            """
            <html>
            <body>
            <p style='font-size:12pt; color:white'><b>Select PARTSBOX Folder Button</b></p>
            <p style='font-size:10pt'>This button is used to select the PARTSBOX folder. The folder should include all of the active PART####.DAT files. This folder will include Take 5 sites. This is usually found in C:\\Auto Data\\AUTOUPDATE on the Auto Data FTP server with the folder name PARTSBOX.</p>
            </body>
            </html>
            """
        )
        
        # Take5 button
        self.take5_button = QPushButton("Select CSV of Take 5 Sites")
        self.take5_button.setFont(QFont("Arial", 12))
        self.take5_button.setFixedWidth(max_button_width)  # Set fixed width
        self.take5_button.clicked.connect(self.select_take5_file)
        self.add_help_button(
            self.take5_button, 
            "Click for help", 
            """
            <html>
            <body>
            <p style='font-size:12pt; color:white'><b>Select CSV of Take 5 Sites Button</b></p>
            <p style='font-size:10pt'>This button is used to select a CSV with all of the Take 5 site numbers. This CSV of site numbers should be a list having 4 digit site numbers of all the Take 5 sites. There should be no header in the CSV. The CSV should have the following format:</p>
            <p style='font-size:10pt'>0105<br>
            0108<br>
            3256<br>
            2548<br>
            3568<br>
            ....</p>
            </body>
            </html>
            """
        )
        
        self.update_thread = None
        self.masterfiles_found = False
        self.partfiles_found = False
        self.check_start_button_state()
        self.check_button_color()

        # Add new button
        self.take5_update_button = QPushButton("Move Update")
        self.take5_update_button.setFont(QFont("Arial", 12))
        self.take5_update_button.setFixedWidth(max_button_width)
        self.take5_update_button.clicked.connect(self.select_update_destination)
        self.add_help_button(
            self.take5_update_button, 
            "Click for help", 
            """
            <html>
            <body>
            <p style='font-size:12pt; color:white'><b>Move Update Button</b></p>
            <p style='font-size:10pt'>This button is used for selecting the folder to move the update to. This will move the contents of (CURRENT DIRECTORY)\\UPDATE\\PROCESSED to whatever folder is selected by the user, when the button is pressed. This is used to send out the update once it has finished building</p>
            </body>
            </html>
            """
        )
        
    def select_update_destination(self):
        # Prompt user to select a folder
        folder_path = QFileDialog.getExistingDirectory(self, "Select Update Destination Folder")
        if folder_path:
            source_folder = os.path.join(os.getcwd(), 'UPDATE', 'PROCESSED')
            # Check if the source folder exists
            if os.path.exists(source_folder):
                # Move contents of PROCESSED directory to the selected folder
                move_folder_contents(source_folder, folder_path)
            else:
                QMessageBox.critical(None, "Error", "The PROCESSED folder does not exist.")

    def add_help_button(self, target_button, tooltip_hover, tooltip_click):
        # Create a container widget for the button and help button
        widget = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Create the help button
        help_button = QPushButton("?")
        help_button.setToolTip(tooltip_hover)
        help_button.setFixedSize(20, 20)
        help_button.setStyleSheet("QPushButton {background-color: #008CBA; color: white; border-radius: 10px;}"
                                   "QPushButton:hover {background-color: #006080;}")
        
        # Add the main button and help button to the layout
        layout.addWidget(target_button)
        layout.addSpacing(1) 
        layout.addWidget(help_button)
        
        # Set the layout for the container widget
        widget.setLayout(layout)
        
        # Connect the click signal to the help popup
        help_button.clicked.connect(lambda: self.show_help_popup(tooltip_click))
        
        # Add the container widget to the main layout
        self.layout.addWidget(widget)

    def show_help_popup(self, text):
        formatted_text = f"<html><body><p style='font-size:12pt'>{text}</p></body></html>"
        QMessageBox.information(None, "Help", formatted_text)


    def check_button_color(self):
        global start_button_enabled
        if start_button_enabled:
            self.start_button.setStyleSheet("QPushButton {background-color: #008CBA; color: white; border-radius: 5px;}"
                                             "QPushButton:hover {background-color: #006080;}")
            self.start_button.setEnabled(True)
        else:
            self.start_button.setStyleSheet("QPushButton {background-color: #FF5733; color: white; border-radius: 5px;}"
                                             "QPushButton:hover {background-color: #FF5733;}")
            self.start_button.setEnabled(False)

    def select_take5_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select CSV of Take 5 Sites", "", "CSV Files (*.csv)")
        if file_path:
            # Store the selected file path
            self.take5_file_path = file_path

    def select_masterfiles_folder(self):
        # Delete PRONTO_ACES folder before selecting files
        masterfiles_dir = os.path.join(os.getcwd(), 'PRONTO_ACES')
        if os.path.exists(masterfiles_dir):
            shutil.rmtree(masterfiles_dir)

        folder_path = QFileDialog.getExistingDirectory(self, "Select PRONTO_ACES Folder")
        if folder_path:
            # Call function to copy contents of selected folder to PRONTO_ACES
            copy_folder_contents(folder_path, "PRONTO_ACES")
            self.masterfiles_found = True
            self.check_start_button_state()
            self.check_button_color()

    def select_partfiles_folder(self):
        # Delete PARTSBOX folder before selecting files
        partfiles_dir = os.path.join(os.getcwd(), 'PARTSBOX')
        if os.path.exists(partfiles_dir):
            shutil.rmtree(partfiles_dir)

        folder_path = QFileDialog.getExistingDirectory(self, "Select PARTSBOX Folder")
        if folder_path:
            # Call function to copy contents of selected folder to PARTSBOX
            copy_folder_contents(folder_path, "PARTSBOX")
            self.partfiles_found = True
            self.check_start_button_state()
            self.check_button_color()

    def check_start_button_state(self):
        global start_button_enabled
        if self.masterfiles_found and self.partfiles_found:
            start_button_enabled = True
        else:
            start_button_enabled = False

    def start_update_process(self):
        if self.update_thread is None or not self.update_thread.isRunning():
            self.update_thread = UpdateProcessThread()
            self.update_thread.update_progress.connect(self.update_progress_bar)
            self.update_thread.update_site_number.connect(self.update_site_number)
            self.update_thread.take5_file_path = self.take5_file_path if hasattr(self, 'take5_file_path') else None
            self.update_thread.start()

    def update_progress_bar(self, value):
        self.progress_bar.setValue(value)

    def update_site_number(self, site_number):
        self.progress_label.setText(site_number)

def move_folder_contents(source_folder, destination_folder):
    # Ensure destination folder exists
    os.makedirs(destination_folder, exist_ok=True)

    # Move entire contents of source folder to destination
    for item in os.listdir(source_folder):
        source_item = os.path.join(source_folder, item)
        destination_item = os.path.join(destination_folder, item)
        if os.path.isdir(source_item):
            shutil.move(source_item, destination_item)
        else:
            shutil.move(source_item, destination_item)

    logger.info(f"Contents of {source_folder} moved to {destination_folder} successfully.")


def parse_dat_files():
    logger.info("Parsing .DAT files...")
    data = []

    partfiles_dir = os.path.join(os.getcwd(), 'PARTSBOX')
    masterfiles_dir = os.path.join(os.getcwd(), 'PRONTO_ACES')

    # Check if directories exist and have files
    if not os.path.exists(partfiles_dir) or not os.listdir(partfiles_dir):
        logger.error("PARTSBOX directory is empty or doesn't exist. Please input the necessary data.")
        return []

    if not os.path.exists(masterfiles_dir) or not os.listdir(masterfiles_dir):
        logger.error("PRONTO_ACES directory is empty or doesn't exist. Please input the necessary data.")
        return []

    # Process DAT files
    for filename in os.listdir(partfiles_dir):
        if filename.endswith('.DAT') and re.match(r'PART\d{4}\.DAT', filename):
            logger.info(f"Processing file: {filename}")
            unique_id = filename[4:8].zfill(4)
            with open(os.path.join(partfiles_dir, filename), 'r') as file:
                next(file)
                for line in file:
                    parts = line.strip().split(',')
                    part_code = parts[2].strip('"').replace("A_", "")
                    brand_code = parts[4].strip('"')
                    description = parts[3].strip('"')
                    brand_name = parts[5].strip('"')
                    data.append((f"A_{part_code}_{brand_code}", description, brand_name, unique_id))

    return data

def start_update_process(self):
    if self.update_thread is None or not self.update_thread.isRunning():
        self.update_thread = UpdateProcessThread()
        self.update_thread.update_progress.connect(self.update_progress_bar)
        self.update_thread.update_site_number.connect(self.update_site_number)
        self.update_thread.take5_file_path = self.take5_file_path if hasattr(self, 'take5_file_path') else None
        self.update_thread.start()

def write_to_csv(data, filename):
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Part_Code_BrandCode', 'Description', 'BrandName', 'Unique_ID'])
        writer.writerows(data)


def parse_master_csv():
    logger.info("Parsing MASTER.CSV...")
    entries_with_files = []
    with open('MASTER.CSV', 'r') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)
        for row in reader:
            entries_with_files.append(row)
    # Sort entries based on Unique_ID
    entries_with_files.sort(key=lambda x: x[-1])
    logger.info(f"Entries found in MASTER.CSV: {len(entries_with_files)}")
    return entries_with_files


def create_update_folders(unique_id):
    logger.info(r"Creating update folders for Unique_ID: {unique_id}...")
    update_folder = os.path.join(os.getcwd(), 'UPDATE', unique_id, 'POS', 'PARTS')
    os.makedirs(update_folder, exist_ok=True)

    with open('MASTER.CSV', 'r') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)
        for row in reader:
            if row[-1] == unique_id:
                part_brand_code = row[0]
                part_code, brand_code = part_brand_code.split('_')[1:]
                source_folder = os.path.join(os.getcwd(), 'PRONTO_ACES', f'A_{part_code}')
                if os.path.exists(source_folder):
                    for filename in os.listdir(source_folder):
                        if filename.startswith(part_brand_code):
                            shutil.copy(os.path.join(source_folder, filename), update_folder)

    partfiles_dir = os.path.join(os.getcwd(), 'PARTSBOX')
    part_dat_files = [f for f in os.listdir(partfiles_dir) if f.endswith('.DAT') and f.startswith('PART' + unique_id)]
    for dat_file in part_dat_files:
        shutil.copy(os.path.join(partfiles_dir, dat_file), os.path.join(update_folder, '..'))
    logger.info("Update folders created successfully.")


def zip_update_folder(unique_id):
    logger.info(f"Zipping update folder for Unique_ID: {unique_id}...")
    update_directory = os.path.join(os.getcwd(), 'UPDATE', unique_id)
    processed_directory = os.path.join(os.getcwd(), 'UPDATE', 'PROCESSED')
    os.makedirs(processed_directory, exist_ok=True)
    zip_filename = f"UP{unique_id}.ZIP"
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        pos_folder = os.path.join(update_directory, 'POS')
        for root_path, _, files in os.walk(pos_folder):
            for file in files:
                zipf.write(os.path.join(root_path, file), os.path.relpath(os.path.join(root_path, file), update_directory), compress_type=zipfile.ZIP_DEFLATED)
    shutil.move(zip_filename, os.path.join(processed_directory, zip_filename))
    shutil.rmtree(update_directory)
    logger.info(f"Zipped and moved {update_directory} to {processed_directory}.")


def mark_unused_csv(entries_with_files):
    logger.info("Marking unused files in UNUSED.CSV...")
    unused_entries = []

    # Check if entries have corresponding files in PRONTO_ACES directory
    for entry in entries_with_files:
        part_code_brand_code = entry[0]
        part_code, brand_code = part_code_brand_code.split('_')[1:]
        source_folder = os.path.join(os.getcwd(), 'PRONTO_ACES', f'A_{part_code}')
        # Create directory if it doesn't exist
        if not os.path.exists(source_folder):
            os.makedirs(source_folder, exist_ok=True)
        source_files = [f for f in os.listdir(source_folder) if os.path.isfile(os.path.join(source_folder, f))]
        logger.debug(f"Checking part code: {part_code}, brand code: {brand_code}, source folder: {source_folder}, source files: {source_files}")
        if not source_files:
            unused_entries.append(entry)

    # Write unused entries to UNUSED.CSV
    if unused_entries:
        with open('UNUSED.CSV', 'w', newline='') as unused_csv:
            writer = csv.writer(unused_csv)
            writer.writerow(['Part_Code_BrandCode', 'Description', 'BrandName', 'Unique_ID'])
            writer.writerows(unused_entries)
        logger.info(f"{len(unused_entries)} entries marked as unused in UNUSED.CSV.")
    else:
        logger.info("No unused files found in MASTER.CSV.")

def move_to_take5_update(take5_sites):
    logger.info("Moving files to TAKE5UPDATE directory...")
    take5_update_dir = os.path.join(os.getcwd(), 'UPDATE', 'TAKE5UPDATE')
    os.makedirs(take5_update_dir, exist_ok=True)
    for site_number in take5_sites:
        zip_filename = f"UP{site_number}.ZIP"
        if os.path.exists(os.path.join(os.getcwd(), 'UPDATE', 'PROCESSED', zip_filename)):
            shutil.move(os.path.join(os.getcwd(), 'UPDATE', 'PROCESSED', zip_filename), take5_update_dir)
            logger.info(f"Moved UP{site_number}.ZIP to TAKE5UPDATE directory.")
        else:
            logger.warning(f"UP{site_number}.ZIP not found in PROCESSED directory.")


def check_PRONTO_ACES():
    logger.info("Checking PRONTO_ACES directory for .dbf files...")
    PRONTO_ACES_dir = os.path.join(os.getcwd(), 'PRONTO_ACES')
    dbf_files = []
    for root, dirs, filenames in os.walk(PRONTO_ACES_dir):
        for filename in filenames:
            if filename.endswith('.DBF'):
                logger.debug(f"Found .dbf file: {filename}")
                dbf_files.append(os.path.join(root, filename))
    if not dbf_files:
        logger.error("No .dbf files found in PRONTO_ACES directory. Please load data into that folder.")
        raise FileNotFoundError("No .dbf files found in PRONTO_ACES directory. Please load data into that folder.")
    else:
        logger.info("At least one .dbf file found in PRONTO_ACES directory.")


def copy_folder_contents(source_folder, destination_folder):
    # Ensure destination folder exists
    os.makedirs(destination_folder, exist_ok=True)

    # Copy entire contents of source folder to destination
    for item in os.listdir(source_folder):
        source_item = os.path.join(source_folder, item)
        destination_item = os.path.join(destination_folder, item)
        if os.path.isdir(source_item):
            shutil.copytree(source_item, destination_item)
        else:
            shutil.copy(source_item, destination_item)

    # If destination is PRONTO_ACES, ensure subfolders start with 'A_'
    if destination_folder == "PRONTO_ACES":
        for item in os.listdir(destination_folder):
            if os.path.isdir(os.path.join(destination_folder, item)) and not item.startswith("A_"):
                os.rename(os.path.join(destination_folder, item), os.path.join(destination_folder, "A_" + item))

    logger.info(f"Contents of {source_folder} copied to {destination_folder} successfully.")


def main():
    global start_button_enabled
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    window = UpdateProcessApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    start_button_enabled = False
    main()