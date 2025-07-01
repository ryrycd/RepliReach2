#!/usr/bin/env python3

import sys
import os
import csv
import time
import subprocess
import threading
from datetime import datetime
from shutil import which
import platform

import pyautogui
import pyperclip  # For clipboard operations

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QLineEdit, QMessageBox, QTextEdit
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QTimer

# Selenium imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


class HeyGenAutomation(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

        self.csv_data = []
        self.is_running = False

        # For thread-safe status logging
        self.thread_lock = threading.Lock()

        # Will hold our Selenium driver reference
        self.driver = None

    def init_ui(self):
        self.setWindowTitle("HeyGen Automation")
        self.resize(600, 400)

        main_layout = QVBoxLayout()

        # Title Label
        title_label = QLabel("HeyGen Automation")
        title_label.setFont(QFont("Arial", 18, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # CSV Selection
        csv_layout = QHBoxLayout()
        self.csv_edit = QLineEdit()
        # Updated placeholder: CSV is now expected to have at least 7 columns,
        # where column 6 (index 5) is the screenshot subfolder name
        # and column 7 (index 6) is the HeyGen script.
        self.csv_edit.setPlaceholderText("Select CSV with at least 7 columns (Headers in first row)")
        csv_button = QPushButton("Browse CSV")
        csv_button.clicked.connect(self.select_csv)
        csv_layout.addWidget(self.csv_edit)
        csv_layout.addWidget(csv_button)
        main_layout.addLayout(csv_layout)

        # Start Button
        self.start_button = QPushButton("Start")
        self.start_button.setFont(QFont("Arial", 14))
        self.start_button.clicked.connect(self.run_process)
        main_layout.addWidget(self.start_button)

        # Status Output
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        main_layout.addWidget(self.status_text)

        self.setLayout(main_layout)

    def select_csv(self):
        """Open a file dialog to select the CSV with headers in the first row."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select CSV File", "", "CSV Files (*.csv)")
        if file_path:
            self.csv_edit.setText(file_path)

    def run_process(self):
        if self.is_running:
            QMessageBox.warning(self, "Busy", "Process is already running.")
            return

        csv_file = self.csv_edit.text().strip()
        if not csv_file or not os.path.isfile(csv_file):
            QMessageBox.critical(self, "Error", "Please select a valid CSV file.")
            return

        # Load CSV
        try:
            with open(csv_file, mode='r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                self.csv_data = list(reader)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read CSV file: {e}")
            return

        if not self.csv_data:
            QMessageBox.critical(self, "Error", "CSV is empty.")
            return

        self.append_status(f"Loaded {len(self.csv_data)} rows (including header) from CSV.")
        self.is_running = True
        self.start_button.setEnabled(False)

        # Run the main loop in a separate thread
        thread = threading.Thread(target=self.process_csv, args=(csv_file,))
        thread.start()
        # Poll the thread
        self.poll_thread(thread)

    def poll_thread(self, thread):
        """Periodically check if the worker thread is alive."""
        if thread.is_alive():
            QTimer.singleShot(500, lambda: self.poll_thread(thread))
        else:
            self.is_running = False
            self.start_button.setEnabled(True)
            self.append_status("Process complete!")

    def process_csv(self, csv_file):
        """
        1) Launches Chrome in fullscreen, hides the automation banner, goes to heygen.com.
        2) Iterates each data row in the CSV (skipping header).
        3) For the first row, do all steps (1–19). For subsequent rows, skip step 1.
        4) Copies a "Video ID" from step 18, appends it to the CSV in a new "Video ID" column.
        5) **After each row**, update (rewrite) the CSV so progress is saved as we go.
        """
        try:
            # 1. Launch Chrome, open heygen.com
            self.launch_chrome_and_open_heygen()

            # 2. Prepare CSV for a new column "Video ID" if not already present.
            headers = self.csv_data[0]
            # Ensure headers has at least 8 columns.
            if len(headers) < 8:
                headers.extend([""] * (8 - len(headers)))
            # Set the 8th column header to "Video ID"
            headers[7] = "Video ID"

            # 3. For each data row (index 1..N-1):
            for idx in range(1, len(self.csv_data)):
                row_data = self.csv_data[idx]
                # Ensure the row has at least 7 columns (for the subfolder name and script)
                if len(row_data) < 7:
                    row_data.extend([""] * (7 - len(row_data)))
                self.append_status(f"\nProcessing row {idx} -> {row_data}")

                # If idx == 1 => first data row => do full steps;
                # if idx > 1 => skip step 1.
                skip_step_1 = (idx > 1)

                # In the new CSV:
                # - Column 6 (index 5) is the screenshot subfolder name.
                # - Column 7 (index 6) is the HeyGen script.
                # We will use row_data[6] for step 10 (pasting the script)
                # and row_data[5] for step 14 (typing the subfolder name).
                video_id = self.perform_heygen_steps(row_data, idx, skip_first_step=skip_step_1)

                # Store the result in our in-memory CSV data as column 8 (index 7)
                if video_id is not None:
                    if len(row_data) < 8:
                        row_data.extend([""] * (8 - len(row_data)))
                    row_data[7] = video_id

                # 4. **Rewrite CSV after each row** so progress is saved incrementally.
                with open(csv_file, "w", newline="", encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerows(self.csv_data)
                self.append_status(f"Row {idx} updated in CSV.")

        except Exception as e:
            self.append_status(f"An error occurred: {e}")
        finally:
            # Close browser if open.
            if self.driver is not None:
                self.driver.quit()
                self.driver = None
                self.append_status("Chrome driver closed.")

    def launch_chrome_and_open_heygen(self):
        """Initialize ChromeDriver with fullscreen, hide automation banner, go to heygen.com."""
        self.append_status("Launching Chrome...")

        chrome_options = Options()
        # On macOS, try to use the default Chrome profile path.
        if platform.system() == "Darwin":
            chrome_profile_path = os.path.expanduser("~/Library/Application Support/Google/Chrome/Default")
            if os.path.exists(chrome_profile_path):
                chrome_options.add_argument(f"user-data-dir={chrome_profile_path}")

        chrome_options.add_argument("--start-fullscreen")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")

        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )
        self.driver.execute_cdp_cmd(
            'Page.addScriptToEvaluateOnNewDocument',
            {
                'source': '''
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    })
                '''
            }
        )
        self.append_status("Chrome driver initialized.")

        self.driver.get("https://heygen.com")
        self.append_status("Navigated to heygen.com.")
        time.sleep(5)  # Let the page load for a few seconds

    def perform_heygen_steps(self, row_data, row_index, skip_first_step=False):
        """
        Performs the 19 steps with pyautogui, pulling data from row_data.
        For subsequent runs (skip_first_step=True), we skip Step 1.
        
        Returns the "video_id" from Step 18, or None if something fails.
        """

        # Helper for smooth clicking.
        def smooth_click(x, y, move_duration=1.0, after_click_wait=2.0):
            pyautogui.moveTo(x, y, duration=move_duration)
            time.sleep(0.3)
            pyautogui.click()
            time.sleep(after_click_wait)

        # Helper for smooth hovering (no click).
        def smooth_hover(x, y, move_duration=1.0, after_hover_wait=2.0):
            pyautogui.moveTo(x, y, duration=move_duration)
            time.sleep(after_hover_wait)

        # Step 1 (only if not skipping).
        if not skip_first_step:
            # Step 1: Click on 830, 830. Then wait 2 seconds.
            smooth_click(830, 830, move_duration=1.0, after_click_wait=2)

        # Step 2: Click on 140, 244. Then wait 2 seconds.
        smooth_click(140, 244, move_duration=1.0, after_click_wait=2)

        # Step 3: Click on 570, 560. Then wait 2 seconds.
        smooth_click(570, 560, move_duration=1.0, after_click_wait=2)

        # Step 4: Click on 1030, 638. Then wait 2 seconds.
        smooth_click(1030, 638, move_duration=1.0, after_click_wait=2)

        # Step 5: Click on 391, 401. Then wait 4 seconds.
        smooth_click(391, 401, move_duration=1.0, after_click_wait=4)

        # Step 6: Click on 391, 348. Then wait 2 seconds.
        smooth_click(391, 348, move_duration=1.0, after_click_wait=2)

        # Step 7: Click on 40, 336. Then wait 2 seconds.
        smooth_click(40, 336, move_duration=1.0, after_click_wait=2)

        # Step 8: Click on 250, 330. Then wait 2 seconds.
        smooth_click(250, 330, move_duration=1.0, after_click_wait=2)

        # Step 9: Press Cmd + A to select all. Then wait 2 seconds.
        if platform.system() == "Darwin":
            pyautogui.hotkey("command", "a")
        else:
            pyautogui.hotkey("ctrl", "a")
        time.sleep(2)

        # Step 10: Paste the value from row_data[6] (the script) then wait 2 seconds.
        text_for_script = row_data[6]
        pyperclip.copy(text_for_script)
        time.sleep(1)  # Allow clipboard to update
        if platform.system() == "Darwin":
            pyautogui.hotkey("command", "v")
        else:
            pyautogui.hotkey("ctrl", "v")
        time.sleep(2)

        # Step 11: Click on 627, 282. Then wait 30 seconds.
        smooth_click(627, 282, move_duration=1.0, after_click_wait=30)

        # Step 12: Click on 1379, 119. Then wait 2 seconds.
        smooth_click(1379, 119, move_duration=1.0, after_click_wait=2)

        # Step 13: Click on 917, 319. Then wait 2 seconds.
        smooth_click(917, 319, move_duration=1.0, after_click_wait=2)

        # Step 14: Type the value from row_data[5] (the subfolder name) then wait 2 seconds.
        text_for_subfolder = row_data[5]
        pyautogui.typewrite(text_for_subfolder, interval=0.02)
        time.sleep(2)

        # Step 15: Click on 902, 744. Then wait 20 seconds.
        smooth_click(902, 744, move_duration=1.0, after_click_wait=20)

        # Step 16: Click on 557, 748. Then wait 2 seconds.
        smooth_click(557, 748, move_duration=1.0, after_click_wait=2)

        # Step 17: Hover to 475, 748. Do not click. Wait 2 seconds.
        smooth_hover(475, 748, move_duration=1.0, after_hover_wait=2)

        # Step 18: Hover to 475, 701. Wait 1 second, then click.
        pyautogui.moveTo(475, 701, duration=1.0)
        time.sleep(1)
        pyautogui.click()
        time.sleep(1)

        # "Copy Video ID" from the clipboard.
        video_id = pyperclip.paste().strip()
        self.append_status(f"Copied Video ID: {video_id}")

        # Step 19: Return the video_id so the caller can store it.
        return video_id

    def append_status(self, msg):
        """Thread-safe method to append a message to the status box."""
        with self.thread_lock:
            def do_append():
                self.status_text.append(msg)
                self.status_text.verticalScrollBar().setValue(
                    self.status_text.verticalScrollBar().maximum()
                )
            QTimer.singleShot(0, do_append)


def main():
    app = QApplication(sys.argv)
    window = HeyGenAutomation()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()