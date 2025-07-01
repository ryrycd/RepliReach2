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
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


class LinkedInProfileRecorder(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

        self.csv_data = []
        self.is_running = False

        # For thread-safe status logging
        self.thread_lock = threading.Lock()

        # We'll store the page's total scroll height after the first big scroll
        self.page_total_height = 0

    def init_ui(self):
        self.setWindowTitle("LinkedIn Profile Recorder")
        self.resize(600, 400)

        main_layout = QVBoxLayout()

        # Title Label
        title_label = QLabel("LinkedIn Profile Recorder")
        title_label.setFont(QFont("Arial", 18, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # CSV Selection
        csv_layout = QHBoxLayout()
        self.csv_edit = QLineEdit()
        # Updated placeholder text to reflect new expected columns (col4 = header, col5 = URL)
        self.csv_edit.setPlaceholderText("Select CSV with LinkedIn Header (col4) and URL (col5)")
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
        """Open a file dialog to select the CSV with LinkedIn header and URL in columns 4 and 5."""
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

        self.append_status(f"Loaded {len(self.csv_data)} rows from CSV.")
        self.is_running = True
        self.start_button.setEnabled(False)

        # Run the main loop in a separate thread
        thread = threading.Thread(target=self.process_rows)
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

    def process_rows(self):
        # Define main folders for screenshots and recordings (created in current working directory)
        screenshots_main_folder = os.path.join(os.getcwd(), "Screenshots")
        recordings_main_folder = os.path.join(os.getcwd(), "Recordings")
        if not os.path.exists(screenshots_main_folder):
            os.mkdir(screenshots_main_folder)
        if not os.path.exists(recordings_main_folder):
            os.mkdir(recordings_main_folder)

        for idx, row in enumerate(self.csv_data, start=1):
            # Expect at least 5 columns (we need columns 4 and 5)
            if len(row) < 5:
                self.append_status(f"Row {idx} invalid (fewer than 5 columns). Skipping.")
                continue

            # Now expect column 4 (index 3) to be the LinkedIn header and column 5 (index 4) the URL.
            title, url = row[3], row[4]
            if not url.startswith("http"):
                self.append_status(f"Row {idx} has invalid URL: {url}")
                continue

            self.append_status(f"\nProcessing row {idx}: {title} => {url}")

            # Create a subfolder for screenshots using the row number and LinkedIn header
            screenshots_subfolder = os.path.join(screenshots_main_folder, f"{idx} - {title}")
            if not os.path.exists(screenshots_subfolder):
                os.mkdir(screenshots_subfolder)

            # Create a subfolder for recordings using the row number and LinkedIn header
            recordings_subfolder = os.path.join(recordings_main_folder, f"{idx} - {title}")
            if not os.path.exists(recordings_subfolder):
                os.mkdir(recordings_subfolder)
            
            # Create an additional subfolder for screen recordings called "Screen Recording"
            screen_recording_subfolder = os.path.join(recordings_subfolder, "Screen Recording")
            if not os.path.exists(screen_recording_subfolder):
                os.mkdir(screen_recording_subfolder)

            # Record & scroll; pass the screen recording folder and the screenshots folder
            self.record_linkedin_profile(url, screen_recording_subfolder, screenshots_subfolder, idx)

        self.append_status("All rows processed.")

    def record_linkedin_profile(self, linkedin_url, recordings_folder, screenshots_folder, row_index):
        """
        Opens Chrome and navigates to the LinkedIn profile URL.
        Starts screen recording and performs scrolling and screenshots.
        The screen recording is saved in the 'recordings_folder' while screenshots are saved in 'screenshots_folder'.
        """
        driver = None
        screen_record_process = None
        output_file = ""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        try:
            self.append_status("Starting recording...")

            # CHROME OPTIONS
            chrome_options = Options()

            # On macOS, Windows, etc.
            if platform.system() == "Darwin":
                chrome_profile_path = os.path.expanduser("~/Library/Application Support/Google/Chrome/Default")
            elif platform.system() == "Windows":
                chrome_profile_path = os.path.join(
                    os.environ['USERPROFILE'],
                    "AppData", "Local", "Google", "Chrome", "User Data", "Default"
                )
            else:
                self.append_status("Unsupported OS for example user-data-dir. Attempting default.")
                chrome_profile_path = None

            if chrome_profile_path and os.path.exists(chrome_profile_path):
                chrome_options.add_argument(f"user-data-dir={chrome_profile_path}")

            chrome_options.add_argument("--start-fullscreen")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")

            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=chrome_options
            )
            self.append_status("Chrome driver initialized.")

            # Stealth
            driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                'source': '''
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    })
                '''
            })

            driver.implicitly_wait(10)

            # Open a new blank tab
            driver.execute_script("window.open('');")
            self.append_status("Opened a new blank tab.")

            # Switch to the new blank tab (index 1)
            if len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[1])
                self.append_status("Switched to the new tab.")
            else:
                self.append_status("Failed to open a new tab.")
                return

            # Navigate to LinkedIn URL
            driver.get(linkedin_url)
            self.append_status(f"Navigating to LinkedIn profile: {linkedin_url}")

            # Wait for page to load
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            self.append_status("LinkedIn profile loaded.")

            # Do a big scroll to load dynamic content and get page height
            self.page_total_height = self.scroll_page(driver, pause_time=2)
            self.append_status("Finished scrolling the profile.")

            # Scroll back to top
            driver.execute_script("window.scrollTo(0, 0);")
            self.append_status("Scrolled back to top.")
            time.sleep(2)

            # Switch focus back to original tab (index 0), then close it
            if len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[0])
                self.append_status("Switched back to original tab.")
                driver.close()
                self.append_status("Closed the original tab.")
                # Now switch back to the remaining tab
                driver.switch_to.window(driver.window_handles[0])
                self.append_status("Switched focus to the remaining tab.")

            # ~~~ Start Screen Recording (Immediately after tab is closed) ~~~
            if platform.system() == "Darwin":
                output_file = os.path.join(recordings_folder, f"recording_{timestamp}.mov")
            elif platform.system() == "Windows":
                output_file = os.path.join(recordings_folder, f"recording_{timestamp}.mp4")
            else:
                self.append_status("Unsupported OS for screen recording in this example.")
                return

            def is_tool(name):
                """Check whether 'name' is on PATH and marked as executable."""
                return which(name) is not None

            if platform.system() == "Darwin":
                if not is_tool("screencapture"):
                    self.append_status("screencapture not found. Install or use ffmpeg.")
                    return
                screen_record_process = subprocess.Popen(
                    ["screencapture", "-v", output_file],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True
                )
            elif platform.system() == "Windows":
                if not is_tool("ffmpeg"):
                    self.append_status("ffmpeg not found in PATH.")
                    return
                screen_record_process = subprocess.Popen(
                    ["ffmpeg", "-y", "-f", "gdigrab", "-framerate", "30",
                     "-i", "desktop", output_file],
                    stdin=subprocess.PIPE,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True
                )
            else:
                self.append_status("Unsupported OS for screen recording.")
                return

            self.append_status(f"Screen recording started => {output_file}")

            # Take the first screenshot immediately
            self.take_screenshot(screenshots_folder, row_index)

            # Wait 5 seconds before starting to scroll
            time.sleep(5)
            self.append_status("Paused for 5 seconds before scrolling.")

            # ~~~ Perform scrolling and take screenshots for up to 25 seconds ~~~
            record_duration = 25
            start_time = time.time()
            bottom_reached = False

            while (time.time() - start_time) < record_duration:
                if not bottom_reached:
                    self.smooth_scroll(driver, duration=1, max_scroll=500)
                    current_scroll_pos = driver.execute_script("return window.pageYOffset")
                    window_height = driver.execute_script("return window.innerHeight")
                    if current_scroll_pos + window_height >= self.page_total_height:
                        bottom_reached = True
                        self.append_status("Bottom of page reached. Scrolling back up.")
                        self.smooth_scroll_up(driver, duration=1, total_scroll=current_scroll_pos)
                time.sleep(1)
                # Only take screenshots if the bottom has not been reached
                if not bottom_reached:
                    self.take_screenshot(screenshots_folder, row_index)
            # End of scrolling/screenshot loop; the recording continues until ~30 seconds are complete.
        except Exception as e:
            self.append_status(f"An error occurred: {str(e)}")
        finally:
            # Stop the screen recording
            if screen_record_process:
                self.append_status("Stopping screen recording...")
                if platform.system() == "Darwin":
                    import signal
                    screen_record_process.send_signal(signal.SIGINT)
                    screen_record_process.wait()
                elif platform.system() == "Windows":
                    try:
                        screen_record_process.communicate(input=b"q", timeout=5)
                    except subprocess.TimeoutExpired:
                        screen_record_process.kill()
                        self.append_status("ffmpeg did not stop in time, process killed.")
                    screen_record_process.wait()

                self.append_status("Screen recording stopped.")

            # Close browser
            if driver:
                driver.quit()
                self.append_status("Chrome driver closed.")

            if output_file:
                self.append_status(f"Screen recording saved to: {output_file}")

    def scroll_page(self, driver, pause_time=2):
        """
        Scroll to the bottom of the page to load dynamic content.
        Returns the final document.body.scrollHeight for reference.
        """
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            self.append_status("Scrolling down...")
            time.sleep(pause_time)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
        return last_height

    def smooth_scroll(self, driver, duration=1, max_scroll=500):
        """
        Smoothly scrolls down by 'max_scroll' px over 'duration' seconds.
        Uses an ease-out cubic function.
        """
        def ease_out_cubic(t):
            return 1 - (1 - t) ** 3

        start_time = time.time()
        last_scroll_amount = 0

        while True:
            elapsed = time.time() - start_time
            if elapsed > duration:
                break

            progress = elapsed / duration
            eased = ease_out_cubic(progress)
            scroll_amount = eased * max_scroll
            inc = scroll_amount - last_scroll_amount
            last_scroll_amount = scroll_amount

            driver.execute_script(f"window.scrollBy(0, {inc});")
            time.sleep(0.01)

    def smooth_scroll_up(self, driver, duration=1, total_scroll=500):
        """
        Smoothly scrolls up by 'total_scroll' px over 'duration' seconds
        using an ease-out cubic function.
        """
        def ease_out_cubic(t):
            return 1 - (1 - t) ** 3

        start_time = time.time()
        last_scroll_amount = 0

        while True:
            elapsed = time.time() - start_time
            if elapsed > duration:
                break

            progress = elapsed / duration
            eased = ease_out_cubic(progress)
            scroll_amount = eased * total_scroll
            inc = scroll_amount - last_scroll_amount
            last_scroll_amount = scroll_amount

            # Negative inc to scroll up
            driver.execute_script(f"window.scrollBy(0, {-inc});")
            time.sleep(0.01)

    def take_screenshot(self, folder, row_index):
        """
        Takes a desktop screenshot using pyautogui, then crops it so that:
          - The top 16.11% of the screen (approximately 290 pixels on a 1800px tall screen)
            is removed.
          - The right 32.9861% of the screen (approximately 965 pixels on a 2880px wide screen)
            is removed.
        This method uses the actual full screenshot dimensions and calculates the crop box
        to avoid any squeezing.
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(folder, f"screenshot_{row_index}_{timestamp}.png")
        try:
            # Capture full screenshot first
            full_img = pyautogui.screenshot()
            full_width, full_height = full_img.size

            # Calculate crop amounts using percentages
            top_percent = 290 / 1800      # ~16.11%
            right_percent = 950 / 2880    # ~32.9861%

            crop_top = int(full_height * top_percent)
            crop_right = int(full_width * right_percent)

            # Define crop box: (left, top, right, bottom)
            # We keep the left edge at 0, the top edge at crop_top, 
            # and crop the right side by crop_right while keeping the bottom edge at full_height.
            crop_box = (0, crop_top, full_width - crop_right, full_height)
            cropped = full_img.crop(crop_box)
            cropped.save(filename)
            self.append_status(f"Screenshot saved: {filename}")
        except Exception as e:
            self.append_status(f"Error taking screenshot: {e}")

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
    window = LinkedInProfileRecorder()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()