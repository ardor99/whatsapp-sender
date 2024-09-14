# whatsapp_sender/controller.py

"""Controller module for WhatsApp Automation application."""

import threading
import time
from tkinter import filedialog
from datetime import datetime
import os
import shutil

from whatsapp_sender.model import WhatsAppModelError  # Now used in exception handling

class WhatsAppController:
    """Controller class that manages interactions between the model and view."""

    def __init__(self, model, view):
        """Initialize the controller with the model and view."""
        self.model = model
        self.view = view
        self.filepath = ''
        self.photo_path = ''
        self.stop_thread = False
        self.pause_thread = False
        self.process_thread = None

        # Set up commands for the view's buttons
        self.view.set_webdriver_command(self.set_webdriver_path)
        self.view.set_login_command(self.login)
        self.view.set_choose_file_command(self.choose_file)
        self.view.set_choose_photo_command(self.choose_photo)
        self.view.set_send_command(self.start_process)
        self.view.set_pause_command(self.pause_messages)
        self.view.set_stop_command(self.stop_messages)

    def set_webdriver_path(self):
        """Set the path to the WebDriver executable."""
        self.model.webdriver_path = filedialog.askopenfilename(
            title="Select WebDriver",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if self.model.webdriver_path and os.path.isfile(self.model.webdriver_path):
            self.view.update_info(f"WebDriver set to: {self.model.webdriver_path}")
            self.view.update_text_area(f"WebDriver set to: {self.model.webdriver_path}")
        else:
            self.view.update_info("Invalid WebDriver path!")
            self.view.update_text_area("Invalid WebDriver path!")

    def login(self):
        """Log in to WhatsApp Web using the WebDriver."""
        self.model.chosen_language = self.view.chosen_language.get()
        try:
            self.model.login()
            self.view.update_info("Status: Please login to WhatsApp...")
            self.view.update_text_area("Opened WhatsApp Web. Please login.")
        except (WhatsAppModelError, ValueError) as e:
            self.view.update_info(f"Error: {str(e)}")
            self.view.update_text_area(f"Error: {str(e)}")

    def choose_file(self):
        """Allow the user to select an Excel file containing messages."""
        self.filepath = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if self.filepath:
            self.view.update_info(f"File chosen: {self.filepath}")
            self.view.update_text_area(f"Chosen Excel File: {self.filepath}")
        else:
            self.view.update_info("No file chosen.")
            self.view.update_text_area("No file chosen.")

    def choose_photo(self):
        """Allow the user to select a photo to send."""
        self.photo_path = filedialog.askopenfilename(
            title="Select Photo",
            filetypes=[("Image files", "*.jpg;*.jpeg;*.png")]
        )
        if self.photo_path:
            target_path = os.path.join(os.getcwd(), os.path.basename(self.photo_path))
            shutil.copy(self.photo_path, target_path)
            self.model.photo_path = target_path
            self.view.update_info(f"Photo chosen: {self.photo_path}")
            self.view.update_text_area(f"Photo chosen: {self.photo_path}")
        else:
            self.view.update_info("No photo chosen.")
            self.view.update_text_area("No photo chosen.")

    def start_process(self):
        """Start the message sending process in a separate thread."""
        if not self.process_thread or not self.process_thread.is_alive():
            self.process_thread = threading.Thread(target=self.send_messages)
            self.process_thread.start()

    def send_messages(self):
        """Send messages to the list of phone numbers."""
        self.stop_thread = False

        if not self.model.driver:
            self.view.update_info("Please login first!")
            self.view.update_text_area("Please login first!")
            return
        if not self.filepath:
            self.view.update_info("Please choose an Excel file first!")
            self.view.update_text_area("Please choose an Excel file first!")
            return

        self.view.update_info("Status: Sending messages...")
        self.view.update_text_area("Started sending messages...")

        self.model.load_excel_file(self.filepath)

        total_rows = len(self.model.messages)
        start_time = datetime.now()

        for idx, (phone_number, message) in enumerate(self.model.messages, start=1):
            if self.stop_thread:
                self.view.update_text_area("Stopping the message sending process...")
                break
            while self.pause_thread:
                self.view.update_info("Status: Paused. Click Resume to continue.")
                time.sleep(1)

            if self.view.send_mode.get() == "message":
                status = self.model.send_message(phone_number, message)
            else:
                if not self.model.photo_path:
                    self.view.update_info(f"No photo chosen for {phone_number}")
                    self.view.update_text_area(f"No photo chosen for {phone_number}")
                    continue
                status = self.model.send_photo(phone_number, message)

            self.view.update_text_area(f"Message/photo to {phone_number}: {status}")

            progress_value = (idx / total_rows) * 100
            self.view.update_progress(progress_value)

            elapsed_time = datetime.now() - start_time
            estimated_total_time = (elapsed_time / idx) * total_rows
            remaining_time = estimated_total_time - elapsed_time
            remaining_time_str = str(remaining_time).split('.', maxsplit=1)[0]
            self.view.time_var.set(f"Estimated Time Remaining: {remaining_time_str}")

            time.sleep(2)

        self.model.cleanup()
        self.view.update_info("Status: Done sending messages!")
        self.view.update_text_area("Finished sending messages.")

    def pause_messages(self):
        """Pause or resume the message sending process."""
        self.pause_thread = not self.pause_thread
        if self.pause_thread:
            self.view.update_info("Paused...")
            self.view.btn_pause.config(text="Resume", bg="blue")
        else:
            self.view.update_info("Resuming...")
            self.view.btn_pause.config(text="Pause", bg="orange")

    def stop_messages(self):
        """Stop the message sending process."""
        self.stop_thread = True
        self.view.update_info("Stopping...")
        self.view.update_text_area("Attempting to stop the message sending process...")
        if self.process_thread and self.process_thread.is_alive():
            self.process_thread.join()
        self.view.update_info("Process Stopped.")
