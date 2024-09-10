import openpyxl
import re
import pyperclip
import tkinter as tk
from tkinter import filedialog, ttk
import threading
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.chrome.options import Options
import os
import shutil
from datetime import datetime, timedelta


class WhatsAppAutomation:

    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Automation")
        self.root.geometry("600x500")  # Adjust window size
        self.root.resizable(True, True)  # Allow resizing

        self.init_gui()

        self.driver = None
        self.wait = None
        self.filepath = ''
        self.webdriver_path = ''
        self.photo_path = ''
        self.stop_thread = False
        self.pause_thread = False

    def init_gui(self):
        # Main Frame for Padding
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)

        # Add Header Label
        header = tk.Label(main_frame, text="WhatsApp Automation", font=("Arial", 18, "bold"), pady=10)
        header.pack()

        # WebDriver Path Button
        btn_webdriver = tk.Button(main_frame, text="Set WebDriver Path", command=self.set_webdriver_path, padx=10, pady=5)
        btn_webdriver.pack(pady=5)

        # Login Button
        btn_login = tk.Button(main_frame, text="Login", command=self.login, padx=10, pady=5)
        btn_login.pack(pady=5)

        # Choose File Button
        btn_choose_file = tk.Button(main_frame, text="Choose Excel File", command=self.choose_file, padx=10, pady=5)
        btn_choose_file.pack(pady=5)

        # Radio Buttons for Language and Send Mode
        self.chosen_language = tk.StringVar(value="en")  # default to English
        self.send_mode = tk.StringVar(value="message")  # default to sending message
        lang_frame = tk.Frame(main_frame)
        lang_frame.pack(pady=5)
        lang_en = tk.Radiobutton(lang_frame, text="English", variable=self.chosen_language, value="en")
        lang_ar = tk.Radiobutton(lang_frame, text="Arabic", variable=self.chosen_language, value="ar")
        lang_en.pack(side="left", padx=5)
        lang_ar.pack(side="right", padx=5)

        # Radio Buttons for Sending Mode (Message/Photo)
        mode_frame = tk.Frame(main_frame)
        mode_frame.pack(pady=5)
        mode_message = tk.Radiobutton(mode_frame, text="Send Message", variable=self.send_mode, value="message", command=self.toggle_photo_options)
        mode_photo = tk.Radiobutton(mode_frame, text="Send Photo", variable=self.send_mode, value="photo", command=self.toggle_photo_options)
        mode_message.pack(side="left", padx=5)
        mode_photo.pack(side="right", padx=5)

        # Photo selection (Initially hidden)
        self.btn_choose_photo = tk.Button(main_frame, text="Choose Photo", command=self.choose_photo, padx=10, pady=5)
        self.btn_choose_photo.pack(pady=5)
        self.btn_choose_photo.pack_forget()  # Hide the button initially

        # Send Button
        btn_send = tk.Button(main_frame, text="Send", command=lambda: threading.Thread(target=self.send_messages).start(), padx=10, pady=5, bg="green", fg="white")
        btn_send.pack(pady=5)

        # Pause Button
        btn_pause = tk.Button(main_frame, text="Pause", command=self.pause_messages, padx=10, pady=5, bg="orange", fg="white")
        btn_pause.pack(pady=5)

        # Stop Button
        btn_stop = tk.Button(main_frame, text="Stop", command=self.stop_messages, padx=10, pady=5, bg="red", fg="white")
        btn_stop.pack(pady=5)

        # Status Label
        self.info_var = tk.StringVar()
        self.info_var.set("Status: Waiting...")
        lbl_info = tk.Label(main_frame, textvariable=self.info_var, pady=10, font=("Arial", 12))
        lbl_info.pack()

        # Text Area with Scrollbar for Logs
        log_frame = tk.Frame(main_frame)
        log_frame.pack(fill="both", expand=True)

        self.text_area = tk.Text(log_frame, height=10, width=50, font=("Arial", 10), padx=10, pady=10)
        self.scrollbar = tk.Scrollbar(log_frame, command=self.text_area.yview)
        self.text_area.configure(yscrollcommand=self.scrollbar.set)
        self.text_area.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Progress Bar
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate", length=500)
        self.progress.pack(pady=10)

        # Time Remaining Label
        self.time_var = tk.StringVar()
        self.time_var.set("Estimated Time Remaining: --:--")
        lbl_time = tk.Label(main_frame, textvariable=self.time_var, pady=10, font=("Arial", 12))
        lbl_time.pack()

    def toggle_photo_options(self):
        """ Show or hide photo selection based on send mode (message or photo). """
        if self.send_mode.get() == "photo":
            self.btn_choose_photo.pack(pady=5)
        else:
            self.btn_choose_photo.pack_forget()

    def choose_photo(self):
        """ Let the user select a photo and copy it to the program's directory. """
        self.photo_path = filedialog.askopenfilename(title="Select Photo", filetypes=[("Image files", "*.jpg;*.jpeg;*.png"), ("All files", "*.*")])
        if self.photo_path:
            # Copy the selected photo to the same directory as the script
            target_path = os.path.join(os.getcwd(), os.path.basename(self.photo_path))
            shutil.copy(self.photo_path, target_path)
            self.photo_path = target_path  # Update photo path to the new location
            self.update_text_area(f"Photo chosen: {self.photo_path}")
        else:
            self.update_info_var("No photo chosen.")
            self.update_text_area("No photo chosen.")

    def update_progress(self, value):
        self.progress["value"] = value
        self.root.update_idletasks()  # Ensure progress bar updates in real-time

    def update_text_area(self, message):
        """ Update the text area and log to file. """
        self.text_area.insert(tk.END, message + "\n")
        self.text_area.see(tk.END)
        with open("log.txt", "a") as log_file:
            log_file.write(f"{time.ctime()}: {message}\n")

    def _safe_update_text_area(self, message):
        """ Safely update the text area in the main thread. """
        self.root.after(0, lambda: self.update_text_area(message))

    def update_info_var(self, message):
        """ Thread-safe update to info_var. """
        self.root.after(0, self._safe_update_info_var, message)

    def _safe_update_info_var(self, message):
        """ Safely update the info_var in the main thread. """
        self.info_var.set(message)

    def set_webdriver_path(self):
        self.webdriver_path = filedialog.askopenfilename(title="Select WebDriver", filetypes=[("Executable files", "*.exe"), ("All files", "*.*")])
        if self.webdriver_path and os.path.isfile(self.webdriver_path):
            self.update_info_var(f"WebDriver set to: {self.webdriver_path}")
            self.update_text_area(f"WebDriver set to: {self.webdriver_path}")
        else:
            self.update_info_var("Invalid WebDriver path!")
            self.update_text_area("Invalid WebDriver path!")

    def login(self):
        if not self.webdriver_path:
            self.update_info_var("Please select WebDriver path first!")
            self.update_text_area("Please select WebDriver path first!")
            return

        chrome_options = Options()
        data_dir = os.path.join(os.getcwd(), 'whatsapp_session')
        chrome_options.add_argument(f"--user-data-dir={data_dir}")

        try:
            service = webdriver.chrome.service.Service(self.webdriver_path)
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 10)  # Initialize WebDriverWait after driver creation
        except WebDriverException as e:
            self.update_info_var(f"Failed to launch WebDriver: {str(e)}")
            self.update_text_area(f"Error launching WebDriver: {str(e)}")
            return

        try:
            self.driver.get("https://web.whatsapp.com/")
            self.update_info_var("Status: Please login to WhatsApp...")
            self.update_text_area("Opened WhatsApp Web. Please login.")
        except Exception as e:
            self.update_info_var(f"Failed to open WhatsApp Web: {str(e)}")
            self.update_text_area(f"Error opening WhatsApp Web: {str(e)}")

    def choose_file(self):
        self.filepath = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if self.filepath:
            self.update_info_var(f"File chosen: {self.filepath}")
            self.update_text_area(f"Chosen Excel File: {self.filepath}")
        else:
            self.update_info_var("No file chosen.")
            self.update_text_area("No file chosen.")

    def send_messages(self):
        self.stop_thread = False

        if not self.driver:
            self.update_info_var("Please login first!")
            self.update_text_area("Please login first!")
            return
        if not self.filepath:
            self.update_info_var("Please choose an Excel file first!")
            self.update_text_area("Please choose an Excel file first!")
            return

        self.update_info_var("Status: Sending messages...")
        self.update_text_area("Started sending messages...")

        try:
            wb = openpyxl.load_workbook(self.filepath)
            sheet = wb.active
            wb_sent = openpyxl.Workbook()
            sheet_sent = wb_sent.active
            sheet_sent.append(["Phone Number", "Message", "Status"])

            unsent_row = 2
            total_rows = sheet.max_row - 1  # Total messages to be sent
            start_time = datetime.now()  # Start time for calculating remaining time

            for idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
                if self.stop_thread:
                    self.update_text_area("Stopping the message sending process...")
                    break
                while self.pause_thread:
                    self.update_info_var("Status: Paused. Click Resume to continue.")
                    time.sleep(1)
                phone_number, message = row
                phone_number = str(phone_number)

                try:
                    self.driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
                    self.update_text_area(f"Attempting to send to {phone_number}...")

                    retries = 0
                    message_box = None
                    while retries < 3:
                        try:
                            if self.chosen_language.get() == "en":
                                message_box = self.wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Type a message']")))
                            elif self.chosen_language.get() == "ar":
                                message_box = self.wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='اكتب رسالة']")))
                            break
                        except TimeoutException:
                            retries += 1
                            self.update_info_var(f"Retry {retries} for {phone_number}")
                            self.update_text_area(f"Retry {retries}: Failed to locate message box for {phone_number}")
                            continue

                    if not message_box:
                        self.update_text_area(f"Failed to find message box for {phone_number} after retries.")
                        unsent_row += 1

                        continue

                    if self.send_mode.get() == "message":
                        # Send message
                        cleaned_message = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\xff]', '', message)
                        pyperclip.copy(cleaned_message)
                        message_box.send_keys(Keys.CONTROL, 'v')
                        message_box.send_keys(Keys.ENTER)
                    else:
                        # Send photo
                        if self.photo_path:
                            pyperclip.copy(self.photo_path)
                            message_box.send_keys(Keys.CONTROL, 'v')
                            message_box.send_keys(Keys.ENTER)
                        else:
                            self.update_info_var(f"No photo chosen for {phone_number}")
                            self.update_text_area(f"No photo chosen for {phone_number}")
                            continue

                    # Check status (sent, delivered, read)
                    status = self.check_message_status()
                    sheet_sent.append([phone_number, message, status])
                    self.update_text_area(f"Message/photo to {phone_number} {status}.")

                except (NoSuchElementException, TimeoutException, WebDriverException) as e:
                    error_msg = f"Error for {phone_number}: {str(e)}"
                    self.update_info_var(f"Error for {phone_number}")
                    self.update_text_area(error_msg)

                time.sleep(2)  # Small delay between messages

                # Update progress after each message
                progress_value = (idx / total_rows) * 100
                self.update_progress(progress_value)

                # Calculate and update estimated time remaining
                elapsed_time = datetime.now() - start_time
                estimated_total_time = (elapsed_time / idx) * total_rows
                remaining_time = estimated_total_time - elapsed_time
                self.time_var.set(f"Estimated Time Remaining: {str(remaining_time).split('.')[0]}")
                wb_sent.save("sent_messages.xlsx")

        except Exception as e:
            self.update_info_var(f"Error: {str(e)}")
            self.update_text_area(f"Error during message sending process: {str(e)}")
        finally:
            self.driver.quit()
            self.update_info_var("Status: Done sending messages!")
            self.update_text_area("Finished sending messages.")

    def check_message_status(self):
        # Step 1: Check if the message is in the 'pending' (clock) state (msg-time)
        try:
            self.wait.until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='msg-time']")))
            self.update_text_area("Message is pending (msg-time). Waiting for final status...")
        except TimeoutException:
            pass  # Message is not pending; move on to check final status

        # Step 2: Wait for the msg-time icon to disappear (indicating the message has been sent)
        try:
            self.wait.until_not(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='msg-time']")))
            self.update_text_area("Message is no longer pending (msg-time disappeared). Checking final status...")
        except TimeoutException:
            self.update_text_area("Message remained in pending state for too long. Assuming not sent.")
            return "Not Sent"

        # Step 3: Check for final status (Sent, Delivered, Read)
        icons = [
            ("//span[@data-icon='msg-check']", "Sent"),  # Single check mark
            ("//span[@data-icon='msg-dblcheck']", "Delivered"),  # Double check mark (gray)
            ("//span[@data-icon='msg-dblcheck-ack']", "Read")  # Blue double check marks
        ]

        retries = 0
        while retries < 5:  # Retry up to 5 times to catch a final status
            for icon, state in icons:
                try:
                    self.wait.until(EC.presence_of_element_located((By.XPATH, icon)))
                    self.update_text_area(f"Message status updated to: {state}")
                    return state
                except TimeoutException:
                    continue  # Try the next icon if this one isn't found

            retries += 1
            time.sleep(2)  # Wait for 2 seconds before re-checking the status

        # If no final state is found after retries, return 'Not Sent'
        self.update_text_area("No final status found after retries. Returning 'Not Sent'.")
        return "Not Sent"

    def pause_messages(self):
        self.pause_thread = not self.pause_thread
        if self.pause_thread:
            self.update_info_var("Paused...")
        else:
            self.update_info_var("Resuming...")

    def stop_messages(self):
        self.stop_thread = True
        self.update_info_var("Stopping...")
        self.update_text_area("Attempting to stop the message sending process...")


if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppAutomation(root)
    root.mainloop()
