# whatsapp_sender/view.py

"""View module for WhatsApp Automation application."""

import tkinter as tk
from tkinter import ttk
import time

class WhatsAppView:
    """View class for the WhatsApp Automation GUI."""

    # pylint: disable=too-many-instance-attributes

    def __init__(self, root):
        """Initialize the GUI components."""
        self.root = root
        self.root.title("WhatsApp Automation")
        self.root.geometry("600x500")
        self.root.resizable(True, True)

        # Initialize attributes
        self.main_frame = tk.Frame(self.root, padx=10, pady=10)
        self.main_frame.pack(fill="both", expand=True)

        self.create_widgets()

    def create_widgets(self):
        """Create all the widgets for the GUI."""
        header = tk.Label(
            self.main_frame, text="WhatsApp Automation",
            font=("Arial", 18, "bold"), pady=10
        )
        header.pack()

        self.btn_webdriver = tk.Button(
            self.main_frame, text="Set WebDriver Path", padx=10, pady=5
        )
        self.btn_webdriver.pack(pady=5)

        self.btn_login = tk.Button(
            self.main_frame, text="Login", padx=10, pady=5
        )
        self.btn_login.pack(pady=5)

        self.btn_choose_file = tk.Button(
            self.main_frame, text="Choose Excel File", padx=10, pady=5
        )
        self.btn_choose_file.pack(pady=5)

        self.chosen_language = tk.StringVar(value="en")
        self.send_mode = tk.StringVar(value="message")

        lang_frame = tk.Frame(self.main_frame)
        lang_frame.pack(pady=5)
        lang_en = tk.Radiobutton(
            lang_frame, text="English", variable=self.chosen_language, value="en"
        )
        lang_ar = tk.Radiobutton(
            lang_frame, text="Arabic", variable=self.chosen_language, value="ar"
        )
        lang_en.pack(side="left", padx=5)
        lang_ar.pack(side="right", padx=5)

        mode_frame = tk.Frame(self.main_frame)
        mode_frame.pack(pady=5)
        mode_message = tk.Radiobutton(
            mode_frame, text="Send Message", variable=self.send_mode,
            value="message", command=self.toggle_photo_options
        )
        mode_photo = tk.Radiobutton(
            mode_frame, text="Send Photo", variable=self.send_mode,
            value="photo", command=self.toggle_photo_options
        )
        mode_message.pack(side="left", padx=5)
        mode_photo.pack(side="right", padx=5)

        self.btn_choose_photo = tk.Button(
            self.main_frame, text="Choose Photo", padx=10, pady=5
        )
        self.btn_choose_photo.pack(pady=5)
        self.btn_choose_photo.pack_forget()

        self.btn_send = tk.Button(
            self.main_frame, text="Send", padx=10, pady=5, bg="green", fg="white"
        )
        self.btn_send.pack(pady=5)

        self.btn_pause = tk.Button(
            self.main_frame, text="Pause", padx=10, pady=5, bg="orange", fg="white"
        )
        self.btn_pause.pack(pady=5)

        self.btn_stop = tk.Button(
            self.main_frame, text="Stop", padx=10, pady=5, bg="red", fg="white"
        )
        self.btn_stop.pack(pady=5)

        self.info_var = tk.StringVar()
        self.info_var.set("Status: Waiting...")
        lbl_info = tk.Label(
            self.main_frame, textvariable=self.info_var, pady=10, font=("Arial", 12)
        )
        lbl_info.pack()

        log_frame = tk.Frame(self.main_frame)
        log_frame.pack(fill="both", expand=True)

        self.text_area = tk.Text(
            log_frame, height=10, width=50, font=("Arial", 10), padx=10, pady=10
        )
        self.scrollbar = tk.Scrollbar(log_frame, command=self.text_area.yview)
        self.text_area.configure(yscrollcommand=self.scrollbar.set)
        self.text_area.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.progress = ttk.Progressbar(
            self.main_frame, orient="horizontal", mode="determinate", length=500
        )
        self.progress.pack(pady=10)

        self.time_var = tk.StringVar()
        self.time_var.set("Estimated Time Remaining: --:--")
        lbl_time = tk.Label(
            self.main_frame, textvariable=self.time_var, pady=10, font=("Arial", 12)
        )
        lbl_time.pack()

    def toggle_photo_options(self):
        """Show or hide photo options based on the selected send mode."""
        if self.send_mode.get() == "photo":
            self.btn_choose_photo.pack(pady=5)
        else:
            self.btn_choose_photo.pack_forget()

    def set_webdriver_command(self, command):
        """Set the command for the WebDriver button."""
        self.btn_webdriver.config(command=command)

    def set_login_command(self, command):
        """Set the command for the Login button."""
        self.btn_login.config(command=command)

    def set_choose_file_command(self, command):
        """Set the command for the Choose File button."""
        self.btn_choose_file.config(command=command)

    def set_choose_photo_command(self, command):
        """Set the command for the Choose Photo button."""
        self.btn_choose_photo.config(command=command)

    def set_send_command(self, command):
        """Set the command for the Send button."""
        self.btn_send.config(command=command)

    def set_pause_command(self, command):
        """Set the command for the Pause button."""
        self.btn_pause.config(command=command)

    def set_stop_command(self, command):
        """Set the command for the Stop button."""
        self.btn_stop.config(command=command)

    def update_info(self, message):
        """Update the status information label."""
        self.info_var.set(message)

    def update_text_area(self, message):
        """Update the text area with a new message."""
        self.text_area.insert(tk.END, message + "\n")
        self.text_area.see(tk.END)
        with open("log.txt", "a", encoding="utf-8") as log_file:
            log_file.write(f"{time.ctime()}: {message}\n")

    def update_progress(self, value):
        """Update the progress bar value."""
        self.progress["value"] = value
        self.root.update_idletasks()
