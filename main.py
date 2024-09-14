# main.py

"""Main module to start the WhatsApp Automation application."""

import tkinter as tk

from whatsapp_sender.controller import WhatsAppController
from whatsapp_sender.model import WhatsAppModel
from whatsapp_sender.view import WhatsAppView

def main():
    """Main function to run the application."""
    root = tk.Tk()
    model = WhatsAppModel()
    view = WhatsAppView(root)
    root.controller = WhatsAppController(model, view)  # Assign to root
    root.mainloop()

if __name__ == "__main__":
    main()
