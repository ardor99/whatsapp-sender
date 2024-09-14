# whatsapp_sender/utils.py

"""Utility functions for the WhatsApp Automation application."""

import platform
import subprocess
from io import BytesIO
from PIL import Image

try:
    import win32clipboard  # For Windows clipboard operations
except ImportError:
    win32clipboard = None  # Handle non-Windows systems

def copy_image_to_clipboard(image_path):
    """Copy an image to the clipboard for pasting into WhatsApp."""
    system_platform = platform.system()

    if system_platform == "Windows":
        if win32clipboard is None:
            raise ImportError("win32clipboard module is required on Windows systems.")

        image = Image.open(image_path)
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        # pylint: disable=c-extension-no-member
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()

    elif system_platform == "Linux":
        image = Image.open(image_path)
        output = BytesIO()
        image.save(output, "PNG")
        output.seek(0)

        subprocess.run(
            ['xclip', '-selection', 'clipboard', '-t', 'image/png'],
            input=output.read(), check=True
        )

    elif system_platform == "Darwin":
        subprocess.run(
            ['osascript', '-e',
             f'set the clipboard to (read (POSIX file "{image_path}") as JPEG picture)'],
            check=True
        )

    else:
        raise NotImplementedError(f"Clipboard copy is not implemented for {system_platform}.")
