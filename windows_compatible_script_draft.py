import keyboard
import pyautogui
import pyperclip  # pip install pyperclip
from docx import Document
from docx.shared import Inches
import time
import os

# Set up your document path (make sure this file exists or create if not!)
# Corrected path to your actual file
docxfile = r"C:\Users\lefte\Downloads\Screenshots_project\Snap-Save-URL-Screenshot-Automation\shots.docx"
shotfile = r"C:\Users\lefte\Downloads\Screenshots_project\Snap-Save-URL-Screenshot-Automation\shots\shot.png"

if not os.path.exists(docxfile):
    # Create a new document if it doesn't exist
    doc = Document()
    doc.save(docxfile)

doc = Document(docxfile)

def do_cap():
    # Take a screenshot and save it as PNG
    shot = pyautogui.screenshot()
    shot.save(shotfile)

    # Add the screenshot to the docx document
    doc.add_picture(shotfile, width=Inches(7))

    # Focus address bar and copy the URL using keyboard shortcuts
    keyboard.press_and_release('ctrl+l')  # Focus address bar
    time.sleep(1)  # Give browser time to focus the address bar

    keyboard.press_and_release('ctrl+c')  # Copy URL to clipboard
    time.sleep(1)  # Give clipboard time to update

    # Get the URL from clipboard
    url = pyperclip.paste()
    print(f"--------------Captured URL: {url}")

    # Add the URL text to the document
    doc.add_paragraph(f"Captured URL: {url}")

    # Save the document
    doc.save(docxfile)
    print('Screenshot and URL saved and appended to docx.')

# Bind a keyboard shortcut to the function
keyboard.add_hotkey('ctrl+q', do_cap)
print("Script running. Press Ctrl+Q to take screenshot and grab URL.")
keyboard.wait()  # Keeps the script running, waiting for hotkey press

#to break press Ctrl + C