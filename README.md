# Snap-Save URL Screenshot Automation  

Automated capture tool for browser screenshots and active page URLs. The script saves each captured screenshot and corresponding URL directly into a Microsoft Word document for easy archival and reporting.  

***

### Features  
- Capture a full-screen screenshot with **Ctrl+Q**.  
- Automatically copy the active browser URL.  
- Append both the screenshot and URL into a Word (.docx) file.  
- Maintain timestamps for each entry.  
- Exit cleanly by typing **exit** in the terminal.  

***

### Requirements  
Ensure the following Python libraries are installed:  
```bash
pip install pyautogui keyboard pyperclip python-docx
```

***

### File Structure  
```
Snap-Save-URL-Screenshot-Automation/
│
├── shots/                        # Folder to store temporary screenshots
│   └── shot.png
│
├── shots.docx                    # Main output file with screenshots and URLs
│
└── main.py                       # The script
```

Edit the paths inside the script (`docxfile` and `shotfile`) to match your local environment before running.  

***

### Usage  
1. Run the script in a terminal or IDE.  
2. Open a browser window with the target webpage.  
3. Press **Ctrl+Q** to capture a screenshot and copy the current URL.  
4. Repeat as needed; each capture is appended to the same Word document.  
5. Type **exit** and press Enter to stop the script safely.  

***

### How It Works  
- When **Ctrl+Q** is pressed, `pyautogui` takes a screenshot and saves it.  
- The script simulates `Ctrl+L` and `Ctrl+C` keystrokes to copy the URL.  
- The timestamp, screenshot, and URL are appended to a `.docx` file via `python-docx`.  
- An input listener running on a background thread allows the user to exit gracefully.  

***

### Notes  
- Works best when a browser window is active and address bar is focused by `Ctrl+L`.  
- Screenshots are temporarily saved as `shot.png` before being embedded in the document.  
- The script runs continuously and uses low CPU resources through timed sleep intervals.  

***

### Example Output in Word  
Each capture entry in `shots.docx` includes:  
- A timestamp  
- The screenshot image  
- The corresponding URL  
