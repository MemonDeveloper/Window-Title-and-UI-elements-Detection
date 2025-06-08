# 🧐 Window Title Monitor & UI elements Detection

A Python automation tool to monitor specific trading windows on Windows OS. The script detects targeted window titles and performs OCR (Optical Character Recognition) inside those windows to identify specific text elements. It alerts the user based on detected content and can automatically close windows if needed.

## 💡 Features

1. **Window Title Detection**  
   - Continuously scans open windows for titles starting with:  
     - `"Order Ticket - ES"`  
     - `"Order Preview"`  
   - Logs window open/close events in real-time.

2. **In-Window Text Detection Using OCR**  
   - Captures a screenshot of the detected window’s visible UI elements (labels, textboxes, buttons).  
   - Uses [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) to extract text from the window image.  
   - Searches the extracted text for the keyword `"ES"`.  
   - If `"ES"` is found, prompts the user with a popup to either continue or close the window.  
   - Automatically closes the window if the user selects “No” in the popup.

## 🛠️ Technologies Used

- Python 3.x  
- Windows API via `ctypes` and `pywin32` (`win32gui`, `win32con`, `win32process`, `win32api`, `win32ui`)  
- [Pillow](https://python-pillow.org/) for image handling  
- [pytesseract](https://github.com/madmaze/pytesseract) for OCR  

## 🔧 Setup Instructions

1. **Install Python dependencies:**
   ```bash
   pip install pywin32 pillow pytesseract
````

2. **Install Tesseract OCR engine:**

   * Download and install from [Tesseract GitHub](https://github.com/tesseract-ocr/tesseract) or [official binaries](https://github.com/UB-Mannheim/tesseract/wiki).
   * Update the `tesseract_cmd` path in the script to the installed location:

     ```python
     pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
     ```

3. **Run the script:**

   ```bash
   python your_script_name.py
   ```

## 🚀 How It Works

The script continuously runs and performs two main tasks:

1. **Window Title Monitoring:**
   Enumerates all open windows on the system and looks for window titles starting with the specified prefixes. Logs when these windows open or close.

2. **In-Window OCR Text Recognition:**
   When a target window is detected, it captures a screenshot of the window's content. Then, it applies OCR to extract visible text from UI elements such as labels, textboxes, and buttons. If the keyword `"ES"` is found in the extracted text, a popup is shown asking the user whether to keep the window open or close it automatically.

## ⚠️ Notes

* This tool works **only on Windows OS**.
* Make sure Tesseract OCR is installed and correctly configured.
* Running with sufficient permissions may be required to access window contents.


## 👨‍💻 Author

*Muhammad Sami Naeem* — contributions and feedback are welcome!
