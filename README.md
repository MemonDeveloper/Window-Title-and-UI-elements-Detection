# 🧐 ES Order Window Monitor

A Python-based automation tool for monitoring trading-related windows on Windows OS. This script detects the presence of specific window titles and performs OCR (Optical Character Recognition) to find the keyword "ES". Based on detection, it prompts the user to continue or automatically closes the window.

## 💡 Features

- Monitors windows titled:
  - `Order Ticket - ES`
  - `Order Preview`
- Captures the window's screen content
- Performs OCR using [Tesseract](https://github.com/tesseract-ocr/tesseract)
- Automatically prompts the user if `"ES"` is detected
- Optionally closes the window based on user input
- Logs timestamps for each action

## 🛠️ Technologies Used

- Python 3.x
- `ctypes`, `datetime`, `win32gui`, `win32con`, `win32api`, `win32ui`, `win32process` (via `pywin32`)
- `PIL` (Python Imaging Library via `Pillow`)
- `pytesseract` for OCR

## 🔧 Setup Instructions

1. **Install dependencies**:
   ```bash
   pip install pywin32 pillow pytesseract
````

2. **Install Tesseract OCR**:

   * Download from: [https://github.com/tesseract-ocr/tesseract](https://github.com/tesseract-ocr/tesseract)
   * Update the path in the script:

     ```python
     pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
     ```

## 🚀 How It Works

* The script continuously checks open windows.
* If it finds a window that matches the title pattern, it captures the window content.
* OCR is applied to detect `"ES"` in the content.
* If `"ES"` is found, a warning popup is shown:

  * Clicking **Yes** continues.
  * Clicking **No** automatically closes the window.

## ⚠️ Notes

* Make sure Tesseract is properly installed and its path is set
* Requires admin privileges in some environments

## 👨‍💻 Author

*Muhammad Sami Memon* — Contributions, issues, and suggestions are welcome!
