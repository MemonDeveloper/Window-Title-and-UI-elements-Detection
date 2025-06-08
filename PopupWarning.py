import ctypes
import datetime
import time
import win32gui
import win32con
import win32process
import win32api
import win32ui
from ctypes import wintypes
from PIL import Image
import pytesseract

# ✅ Set your Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Constants
prefix_order_ES = "Order Ticket - ES"
prefix_order_preview_ES = "Order Preview"
MB_YESNO = 0x04
MB_ICONQUESTION = 0x20
MB_TOPMOST = 0x40000
IDYES = 6
IDNO = 7
WM_CLOSE = 0x0010

# Global flags
window_currently_open = False
window_found = False
popup_shown = False
es_popup_shown = False

# Windows API setup
user32 = ctypes.WinDLL('user32', use_last_error=True)

EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
EnumWindows.argtypes = [EnumWindowsProc, wintypes.LPARAM]

GetWindowTextW = user32.GetWindowTextW
GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

MessageBoxW = user32.MessageBoxW
MessageBoxW.argtypes = [wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.UINT]
MessageBoxW.restype = ctypes.c_int

FindWindowW = user32.FindWindowW
FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
FindWindowW.restype = wintypes.HWND

PostMessageW = user32.PostMessageW
PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
PostMessageW.restype = wintypes.BOOL

# Timestamp logger
def get_current_timestamp():
    now = datetime.datetime.now()
    return now.strftime("%b %d %Y %H:%M:%S.") + f"{now.microsecond // 1000:03d}"

def print_with_timestamp(message):
    print(f"{message} [{get_current_timestamp()}]")

# Forced popup
def show_forced_popup(message="Do you want to work on this window?", title="ES - Warning"):
    fg_hwnd = win32gui.GetForegroundWindow()
    fg_thread, _ = win32process.GetWindowThreadProcessId(fg_hwnd)
    this_thread = win32api.GetCurrentThreadId()
    user32.AttachThreadInput(this_thread, fg_thread, True)

    try:
        dummy_hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_TOPMOST,
            "Static",
            "",
            win32con.WS_POPUP,
            0, 0, 0, 0,
            None, None, None, None
        )
        win32gui.ShowWindow(dummy_hwnd, win32con.SW_SHOWMINIMIZED)
        win32gui.SetForegroundWindow(dummy_hwnd)

        result = MessageBoxW(
            dummy_hwnd,
            message,
            title,
            MB_YESNO | MB_ICONQUESTION | MB_TOPMOST
        )

        win32gui.DestroyWindow(dummy_hwnd)
        return result
    finally:
        user32.AttachThreadInput(this_thread, fg_thread, False)

                ##### OCR ####

# OCR from target window image
def ocr_check_es_on_screen(hwnd):
    global es_popup_shown
    if es_popup_shown:
        return

    print_with_timestamp("OCR Step 1: Capturing image of target window...")

    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top

        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)

        saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1
        )

        print_with_timestamp("OCR Step 2: Extracting text...")
        text = pytesseract.image_to_string(img)

        print_with_timestamp(f"OCR Step 3: Text extracted:\n{text.strip()}")

        if "ES" in text:
            print_with_timestamp('OCR Step 4: "ES" FOUND in text.')
            print_with_timestamp('OCR Step 5: Showing popup...')
            result = show_forced_popup(
                message='"ES" detected in window.\nDo you want to continue?',
                title='OCR Alert'
            )
            if result == IDNO:
                PostMessageW(hwnd, WM_CLOSE, 0, 0)
            es_popup_shown = True
        else:
            print_with_timestamp('OCR Step 4: "ES" NOT found.')
    except Exception as e:
        print_with_timestamp(f"OCR ERROR: {e}")

# Window monitor callback
def ocr_enum_windows_proc(hwnd, lParam):
    global window_found, window_currently_open, popup_shown, es_popup_shown

    length = 512
    buffer = ctypes.create_unicode_buffer(length)

    if GetWindowTextW(hwnd, buffer, length):
        title = buffer.value
        if title.startswith(prefix_order_preview_ES):
            window_found = True

            if not window_currently_open:
                window_currently_open = True
                popup_shown = False
                es_popup_shown = False
                print_with_timestamp("Window is OPEN")

            # ✅ Directly run OCR on window (no popup here)
            ocr_check_es_on_screen(hwnd)
            return False
    return True

# Main loop
def ocr_main():
    global window_found, window_currently_open, popup_shown, es_popup_shown

    while True:
        window_found = False
        EnumWindows(EnumWindowsProc(ocr_enum_windows_proc), 0)

        if not window_found and window_currently_open:
            window_currently_open = False
            popup_shown = False
            es_popup_shown = False
            print_with_timestamp("Window is CLOSED")

def enum_windows_proc(hwnd, lParam):
    global window_found, window_currently_open, popup_shown, es_popup_shown

    length = 512
    buffer = ctypes.create_unicode_buffer(length)

    if GetWindowTextW(hwnd, buffer, length):
        title = buffer.value
        if title.startswith(prefix_order_preview_ES):
            window_found = True
            if not window_currently_open:
                window_currently_open = True
                popup_shown = False
                es_popup_shown = False
                print_with_timestamp("Window is OPEN")
            ocr_main()
            return False

        if title.startswith(prefix_order_ES):
            window_found = True
            if not window_currently_open:
                window_currently_open = True
                popup_shown = False
                es_popup_shown = False
                print_with_timestamp("Window is OPEN")
            
                if not popup_shown:
                    popup_shown = True
                    print_with_timestamp("Popup shown")
                    
                    while True:
                        # Close previous popup if exists
                        existing_popup = FindWindowW(None, "ES - Warning")
                        if existing_popup:
                            PostMessageW(existing_popup, WM_CLOSE, 0, 0)
                        
                        result = show_forced_popup()
                        
                        if result == IDYES:
                            break
                        elif result == IDNO:
                            PostMessageW(hwnd, WM_CLOSE, 0, 0)
                            break
                        else:
                            break
            return False

    return True

def main():
    global window_found, window_currently_open, popup_shown, es_popup_shown
    print_with_timestamp("Started")
    while True:
        window_found = False
        EnumWindows(EnumWindowsProc(enum_windows_proc), 0)

        if not window_found and window_currently_open:
            window_currently_open = False
            popup_shown = False
            es_popup_shown = False
            print_with_timestamp("Window is CLOSED")

if __name__ == "__main__":
    main()