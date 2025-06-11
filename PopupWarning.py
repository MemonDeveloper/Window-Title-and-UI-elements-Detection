import ctypes
import datetime
import win32gui
import win32con
import win32ui
import win32api
import tkinter as tk
import pytesseract
import threading
from PIL import Image
from ctypes import wintypes, windll
from ctypes.wintypes import HWND, LPARAM

# ✅ Set your Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Constants
prefix_order_ES = "Order Ticket - ES"
prefix_order_preview_ES = "Order Preview"
MB_YESNO = 0x00000004
MB_ICONQUESTION = 0x00000020
MB_TOPMOST = 0x00040000
MB_SYSTEMMODAL = 0x00001000  # Highest priority modal dialog
IDYES = 6
IDNO = 7
WM_CLOSE = 0x0010
user32 = windll.user32
MessageBoxW = windll.user32.MessageBoxW
FindWindowW = windll.user32.FindWindowW
GetForegroundWindow = win32gui.GetForegroundWindow

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

# # Forced popup
def show_forced_popup(message="Do you want to work on this window?", title="ES - Warning"):
    # Create owner window
    hwndOwner = win32gui.CreateWindowEx(
        win32con.WS_EX_TOPMOST,
        "Static", "", 0, 0, 0, 0, 0, None, None, None, None
    )
    
    # Use an Event for thread control
    stop_event = threading.Event()
    
    # Thread to constantly force window to front
    def force_front():
        while not stop_event.is_set():
            try:
                win32gui.SetWindowPos(
                    hwndOwner,
                    win32con.HWND_TOPMOST,
                    0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
                )
            except:
                break
    
    t = threading.Thread(target=force_front)
    t.daemon = True
    t.start()
    
    try:
        # Show message box
        result = user32.MessageBoxW(
            hwndOwner,
            message,
            title,
            0x00000004 | 0x00000020 | 0x00040000 | 0x00001000  # MB_YESNO | MB_ICONQUESTION | MB_TOPMOST | MB_SYSTEMMODAL
        )
        return result
    finally:
        # Proper cleanup sequence
        stop_event.set()  # Signal thread to stop
        t.join(timeout=0.5)  # Wait for thread to finish
        
        # Final cleanup
        try:
            win32gui.DestroyWindow(hwndOwner)
        except:
            pass

                ##### OCR ####
def show_waiting_popup():
    stop_event = threading.Event()

    def create_popup():
        global wait_hwnd

        # Calculate center position
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)
        width = 250
        height = 100
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        # Create the popup window
        wait_hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_TOPMOST,
            "Static",
            "Scanning...\nPlease wait.",
            win32con.WS_OVERLAPPEDWINDOW | win32con.WS_VISIBLE,
            x, y, width, height,
            0, 0, 0, None
        )
        win32gui.UpdateWindow(wait_hwnd)

        # Thread to keep it always on top
        def force_front():
            while not stop_event.is_set():
                try:
                    win32gui.SetWindowPos(
                        wait_hwnd,
                        win32con.HWND_TOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
                    )
                except:
                    break

        threading.Thread(target=force_front, daemon=True).start()

        try:
            win32gui.PumpMessages()
        except:
            pass

    t = threading.Thread(target=create_popup, daemon=True)
    t.start()

    return stop_event

def close_waiting_popup(stop_event):
    global wait_hwnd
    if wait_hwnd:
        try:
            win32gui.PostMessage(wait_hwnd, win32con.WM_CLOSE, 0, 0)
        except:
            pass
        wait_hwnd = None

    stop_event.set()

# # OCR from target window image
def ocr_check_es_on_screen(hwnd):
    global es_popup_shown
    if es_popup_shown:
        return
    
    # Show the popup
    stop_event = show_waiting_popup()
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

        # Use PrintWindow instead of BitBlt
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)

        if not result:
            print_with_timestamp("PrintWindow failed. Falling back to BitBlt.")
            saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1
        )

        # Optional: save raw capture for debugging
        img.save("captured_window_raw.png")
        print_with_timestamp("OCR Step 2: Preprocessing image...")

        # Preprocess: grayscale → threshold → optional upscale
        gray = img.convert('L')
        bw = gray.point(lambda x: 0 if x < 180 else 255, '1')
        upscaled = bw.resize((width * 2, height * 2), Image.LANCZOS)

        # Optional: save preprocessed image for debugging
        upscaled.save("captured_window_preprocessed.png")

        print_with_timestamp("OCR Step 3: Extracting text...")
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(upscaled, config=custom_config)

        print_with_timestamp(f"OCR Step 4: Text extracted:\n{text.strip()}")

        # Step 4: Decision popup
        if "ES" in text:
            close_waiting_popup(stop_event)
            result = show_forced_popup()
            if result == IDNO:
                win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)
            es_popup_shown = True
        else:
            # Clean up any leftover popups
            close_waiting_popup(stop_event)

    except Exception as e:
        print_with_timestamp(f"OCR ERROR: {e}")
        close_waiting_popup()

def enum_windows_proc(hwnd, lParam):
    global window_found, window_currently_open, popup_shown, es_popup_shown

    length = 512
    buffer = ctypes.create_unicode_buffer(length)

    if GetWindowTextW(hwnd, buffer, length):
        title = buffer.value
        if prefix_order_preview_ES.lower() in title.lower():
            window_found = True
            if not window_currently_open:
                window_currently_open = True
                popup_shown = False
                es_popup_shown = False
                print_with_timestamp("✅ Window is OPEN")
                
                if not popup_shown:
                    popup_shown = True
                    print_with_timestamp("Popup shown")
                    
                    while True:
                        # Close previous popup if exists
                        existing_popup = FindWindowW(None, "ES - Warning")
                        if existing_popup:
                            PostMessageW(existing_popup, WM_CLOSE, 0, 0)
                    
                        result = ocr_check_es_on_screen(hwnd)
                        print_with_timestamp("Popup close")
                        # Handle result
                        if result == IDYES:
                            print_with_timestamp("✅ User Selected Yes")
                            break
                        elif result == IDNO:
                            PostMessageW(hwnd, WM_CLOSE, 0, 0)
                            print_with_timestamp("❌ User Selected No")
                            break
                        else:
                            break
            return False

        if prefix_order_ES.lower() in title.lower():
            window_found = True
            if not window_currently_open:
                window_currently_open = True
                popup_shown = False
                es_popup_shown = False
                print_with_timestamp("✅ Window is OPEN")
                
                if not popup_shown:
                    popup_shown = True
                    print_with_timestamp("Popup shown")
                    
                    while True:
                        # Close previous popup if exists
                        existing_popup = FindWindowW(None, "ES - Warning")
                        if existing_popup:
                            PostMessageW(existing_popup, WM_CLOSE, 0, 0)
                        
                        # Show popup and get result
                        result = show_forced_popup()
                        print_with_timestamp("Popup close")
                        # Handle result
                        if result == IDYES:
                            print_with_timestamp("✅ User Selected Yes")
                            break
                        elif result == IDNO:
                            PostMessageW(hwnd, WM_CLOSE, 0, 0)
                            print_with_timestamp("❌ User Selected No")
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
            print_with_timestamp("❌ Window is CLOSED")

if __name__ == "__main__":
    main()
