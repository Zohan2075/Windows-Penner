import ctypes
import pyautogui
import win32con
import win32gui
import tkinter as tk
from tkinter import messagebox

user32 = ctypes.windll.user32

pinned_windows = []

def list_open_windows():
    windows = []

    def enum_windows_callback(hwnd, lParam):
        if user32.IsWindowVisible(hwnd) and user32.GetParent(hwnd) == 0:
            length = user32.GetWindowTextLengthW(hwnd) + 1
            window_title = ctypes.create_unicode_buffer(length)
            user32.GetWindowTextW(hwnd, window_title, length)
            windows.append((hwnd, window_title.value))

        return True

    # Define the WNDENUMPROC callback function using WINFUNCTYPE
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(enum_windows_callback), 0)
    return windows

def pin_window(app_title):
    try:
        hwnd_to_pin = None
        for hwnd, title in list_open_windows():
            if title.lower() == app_title.lower():
                hwnd_to_pin = hwnd
                break

        if not hwnd_to_pin:
            messagebox.showerror("Window Not Found", f"Window with title '{app_title}' not found.")
            return

        # Restore and bring the window to the foreground
        user32.ShowWindow(hwnd_to_pin, win32con.SW_RESTORE)
        user32.SetForegroundWindow(hwnd_to_pin)

        # Pin the window on top of all other applications
        win32gui.SetWindowPos(hwnd_to_pin, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

        pinned_windows.append(hwnd_to_pin)

        messagebox.showinfo("Window Pinned", f"Window with title '{app_title}' pinned over other applications.")
    except Exception as e:
        messagebox.showerror("Error", f"Error: {e}")

def stop_pinning():
    if pinned_windows:
        for hwnd in pinned_windows:
            # Unpin the window and bring it to the foreground
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            user32.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
            user32.SetForegroundWindow(hwnd)
        pinned_windows.clear()
    else:
        messagebox.showinfo("No Pinned Windows", "No windows are currently pinned.")

def on_pin_button_click():
    app_title_or_number = window_title_entry.get().strip()
    try:
        # Try to convert the input to an integer (window number)
        window_number = int(app_title_or_number)
        if 1 <= window_number <= len(pinned_windows):
            hwnd_to_unpin = pinned_windows.pop(window_number - 1)
            # Unpin the window and bring it to the foreground
            win32gui.SetWindowPos(hwnd_to_unpin, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            user32.ShowWindow(hwnd_to_unpin, win32con.SW_SHOWNORMAL)
            user32.SetForegroundWindow(hwnd_to_unpin)
        else:
            messagebox.showerror("Invalid Window Number", "Invalid window number.")
    except ValueError:
        # If conversion to int fails, treat the input as a window title
        pin_window(app_title_or_number)

def display_open_windows():
    windows = list_open_windows()
    window_list.delete(1, tk.END)
    for i, (hwnd, title) in enumerate(windows, start=1):
        window_list.insert(tk.END, f"{i}. {title}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Window Pinner")

    window_title_label = tk.Label(root, text="Enter the title or number of the window to pin/unpin:")
    window_title_label.pack(pady=10)

    window_title_entry = tk.Entry(root, width=50)
    window_title_entry.pack(pady=5)

    pin_button = tk.Button(root, text="Pin/Unpin Window", command=on_pin_button_click)
    pin_button.pack(pady=10)

    windows_label = tk.Label(root, text="List of Open Windows:")
    windows_label.pack(pady=5)

    window_list = tk.Listbox(root, width=70)
    window_list.pack()

    refresh_button = tk.Button(root, text="Refresh List", command=display_open_windows)
    refresh_button.pack(pady=10)

    stop_pin_button = tk.Button(root, text="Stop Pinning", command=stop_pinning)
    stop_pin_button.pack(pady=5)

    display_open_windows()

    root.mainloop()
