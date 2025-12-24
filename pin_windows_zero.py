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
            set_message(f"Window with title '{app_title}' not found.")
            return

        # Restore and bring the window to the foreground
        user32.ShowWindow(hwnd_to_pin, win32con.SW_RESTORE)
        user32.SetForegroundWindow(hwnd_to_pin)

        if hwnd_to_pin in pinned_windows:
            # Unpin the window and bring it to the foreground
            win32gui.SetWindowPos(hwnd_to_pin, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            user32.ShowWindow(hwnd_to_pin, win32con.SW_SHOWNORMAL)
            user32.SetForegroundWindow(hwnd_to_pin)
            pinned_windows.remove(hwnd_to_pin)
            set_message(f"Window with title '{app_title}' unpinned.")
        else:
            # Pin the window on top of all other applications
            win32gui.SetWindowPos(hwnd_to_pin, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            pinned_windows.append(hwnd_to_pin)
            set_message(f"Window with title '{app_title}' pinned over other applications.")
    except Exception as e:
        set_message(f"Error: {e}")

def stop_pinning():
    if pinned_windows:
        for hwnd in pinned_windows:
            # Unpin the window and bring it to the foreground
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            user32.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
            user32.SetForegroundWindow(hwnd)
        pinned_windows.clear()
        set_message("All windows unpinned.")
    else:
        set_message("No windows are currently pinned.")

def move_to_pinned():
    selected_index = unpinned_list.curselection()
    if selected_index:
        window_title = unpinned_list.get(selected_index)
        window_title = window_title.split(". ", 1)[1]  # Remove the number from the window title
        pin_window(window_title)
    display_open_windows()

def move_to_unpinned():
    selected_index = pinned_list.curselection()
    if selected_index:
        window_title = pinned_list.get(selected_index)
        window_title = window_title.split(". ", 1)[1]  # Remove the number from the window title
        pin_window(window_title)
    display_open_windows()

def on_list_double_click(event):
    selected_index = pinned_list.curselection() or unpinned_list.curselection()
    if selected_index:
        window_title = pinned_list.get(selected_index) or unpinned_list.get(selected_index)
        window_title = window_title.split(". ", 1)[1]  # Remove the number from the window title
        pin_window(window_title)
    display_open_windows()

def display_open_windows():
    pinned_list.delete(0, tk.END)
    unpinned_list.delete(0, tk.END)
    windows = list_open_windows()
    pinned_set = set(pinned_windows)
    for i, (hwnd, title) in enumerate(windows, start=1):
        if hwnd in pinned_set:
            pinned_list.insert(tk.END, f"{i}. {title}")
        else:
            unpinned_list.insert(tk.END, f"{i}. {title}")

def set_message(message):
    message_label.config(text=message)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Window Pinner")

    windows_frame = tk.Frame(root)
    windows_frame.pack(pady=5)

    pinned_list_label = tk.Label(windows_frame, text="Pinned Windows:")
    pinned_list_label.pack(side=tk.LEFT, padx=10)

    scrollbar = tk.Scrollbar(windows_frame, orient=tk.VERTICAL)
    pinned_list = tk.Listbox(windows_frame, width=40, yscrollcommand=scrollbar.set)
    pinned_list.pack(side=tk.LEFT)
    scrollbar.config(command=pinned_list.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    unpinned_list_label = tk.Label(windows_frame, text="Unpinned Windows:")
    unpinned_list_label.pack(side=tk.LEFT, padx=10)

    scrollbar = tk.Scrollbar(windows_frame, orient=tk.VERTICAL)
    unpinned_list = tk.Listbox(windows_frame, width=40, yscrollcommand=scrollbar.set)
    unpinned_list.pack(side=tk.LEFT)
    scrollbar.config(command=unpinned_list.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    pinned_list.bind("<Double-Button-1>", on_list_double_click)
    unpinned_list.bind("<Double-Button-1>", on_list_double_click)

    button_frame = tk.Frame(root)
    button_frame.pack(pady=5)

    pin_button = tk.Button(button_frame, text="Pin", command=move_to_pinned)
    pin_button.pack(side=tk.LEFT, padx=10)

    unpin_button = tk.Button(button_frame, text="Unpin", command=move_to_unpinned)
    unpin_button.pack(side=tk.LEFT)

    stop_pin_button = tk.Button(root, text="Stop Pinning", command=stop_pinning)
    stop_pin_button.pack(pady=5)

    message_label = tk.Label(root, text="", fg="blue")
    message_label.pack(pady=5)

    display_open_windows()

    root.mainloop()
