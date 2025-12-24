import ctypes
import tkinter as tk
import customtkinter as ctk
import win32con
import win32gui

# Configure customtkinter appearance
ctk.set_appearance_mode("System")  # "Light", "Dark", or "System"
ctk.set_default_color_theme("blue")  # Other options: "green", "dark-blue"

# Initialize user32
user32 = ctypes.windll.user32
pinned_windows = []

def list_open_windows():
    windows = []

    def enum_windows_callback(hwnd, lParam):
        if user32.IsWindowVisible(hwnd) and user32.GetParent(hwnd) == 0:
            length = user32.GetWindowTextLengthW(hwnd) + 1
            window_title = ctypes.create_unicode_buffer(length)
            user32.GetWindowTextW(hwnd, window_title, length)
            if window_title.value:
                windows.append((hwnd, window_title.value))
        return True

    callback = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)(enum_windows_callback)
    user32.EnumWindows(callback, 0)
    return windows

def pin_window(app_title):
    try:
        hwnd_to_pin = next((hwnd for hwnd, title in list_open_windows() if title.lower() == app_title.lower()), None)

        if not hwnd_to_pin:
            set_message(f"Window '{app_title}' not found.")
            return

        user32.ShowWindow(hwnd_to_pin, win32con.SW_RESTORE)
        user32.SetForegroundWindow(hwnd_to_pin)

        if hwnd_to_pin in pinned_windows:
            win32gui.SetWindowPos(hwnd_to_pin, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            pinned_windows.remove(hwnd_to_pin)
            set_message(f"Window '{app_title}' unpinned.")
        else:
            win32gui.SetWindowPos(hwnd_to_pin, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            pinned_windows.append(hwnd_to_pin)
            set_message(f"Window '{app_title}' pinned on top.")
    except Exception as e:
        set_message(f"Error: {e}")

def stop_pinning():
    if pinned_windows:
        for hwnd in pinned_windows:
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            user32.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
            user32.SetForegroundWindow(hwnd)
        pinned_windows.clear()
        set_message("All windows unpinned.")
    else:
        set_message("No windows are currently pinned.")

def move_selected_window(source_listbox):
    selected_index = source_listbox.curselection()
    if selected_index:
        item = source_listbox.get(selected_index)
        window_title = item.split(". ", 1)[1]
        pin_window(window_title)
        display_open_windows()

def on_list_double_click(event):
    move_selected_window(event.widget)

def display_open_windows():
    pinned_listbox.delete(0, "end")
    unpinned_listbox.delete(0, "end")

    windows = list_open_windows()
    pinned_set = set(pinned_windows)

    for i, (hwnd, title) in enumerate(windows, start=1):
        formatted_title = f"{i}. {title}"
        if hwnd in pinned_set:
            pinned_listbox.insert("end", formatted_title)
        else:
            unpinned_listbox.insert("end", formatted_title)

def set_message(message):
    message_label.configure(text=message)

if __name__ == "__main__":
    # App initialization
    app = ctk.CTk()
    app.title("Window Pinner")
    app.geometry("800x500")

    # Frame setup
    frame = ctk.CTkFrame(app)
    frame.pack(padx=20, pady=20, fill="both", expand=True)

    # Pinned Windows List
    pinned_label = ctk.CTkLabel(frame, text="Pinned Windows")
    pinned_label.grid(row=0, column=0, padx=10, pady=5)

    pinned_listbox = tk.Listbox(frame, width=40, height=20)
    pinned_listbox.grid(row=1, column=0, padx=10, pady=10)
    pinned_listbox.bind("<Double-Button-1>", on_list_double_click)

    # Unpinned Windows List
    unpinned_label = ctk.CTkLabel(frame, text="Unpinned Windows")
    unpinned_label.grid(row=0, column=1, padx=10, pady=5)

    unpinned_listbox = tk.Listbox(frame, width=40, height=20)
    unpinned_listbox.grid(row=1, column=1, padx=10, pady=10)
    unpinned_listbox.bind("<Double-Button-1>", on_list_double_click)

    # Buttons
    buttons_frame = ctk.CTkFrame(app)
    buttons_frame.pack(pady=10)

    pin_button = ctk.CTkButton(buttons_frame, text="Pin", command=lambda: move_selected_window(unpinned_listbox))
    pin_button.grid(row=0, column=0, padx=10)

    unpin_button = ctk.CTkButton(buttons_frame, text="Unpin", command=lambda: move_selected_window(pinned_listbox))
    unpin_button.grid(row=0, column=1, padx=10)

    stop_pinning_button = ctk.CTkButton(buttons_frame, text="Stop All Pinning", command=stop_pinning)
    stop_pinning_button.grid(row=0, column=2, padx=10)

    # Message Label
    message_label = ctk.CTkLabel(app, text="", text_color="blue")
    message_label.pack(pady=5)

    # Start displaying windows
    display_open_windows()
    app.mainloop()
