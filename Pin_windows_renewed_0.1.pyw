import ctypes
import tkinter as tk
import customtkinter as ctk
import win32con
import win32gui

# Configure customtkinter appearance
ctk.set_appearance_mode("System")  # "Light", "Dark", or "System"
ctk.set_default_color_theme("dark-blue")  # Better theme for modern look

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
            set_message(f"⚠️ Window '{app_title}' not found.", "warning")
            return

        user32.ShowWindow(hwnd_to_pin, win32con.SW_RESTORE)
        user32.SetForegroundWindow(hwnd_to_pin)

        if hwnd_to_pin in pinned_windows:
            win32gui.SetWindowPos(hwnd_to_pin, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            pinned_windows.remove(hwnd_to_pin)
            set_message(f"✅ Window '{app_title}' unpinned.", "info")
        else:
            win32gui.SetWindowPos(hwnd_to_pin, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            pinned_windows.append(hwnd_to_pin)
            set_message(f"📌 Window '{app_title}' pinned on top.", "success")
    except Exception as e:
        set_message(f"❌ Error: {e}", "error")

def stop_pinning():
    if pinned_windows:
        for hwnd in pinned_windows:
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            user32.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
            user32.SetForegroundWindow(hwnd)
        pinned_windows.clear()
        set_message("🧹 All windows unpinned.", "success")
    else:
        set_message("ℹ️ No windows are currently pinned.", "info")

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

def set_message(message, level="info"):
    colors = {
        "info": "#3498db",      # Blue
        "success": "#2ecc71",   # Green
        "warning": "#f1c40f",   # Yellow
        "error": "#e74c3c"      # Red
    }
    message_label.configure(text=message, text_color=colors.get(level, "#3498db"))

if __name__ == "__main__":
    # App initialization
    app = ctk.CTk()
    app.title("✨ Window Pinner")
    app.geometry("900x600")
    app.configure(fg_color="#1e1e2f")  # Set background to dark purple-blue

    # Frame setup
    frame = ctk.CTkFrame(app, fg_color="#29293d", corner_radius=15)
    frame.pack(padx=20, pady=20, fill="both", expand=True)

    # Titles
    title_label = ctk.CTkLabel(frame, text="Window Pinner", font=("Arial", 26, "bold"))
    title_label.grid(row=0, column=0, columnspan=2, pady=10)

    # Pinned Windows List
    pinned_label = ctk.CTkLabel(frame, text="📌 Pinned Windows", font=("Arial", 18))
    pinned_label.grid(row=1, column=0, padx=10, pady=5)

    pinned_listbox = tk.Listbox(frame, width=40, height=20, bg="#2f2f46", fg="white",
                                selectbackground="#5dade2", selectforeground="black", font=("Arial", 12))
    pinned_listbox.grid(row=2, column=0, padx=10, pady=10)
    pinned_listbox.bind("<Double-Button-1>", on_list_double_click)

    # Unpinned Windows List
    unpinned_label = ctk.CTkLabel(frame, text="🗂️ Unpinned Windows", font=("Arial", 18))
    unpinned_label.grid(row=1, column=1, padx=10, pady=5)

    unpinned_listbox = tk.Listbox(frame, width=40, height=20, bg="#2f2f46", fg="white",
                                  selectbackground="#5dade2", selectforeground="black", font=("Arial", 12))
    unpinned_listbox.grid(row=2, column=1, padx=10, pady=10)
    unpinned_listbox.bind("<Double-Button-1>", on_list_double_click)

    # Buttons
    buttons_frame = ctk.CTkFrame(app, fg_color="transparent")
    buttons_frame.pack(pady=10)

    button_style = {
        "font": ("Arial", 16, "bold"),
        "corner_radius": 12,
        "height": 45,
        "width": 160,
        "hover_color": "#2980b9"
    }

    pin_button = ctk.CTkButton(buttons_frame, text="📌 Pin Window", command=lambda: move_selected_window(unpinned_listbox), **button_style)
    pin_button.grid(row=0, column=0, padx=15)

    unpin_button = ctk.CTkButton(buttons_frame, text="📤 Unpin Window", command=lambda: move_selected_window(pinned_listbox), **button_style)
    unpin_button.grid(row=0, column=1, padx=15)

    stop_pinning_button = ctk.CTkButton(buttons_frame, text="🛑 Unpin All", command=stop_pinning, **button_style)
    stop_pinning_button.grid(row=0, column=2, padx=15)

    # Message Label
    message_label = ctk.CTkLabel(app, text="", font=("Arial", 16))
    message_label.pack(pady=10)

    # Start displaying windows
    display_open_windows()
    app.mainloop()
