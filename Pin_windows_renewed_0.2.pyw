import sys
import os
import ctypes
import ctypes.wintypes  # For HWND and LPARAM types

def is_admin() -> bool:
    """Check if the current process is running with administrative privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

# If not running as admin, relaunch the script with admin rights.
if os.name == "nt" and not is_admin():
    # Try to use pythonw.exe instead of python.exe (to avoid a console window)
    executable = sys.executable
    if "python.exe" in executable.lower():
        executable_candidate = executable.lower().replace("python.exe", "pythonw.exe")
        if os.path.exists(executable_candidate):
            executable = executable_candidate
    # Build command line parameters. Surround each argument with quotes.
    params = " ".join(f'"{arg}"' for arg in sys.argv)
    # ShellExecuteW with "runas" verb will prompt the user with UAC.
    ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, params, None, 1)
    if int(ret) <= 32:
        import tkinter.messagebox as messagebox
        messagebox.showerror("Elevation Error", "Failed to elevate privileges.")
    sys.exit(0)

# Reset the working directory to the script's location.
os.chdir(os.path.dirname(os.path.realpath(sys.argv[0])))

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    print(f"Failed to set DPI awareness: {e}")

import tkinter as tk
import customtkinter as ctk
import win32con
import win32gui

# Configure customtkinter appearance
ctk.set_appearance_mode("System")       # Options: "Light", "Dark", or "System"
ctk.set_default_color_theme("dark-blue")  # A modern dark-blue theme

# Initialize user32 and global storage
user32 = ctypes.windll.user32
pinned_windows = []  # List to track pinned window handles

def list_open_windows() -> list:
    """Return a list of tuples (hwnd, window_title) for all visible top-level windows."""
    windows = []

    def enum_windows_callback(hwnd, lParam):
        if user32.IsWindowVisible(hwnd) and user32.GetParent(hwnd) == 0:
            length = user32.GetWindowTextLengthW(hwnd) + 1
            window_title = ctypes.create_unicode_buffer(length)
            user32.GetWindowTextW(hwnd, window_title, length)
            if window_title.value:
                windows.append((hwnd, window_title.value))
        return True

    CALLBACK = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    callback = CALLBACK(enum_windows_callback)
    user32.EnumWindows(callback, 0)
    return windows

def pin_window(app_title: str) -> None:
    """
    Toggle the pin (always-on-top) status of a window given its title.
    If the window is visible and not pinned, it will be made topmost.
    If already pinned, its topmost status will be removed.
    """
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

def stop_pinning() -> None:
    """Unpin all currently pinned windows."""
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

def toggle_selected_window(listbox: tk.Listbox) -> None:
    """Toggle (pin/unpin) the window for the selected listbox entry."""
    selected_indices = listbox.curselection()
    if not selected_indices:
        set_message("⚠️ No window selected.", "warning")
        return
    item_text = listbox.get(selected_indices[0])
    try:
        window_title = item_text.split(". ", 1)[1]
    except IndexError:
        set_message("⚠️ Unexpected item format.", "error")
        return
    pin_window(window_title)
    display_open_windows()

def on_list_double_click(event) -> None:
    """Double-click callback to toggle the selected window."""
    toggle_selected_window(event.widget)

def display_open_windows() -> None:
    """Update the pinned and unpinned windows listboxes."""
    pinned_listbox.delete(0, "end")
    unpinned_listbox.delete(0, "end")
    current_windows = list_open_windows()
    pinned_set = set(pinned_windows)
    for i, (hwnd, title) in enumerate(current_windows, start=1):
        formatted_title = f"{i}. {title}"
        if hwnd in pinned_set:
            pinned_listbox.insert("end", formatted_title)
        else:
            unpinned_listbox.insert("end", formatted_title)

def set_message(message: str, level: str = "info") -> None:
    """
    Update the message label with a given message.
    Level can be: info (blue), success (green), warning (yellow), or error (red).
    """
    colors = {
        "info": "#3498db",
        "success": "#2ecc71",
        "warning": "#f1c40f",
        "error": "#e74c3c"
    }
    message_label.configure(text=message, text_color=colors.get(level, "#3498db"))

# --- Main Application Setup ---
if __name__ == "__main__":
    app = ctk.CTk()
    app.title("✨ Window Pinner")
    app.geometry("900x600")
    app.configure(fg_color="#1e1e2f")  # Dark purple-blue background

    # Main Frame
    frame = ctk.CTkFrame(app, fg_color="#29293d", corner_radius=15)
    frame.pack(padx=20, pady=20, fill="both", expand=True)

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

    # Controls Frame
    buttons_frame = ctk.CTkFrame(app, fg_color="transparent")
    buttons_frame.pack(pady=10)
    button_style = {
        "font": ("Arial", 16, "bold"),
        "corner_radius": 12,
        "height": 45,
        "width": 160,
        "hover_color": "#2980b9"
    }
    pin_button = ctk.CTkButton(buttons_frame, text="📌 Pin Window", command=lambda: toggle_selected_window(unpinned_listbox), **button_style)
    pin_button.grid(row=0, column=0, padx=15)
    unpin_button = ctk.CTkButton(buttons_frame, text="📤 Unpin Window", command=lambda: toggle_selected_window(pinned_listbox), **button_style)
    unpin_button.grid(row=0, column=1, padx=15)
    stop_pinning_button = ctk.CTkButton(buttons_frame, text="🛑 Unpin All", command=stop_pinning, **button_style)
    stop_pinning_button.grid(row=0, column=2, padx=15)
    refresh_button = ctk.CTkButton(buttons_frame, text="🔄 Refresh", command=display_open_windows, **button_style)
    refresh_button.grid(row=0, column=3, padx=15)

    # Message Label for user feedback
    message_label = ctk.CTkLabel(app, text="", font=("Arial", 16))
    message_label.pack(pady=10)

    # Initially display open windows
    display_open_windows()
    app.mainloop()
