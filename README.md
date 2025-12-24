Window Pinner

A Windows desktop application that allows you to pin and unpin application windows as "Always on Top" with administrative privileges.



Features

* Window Pinning: Make any window stay on top of other windows
* Visual Interface: Dual-pane display showing pinned vs. unpinned windows
* Admin Privileges: Automatically requests elevation for full system control
* Modern UI: Dark theme with custom styling using CustomTkinter
* Quick Actions: Double-click to pin/unpin, dedicated buttons for all operations



Tech Stack

* Python 3
* CustomTkinter (modern UI framework)
* Windows API (ctypes, win32gui, win32con)
* Pandas (optional, not used in current version)



How to Run

bash

python Pin\_windows\_renewed.pyw



**Note: On Windows, this will trigger a UAC prompt for administrative privileges.**



Key Capabilities

* Window Detection: Lists all visible top-level windows on the system
* Pin/Unpin Toggle: Click or double-click to toggle window "always on top" status
* Batch Operations: Unpin all windows with a single button
* Auto-Refresh: Refresh button to update window list in real-time
* DPI Awareness: Proper scaling for high-resolution displays



Controls

📌 Pin Window: Pin selected window from unpinned list
📤 Unpin Window: Unpin selected window from pinned list
🛑 Unpin All: Remove "always on top" from all pinned windows
🔄 Refresh: Update the window lists
* Double-Click: Toggle pin status on any window in either list



Technical Details

* Elevation: Automatically relaunches with admin rights using ShellExecuteW with "runas" verb
* Window Management: Uses Windows API functions (EnumWindows, SetWindowPos, GetWindowTextW)



UI Components:

* Two listboxes for pinned/unpinned windows
* Color-coded status messages (blue=info, green=success, yellow=warning, red=error)
* Dark purple-blue theme for reduced eye strain



System Requirements

* Windows 10/11
* Python 3.7+
* Administrative privileges (app will request elevation)



Safety Features

* Error Handling: Graceful error messages for failed operations
* Resource Management: Proper cleanup of window handles
* User Feedback: Clear status messages for all actions





