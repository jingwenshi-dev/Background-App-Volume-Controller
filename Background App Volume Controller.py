import tkinter as tk
from tkinter import ttk
from tkinter import font
import win32process
import psutil
import pygetwindow
from pygetwindow import getActiveWindow
from pycaw.pycaw import AudioUtilities


class AppVolumeController(tk.Tk):
    def __init__(self):
        super().__init__()

        # Set default font size for all widgets
        default_font = tk.font.nametofont("TkDefaultFont")
        default_font.configure(size=14)
        self.option_add("*Font", default_font)

        self.title("Background App Volume Controller")

        self.stop_button = None
        self.start_button = None
        self.APP_TITLE_entry = None
        self.app_process_name_entry = None
        self.apps_combo = None
        self.create_widgets()

        self.running = False
        self.APP_TITLE = ""
        self.APP_PROCESS_NAME = ""

    def create_widgets(self):
        # app title input
        ttk.Label(self, text="App Title:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.APP_TITLE_entry = ttk.Entry(self)
        self.APP_TITLE_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=5)

        # app process name input
        ttk.Label(self, text="App Process Name:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.app_process_name_entry = ttk.Entry(self)
        self.app_process_name_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=5)

        # Volume control
        ttk.Label(self, text="Volume:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.volume_scale = ttk.Scale(self, from_=0, to=100, orient=tk.HORIZONTAL)
        self.volume_scale.grid(row=2, column=1, sticky="ew", padx=10, pady=5)
        self.volume_scale.set(100)  # default volume

        # Start and stop buttons
        self.start_button = ttk.Button(self, text="Start", command=self.start)
        self.start_button.grid(row=3, column=0, sticky="ew", padx=10, pady=5)

        self.stop_button = ttk.Button(self, text="Stop", command=self.stop, state=tk.DISABLED)
        self.stop_button.grid(row=3, column=1, sticky="ew", padx=10, pady=5)

        # Dropdown for selecting game from running processes
        self.apps_combo = ttk.Combobox(self, values=self.get_running_apps(), postcommand=self.update_apps_list, state="readonly")
        self.apps_combo.grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        self.apps_combo.bind("<<ComboboxSelected>>", self.set_selected_app)

    def get_running_apps(self):
        windows = pygetwindow.getAllWindows()
        apps = []
        for window in windows:
            if window.title:  # Only consider windows with titles
                try:
                    hwnd = window._hWnd
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)  # Get the PID of the window
                    process_name = psutil.Process(pid).name()  # Retrieve the process name using PID
                    apps.append(f"{window.title} ({process_name})")
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, AttributeError):
                    pass
        return apps

    def update_apps_list(self):
        """Update list of running apps in the combobox."""
        self.apps_combo['values'] = self.get_running_apps()

    def set_selected_app(self, event):
        """Set the app title and process name based on the selected item in the combobox."""
        selected = self.apps_combo.get()
        if "(" in selected and ")" in selected:
            title, process = selected.rsplit(" (", 1)
            self.APP_TITLE = title.strip()
            self.APP_PROCESS_NAME = process.rstrip(')').strip()
            self.APP_TITLE_entry.delete(0, tk.END)
            self.APP_TITLE_entry.insert(0, self.APP_TITLE)
            self.app_process_name_entry.delete(0, tk.END)
            self.app_process_name_entry.insert(0, self.APP_PROCESS_NAME)

    def start(self):
        self.APP_TITLE = self.APP_TITLE_entry.get()
        self.APP_PROCESS_NAME = self.app_process_name_entry.get()

        self.start_button['state'] = tk.DISABLED
        self.stop_button['state'] = tk.NORMAL

        self.running = True
        self.monitor_app()

    def stop(self):
        self.running = False
        self.start_button['state'] = tk.NORMAL
        self.stop_button['state'] = tk.DISABLED

    def monitor_app(self):
        if not self.running:
            return

        if not self.is_app_in_foreground(self.APP_TITLE):
            self.set_volume(self.APP_PROCESS_NAME, self.volume_scale.get() / 100.0)
        else:
            self.set_volume(self.APP_PROCESS_NAME, 1.0)

        self.after(1000, self.monitor_app)  # check every second

    def set_volume(self, app_name, volume=0.0):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process and session.Process.name() == app_name:
                volume_control = session.SimpleAudioVolume
                volume_control.SetMasterVolume(volume, None)

    def is_app_in_foreground(self, APP_TITLE):
        foreground_window = getActiveWindow()
        if foreground_window is None:
            return False
        return APP_TITLE in foreground_window.title


if __name__ == "__main__":
    app = AppVolumeController()
    app.mainloop()
