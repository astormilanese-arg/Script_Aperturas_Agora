# UI.py

import sys
import os
import ast
import subprocess
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog

# --- Configuration for Look & Feel ---
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

print(">>> Running with Python:", sys.executable)

CONFIG_PATH = "Dependencies/config.py"
MAIN_SCRIPT = "Main.py"

config_data = {}
entries = {}  # Dictionary to hold the UI entry widgets

def parse_config_py():
    """Parses the config file into a dictionary."""
    global config_data
    config_data = {}
    
    # Ensure directory exists to avoid errors
    if not os.path.exists(os.path.dirname(CONFIG_PATH)):
        os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    
    # Create file if missing
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w") as f: pass

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        for line in f:
            if "=" in line:
                key, val = line.split("=", 1)
                key = key.strip()
                try:
                    val = ast.literal_eval(val.strip())
                except:
                    # Fallback if eval fails (e.g. simple string without quotes)
                    val = val.strip()
                config_data[key] = val

def write_config_py():
    """Writes the UI values back to the config file."""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        for key, entry_widget in entries.items():
            val = entry_widget.get()
            
            # Optional: Try to convert numeric strings back to int/float
            # so they aren't saved as strings like "10" instead of 10.
            if val.isdigit():
                val_to_save = int(val)
            else:
                try:
                    val_to_save = float(val)
                except ValueError:
                    val_to_save = val

            val_str = repr(val_to_save)
            f.write(f"{key} = {val_str}\n")

def on_start_pressed():
    """Saves config and runs the main script."""
    write_config_py()
    print(">>> Saving config and starting Main.py...")
    # Use sys.executable to ensure we use the same python interpreter
    subprocess.Popen([sys.executable, MAIN_SCRIPT])

def browse_path(key, entry_widget):
    """Opens a file or folder picker based on the config key name."""
    key_lower = key.lower()
    
    # 1. Save locations -> Ask where to save
    if 'save' in key_lower or 'export' in key_lower or 'output' in key_lower:
        # Default extension based on context
        def_ext = ".xlsx"
        filetypes = [("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        
        if "ppt" in key_lower or "presentation" in key_lower:
            def_ext = ".pptx"
            filetypes = [("PowerPoint Files", "*.pptx"), ("All Files", "*.*")]

        result = filedialog.asksaveasfilename(
            defaultextension=def_ext,
            filetypes=filetypes
        )
        
    # 2. Input files -> Ask to open existing file
    elif any(x in key_lower for x in ['app', 'template', 'raw', 'titulos', 'file']):
        filetypes = [("All Files", "*.*")]
        # Heuristics for specific types
        if "xlsx" in key_lower or "excel" in key_lower or "app" in key_lower or "titulos" in key_lower:
            filetypes = [("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
        elif "ppt" in key_lower:
            filetypes = [("PowerPoint Files", "*.pptx;*.ppt"), ("All Files", "*.*")]
            
        result = filedialog.askopenfilename(filetypes=filetypes)
        
    # 3. Directories -> Ask to pick folder
    elif 'dir' in key_lower or 'folder' in key_lower:
        result = filedialog.askdirectory()
        
    # 4. Fallback -> Open generic file
    else:
         result = filedialog.askopenfilename()

    if result:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, result)

import threading

# ... existing code ...

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Config Editor")
        self.geometry("850x800")  # Increased height for console

        # Title Label
        self.label = ctk.CTkLabel(self, text="Configuration Settings", font=("Roboto", 24, "bold"))
        self.label.pack(pady=(20, 10))

        # Scrollable Frame for config items
        self.scroll_frame = ctk.CTkScrollableFrame(self, width=800, height=400)
        self.scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)

        parse_config_py()

        # Dynamically create rows for each config item
        for key, val in config_data.items():
            # Container for each row
            row = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
            row.pack(fill="x", pady=5)

            # Label
            lbl = ctk.CTkLabel(row, text=key, width=150, anchor="w", font=("Roboto", 14))
            lbl.pack(side="left", padx=(10, 5))

            # Input Entry
            entry = ctk.CTkEntry(row, width=400)
            entry.insert(0, str(val))
            entry.pack(side="left", fill="x", expand=True, padx=5)
            
            # Store entry reference to retrieve later
            entries[key] = entry

            # Heuristic: If key looks like a path, add a "Browse" button
            is_path = any(x in key.lower() for x in ['path', 'dir', 'folder', 'location', 'titulos'])
            
            if is_path:
                btn_browse = ctk.CTkButton(
                    row, 
                    text="📂", 
                    width=40, 
                    command=lambda k=key, e=entry: browse_path(k, e)
                )
                btn_browse.pack(side="right", padx=(5, 10))

        # Start Button
        self.btn_start = ctk.CTkButton(
            self, 
            text="Save & Start", 
            command=self.on_start_pressed, 
            height=40,
            width=200,
            font=("Roboto", 16, "bold"),
            fg_color="#2CC985", hover_color="#229C68"
        )
        self.btn_start.pack(pady=10)

        # Console Label
        self.console_label = ctk.CTkLabel(self, text="Process Output", font=("Roboto", 14, "bold"), anchor="w")
        self.console_label.pack(fill="x", padx=20, pady=(10, 0))

        # Console Output (Text Box)
        self.console_box = ctk.CTkTextbox(self, width=800, height=200, font=("Consolas", 12))
        self.console_box.pack(pady=(5, 20), padx=20, fill="x")
        self.console_box.configure(state="disabled")

    def append_log(self, text):
        """Thread-safe way to append text to the console box."""
        self.console_box.configure(state="normal")
        self.console_box.insert("end", text)
        self.console_box.see("end")
        self.console_box.configure(state="disabled")

    def run_script_thread(self):
        """Runs the script in a subprocess and captures output."""
        try:
            # Prepare startup info to hide console window on Windows
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            process = subprocess.Popen(
                [sys.executable, MAIN_SCRIPT],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True,
                encoding="utf-8",
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            for line in process.stdout:
                # Schedule the UI update on the main thread
                self.after(0, self.append_log, line)
            
            process.wait()
            self.after(0, self.append_log, f"\n>>> Process finished with exit code {process.returncode}\n")
            
            # Re-enable button
            self.after(0, lambda: self.btn_start.configure(state="normal", text="Save & Start"))

        except Exception as e:
            self.after(0, self.append_log, f"\n>>> Error running script: {e}\n")
            self.after(0, lambda: self.btn_start.configure(state="normal", text="Save & Start"))

    def on_start_pressed(self):
        """Saves config and starts the thread."""
        write_config_py()
        self.console_box.configure(state="normal")
        self.console_box.delete("1.0", "end")
        self.console_box.insert("end", ">>> Saving config and starting Main.py...\n")
        self.console_box.configure(state="disabled")
        
        # Disable button while running
        self.btn_start.configure(state="disabled", text="Running...")

        # Start execution in a separate thread
        threading.Thread(target=self.run_script_thread, daemon=True).start()

if __name__ == "__main__":
    app = App()
    app.mainloop()
