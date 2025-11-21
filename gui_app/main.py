import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
from compressor import compress_excel_images
import os
import sys

import time

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("EIC")
        self.iconphoto(True, tk.PhotoImage(file=resource_path('logo.png')))
        self.geometry("280x180")
        self.resizable(False, False)

        # Style
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("green.Horizontal.TProgressbar", troughcolor='white', background='green')

        # --- Options Frame ---
        options_frame = ttk.Frame(self, padding=(10, 12, 10, 6))
        options_frame.pack(fill=tk.X)

        ttk.Label(options_frame, text="JPEG Quality (0-100):").pack(anchor=tk.W, side=tk.LEFT)
        self.quality_var = tk.IntVar(value=20)
        quality_spinbox = ttk.Spinbox(options_frame, from_=0, to=100, textvariable=self.quality_var, width=5)
        quality_spinbox.pack(anchor=tk.W, side=tk.LEFT, padx=5)

        # --- Main Frame ---
        main_frame = ttk.Frame(self, padding=(10, 6, 10, 10))
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Drop target
        self.drop_target = ttk.Label(
            main_frame, 
            text="Drag and drop .xlsx files here", 
            relief="solid",
            borderwidth=1,
            padding="20",
            anchor=tk.CENTER
        )
        self.drop_target.pack(fill=tk.X, pady=0)

        # Status Label
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, anchor=tk.CENTER)
        status_label.pack(fill=tk.X, pady=(5, 2))

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame, 
            orient="horizontal", 
            length=100, 
            mode='determinate', 
            variable=self.progress_var,
            style="green.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, pady=(2, 5))

        # Register drop target
        self.drop_target.drop_target_register(DND_FILES)
        self.drop_target.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        files = self.tk.splitlist(event.data)
        for file in files:
            if file.endswith('.xlsx'):
                # Run compression in a separate thread to avoid freezing the GUI
                thread = threading.Thread(
                    target=self.run_compression, 
                    args=(file,)
                )
                thread.start()
            else:
                self.update_status(f"Skipped: {file} (not a .xlsx file)")

    def run_compression(self, file_path):
        self.update_status(f"Compressing: {file_path}...")
        self.update_progress(0)
        try:
            # The logger will now be a function that updates the progress bar
            compress_excel_images(
                file_path,
                compression_level=self.quality_var.get(),
                logger=self.update_status,
                progress_callback=self.update_progress
            )
            self.update_status(f"Successfully compressed: {file_path}")
            self.update_progress(100)
            time.sleep(1.5)
            self.update_status("Ready")
            self.update_progress(0)
        except Exception as e:
            self.update_status(f"Error: {e}")
            self.update_progress(0)

    def update_status(self, message):
        # This method needs to be thread-safe
        self.after(0, self.status_var.set, message)

    def update_progress(self, value):
        # This method needs to be thread-safe
        self.after(0, self.progress_var.set, value)

if __name__ == "__main__":
    app = App()
    app.mainloop()
