import os
import tkinter as tk
from tkinter import filedialog

def prompt_for_file(title: str, filetypes: list) -> str:
    """
    Opens a native Windows file selection dialog.
    This is extremely user-friendly for non-technical users.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the small empty tkinter window
    root.attributes('-topmost', True)  # Force dialog to appear on top of other windows
    
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=filetypes
    )
    
    root.destroy()
    return file_path
