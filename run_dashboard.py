"""
Launcher for IMT Dashboard Application
Starts the Streamlit server and automatically opens the user's browser.
"""

import os
import sys
import webbrowser
import threading
import time
from streamlit.web import bootstrap

def main():
    if getattr(sys, "frozen", False):
        # Running as PyInstaller executable
        if hasattr(sys, "_MEIPASS"):
            app_path = os.path.join(sys._MEIPASS, "app.py")
            if not os.path.exists(app_path):
                app_path = os.path.join(os.path.dirname(sys.executable), "app.py")
        else:
            app_path = os.path.join(os.path.dirname(sys.executable), "app.py")
    else:
        app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")

    flag_options = {
        "server.headless": False,
        "server.port": 8501,
        "global.developmentMode": False,
        "browser.gatherUsageStats": False,
    }

    bootstrap.run(app_path, False, [], flag_options)

if __name__ == "__main__":
    main()
