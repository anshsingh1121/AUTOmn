"""
Launcher for IMT Dashboard Application
Starts the Streamlit server and automatically opens the user's browser.
"""

import os
import sys

# Force environment variables before Streamlit is imported or initialized
os.environ["STREAMLIT_SERVER_PORT"] = "8501"
os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
os.environ["STREAMLIT_BROWSER_SERVER_PORT"] = "8501"
os.environ["STREAMLIT_BROWSER_SERVER_ADDRESS"] = "127.0.0.1"
os.environ["STREAMLIT_SERVER_HEADLESS"] = "false"

import webbrowser
import threading
import time
from streamlit.web import bootstrap

def open_browser():
    time.sleep(2)
    webbrowser.open("http://127.0.0.1:8501")

def main():
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            app_path = os.path.join(sys._MEIPASS, "app.py")
            if not os.path.exists(app_path):
                app_path = os.path.join(os.path.dirname(sys.executable), "app.py")
        else:
            app_path = os.path.join(os.path.dirname(sys.executable), "app.py")
    else:
        app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")

    # Launch browser automatically to port 8501
    threading.Thread(target=open_browser, daemon=True).start()

    flag_options = {
        "server.port": 8501,
        "server.address": "127.0.0.1",
        "browser.serverAddress": "127.0.0.1",
        "browser.serverPort": 8501,
        "server.headless": False,
        "global.developmentMode": False,
        "browser.gatherUsageStats": False,
    }

    bootstrap.run(app_path, False, [], flag_options)

if __name__ == "__main__":
    main()
