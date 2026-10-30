"""
Launcher for IMT Dashboard Application
Starts the Streamlit server on localhost and automatically opens the user's default browser
as soon as the server is ready.
"""

import os
import sys
import socket
import urllib.request
import threading
import time
import webbrowser


def find_free_port(start_port: int = 8501, max_attempts: int = 30) -> int:
    """Find the first available TCP port on localhost starting at start_port."""
    for port in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start_port


def wait_and_open_browser(port: int):
    """
    Polls the local server health endpoint until it responds to HTTP requests,
    then immediately launches the default browser.
    Prevents white-screen / connection-refused errors.
    """
    health_url = f"http://127.0.0.1:{port}/_stcore/health"
    browser_url = f"http://127.0.0.1:{port}"
    for _ in range(40):  # Wait up to 20 seconds
        try:
            req = urllib.request.Request(health_url, headers={"User-Agent": "IMT-Dashboard-Launcher"})
            with urllib.request.urlopen(req, timeout=0.5) as resp:
                if resp.status in (200, 304, 404):
                    break
        except Exception:
            time.sleep(0.5)
    webbrowser.open(browser_url)


def locate_app_py() -> str:
    """Locate app.py regardless of whether running as script or PyInstaller onedir executable."""
    if getattr(sys, "frozen", False):
        internal_dir = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        app_dir = os.path.dirname(sys.executable)
        candidates = [
            os.path.join(internal_dir, "app.py"),
            os.path.join(app_dir, "app.py"),
            os.path.join(internal_dir, "_internal", "app.py"),
            os.path.join(app_dir, "_internal", "app.py"),
        ]
        for p in candidates:
            if os.path.isfile(p):
                return p
        raise FileNotFoundError(f"Could not locate 'app.py' in bundled application paths: {candidates}")
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")


def main():
    app_path = locate_app_py()

    port = find_free_port(8501)

    # Force corporate security environment variables
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_BROWSER_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_BROWSER_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "false"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_ENABLE_CORS"] = "false"
    os.environ["STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION"] = "true"

    # Start browser opener in background thread
    threading.Thread(target=wait_and_open_browser, args=(port,), daemon=True).start()

    from streamlit.web import bootstrap
    from streamlit import config

    flag_options = {
        "server.port": port,
        "server.address": "127.0.0.1",
        "browser.serverAddress": "127.0.0.1",
        "browser.serverPort": port,
        "server.headless": False,
        "global.developmentMode": False,
        "browser.gatherUsageStats": False,
    }

    bootstrap.load_config_options(flag_options)
    config.set_option("server.address", "127.0.0.1")
    config.set_option("server.port", port)
    config.set_option("server.headless", False)
    config.set_option("browser.gatherUsageStats", False)
    config.set_option("server.enableCORS", False)
    config.set_option("server.enableXsrfProtection", True)

    bootstrap.run(app_path, False, [], flag_options)


if __name__ == "__main__":
    main()
