import sys
import os
import logging
import webbrowser
import time
import threading

# Set environment variables to disable Streamlit telemetry and prompts
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
os.environ["STREAMLIT_TELEMETRY_OPT_OUT"] = "true"
os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"
os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

# Configure logging for debugging
log_file = os.path.join(os.path.expanduser("~"), "ProductMatcher.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
    ]
)
logging.info("Starting ProductMatcher executable")

try:
    from streamlit.web import cli as stcli
    logging.info("Streamlit imported successfully")
except Exception as e:
    logging.error(f"Failed to import Streamlit: {e}")
    sys.exit(1)

def get_path(filename):
    """Gets the absolute path to a file, handling PyInstaller's temporary folder."""
    if hasattr(sys, "_MEIPASS"):
        # Running in a PyInstaller bundle
        path = os.path.join(sys._MEIPASS, filename)
        logging.info(f"PyInstaller path for {filename}: {path}")
        return path
    else:
        # Running in a normal Python environment
        return filename

def check_server_and_open_browser():
    """Check if Streamlit server is running and open browser ONCE."""
    import urllib.request
    import urllib.error
    
    # Flag to ensure we only open browser once
    browser_opened = False
    
    max_attempts = 30  # Wait up to 30 seconds
    for attempt in range(max_attempts):
        try:
            time.sleep(1)
            # Try to connect to the Streamlit server
            response = urllib.request.urlopen("http://localhost:8501", timeout=2)
            if response.status == 200 and not browser_opened:
                logging.info(f"Streamlit server is running (attempt {attempt + 1})")
                time.sleep(1)  # Give it one more second to fully load
                webbrowser.open("http://localhost:8501")
                logging.info("Browser opened successfully")
                browser_opened = True
                return
        except (urllib.error.URLError, urllib.error.HTTPError, OSError) as e:
            logging.info(f"Attempt {attempt + 1}: Server not ready yet ({e})")
            continue
    
    if not browser_opened:
        logging.error("Streamlit server failed to start after 30 seconds")
        # Try to open browser anyway, but only once
        try:
            webbrowser.open("http://localhost:8501")
            logging.info("Browser opened anyway")
        except Exception as e:
            logging.error(f"Failed to open browser: {e}")

def main():
    """The main entry point for the launcher."""
    try:
        # PyInstaller creates a temporary folder and unpacks the executable there.
        # We MUST change the current working directory to that folder so that
        # Streamlit can find the app.py script and any other assets.
        if hasattr(sys, '_MEIPASS'):
            os.chdir(sys._MEIPASS)
            logging.info(f"Changed directory to: {sys._MEIPASS}")

        # Verify app.py exists
        app_py_path = get_path("app.py")
        if not os.path.exists(app_py_path):
            logging.error(f"app.py not found at: {app_py_path}")
            sys.exit(1)
        
        logging.info(f"Found app.py at: {app_py_path}")

        # Start browser opening in a separate thread
        browser_thread = threading.Thread(target=check_server_and_open_browser, daemon=True)
        browser_thread.start()

        # This is the critical part. We are modifying the command-line arguments
        # that the script sees, to make it look like `streamlit run app.py` was called.
        sys.argv = [
            "streamlit", 
            "run", 
            app_py_path, 
            "--server.port=8501", 
            "--global.developmentMode=false",
            "--browser.gatherUsageStats=false",
            "--server.fileWatcherType=none",
            "--server.headless=true"
        ]
        
        logging.info(f"Starting Streamlit with args: {sys.argv}")
        
        # We then call Streamlit's own main function, and exit with its return code.
        sys.exit(stcli.main())
        
    except Exception as e:
        logging.error(f"Error in main: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
