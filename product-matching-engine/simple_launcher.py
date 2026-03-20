import os
import sys
import subprocess
import webbrowser
import time
import threading

def open_browser_delayed():
    """Open browser after delay"""
    time.sleep(5)
    webbrowser.open("http://localhost:8501")

def main():
    # Get the directory where the executable is located
    if getattr(sys, 'frozen', False):
        # Running as executable
        app_dir = os.path.dirname(sys.executable)
    else:
        # Running as script
        app_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Path to app.py
    app_path = os.path.join(app_dir, "app.py")
    
    # Start browser in background
    browser_thread = threading.Thread(target=open_browser_delayed, daemon=True)
    browser_thread.start()
    
    # Run streamlit directly using subprocess
    cmd = [
        sys.executable, "-m", "streamlit", "run", app_path,
        "--server.port=8501",
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        "--server.headless=true"
    ]
    
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running Streamlit: {e}")
        input("Press Enter to exit...")
    except KeyboardInterrupt:
        print("Application stopped by user")

if __name__ == "__main__":
    main()
