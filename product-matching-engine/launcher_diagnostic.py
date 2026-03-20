import sys
import os
import traceback
from streamlit.web import cli as stcli

def get_path(filename):
    """Gets the absolute path to a file, handling PyInstaller's temporary folder."""
    if hasattr(sys, "_MEIPASS"):
        # Running in a PyInstaller bundle
        return os.path.join(sys._MEIPASS, filename)
    else:
        # Running in a normal Python environment
        return filename

def main():
    """The main entry point for the diagnostic launcher."""
    try:
        print("=== ProductMatcher Diagnostic Launcher ===")
        print(f"Python version: {sys.version}")
        print(f"Platform: {sys.platform}")
        print(f"Executable: {sys.executable}")
        
        # Check if running in PyInstaller bundle
        if hasattr(sys, '_MEIPASS'):
            print(f"PyInstaller temp directory: {sys._MEIPASS}")
            os.chdir(sys._MEIPASS)
            print(f"Changed working directory to: {os.getcwd()}")
        else:
            print("Running in normal Python environment")
        
        # Check for app.py
        app_py_path = get_path("app.py")
        print(f"Looking for app.py at: {app_py_path}")
        
        if os.path.exists(app_py_path):
            print("✓ app.py found")
        else:
            print("✗ app.py NOT found")
            print("Contents of current directory:")
            for item in os.listdir('.'):
                print(f"  - {item}")
            input("Press Enter to exit...")
            return
        
        # Test imports
        print("\nTesting imports...")
        try:
            import streamlit
            print("✓ streamlit imported successfully")
        except Exception as e:
            print(f"✗ streamlit import failed: {e}")
            
        try:
            import pandas
            print("✓ pandas imported successfully")
        except Exception as e:
            print(f"✗ pandas import failed: {e}")
            
        try:
            from thefuzz import fuzz
            print("✓ thefuzz.fuzz imported successfully")
        except Exception as e:
            print(f"✗ thefuzz.fuzz import failed: {e}")
            
        try:
            from sklearn.feature_extraction.text import TfidfVectorizer
            print("✓ sklearn imported successfully")
        except Exception as e:
            print(f"✗ sklearn import failed: {e}")
            
        try:
            import numpy
            print("✓ numpy imported successfully")
        except Exception as e:
            print(f"✗ numpy import failed: {e}")
            
        try:
            import multiprocessing
            print(f"✓ multiprocessing imported successfully (CPU count: {multiprocessing.cpu_count()})")
        except Exception as e:
            print(f"✗ multiprocessing import failed: {e}")
            
        try:
            from concurrent.futures import ProcessPoolExecutor
            print("✓ concurrent.futures imported successfully")
        except Exception as e:
            print(f"✗ concurrent.futures import failed: {e}")
        
        print("\nStarting Streamlit...")
        print("Arguments:", ["streamlit", "run", app_py_path, "--server.port=8501", "--global.developmentMode=false"])
        
        # Set up Streamlit arguments
        sys.argv = ["streamlit", "run", app_py_path, "--server.port=8501", "--global.developmentMode=false"]
        
        # Call Streamlit
        sys.exit(stcli.main())
        
    except Exception as e:
        print(f"\n!!! FATAL ERROR !!!")
        print(f"Error: {e}")
        print(f"Error type: {type(e).__name__}")
        print("\nFull traceback:")
        traceback.print_exc()
        print("\nPress Enter to exit...")
        input()

if __name__ == "__main__":
    main()