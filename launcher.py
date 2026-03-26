import threading
import webbrowser
import time
import sys
import os

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def open_browser():
    time.sleep(2.5)
    webbrowser.open("http://localhost:5000")

if __name__ == "__main__":
    try:
        from app import app
        app.template_folder = resource_path("templates")
        t = threading.Thread(target=open_browser)
        t.daemon = True
        t.start()
        app.run(port=5000, debug=False, use_reloader=False)
    except Exception as e:
        print("\nERROR:", e)
        import traceback
        traceback.print_exc()
        input("\nPress Enter to close...")
