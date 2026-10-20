"""
Flipkart Bulk Listing Tool - Launcher
Launches the FastAPI dev server and opens the web application in your default browser.
"""

import sys
import webbrowser
import threading
import time
import uvicorn

def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:8000")

if __name__ == "__main__":
    print("=" * 65)
    print("  🚀 Flipkart Bulk Listing Automation Tool")
    print("  Master Seed Row 3 Mirroring + Gemini AI Unique Listing Engine")
    print("=" * 65)
    print("  Server starting at: http://127.0.0.1:8000")
    print("  Flipkart Fast Validate (CTRL+SHIFT+S) Compliance Engine Active")
    print("=" * 65)

    threading.Thread(target=open_browser, daemon=True).start()
    uvicorn.run("backend.main:app", host="127.0.0.1", port=8000, log_level="info")
