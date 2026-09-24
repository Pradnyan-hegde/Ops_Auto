"""
Launch Script for Ops_Auto Reconciliation Web Dashboard.
Starts FastAPI server on http://127.0.0.1:8000 and opens browser automatically.
"""
import os
import sys
import webbrowser
import threading
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import uvicorn
from dashboard.server import app


def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:8000")


def main():
    print("=" * 80)
    print("STARTING OPS_AUTO RECONCILIATION DASHBOARD")
    print("=" * 80)
    print("Dashboard URL: http://127.0.0.1:8000")
    print("Press CTRL+C to stop the server.\n")

    threading.Thread(target=open_browser, daemon=True).start()
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")


if __name__ == "__main__":
    main()
