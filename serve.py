#!/usr/bin/env python3
"""
Simple local server for the portfolio.
Run:  python serve.py
Then open:  http://localhost:8080
"""
import http.server, socketserver, os, webbrowser, threading

PORT = 8080
os.chdir(os.path.dirname(os.path.abspath(__file__)))

class Handler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"  {self.address_string()} → {fmt % args}")

def open_browser():
    import time; time.sleep(0.8)
    webbrowser.open(f"http://localhost:{PORT}")

print(f"\n🌐  Portfolio server running at http://localhost:{PORT}")
print("    Press Ctrl+C to stop\n")

threading.Thread(target=open_browser, daemon=True).start()

with socketserver.TCPServer(("", PORT), Handler) as httpd:
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n\nServer stopped.")
