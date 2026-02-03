#!/usr/bin/env python3
"""
Simple HTTP server with CORS support for Word Add-in development
Usage: python3 server-cors.py
"""

from http.server import HTTPServer, SimpleHTTPRequestHandler
import os

class CORSRequestHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        # Add CORS headers
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, HEAD')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')
        self.send_header('Access-Control-Max-Age', '3600')
        super().end_headers()
    
    def do_OPTIONS(self):
        # Handle preflight requests
        self.send_response(200)
        self.end_headers()
    
    def log_message(self, format, *args):
        # Custom logging
        print(f"[{self.address_string()}] {format % args}")

if __name__ == '__main__':
    port = 3000
    server_address = ('localhost', port)
    httpd = HTTPServer(server_address, CORSRequestHandler)
    
    print(f"""
╔══════════════════════════════════════════════════════════════╗
║  CodeText Word Add-in Development Server                    ║
╚══════════════════════════════════════════════════════════════╝

Server running on: http://localhost:{port}
CORS enabled: Yes

Available files:
  - http://localhost:{port}/taskpane.html
  - http://localhost:{port}/debug.html
  - http://localhost:{port}/manifest-http.xml

Press Ctrl+C to stop the server
""")
    
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n\nServer stopped.")

