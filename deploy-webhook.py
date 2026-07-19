#!/usr/bin/env python3
"""
Simple GitHub webhook listener for deployment.
Run on server: python3 deploy-webhook.py
Listens on localhost:5555 for GitHub push events and auto-deploys.
"""

import os
import subprocess
import json
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

REPO_PATH = "/root/migbot_max"
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "test-secret")
WEBHOOK_PORT = int(os.environ.get("WEBHOOK_PORT", 5555))


class DeployHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        """Handle POST requests from GitHub."""
        if self.path != "/deploy":
            self.send_response(404)
            self.end_headers()
            return
        
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        
        try:
            payload = json.loads(body)
            if payload.get('ref') == 'refs/heads/main':
                self.deploy()
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"Deployment triggered")
        except Exception as e:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(f"Error: {e}".encode())
    
    def deploy(self):
        """Execute git pull and restart services."""
        try:
            os.chdir(REPO_PATH)
            subprocess.run(["git", "pull", "origin", "main"], check=True)
            subprocess.run(["systemctl", "restart", "migbot", "migbot-bot", "webforms"], check=True)
            print("✓ Deployment successful")
        except Exception as e:
            print(f"✗ Deployment failed: {e}")


if __name__ == "__main__":
    server = HTTPServer(("127.0.0.1", WEBHOOK_PORT), DeployHandler)
    print(f"Webhook listener running on port {WEBHOOK_PORT}")
    server.serve_forever()
