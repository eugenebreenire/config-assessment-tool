import logging
import os
import subprocess
import sys
import json
from http.server import BaseHTTPRequestHandler, HTTPServer

def getPlatform():
    platform = sys.platform
    if sys.platform == "linux":
        proc_version = open("/proc/version").read()
        if "microsoft" in proc_version:
            platform = "wsl"
    return platform

def openFolder(path):
    logging.info("Opening folder: " + path)
    platform = getPlatform()
    try:
        if platform == "darwin":
            subprocess.call(["open", "--", path])
        elif platform in ["win64", "win32"]:
            subprocess.call(["start", path])
        elif platform == "wsl":
            command = "explorer.exe `wslpath -w " + path + "`"
            subprocess.run(["bash", "-c", command])
        else:
            subprocess.call(["xdg-open", path])
    except Exception as e:
        logging.error("Error opening folder: " + str(e))

class MyServer(BaseHTTPRequestHandler):
    def do_POST(self):
        if self.path == "/open_folder":
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            try:
                data = json.loads(post_data)
                folder_path = data.get("path")
                if folder_path:
                    openFolder(folder_path)
                    self.send_response(200)
                    self.end_headers()
                    self.wfile.write(b"Folder opened")
                else:
                    self.send_response(400)
                    self.end_headers()
                    self.wfile.write(b"Missing folder path")
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode())
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        logging.info("%s - - [%s] %s" % (self.address_string(), self.log_date_time_string(), format % args))

if __name__ == "__main__":
    if not os.path.exists("logs"):
        os.makedirs("logs")
    if not os.path.exists("output"):
        os.makedirs("output")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler("logs/config-assessment-tool-frontend.log"), logging.StreamHandler()],
    )
    hostName = "localhost"
    serverPort = 16225
    try:
        logging.info("Starting FileHandler on " + hostName + ":" + str(serverPort))
        webServer = HTTPServer((hostName, serverPort), MyServer)
        webServer.serve_forever()
    except KeyboardInterrupt:
        logging.info("Stopping FileHandler")