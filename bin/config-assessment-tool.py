#!/usr/bin/env python3
import logging
import os
import subprocess
import sys
import time
import zipfile
from http.client import RemoteDisconnected
from platform import uname
from urllib.error import URLError
from urllib.request import urlopen

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from frontend.utils.docker_utils import get_image_tags, get_arch, get_version

def run(path: str):
    arch = get_arch()
    version = get_version()
    backend_image_tag, frontend_image_tag = get_image_tags()

    splash = """
                          __ _                                                              _        _              _
          ___ ___  _ __  / _(_) __ _        __ _ ___ ___  ___  ___ ___ _ __ ___   ___ _ __ | |_     | |_ ___   ___ | |
         / __/ _ \| '_ \| |_| |/ _` |_____ / _` / __/ __|/ _ \/ __/ __| '_ ` _ \ / _ \ '_ \| __|____| __/ _ \ / _ \| |
        | (_| (_) | | | |  _| | (_| |_____| (_| \__ \__ \  __/\__ \__ \ | | | | |  __/ | | | ||_____| || (_) | (_) | |
         \___\___/|_| |_|_| |_|\__, |      \__,_|___/___/\___||___/___/_| |_| |_|\___|_| |_|\__|     \__\___/ \___/|_|
                               |___/
    """
    logging.info(splash)

    # Check if config-assessment-tool images exist
    if (
        runBlockingCommand(f"docker images -q {frontend_image_tag}") == ""
        or runBlockingCommand(f"docker images -q {backend_image_tag}") == ""
    ):
        logging.info("Necessary Docker images not found.")
        build()
    else:
        logging.info("Necessary Docker images found.")

    # stop FileHandler
    logging.info("Terminating FileHandler if already running")
    if sys.platform == "win32":
        runBlockingCommand("WMIC path win32_process Where \"name like '%python%' and CommandLine like '%FileHandler.py%'\" CALL TERMINATE")
    else:
        runBlockingCommand("pgrep -f 'FileHandler.py' | xargs kill")

    # stop config-assessment-tool-frontend
    logging.info(f"Terminating config-assessment-tool-frontend-{arch} container if already running")
    containerId = runBlockingCommand("docker ps -f name=config-assessment-tool-frontend-" + arch + ' --format "{{.ID}}"')
    if containerId:
        runBlockingCommand(f"docker container stop {containerId}")

    # start FileHandler
    logging.info("Starting FileHandler")
    runNonBlockingCommand(f"{sys.executable} frontend/FileHandler.py")

    # wait for file handler to start
    while True:
        logging.info("Waiting for FileHandler to start on http://localhost:16225")
        try:
            if urlopen("http://localhost:16225/ping").read() == b"pong":
                logging.info("FileHandler started")
                break
        except URLError:
            pass
        time.sleep(1)

    # start config-assessment-tool-frontend
    logging.info(f"Starting config-assessment-tool-frontend-{arch} container")
    runNonBlockingCommand(
        f"docker run "
        f'--name "config-assessment-tool-frontend-{arch}" '
        f"-v /var/run/docker.sock:/var/run/docker.sock "
        f'-v "{path}/logs:/logs" '
        f'-v "{path}/output:/output" '
        f'-v "{path}/input:/input" '
        f'-e HOST_ROOT="{path}" '
        f'-e PLATFORM_STR="{arch}" '
        f'-e TAG="{version}" '
        f"-p 8501:8501 "
        f"--rm "
        f"{frontend_image_tag} &"
    )

    # wait for config-assessment-tool-frontend to start
    while True:
        logging.info(f"Waiting for config-assessment-tool-frontend-{arch}:{version} to start")
        try:
            if urlopen("http://localhost:8501").status == 200:
                logging.info(f"config-assessment-tool-frontend-{arch}:{version} started")
                break
        except (URLError, RemoteDisconnected):
            pass
        time.sleep(1)

    # open web browser platform specific
    if arch == "windows":
        runBlockingCommand("start http://localhost:8501")
    elif "microsoft" in uname().release.lower():
        runBlockingCommand("wslview http://localhost:8501")
    elif "macos" in arch:
        runBlockingCommand("open http://localhost:8501")
    elif arch == "linux":
        runBlockingCommand("xdg-open http://localhost:8501")
    else:
        logging.info("Unsupported platform, trying to open web browser to http://localhost:8501")
        runBlockingCommand("open http://localhost:8501")

    # Loop until user exits
    logging.info("Press Ctrl-C to stop the config-assessment-tool")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("Received KeyboardInterrupt")
        logging.info(f"Terminating config-assessment-tool-frontend-{arch}:{version} container if still running")
        containerId = runBlockingCommand(f"docker ps -f name=config-assessment-tool-frontend-{arch}:{version}" + ' --format "{{.ID}}"')
        if containerId:
            runBlockingCommand(f"docker container stop {containerId}")

def build():
    backend_image_tag, frontend_image_tag = get_image_tags()

    if os.path.isfile("backend/Dockerfile") and os.path.isfile("frontend/Dockerfile"):
        logging.info(f"Building {backend_image_tag} from Dockerfile")
        runBlockingCommand(f"docker build --no-cache -t {backend_image_tag} -f backend/Dockerfile .")
        logging.info(f"Building {frontend_image_tag} from Dockerfile")
        runBlockingCommand(f"docker build --no-cache -t {frontend_image_tag} -f frontend/Dockerfile .")
    else:
        logging.info("Dockerfiles not found in either backend/ or frontend/.")
        logging.info("Please either clone the full repository to build the images manually.")

    # Check if images exist
    if (
        runBlockingCommand(f"docker images -q {frontend_image_tag}") == ""
        or runBlockingCommand(f"docker images -q {backend_image_tag}") == ""
    ):
        logging.info("Failed to build Docker images.")
        sys.exit(1)

def pull():
    backend_image_tag, frontend_image_tag = get_image_tags()
    logging.info(f"Pulling {backend_image_tag}")
    runBlockingCommand(f"docker pull {backend_image_tag}")
    logging.info(f"Pulling {frontend_image_tag}")
    runBlockingCommand(f"docker pull {frontend_image_tag}")

def package():
    logging.info("Creating zip file")
    with zipfile.ZipFile("config-assessment-tool-dist.zip", "w") as zip_file:
        zip_file.write("README.md")
        zip_file.write("VERSION")
        zip_file.write("bin/config-assessment-tool.py")
        zip_file.write("input/jobs/DefaultJob.json")
        zip_file.write("input/thresholds/DefaultThresholds.json")
        zip_file.write("frontend/FileHandler.py")
    logging.info("Created config-assessment-tool-dist.zip")

def runBlockingCommand(command: str):
    output = ""
    with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=None, shell=True) as process:
        line = process.communicate()[0].decode("ISO-8859-1").strip()
        if line:
            logging.debug(line)
        output += line
    return output.strip()

def runNonBlockingCommand(command: str):
    subprocess.Popen(command, stdout=None, stderr=None, shell=True)

def verifySoftwareVersion() -> str:
    if sys.platform == "win32":
        latestTag = runBlockingCommand(
            'powershell -Command "(Invoke-WebRequest https://api.github.com/repos/appdynamics/config-assessment-tool/tags | ConvertFrom-Json)[0].name"'
        )
    else:
        latestTag = runBlockingCommand(
            "curl -s https://api.github.com/repos/appdynamics/config-assessment-tool/tags | grep 'name' | head -n 1 | cut -d ':' -f 2 | cut -d '\"' -f 2"
        )

    logging.info(f"Latest release tag from https://api.github.com/repos/appdynamics/config-assessment-tool/tags is {latestTag}")

    localTag = get_version()

    if latestTag != localTag:
        logging.warning(f"You are using an outdated version of the software. Current {localTag} Target {latestTag}")
        logging.warning("You can get the latest version from https://github.com/Appdynamics/config-assessment-tool/releases")
    else:
        logging.info(f"You are using the latest version of the software. Current {localTag}")

    return localTag

if __name__ == "__main__":
    assert sys.version_info >= (3, 5), "Python 3.5 or higher required"

    path = os.path.realpath(f"{__file__}/../..")
    os.chdir(path)

    if not os.path.exists("logs"):
        os.makedirs("logs")
    if not os.path.exists("output"):
        os.makedirs("output")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("logs/config-assessment-tool-frontend.log"),
            logging.StreamHandler(),
        ],
    )

    logging.info(f"Working directory is {os.getcwd()}")

    verifySoftwareVersion()

    if len(sys.argv) == 1 or sys.argv[1] == "--help":
        msg = """
    Usage: config-assessment-tool.py [OPTIONS]
    Options:
      --run, Run the config-assessment-tool
      --build, Build frontend and backend from Dockerfile
      --package, Create lightweight package for distribution
      --help, Show this message and exit.
              """.strip()
        print(msg)
        sys.exit(1)
    if sys.argv[1] == "--run":
        run(path)
    elif sys.argv[1] == "--build":
        build()
    elif sys.argv[1] == "--package":
        package()
    else:
        print(f"Unknown option: {sys.argv[1]}")
        print("Use --help for usage information")
        sys.exit(1)

    sys.exit(0)