import os
import json
import time
import asyncio
import logging
from datetime import datetime

import requests
import streamlit as st
from tzlocal import get_localzone
from streamlit_modal import Modal

from utils.streamlit_utils import rerun

# --- Helper Functions ---

def is_running_in_container():
    """
    Checks if the code is running inside a container by inspecting
    the environment. This works for Docker, containerd, and other
    runtimes that use standard cgroup paths.
    """
    # Check for a common file created by Docker.
    if os.path.exists('/.dockerenv'):
        return True

    # Check the cgroup of the init process for container-specific keywords.
    try:
        with open('/proc/1/cgroup', 'rt') as f:
            cgroup_content = f.read()
            if 'docker' in cgroup_content or 'kubepods' in cgroup_content:
                return True
    except FileNotFoundError:
        # /proc/1/cgroup does not exist on non-Linux systems.
        pass

    return False

def get_file_path(base, name):
    return f"input/{base}/{name}.json"

def handle_open_jobfile(file_path, title):
    # This function now only displays the file content
    if os.path.exists(file_path):
        with st.expander(f"📂 {title}", expanded=True):
            with open(file_path) as f:
                data = json.load(f)
                st.json(data)
    else:
        st.warning(f"File not found: {file_path}")

def show_thresholds_file(thresholds):
    file_path = get_file_path("thresholds", thresholds)
    if os.path.exists(file_path):
        with st.expander(f"📂 {thresholds}.json", expanded=True):
            with open(file_path) as f:
                data = json.load(f)
                formatted = json.dumps(data, indent=2)
                st.markdown(
                    f"""
                    <div style="max-height: 240px; overflow-y: scroll; border: 1px solid #ccc; padding: 8px; background-color: #f9f9f9;">
                    <pre>{formatted}</pre>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
    else:
        st.warning(f"File not found: {file_path}")

def open_output_folder(jobName):
    """Sends a POST request to the FileHandler server to open a folder."""
    relative_path = f"output/{jobName}"

    # Dynamically determine the host based on the environment
    if is_running_in_container():
        # 'host.docker.internal' is a special DNS name that resolves to the host's IP.
        default_host = "host.docker.internal"
    else:
        # When running locally, connect to localhost.
        default_host = "localhost"

    # Allow overriding with an environment variable for flexibility
    file_handler_host = os.getenv("FILE_HANDLER_HOST", default_host)

    url = f"http://{file_handler_host}:16225/open_folder"
    try:
        response = requests.post(url, json={"path": relative_path}, timeout=5)
        response.raise_for_status()
        logging.info(f"Successfully requested to open folder: {relative_path}")
    except requests.exceptions.RequestException as e:
        st.error(f"Could not connect to FileHandler server at {url}. Is it running?")
        logging.error(f"Failed to open folder via FileHandler server: {e}")

def dynamic_credentials_section(job_executed, jobName):
    dynamicCredentials = st.expander("Pass credentials dynamically (optional)")
    dynamicCredentials.write("Credentials will be changed for all jobs in the job file.")
    usrNameCol, pwdCol, authTypeCol, dynChckCol = dynamicCredentials.columns(4)
    authType = authTypeCol.selectbox(
        label="Auth Type",
        options=["basic", "secret", "token"],
        key=f"JobFile:{jobName}-authType"
    )
    labels = {
        "token": ("API Client Username", "API Client Token"),
        "secret": ("Client ID", "Client Secret"),
        "basic": ("New Username", "New Password")
    }
    username_label, password_label = labels.get(authType, labels["basic"])
    newUsrName = usrNameCol.text_input(
        label=username_label,
        value="AzureDiamond",
        key=f"JobFile:{jobName}-usrCol"
    )
    newPwd = pwdCol.text_input(
        label=password_label,
        value="hunter2",
        type="password",
        key=f"JobFile:{jobName}-pwdCol"
    )
    dynChckCol.text("")
    dynChckCol.text("")
    dynamicCheck = dynChckCol.checkbox("Dynamic Credentials", key=f"JobFile:{jobName}-chckCol")
    return newUsrName, newPwd, authType, dynamicCheck

def handle_run(runColumn, jobName, thresholds, debug, concurrentConnections, newUsrName, newPwd, authType, dynamicCheck):
    runColumn.text("")
    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        username = newUsrName if dynamicCheck else None
        password = newPwd if dynamicCheck else None
        auth_method = authType if dynamicCheck else None

        async def run_main():
            from backend.core.Engine import Engine
            engine = Engine(jobName, thresholds, concurrentConnections, username, password, auth_method)
            await engine.run()

        try:
            st.session_state.running_job = jobName
            asyncio.run(run_main())
        except SystemExit as e:
            if e.code == 0:
                st.success(f"Job '{jobName}' executed successfully.")
                time.sleep(1)
            else:
                st.error(f"Job execution failed with exit code: {e.code}")
                st.exception(e)
        except Exception as e:
            st.error(f"Job execution failed: {e}")
            st.exception(e)
        finally:
            if 'running_job' in st.session_state:
                del st.session_state.running_job
            rerun()

def tail_file(filepath, n_lines=50):
    """Reads the last N lines from a file."""
    try:
        with open(filepath, "r") as f:
            lines = f.readlines()
            return "".join(lines[-n_lines:])
    except FileNotFoundError:
        return "Log file not found."
    except Exception as e:
        return f"Error reading log file: {e}"

# --- Main Component ---

def jobHandler(jobName: str, debug: bool, concurrentConnections: int):
    st.header(f"{jobName}")

    col_job_file, col_thresholds_file, col_output_folder = st.columns([1, 1, 1])

    # Column 1: Job File
    col_job_file.text("")
    col_job_file.text("")
    if col_job_file.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        handle_open_jobfile(f"input/jobs/{jobName}.json", f"{jobName}.json")

    # Column 2: Thresholds File
    col_thresholds_file.text("")
    col_thresholds_file.text("")

    # Column 3: Output Folder
    info_path = f"output/{jobName}/info.json"
    job_executed = os.path.exists(info_path)
    if job_executed:
        col_output_folder.text("")
        col_output_folder.text("")
        if col_output_folder.button(f"Open Output Folder", key=f"{jobName}-outputFolder"):
            open_output_folder(jobName)

    # Dynamic Credentials
    newUsrName, newPwd, authType, dynamicCheck = dynamic_credentials_section(job_executed, jobName)

    # Thresholds Selection
    thresholds_dir = "input/thresholds"
    thresholdsFiles = []
    if os.path.exists(thresholds_dir) and os.path.isdir(thresholds_dir):
        thresholdsFiles = [f[:-5] for f in os.listdir(thresholds_dir) if f.endswith('.json')]

    if jobName in thresholdsFiles:
        thresholdsFiles.remove(jobName)
        thresholdsFiles.insert(0, jobName)
    elif "DefaultThresholds" in thresholdsFiles:
        thresholdsFiles.remove("DefaultThresholds")
        thresholdsFiles.insert(0, "DefaultThresholds")

    # Main Action Row
    thresholdsColumn, infoColumn, runColumn = st.columns([1, 1, 0.3])

    if job_executed:
        try:
            with open(info_path) as f:
                info = json.load(f)
            last_run_str = datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")
            infoColumn.text("")
            infoColumn.info(f'Last Run: {last_run_str}')
        except (IOError, json.JSONDecodeError, KeyError):
            infoColumn.text("")
            infoColumn.warning("Job has not yet been run or info file is invalid.")
    else:
        infoColumn.text("")
        infoColumn.warning("Job has not yet been run")

    if thresholdsFiles:
        thresholds = thresholdsColumn.selectbox("Specify Thresholds File", thresholdsFiles, index=0, key=f"{jobName}-new")
        # Connect the button in col_thresholds_file to its action here
        if col_thresholds_file.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
            show_thresholds_file(thresholds)
        handle_run(runColumn, jobName, thresholds, debug, concurrentConnections, newUsrName, newPwd, authType, dynamicCheck)
    else:
        thresholdsColumn.warning("No threshold files found in `input/thresholds`.")

    # Log viewer Modal
    log_modal = Modal(f"Logs for {jobName}", key=f"logs-modal-{jobName}", max_width=1000)
    if st.button("Show Logs", key=f"show-logs-{jobName}"):
        log_modal.open()

    if log_modal.is_open():
        with log_modal.container():
            log_file = "logs/config-assessment-tool.log"
            log_placeholder = st.empty()
            is_running = st.session_state.get("running_job") == jobName

            log_container_id = f"log-container-{jobName.replace(' ', '-')}"
            js_autoscroll = f"""
                <script>
                    var elem = document.getElementById('{log_container_id}');
                    if (elem) {{
                        elem.scrollTop = elem.scrollHeight;
                    }}
                </script>
            """

            while is_running:
                log_content = tail_file(log_file, 100)
                log_html = f"""
                    <div id="{log_container_id}" style="height: 400px; overflow-y: scroll; border: 1px solid #ccc; padding: 10px; background-color: #f0f2f6; font-family: monospace; white-space: pre-wrap;">{log_content}</div>
                    {js_autoscroll}
                """
                log_placeholder.markdown(log_html, unsafe_allow_html=True)

                if st.session_state.get("running_job") != jobName:
                    is_running = False
                else:
                    time.sleep(2)

            # Display final log state after run
            log_content = tail_file(log_file, 200)
            log_html = f"""
                <div style="height: 400px; overflow-y: scroll; border: 1px solid #ccc; padding: 10px; background-color: #f0f2f6; font-family: monospace; white-space: pre-wrap;">{log_content}</div>
            """
            log_placeholder.markdown(log_html, unsafe_allow_html=True)