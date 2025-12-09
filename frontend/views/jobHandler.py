import os
import json
from datetime import datetime
from urllib import parse

import requests
import streamlit as st
from docker import APIClient
from tzlocal import get_localzone
from utils.docker_utils import isDocker, runConfigAssessmentTool
from FileHandler import openFolder

def get_file_path(base, name):
    return f"{base}/{name}.json" if isDocker() else f"../{base}/{name}.json"

def handle_open_jobfile(file_path, title):
    if os.path.exists(file_path):
        with st.expander(f"📂 {title}", expanded=True):
            with open(file_path) as f:
                data = json.load(f)
                st.json(data)
    else:
        st.warning(f"File not found: {file_path}")

def show_thresholds_file(thresholds):
    file_path = get_file_path("input/thresholds", thresholds)
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
    if not isDocker():
        openFolder(f"../output/{jobName}")
    else:
        payload = {"type": "folder", "path": f"output/{jobName}"}
        payload = parse.urlencode(payload)
        requests.get(f"http://host.docker.internal:16225?{payload}")

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
        "token": ("API Client Username" if job_executed else "Client Name", "API Client Token" if job_executed else "Temporary Access Token"),
        "secret": ("Client ID" if job_executed else "Client Name", "Client Secret"),
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

def handle_run(runColumn, client, jobName, thresholds, debug, concurrentConnections, newUsrName, newPwd, authType, dynamicCheck, platformStr, tag):
    runColumn.text("")
    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        username = newUsrName if dynamicCheck else None
        password = newPwd if dynamicCheck else None
        auth_method = authType if dynamicCheck else None
        runConfigAssessmentTool(client, jobName, thresholds, debug,
                                concurrentConnections, username, password,
                                auth_method)

def handle_open_thresholds(col_thresholds_file, jobName, thresholds):
    if col_thresholds_file.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
        show_thresholds_file(thresholds)

def jobHandler(client: APIClient, jobName: str, debug: bool, concurrentConnections: int, platformStr: str, tag: str):
    st.header(f"{jobName}")

    col_job_file, col_thresholds_file, col_output_folder = st.columns([1, 1, 1])
    col_job_file.text("")
    col_job_file.text("")
    if col_job_file.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        handle_open_jobfile(get_file_path("input/jobs", jobName), f"{jobName}.json")

    col_thresholds_file.text("")
    col_thresholds_file.text("")

    info_path = f"../output/{jobName}/info.json"
    job_executed = os.path.exists(info_path)
    if job_executed:
        col_output_folder.text("")
        col_output_folder.text("")
        if col_output_folder.button(f"Open Output Folder", key=f"{jobName}-outputFolder"):
            open_output_folder(jobName)

    newUsrName, newPwd, authType, dynamicCheck = dynamic_credentials_section(job_executed, jobName)

    thresholdsFiles = [f[:-5] for f in os.listdir("../input/thresholds")]
    if jobName in thresholdsFiles:
        thresholdsFiles.remove(jobName)
        thresholdsFiles.insert(0, jobName)
    elif "DefaultThresholds" in thresholdsFiles:
        thresholdsFiles.remove("DefaultThresholds")
        thresholdsFiles.insert(0, "DefaultThresholds")

    thresholdsColumn, infoColumn, runColumn = st.columns([1, 1, 0.3])
    if job_executed:
        try:
            info = json.loads(open(info_path).read())
            last_run_str = datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")
            infoColumn.text("")
            infoColumn.info(f'Last Run: {last_run_str}')
        except Exception:
            infoColumn.text("")
            infoColumn.warning("Job has not yet been run")
    else:
        infoColumn.text("")
        infoColumn.warning("Job has not yet been run")

    thresholds = thresholdsColumn.selectbox("Specify Thresholds File", thresholdsFiles, index=0, key=f"{jobName}-new")
    handle_open_thresholds(col_thresholds_file, jobName, thresholds)
    handle_run(runColumn, client, jobName, thresholds, debug, concurrentConnections, newUsrName, newPwd, authType, dynamicCheck, platformStr, tag)