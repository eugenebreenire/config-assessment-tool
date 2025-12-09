import os
import sys
from pathlib import Path
from utils.docker_utils import get_image_tags

import docker
import streamlit as st

from utils.docker_utils import getImage, isDocker
from views.header import header
from views.jobHandler import jobHandler


def main():
    client = docker.from_env().api

    if not os.path.exists("../output"):
        os.makedirs("../output")

    debug, throttleNetworkConnections = header()

    if throttleNetworkConnections:
        concurrentNetworkConnections = st.sidebar.number_input("Concurrent Network Connections", min_value=1, max_value=100, value=50)
    else:
        concurrentNetworkConnections = 50

    # Use get_image_tags utility function
    backend_image_tag, _ = get_image_tags()

    # does docker image 'config_assessment_tool:latest' exist
    if getImage(client, backend_image_tag) is None:
        st.write(f"Image {backend_image_tag} not found")
        st.write("Please build from source with --build or use 'make' with included Makefile.")
        st.write("In order to --build you will need to download the full source")
    else:
        orderedJobs = []
        for jobName in os.listdir("../input/jobs"):
            if jobName.startswith("."):
                continue

            jobName = jobName[: len(jobName) - 5]
            if Path(f"../output/{jobName}/info.json").exists():
                orderedJobs.insert(0, jobName)
            else:
                orderedJobs.append(jobName)

        for jobName in orderedJobs:
            backend_image_tag, frontend_image_tag = get_image_tags()
            jobHandler(client, jobName, debug, concurrentNetworkConnections, frontend_image_tag, backend_image_tag)
            st.markdown("""---""")

main()