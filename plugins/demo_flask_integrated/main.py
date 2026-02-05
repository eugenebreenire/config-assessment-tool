import logging
import sys
import os
import json
from datetime import datetime
# Note: We do NOT import flask here ensures compatibility with the main Engine process

ASSESSMENT_DATA = {}

def run_plugin(context):
    """
    Integrated Entry Point.
    Called by Engine.py (Main Process).
    """
    logging.info("Demo Flask Integrated: Bootstrapping...")

    # 1. Capture context into global state
    ASSESSMENT_DATA["jobFileName"] = context.get("jobFileName", "Unknown")
    ASSESSMENT_DATA["outputDir"] = context.get("outputDir", "Unknown")
    ASSESSMENT_DATA["controllerData"] = context.get("controllerData", {})
    ASSESSMENT_DATA["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 2. Determine Port
    port = 5002  # Different default to avoid conflict with others

    # 3. Log and Print URL
    url = f"http://127.0.0.1:{port}"
    logging.info(f"Demo Flask Integrated Plugin: Launching Web UI at {url}")
    print(f"\n[Plugin] Assessment Complete! View report at: {url}")

    # 4. Start Blocking Server
    # Note: access logs are printed to stderr by default.
    try:
        # Import defined Flask app (assuming it is in the same directory and Flask is installed)
        try:
            from .server import app, ASSESSMENT_DATA as SERVER_DATA
            # Sync context
            SERVER_DATA.update(ASSESSMENT_DATA)
            app.run(host='0.0.0.0', port=port, debug=False)
        except ImportError:
             # Fallback if running relative
             try:
                 from server import app, ASSESSMENT_DATA as SERVER_DATA
                 SERVER_DATA.update(ASSESSMENT_DATA)
                 app.run(host='0.0.0.0', port=port, debug=False)
             except ImportError as e:
                 logging.error(f"Could not import Flask server app: {e}")

    except OSError as e:
        logging.error(f"Could not start plugin server on port {port}: {e}")

    logging.info("Demo Flask Integrated Plugin: Finished.")
    return "Finished"

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    # Create dummy data for standalone testing
    dummy_context = {
        "jobFileName": "Standalone_Test_Job",
        "outputDir": "/tmp/cat/output/test",
        "controllerData": {}
    }

    print("Running in Standalone Mode with dummy data...")
    run_plugin(dummy_context)
