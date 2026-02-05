import json
import sys
from datetime import datetime

try:
    from flask import Flask, render_template_string, jsonify
except ImportError:
    print("This plugin requires Flask. Please install it with: pip install flask")
    sys.exit(1)

# Global storage for the context passed from the Engine
ASSESSMENT_DATA = {
    "jobFileName": "Unknown",
    "outputDir": "Unknown",
    "controllerData": {},
    "timestamp": None
}

app = Flask(__name__)

# Load template (omitted for brevity, same as before but minimal)
# We will read it from main.py or redefine it here.
# Better to redefine minimal or share. Let's redefine.

TEMPLATE = """
<!doctype html>
<html>
<head>
    <title>Assessment Results - {{ job_name }}</title>
    <style>
        body { font-family: sans-serif; padding: 40px; background-color: #f4f7f6; color: #333; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; }
        h1 { color: #2c3e50; border-bottom: 2px solid #eee; padding-bottom: 10px; }
        pre { background: #f8f9fa; padding: 15px; border-radius: 4px; overflow-x: auto; border: 1px solid #eee; }
        .footer { margin-top: 40px; text-align: center; color: #aaa; font-size: 12px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Assessment Report: {{ job_name }}</h1>
        <p><strong>Output Directory:</strong> {{ output_dir }}</p>

        <h2>Controllers</h2>
        <pre>{{ controllers | tojson(indent=2) }}</pre>

        <div class="footer">
            Powered by Config Assessment Tool &bull; Integrated Flask Plugin (Isolated Process)
        </div>
    </div>
</body>
</html>
"""

def load_context(filepath):
    with open(filepath, 'r') as f:
        return json.load(f)

@app.route('/')
def index():
    return render_template_string(
        TEMPLATE,
        job_name=ASSESSMENT_DATA.get("jobFileName"),
        output_dir=ASSESSMENT_DATA.get("outputDir"),
        controllers=ASSESSMENT_DATA.get("controllerData")
    )

def run_server(data_file, port=5002):
    print(f"[Server] Loading context from {data_file}")
    ctx = load_context(data_file)
    with open(ctx["controllerDataFile"], 'r', encoding='ISO-8859-1') as f:
        ctx['controllerData'] = json.load(f)

    # app.config['CTX'] = ctx # Revert this
    ASSESSMENT_DATA.update(ctx) # Use global var

    url = f"http://127.0.0.1:{port}"
    print(f"\n[Plugin] Web UI running at: {url}")
    print("[Plugin] Press CTRL+C to stop.\n")

    app.run(host='0.0.0.0', port=port, debug=False)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python server.py <json_context_file>")
        sys.exit(1)

    run_server(sys.argv[1])
