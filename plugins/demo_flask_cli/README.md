# Demo Flask CLI Plugin

This plugin serves as a demonstration of how to integrate a Flask web application within the CAT CLI plugin architecture.

## Overview

- **Name**: `demo_flask_cli`
- **Type**: Python Plugin (with Virtual Environment support)
- **Entry Point**: `main.py`

## Purpose

The purpose of this plugin is to show developers how to:
1.  Bundle a Flask application inside a plugin directory.
2.  Define dependencies in a `requirements.txt` file (e.g., specific Flask versions).
3.  Launch a web server or CLI command triggered by the main application.

## How to Run

You can start this plugin using the main `cat.sh` script from the project root:

```bash
./cat.sh --plugin start demo_flask_cli
```

Or, if passing arguments is supported by the plugin's `main.py`:

```bash
./cat.sh --plugin start demo_flask_cli --port 5000
```
