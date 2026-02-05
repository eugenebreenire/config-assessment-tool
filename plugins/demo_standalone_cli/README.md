# Demo Standalone CLI Plugin

This plugin demonstrates a purely standalone tool that does not participate in the main assessment lifecycle.

## Overview

- **Name**: `demo_standalone_cli`
- **Type**: Standalone CLI
- **Entry Point**: `main.py`

## Features

-   **No `run_plugin` function**: The Engine ignores this plugin during normal assessment runs.
-   **CLI Only**: Designated for utility scripts, specific one-off tasks, or tools that don't need assessment context.
-   **Argument Parsing**: Demonstrates how to receive arguments passed from the CLI wrapper.

## How to Run

Run via the main script:
```bash
./cat.sh --plugin start demo_standalone_cli
```

Pass arguments:
```bash
./cat.sh --plugin start demo_standalone_cli arg1 arg2
```
