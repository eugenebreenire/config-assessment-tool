#!/bin/bash

OS=$(uname -s | tr '[:upper:]' '[:lower:]')
ARCH=$(uname -m)

if [[ "$OS" == "darwin" ]]; then
  OS_TAG="macos"
elif [[ "$OS" == "linux" ]]; then
  OS_TAG="linux"
else
  echo "Unsupported OS: $OS"
  exit 1
fi

if [[ "$ARCH" == "arm64" || "$ARCH" == "aarch64" ]]; then
  ARCH_TAG="arm"
elif [[ "$ARCH" == "x86_64" ]]; then
  ARCH_TAG="amd64"
else
  echo "Unsupported architecture: $ARCH"
  exit 1
fi

REPO="ghcr.io/alexafshar/config-assessment-tool-${OS_TAG}-${ARCH_TAG}"

if [[ -f VERSION ]]; then
  VERSION=$(cat VERSION)
else
  echo "VERSION file not found."
  exit 1
fi

IMAGE="$REPO:$VERSION"
PORT="8501"
LOG_DIR="logs"
LOG_FILE="$LOG_DIR/config-assessment-tool.log"
MOUNTS="-v $(pwd)/input/jobs:/app/input/jobs -v $(pwd)/input/thresholds:/app/input/thresholds -v $(pwd)/output/archive:/app/output/archive -v $(pwd)/$LOG_DIR:/app/$LOG_DIR"
CONTAINER_NAME="cat-tool-container"

mkdir -p "$LOG_DIR"

start_filehandler() {
  if [ ! -f "frontend/FileHandler.py" ]; then
    echo "Error: frontend/FileHandler.py not found."
    exit 1
  fi
  echo "Starting FileHandler service on host..."
  pkill -f "python.*FileHandler.py" 2>/dev/null
  pipenv run python frontend/FileHandler.py >> "$LOG_FILE" 2>&1 &
  echo "FileHandler.py started with PID $!"
  sleep 2
}

case "$1" in
  --start)
    if [[ "$2" == "docker" ]]; then
      export FILE_HANDLER_HOST=host.docker.internal
      start_filehandler
      docker stop $CONTAINER_NAME >/dev/null 2>&1
      docker rm $CONTAINER_NAME >/dev/null 2>&1

      if [[ $# -eq 2 ]]; then
        echo "Starting container in UI mode..."
        CONTAINER_ID=$(docker run --add-host=host.docker.internal:host-gateway -d --name $CONTAINER_NAME -e FILE_HANDLER_HOST=$FILE_HANDLER_HOST -p $PORT:$PORT $MOUNTS $IMAGE streamlit run frontend/frontend.py --server.headless=true)
        if [ $? -eq 0 ]; then
          echo "Container started successfully with ID: $CONTAINER_ID"
          echo "UI available at http://localhost:$PORT"
          docker logs -f $CONTAINER_ID
        else
          echo "Failed to start container."
          exit 1
        fi
      else
        echo "Starting container in backend mode with args: ${@:3}"
        docker run --add-host=host.docker.internal:host-gateway --rm --name $CONTAINER_NAME -e FILE_HANDLER_HOST=$FILE_HANDLER_HOST -p $PORT:$PORT $MOUNTS $IMAGE backend "${@:3}"
        EXIT_CODE=$?
        if [ $EXIT_CODE -ne 0 ]; then
          echo "Container failed with exit code: $EXIT_CODE"
          exit $EXIT_CODE
        fi
      fi
    else
      export PYTHONPATH="$(pwd):$(pwd)/backend"

      # Ensure dependencies are installed before running
      echo "Checking/Installing dependencies..."
      pipenv install

      if [[ $# -eq 1 ]]; then
        echo "PYTHONPATH is: $PYTHONPATH"
        echo "Running application in UI mode from source..."
        echo "UI available at http://localhost:$PORT"
        pipenv run streamlit run frontend/frontend.py
      else
        echo "PYTHONPATH is: $PYTHONPATH"
        echo "Running application in backend mode from source with args: ${@:2}"
        pipenv run python backend/backend.py "${@:2}"
      fi
    fi
    ;;

  --plugin)
    if [[ "$2" == "list" ]]; then
       export PYTHONPATH="$(pwd):$(pwd)/backend"
       pipenv run python backend/plugin_manager.py list
       exit 0
    elif [[ "$2" == "docs" ]]; then
       PLUGIN_NAME="$3"
       if [[ -z "$PLUGIN_NAME" ]]; then
         echo "Error: Plugin name required."
         exit 1
       fi
       export PYTHONPATH="$(pwd):$(pwd)/backend"
       pipenv run python backend/plugin_manager.py docs "$PLUGIN_NAME"
       exit 0
    elif [[ "$2" == "start" ]]; then
       PLUGIN_NAME="$3"
       if [[ -z "$PLUGIN_NAME" ]]; then
         echo "Error: Plugin name required."
         exit 1
       fi
       export PYTHONPATH="$(pwd):$(pwd)/backend"
       # Pass remaining args to the plugin manager
       pipenv run python backend/plugin_manager.py start "$PLUGIN_NAME" "${@:4}"
       exit 0
    fi
    ;;

  shutdown)
    echo "Shutting down container: $CONTAINER_NAME"
    docker stop $CONTAINER_NAME >/dev/null 2>&1
    docker rm $CONTAINER_NAME >/dev/null 2>&1
    echo "Container stopped and removed."
    echo "Stopping FileHandler process..."
    pkill -f "python.*FileHandler.py" 2>/dev/null
    echo "FileHandler stopped."
    echo "Stopping backend process..."
    pkill -f "python.*backend.py" 2>/dev/null
    echo "Backend process stopped."
    echo "Stopping Streamlit process..."
    pkill -f "streamlit run frontend/frontend.py" 2>/dev/null
    echo "Streamlit stopped."
    ;;
  *)
    echo "Usage:"
    echo "  cat --start                # Starts CAT UI. Requires Python 3.12 and pipenv installed.  UI accessible at http://localhost:8501"
    echo "  cat --start [args]         # Starts CAT headless mode from source with [args].  Requires Python 3.12 & pipenv installed".
    echo "  cat --start docker         # Starts CAT UI using Docker. requires Docker. UI accessible at http://localhost:8501"
    echo "  cat --start docker [args]  # Starts CAT headless mode using Docker with [args]. Requires Docker installed."
    echo "  cat --plugin [list|start <plugin>|docs <plugin>]  # list plugins|run <plugin>|show docs for <plugin>"
    echo "  cat shutdown               # Stop and remove the running container and FileHandler"
    echo ""
    echo "Arguments [args]:"
    echo "  -j, --job-file <name>             Job file name (default: DefaultJob)"
    echo "  -t, --thresholds-file <name>      Thresholds file name (default: DefaultThresholds)"
    echo "  -d, --debug                       Enable debug logging"
    echo "  -c, --concurrent-connections <n>  Number of concurrent connections"
    echo "  "
    echo "  "
    exit 1
    ;;
esac