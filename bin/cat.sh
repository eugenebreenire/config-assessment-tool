#!/bin/bash

OS=$(uname -s | tr '[:upper:]' '[:lower:]')
ARCH=$(uname -m)

# uses host.docker.internal for Docker container to access host services.
# for others like containerd, use host IP directly.
export FILE_HANDLER_HOST=host.docker.internal

# Map OS and ARCH to repo format
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

REPO="ghcr.io/appdynamics/config-assessment-tool-${OS_TAG}-${ARCH_TAG}"
VERSION=$(cat VERSION)
IMAGE="$REPO:$VERSION"
PORT="8501"
LOG_DIR="logs"
LOG_FILE="$LOG_DIR/config-assessment-tool.log"
MOUNTS="-v $(pwd)/input/jobs:/app/input/jobs -v $(pwd)/input/thresholds:/app/input/thresholds -v $(pwd)/output/archive:/app/output/archive -v $(pwd)/$LOG_DIR:/app/$LOG_DIR"
CONTAINER_NAME="cat-tool-container"

# Create log directory if it doesn't exist
mkdir -p "$LOG_DIR"

start_filehandler() {
  if [ ! -f "frontend/FileHandler.py" ]; then
    echo "Error: frontend/FileHandler.py not found."
    exit 1
  fi
  echo "Starting FileHandler service on host..."
  pkill -f "python.*FileHandler.py" 2>/dev/null
  # Append FileHandler output to the log file
  python frontend/FileHandler.py >> "$LOG_FILE" 2>&1 &
  echo "FileHandler.py started with PID $!"
  sleep 2 # Give the server a moment to start
}

if [[ "$1" == "run" && "$2" == "--docker" ]]; then
  start_filehandler

  # Stop and remove any existing container with the same name to avoid conflicts
  docker stop $CONTAINER_NAME >/dev/null 2>&1
  docker rm $CONTAINER_NAME >/dev/null 2>&1

  if [[ "$3" == "ui" ]]; then
    echo "Starting container in UI mode..."
    CONTAINER_ID=$(docker run -d --name $CONTAINER_NAME -e FILE_HANDLER_HOST=$FILE_HANDLER_HOST -p $PORT:$PORT $MOUNTS $IMAGE streamlit run frontend/frontend.py --server.headless=true)
  else
    echo "Starting container in CLI mode..."
    # Pass all arguments after 'run --docker' to the backend
    CONTAINER_ID=$(docker run -d --name $CONTAINER_NAME -e FILE_HANDLER_HOST=$FILE_HANDLER_HOST $MOUNTS $IMAGE backend "${@:3}")
  fi

  if [ $? -eq 0 ]; then
    echo "Container started successfully with ID: $CONTAINER_ID"
    if [[ "$3" == "ui" ]]; then
      echo "You can now view your Streamlit app in your browser."
      echo "Local URL: http://localhost:$PORT"
    fi
    echo "Tailing container logs... (Press Ctrl+C to stop tailing the log. The container will keep running.)"
    echo "Log output is also being saved to $LOG_FILE"
    # Tail logs to console
    docker logs -f $CONTAINER_ID
  else
    echo "Failed to start container."
    exit 1
  fi

elif [[ "$1" == "run" && "$2" == "--source" ]]; then
  echo "Running application from source..."
  # Ensure dependencies are installed
  if ! pip show streamlit > /dev/null 2>&1; then
      echo "Dependencies not found. Please install them first, e.g., using 'pip install -r requirements.txt' or 'pipenv install'"
      exit 1
  fi

  # Set PYTHONPATH to the project root directory
  export PYTHONPATH=$(pwd)

  if [[ "$3" == "ui" ]]; then
    echo "Starting Streamlit UI from source..."
    echo "You can now view your Streamlit app in your browser."
    echo "Local URL: http://localhost:$PORT"
    echo "Press Ctrl+C to stop."
    streamlit run frontend/frontend.py
  else
    echo "Starting backend from source..."
    python backend/backend.py "${@:3}"
  fi

elif [[ "$1" == "run" && "$2" == "--filehandler" ]]; then
  start_filehandler
elif [[ "$1" == "shutdown" ]]; then
  echo "Shutting down container: $CONTAINER_NAME"
  docker stop $CONTAINER_NAME >/dev/null 2>&1
  docker rm $CONTAINER_NAME >/dev/null 2>&1
  echo "Container stopped and removed."

  echo "Stopping FileHandler process..."
  pkill -f "python.*FileHandler.py" 2>/dev/null
  echo "FileHandler stopped."
else
  echo "Usage:"
  echo "  ./cat.sh run --docker ui        # Run Streamlit UI in Docker (starts FileHandler)"
  echo "  ./cat.sh run --docker [args]    # Run backend in Docker with optional args (starts FileHandler)"
  echo "  ./cat.sh run --source ui        # Run Streamlit UI from source"
  echo "  ./cat.sh run --source [args]    # Run backend from source with optional args"
  echo "  ./cat.sh run --filehandler      # Restart FileHandler.py"
  echo "  ./cat.sh shutdown               # Stop and remove the running container and FileHandler"
  echo ""
  echo "Backend arguments [args]:"
  echo "  -j, --job-file <name>             Job file name (default: DefaultJob)"
  echo "  -t, --thresholds-file <name>    Thresholds file name (default: DefaultThresholds)"
  echo "  -d, --debug                       Enable debug logging"
  echo "  -c, --concurrent-connections <n>  Number of concurrent connections"
  echo "  -u, --username <user>             Overwrite job file username"
  echo "  -p, --password <pass>             Overwrite job file password"
  echo "  -m, --auth-method <method>        Overwrite job file auth method (basic,secret,token)"
  exit 1
fi