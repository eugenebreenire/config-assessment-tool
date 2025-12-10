# Variables
PYTHON := python3
DOCKER := docker
VERSION := $(shell cat VERSION)
PLATFORM := $(shell uname | tr '[:upper:]' '[:lower:]')
LOG_DIR := logs
OUTPUT_DIR := output
DIST_FILE := config-assessment-tool-dist.zip
FRONTEND_IMAGE := ghcr.io/appdynamics/config-assessment-tool-frontend-$(PLATFORM):$(VERSION)
BACKEND_IMAGE := ghcr.io/appdynamics/config-assessment-tool-backend-$(PLATFORM):$(VERSION)
LOG_FILES := logs/*.log config-assessment-tool-backend-joshua.log
BACKEND_SCRIPT := backend/backend.py
INPUT_FILE := input/jobs/DefaultJob.json

# Determine architecture for Docker image tagging (OS-Chip)
# This ARCH is used ONLY for DOCKER_IMAGE_TAG, not for the bundle target's internal logic.
ARCH := $(shell UNAME_S=$$(uname -s | tr -d '[:space:]'); UNAME_M=$$(uname -m | tr -d '[:space:]'); OS_PART="unknown_os"; if [ "$$UNAME_S" = "Darwin" ]; then OS_PART="macos"; fi; if [ "$$UNAME_S" = "Linux" ]; then OS_PART="linux"; fi; if echo "$$UNAME_S" | grep -q "CYGWIN" || echo "$$UNAME_S" | grep -q "MINGW"; then OS_PART="windows"; fi; ARCH_PART="unknown_arch"; if [ "$$UNAME_M" = "x86_64" ]; then ARCH_PART="x86"; fi; if [ "$$UNAME_M" = "arm64" ] || [ "$$UNAME_M" = "aarch64" ]; then ARCH_PART="arm"; fi; echo "$$OS_PART-$$ARCH_PART")
# Docker image tag
DOCKER_IMAGE_TAG_BACKEND := ghcr.io/appdynamics/config-assessment-tool-backend-$(ARCH):$(VERSION)
DOCKER_IMAGE_TAG_FRONTEND := ghcr.io/appdynamics/config-assessment-tool-frontend-$(ARCH):$(VERSION)

# Default target
.DEFAULT_GOAL := help

# Help target
help:
	@echo "Available targets:"
	@echo "  run                         - Run the config-assessment-tool with UI (requires Docker)"
	@echo "  run-backend                 - Run the non-UI version"
	@echo "  build-images                - Build Docker images required to run config-assessment-tool"
	@echo "  install                     - Install Python dependencies"


# Run the full tool (backend and frontend)
#run: $(LOG_DIR) $(OUTPUT_DIR)
	#$(PYTHON) bin/config-assessment-tool.py --run

run: $(LOG_DIR) $(OUTPUT_DIR)
	PYTHONPATH=. $(PYTHON) bin/config-assessment-tool.py --run

.PHONY: run-backend test lint clean build

.PHONY: run-backend

run-backend: install
	@if [ -z "$(ARGS)" ]; then \
		echo ""; \
		echo "Usage:  make run-backend ARGS=\"<Options>\""; \
		echo "Options:"; \
		echo "  -j, --job-file TEXT"; \
		echo "  -t, --thresholds-file TEXT"; \
		echo "  -d, --debug"; \
		echo "  -c, --concurrent-connections INTEGER"; \
		echo "  -u, --username TEXT             overwrite job file with this username"; \
		echo "  -p, --password TEXT             overwrite job file with this password"; \
		echo "  -m, --auth-method TEXT          overwrite job file with this auth-"; \
		echo "                                  method(basic,secret,token)"; \
		echo ""; \
		echo "  --car                           Generate the configuration analysis report as part of the output"; \
		echo "  --help                          Show this message and exit."; \
		echo "";\
	else \
		if [ -f $(INPUT_FILE) ]; then \
			pipenv run $(PYTHON) $(BACKEND_SCRIPT) $(ARGS); \
		else \
			echo "Input file '$(INPUT_FILE)' not found. Please ensure it exists."; \
			exit 1; \
		fi \
	fi


SHELL := /bin/bash

# Detect platform string
PLATFORM := $(shell \
	unameOut=$$(uname -s); \
	if [[ "$$unameOut" == "Darwin" ]]; then \
		if [[ "$$(uname -m)" == "arm64" ]]; then \
			echo "mac-m1"; \
		else \
			echo "mac"; \
		fi; \
	elif [[ "$$unameOut" == "Linux" ]]; then \
		if grep -qi microsoft /proc/version; then \
			echo "linux"; \
		else \
			echo "linux"; \
		fi; \
	elif [[ "$$unameOut" =~ MINGW* || "$$unameOut" =~ CYGWIN* ]]; then \
		echo "windows"; \
	else \
		echo "unknown"; \
	fi)


check-version:
	@echo "Checking software version..."
	@unameOut=$$(uname -s 2>/dev/null || echo "unknown"); \
	if [[ "$$unameOut" == "MINGW"* || "$$unameOut" == "CYGWIN"* || "$$unameOut" == "MSYS"* ]]; then \
		echo "⚠️  Version check is not supported on native Windows shell (without WSL or Git Bash)."; \
		echo "   Cannot determine latest version from GitHub."; \
		exit 0; \
	fi; \
	latest_tag=$$(curl -s https://api.github.com/repos/appdynamics/config-assessment-tool/tags | grep 'name' | head -n 1 | cut -d ':' -f2 | cut -d '"' -f2); \
	local_tag=$$(cat VERSION 2>/dev/null || echo "unknown"); \
	echo "Local version : $$local_tag"; \
	echo "Latest version: $$latest_tag"; \
	if [ "$$latest_tag" != "$$local_tag" ]; then \
		echo "⚠️  You are using (or building docker images from source) using an outdated version."; \
		echo "   Local version: $$local_tag"; \
		echo "   Latest version: $$latest_tag"; \
		echo "👉  Get the latest at: https://github.com/Appdynamics/config-assessment-tool/releases"; \
	else \
		echo "✅ You are using the latest version."; \
	fi


# Get tag from VERSION file
TAG := $(shell cat VERSION 2>/dev/null || echo "unknown")

# Set docker build options
DOCKER_BUILD_OPTS := $(if $(NO_CACHE),--no-cache,)

.PHONY: build-images

build:
	@echo "Usage: make build-images COMPONENT=frontend|backend|all [NO_CACHE=true]"
	@echo "Example: make build-images COMPONENT=frontend NO_CACHE=true"


build-images:
	@echo "Building for platform: $(PLATFORM)"
	@echo "Using tag: $(TAG)"
	@echo "No cache: $(NO_CACHE)"
	@$(MAKE) check-version
ifeq ($(COMPONENT),frontend)
	@echo "Building frontend image..."
	docker build $(DOCKER_BUILD_OPTS) -t $(DOCKER_IMAGE_TAG_FRONTEND) -f frontend/Dockerfile .
else ifeq ($(COMPONENT),backend)
	@echo "Building backend image..."
	docker build $(DOCKER_BUILD_OPTS) -t $(DOCKER_IMAGE_TAG_BACKEND) -f backend/Dockerfile .
else
	@echo "Building both frontend and backend images..."
	docker build $(DOCKER_BUILD_OPTS) -t $(DOCKER_IMAGE_TAG_FRONTEND) -f frontend/Dockerfile .
	docker build $(DOCKER_BUILD_OPTS) -t $(DOCKER_IMAGE_TAG_BACKEND) -f backend/Dockerfile .
endif

# Install Python dependencies
install:
	@if [ -f Pipfile ]; then \
		$(PYTHON) -m pip install pipenv && pipenv install; \
	elif [ -f requirements.txt ]; then \
		$(PYTHON) -m pip install -r requirements.txt; \
	else \
		echo "No requirements.txt or Pipfile found!"; \
	fi

# Run lint checks on the project
lint:
	@echo "Running lint checks..."
	@flake8 backend/ frontend/ bin/ tests/ || true

# Run unit tests
test:
	@echo "Running tests..."
	$(PYTHON) -m unittest discover -s tests -p "*.py"