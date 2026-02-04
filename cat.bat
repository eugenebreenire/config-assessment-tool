@echo off
setlocal

REM ==========================================
REM Configuration
REM ==========================================
set REPO=ghcr.io/appdynamics/config-assessment-tool-windows

if exist VERSION (
    set /p VERSION=<VERSION
) else (
    echo Error: VERSION file not found.
    exit /b 1
)

set IMAGE=%REPO%:%VERSION%
set PORT=8501
set LOG_DIR=logs
set LOG_FILE=%LOG_DIR%\config-assessment-tool.log
set CONTAINER_NAME=cat-tool-container
set FILE_HANDLER_HOST=host.docker.internal

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM Define Mounts for Windows (Using %cd% for current directory)
set MOUNTS=-v "%cd%\input\jobs:C:\app\input\jobs" -v "%cd%\input\thresholds:C:\app\input\thresholds" -v "%cd%\output\archive:C:\app\output\archive" -v "%cd%\logs:C:\app\logs"

REM ==========================================
REM Argument Parsing
REM ==========================================
if "%1"=="--start" goto start
if "%1"=="shutdown" goto shutdown
goto usage

:start
shift
if "%1"=="docker" goto start_docker
goto start_local

REM ==========================================
REM Docker Startup Mode
REM ==========================================
:start_docker
REM 1. Start FileHandler on Host (Using python directly, NO pipenv)
if not exist "frontend\FileHandler.py" (
    echo Error: frontend\FileHandler.py not found.
    exit /b 1
)

echo Starting FileHandler service on host...
REM Kill any previous instance (rudimentary check by window title or just blindly)
taskkill /F /FI "WINDOWTITLE eq CAT-FileHandler" >nul 2>&1

REM Start in background with a specific title so we can kill it later
start "CAT-FileHandler" /min cmd /c "python frontend\FileHandler.py >> %LOG_FILE% 2>&1"
echo FileHandler started.
timeout /t 2 /nobreak >nul

REM 2. Reset Container
docker stop %CONTAINER_NAME% >nul 2>&1
docker rm %CONTAINER_NAME% >nul 2>&1

shift
REM Check if there are additional arguments (Backend vs UI)
if "%1"=="" (
    echo Starting container in UI mode...
    REM Note: We do NOT pass "streamlit run..." args here.
    REM The Dockerfile ENTRYPOINT defaults to frontend if no "backend" arg is present.
    REM We pass --server.headless=true to ensure non-interactive start if needed,
    REM though the Entrypoint handles 'python -m streamlit run ...' automatically.

    docker run -d --name %CONTAINER_NAME% -e FILE_HANDLER_HOST=%FILE_HANDLER_HOST% -p %PORT%:%PORT% %MOUNTS% %IMAGE% --server.headless=true

    if %ERRORLEVEL% EQU 0 (
        echo Container started successfully.
        echo UI available at http://localhost:%PORT%
        docker logs -f %CONTAINER_NAME%
    ) else (
        echo Failed to start container.
        exit /b 1
    )
) else (
    echo Starting container in backend mode with args: %1 %2 %3 %4 %5 %6 %7 %8 %9
    REM Pass "backend" + user arguments to trigger the backend branch in entrypoint.bat
    docker run --rm --name %CONTAINER_NAME% -e FILE_HANDLER_HOST=%FILE_HANDLER_HOST% -p %PORT%:%PORT% %MOUNTS% %IMAGE% backend %1 %2 %3 %4 %5 %6 %7 %8 %9

    if %ERRORLEVEL% EQU 0 (
        echo Container finished.
    ) else (
        echo Failed to start container.
        exit /b 1
    )
)
goto end


REM ==========================================
REM Local Source Startup Mode
REM ==========================================
:start_local
REM Setting PYTHONPATH for local execution
set PYTHONPATH=%cd%;%cd%\backend

if "%1"=="" (
    echo Running application in UI mode from source...
    echo UI available at http://localhost:%PORT%
    pipenv run streamlit run frontend\frontend.py
) else (
    echo Running application in backend mode from source with args: %1 %2 %3 %4 %5 %6 %7 %8 %9
    pipenv run python backend\backend.py %1 %2 %3 %4 %5 %6 %7 %8 %9
)
goto end


REM ==========================================
REM Shutdown Mode
REM ==========================================
:shutdown
echo Shutting down container: %CONTAINER_NAME%
docker stop %CONTAINER_NAME% >nul 2>&1
docker rm %CONTAINER_NAME% >nul 2>&1
echo Container stopped and removed.

echo Stopping FileHandler process...
taskkill /F /FI "WINDOWTITLE eq CAT-FileHandler" >nul 2>&1
echo FileHandler stopped.

echo Stopping generic python processes (backend/streamlit)...
REM Warning: This might be too aggressive on a developer machine,
REM but matches the aggressive pkill in cat.sh
taskkill /F /IM python.exe /T >nul 2>&1
taskkill /F /IM streamlit.exe /T >nul 2>&1
echo Processes stopped.
goto end


REM ==========================================
REM Usage / Help
REM ==========================================
:usage
echo Usage:
echo   cat.bat --start                # Starts CAT UI locally (Uses pipenv)
echo   cat.bat --start [args]         # Starts CAT backend locally (Uses pipenv)
echo   cat.bat --start docker         # Starts CAT UI in Docker (Uses host python for FileHandler)
echo   cat.bat --start docker [args]  # Starts CAT backend in Docker
echo   cat.bat shutdown               # Stop container and processes
echo.
echo Arguments:
echo   --job-file name, --thresholds-file name, --debug, etc.
goto end

:end
endlocal