@echo off
rem Check if the virtual environment directory exists
if not exist venv\Scripts\activate.bat (
    echo Virtual environment not found. Please run setup.bat to create the virtual environment.
    pause
    exit /b
)
rem Activate the virtual environment
call venv\Scripts\activate.bat

rem Check if Python is installed
where /q python
if %ERRORLEVEL% neq 0 (
    echo Python is not installed. Please install Python and try again.
    pause
    exit /b
)

rem Check if app.py exists
if not exist app.py (
    echo app.py does not exist. Please verify the file path and try again.
    pause
    exit /b
)

rem Launch the application
python app.py
pause
