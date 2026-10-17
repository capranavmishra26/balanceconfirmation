@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Not set up yet - running setup.bat first...
    call setup.bat || exit /b 1
)
rem Skip Streamlit's first-run email prompt.
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    mkdir "%USERPROFILE%\.streamlit" 2>nul
    > "%USERPROFILE%\.streamlit\credentials.toml" (
        echo [general]
        echo email = ""
    )
)
".venv\Scripts\python.exe" -m streamlit run app.py
