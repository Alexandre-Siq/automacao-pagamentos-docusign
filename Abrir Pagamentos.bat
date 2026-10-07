@echo off
cd /d "%~dp0"
where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw "%~dp0abrir_interface.py"
    exit /b 0
)
where pyw >nul 2>&1
if %errorlevel%==0 (
    start "" pyw -3 "%~dp0abrir_interface.py"
    exit /b 0
)
echo Nao encontrei o Python. Instale o Python e marque a opcao "Add python.exe to PATH".
pause
