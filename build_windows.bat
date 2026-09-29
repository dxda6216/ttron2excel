@echo off
rem Build TTRON2Excel.exe on your own Windows PC (Python 3.10-3.12 must be installed).
rem The EXE is written to the "dist" folder.
setlocal
cd /d "%~dp0"
if not exist .venv (
    py -3 -m venv .venv || python -m venv .venv || goto :error
)
call .venv\Scripts\activate.bat || goto :error
python -m pip install --upgrade pip || goto :error
pip install -r requirements.txt "pyinstaller>=6.6" || goto :error
pyinstaller --noconfirm --clean ttron2excel.spec || goto :error
echo.
echo Done: dist\TTRON2Excel.exe
pause
exit /b 0
:error
echo.
echo Build failed.
pause
exit /b 1
