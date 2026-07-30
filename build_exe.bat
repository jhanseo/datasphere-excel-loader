@echo off
setlocal

REM ===================================================================
REM  Excel -> HANA Datasphere Upload Wizard
REM  Windows EXE build script
REM
REM  Usage : double-click this file, or run  build_exe.bat  in cmd
REM  Output: dist\ExcelUploadWizard.exe  +  dist\config.ini
REM
REM  NOTE: this file must keep CRLF line endings and ASCII-only text.
REM        Do not add "chcp" here - it breaks batch parsing.
REM ===================================================================

cd /d "%~dp0"

echo.
echo ============================================================
echo  ExcelUploadWizard - EXE build
echo  Folder : %CD%
echo ============================================================
echo.

REM ---------- [1/6] check source files ------------------------------
if not exist "wizard.py"       goto :no_source
if not exist "excel_loader.py" goto :no_source
if not exist "wizard.spec"     goto :no_source

REM config.ini is gitignored, so a fresh clone has only config.ini.example.
if not exist "config.ini" if exist "config.ini.example" (
    copy /y "config.ini.example" "config.ini" >nul
    echo [1/6] Created config.ini from config.ini.example
)
if not exist "config.ini" goto :no_config
echo [1/6] Source files OK

REM ---------- [2/6] locate Python -----------------------------------
set "PY="
py -3 --version >nul 2>&1
if not errorlevel 1 set "PY=py -3"
if defined PY goto :got_python
python --version >nul 2>&1
if not errorlevel 1 set "PY=python"
if defined PY goto :got_python
goto :no_python

:got_python
echo [2/6] Python found:
%PY% -c "import sys,struct;print('      ',sys.version.split()[0],struct.calcsize('P')*8,'bit');print('      ',sys.executable)"
if errorlevel 1 goto :no_python

REM ---------- [3/6] install dependencies ----------------------------
echo.
echo [3/6] Installing dependencies (this may take a few minutes)...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto :pip_failed

REM ---------- [4/6] clean previous build ----------------------------
echo.
echo [4/6] Cleaning previous build

REM A running instance keeps the EXE locked and breaks the build.
taskkill /f /im ExcelUploadWizard.exe >nul 2>&1

if exist "build" rmdir /s /q "build"
if exist "dist"  rmdir /s /q "dist"
if exist "dist" goto :dist_locked

REM ---------- [5/6] run PyInstaller ---------------------------------
echo.
echo [5/6] Running PyInstaller
%PY% -m PyInstaller --clean --noconfirm wizard.spec
if errorlevel 1 goto :build_failed

REM The spec builds one-folder by default, one-file if ONEFILE=True.
set "OUTDIR="
if exist "dist\ExcelUploadWizard\ExcelUploadWizard.exe" set "OUTDIR=dist\ExcelUploadWizard"
if not defined OUTDIR if exist "dist\ExcelUploadWizard.exe" set "OUTDIR=dist"
if not defined OUTDIR goto :build_failed

REM ---------- [6/6] copy config.ini ---------------------------------
echo.
echo [6/6] Copying config.ini
if not exist "%OUTDIR%\config.ini" copy /y "config.ini" "%OUTDIR%\config.ini" >nul

echo.
echo ============================================================
echo  BUILD SUCCEEDED
echo.
echo    Output : %CD%\%OUTDIR%
echo    EXE    : %OUTDIR%\ExcelUploadWizard.exe
echo    Config : %OUTDIR%\config.ini
echo.
echo  Keep ExcelUploadWizard.exe and config.ini together.
echo  For a one-folder build, zip the whole folder and ship that.
echo  Microsoft Excel must be installed on the target PC
echo  in order to preview DRM-protected files.
echo ============================================================
echo.
pause
exit /b 0

REM ------------------------------------------------------------------
:no_config
echo [ERROR] config.ini is missing and config.ini.example was not found.
echo         Re-clone the repository, or create config.ini by hand.
goto :fail

:no_source
echo [ERROR] Source files are missing.
echo         wizard.py, excel_loader.py and wizard.spec must be in
echo         the same folder as this script:
echo         %CD%
goto :fail

:no_python
echo [ERROR] Python was not found.
echo.
echo   1) Install Python 3.9 or newer from https://www.python.org/downloads/
echo   2) Tick "Add python.exe to PATH" during installation.
echo   3) Open a NEW command prompt and run this script again.
echo.
echo   If the Microsoft Store opens when you type "python", turn off
echo   the app alias: Settings ^> Apps ^> Advanced app settings ^>
echo   App execution aliases ^> disable "python.exe" / "python3.exe".
goto :fail

:pip_failed
echo [ERROR] Failed to install dependencies.
echo.
echo   Behind a corporate proxy, run this first:
echo     %PY% -m pip install -r requirements.txt --proxy http://HOST:PORT
echo.
echo   For an SSL inspection error, add:
echo     --trusted-host pypi.org --trusted-host files.pythonhosted.org
goto :fail

:dist_locked
echo [ERROR] The "dist" folder could not be deleted.
echo         Something is holding a file inside it open.
echo         Close any running ExcelUploadWizard and any Explorer window
echo         pointing at that folder, then run this script again.
goto :fail

:build_failed
echo [ERROR] PyInstaller build failed.
echo.
echo   If the output above says
echo     "Failed to compute PE checksum"   or   "Error code: 32"
echo     "EndUpdateResourceW"              or   sharing violation
echo   then antivirus or DRM software locked the new EXE while
echo   PyInstaller was still writing it. Fix it like this:
echo.
echo     1) Exclude this folder from real-time scanning.
echo        Windows Defender, in an ADMIN PowerShell:
echo          Add-MpPreference -ExclusionPath '%CD%'
echo        For a corporate AV (V3, Symantec, McAfee...) ask your
echo        IT admin to exclude the folder, or pause protection.
echo.
echo     2) Build on a local disk, not OneDrive or a network drive.
echo.
echo     3) Make sure ExcelUploadWizard.exe is not already running.
echo.
echo   Otherwise scroll up and read the last lines of the output.
echo   The full warning list is written to:
echo     %CD%\build\ExcelUploadWizard\warn-ExcelUploadWizard.txt
goto :fail

:fail
echo.
echo ============================================================
echo  BUILD FAILED
echo ============================================================
echo.
pause
exit /b 1
