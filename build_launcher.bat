@echo off
setlocal enabledelayedexpansion
title Build Python to EXE

echo.
echo ========================================
echo   Build generica Python to EXE
echo ========================================
echo.

REM ── File .py ──────────────────────────────────────────────────────────────
set "MAIN=%~1"
if not "!MAIN!"=="" set "MAIN=!MAIN:"=!"

if "!MAIN!"=="" (
    set /p "MAIN=Inserisci il nome del file .py: "
    set "MAIN=!MAIN:"=!"
)

if not exist "!MAIN!" (
    echo ERRORE: File "!MAIN!" non trovato!
    goto :error
)

for %%f in ("!MAIN!") do set "NAME=%%~nf"

REM ── Icona ─────────────────────────────────────────────────────────────────
set "ICON=%~2"
if not "!ICON!"=="" set "ICON=!ICON:"=!"

if "!ICON!"=="" (
    echo.
    echo Vuoi aggiungere un'icona?
    echo   [1] Si
    echo   [2] No
    echo.
    set /p "SCELTA=Scelta (1/2): "
    if "!SCELTA!"=="1" (
        set /p "ICON=Percorso file .ico: "
        set "ICON=!ICON:"=!"
        if not exist "!ICON!" (
            echo ATTENZIONE: icona non trovata, ignorata.
            set "ICON="
        )
    )
)

REM ── Riepilogo ─────────────────────────────────────────────────────────────
echo.
echo File   : !MAIN!
echo Nome   : !NAME!.exe
if not "!ICON!"=="" (echo Icona  : !ICON!) else (echo Icona  : nessuna)
echo.

REM ── Installa PyInstaller se mancante ──────────────────────────────────────
echo [1/3] Verifica PyInstaller...
python -c "import PyInstaller" 2>nul || pip install pyinstaller
echo.

REM ── Argomenti icona ───────────────────────────────────────────────────────
set "ICON_ARGS="
if not "!ICON!"=="" (
    if exist "!ICON!" (
        set "ICON_ARGS=--icon "!ICON!" --add-data "!ICON!;.""
    )
)

REM ── Build ─────────────────────────────────────────────────────────────────
echo [2/3] Build in corso...
echo.

python -m PyInstaller --onefile --windowed --name "!NAME!" !ICON_ARGS! --hidden-import tkinter --hidden-import tkinter.ttk --hidden-import tkinter.messagebox --hidden-import tkinter.filedialog --hidden-import tkcalendar --hidden-import win32com --hidden-import win32com.client --hidden-import win32process --hidden-import win32gui --hidden-import win32timezone --hidden-import pythoncom --hidden-import pywintypes --hidden-import pandas --hidden-import openpyxl --hidden-import PIL._tkinter_finder --hidden-import sqlite3 --hidden-import babel.numbers --collect-all tkcalendar --collect-all babel --clean --noconfirm "!MAIN!"

if errorlevel 1 goto :error

REM ── Fine ──────────────────────────────────────────────────────────────────
echo.
echo [3/3] Completato!
echo.
echo EXE: .\dist\!NAME!.exe
echo.
pause
exit /b 0

:error
echo.
echo ========================================
echo   ERRORE - Build fallita
echo ========================================
echo.
pause
exit /b 1