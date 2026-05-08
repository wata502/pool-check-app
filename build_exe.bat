@echo off
chcp 932 > nul
echo ============================================
echo  Build PoolWriter.exe
echo ============================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install from https://www.python.org/
    pause
    exit /b 1
)

echo [1/3] Installing libraries...
python -m pip install requests pywin32 pyinstaller pillow pystray sseclient-py
if errorlevel 1 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)
echo OK.

echo [2/3] Building exe...
python -m PyInstaller --onefile --windowed --name "PoolWriter" ^
  --hidden-import=pystray._win32 ^
  --hidden-import=PIL._tkinter_finder ^
  --hidden-import=win32com.client ^
  --hidden-import=win32com ^
  --hidden-import=pywintypes ^
  --hidden-import=pythoncom ^
  --hidden-import=requests ^
  --hidden-import=sseclient ^
  --hidden-import=win32timezone ^
  pool_writer.py
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)
echo OK.

echo [3/3] Copying exe...
:: 実行中の PoolWriter.exe があると上書きに失敗する。事前停止確認。
tasklist /FI "IMAGENAME eq PoolWriter.exe" 2>nul | find /I "PoolWriter.exe" >nul
if not errorlevel 1 (
    echo [ERROR] PoolWriter.exe is still running. Stop it from the tray icon and Task Manager, then retry.
    pause
    exit /b 1
)

copy /Y "dist\PoolWriter.exe" "PoolWriter.exe" > nul
if errorlevel 1 (
    echo [ERROR] Copy failed. PoolWriter.exe may be locked by a running instance.
    pause
    exit /b 1
)

:: コピー後のタイムスタンプを表示して目視確認できるようにする
for %%F in ("PoolWriter.exe") do echo Updated: %%~tF  Size: %%~zF
echo OK.

echo.
echo ============================================
echo  Build SUCCESS! PoolWriter.exe is ready.
echo ============================================
echo.
pause
