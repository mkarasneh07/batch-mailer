@echo off
rem Batch Mailer for Windows. Double-click to open. No Python or admin rights needed.
rem First run downloads a small engine (uv) that fetches Python and everything else (1-2 minutes, once).
cd /d "%~dp0"
title Batch Mailer - keep this window open while sending

if not exist "tools\uv.exe" (
  echo First-time setup. This takes a minute or two, only once...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; New-Item -ItemType Directory -Force 'tools' | Out-Null; Invoke-WebRequest -UseBasicParsing 'https://github.com/astral-sh/uv/releases/latest/download/uv-x86_64-pc-windows-msvc.zip' -OutFile 'tools\uv.zip'; Expand-Archive -Force 'tools\uv.zip' 'tools'; Remove-Item 'tools\uv.zip'"
  if not exist "tools\uv.exe" (
    echo.
    echo Setup couldn't download what it needs. Check the internet connection and try again.
    echo On a work laptop, IT may be blocking downloads from github.com.
    pause
    exit /b 1
  )
)

rem Skip Streamlit's one-time "enter your email" question.
set "ST=%USERPROFILE%\.streamlit"
if not exist "%ST%" mkdir "%ST%"
if not exist "%ST%\credentials.toml" (
  >"%ST%\credentials.toml" echo [general]
  >>"%ST%\credentials.toml" echo email = ""
)

rem Put a "Batch Mailer" shortcut on the desktop, once. It opens this window minimized.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$p = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Batch Mailer.lnk'; if (-not (Test-Path $p)) { $s = (New-Object -ComObject WScript.Shell).CreateShortcut($p); $s.TargetPath = '%~f0'; $s.WorkingDirectory = '%~dp0'; $s.WindowStyle = 7; $s.Save() }"

echo Opening Batch Mailer in your browser...
"tools\uv.exe" run --quiet --python 3.12 --with-requirements requirements.txt streamlit run app.py
if errorlevel 1 (
  echo.
  echo Batch Mailer stopped because of a problem. Take a photo of this window and send it to whoever set it up.
  pause
)
