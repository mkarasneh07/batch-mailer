# Batch Mailer installer and updater for Windows.
#
# Paste this into PowerShell and press Enter:
#   irm https://raw.githubusercontent.com/mkarasneh07/batch-mailer/main/install.ps1 | iex
#
# Installs into your own user folder: no admin rights and no Python needed.
# Run the same line again to update. Your connected email, sent history and
# do-not-email list are kept.

& {  # everything runs in its own scope, so nothing is left behind in your PowerShell window

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
# Older Windows PowerShell versions need this to talk to GitHub.
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$Repo     = 'mkarasneh07/batch-mailer'
$Source   = if ($env:BATCH_MAILER_SOURCE) { $env:BATCH_MAILER_SOURCE } else { "https://github.com/$Repo/archive/refs/heads/main.zip" }
$Dir      = if ($env:BATCH_MAILER_DIR) { $env:BATCH_MAILER_DIR } else { Join-Path $env:LOCALAPPDATA 'BatchMailer' }
$HomeDir  = if ($env:USERPROFILE) { $env:USERPROFILE } else { $HOME }
$Tools    = Join-Path $Dir 'tools'
$Uv       = Join-Path $Tools 'uv.exe'
$UvSource = 'https://github.com/astral-sh/uv/releases/latest/download/uv-x86_64-pc-windows-msvc.zip'
$RunArgs  = 'run --quiet --python 3.12 --with-requirements requirements.txt streamlit run app.py'
$Tmp      = Join-Path ([IO.Path]::GetTempPath()) ('batch-mailer-' + [guid]::NewGuid().ToString('N'))

function Get-File([string]$From, [string]$To) {
    if (Test-Path -LiteralPath $From) { Copy-Item -LiteralPath $From -Destination $To }
    else { Invoke-WebRequest -UseBasicParsing -Uri $From -OutFile $To }
}

Write-Host ''
Write-Host 'Installing Batch Mailer...' -ForegroundColor Cyan

try {
    New-Item -ItemType Directory -Force -Path $Dir, $Tools, $Tmp | Out-Null

    # 1. The app itself. Files are copied one by one so an update never touches your own data.
    Write-Host '  1/4  Getting the app'
    $zip = Join-Path $Tmp 'app.zip'
    Get-File $Source $zip
    $unpacked = Join-Path $Tmp 'app'
    Expand-Archive -LiteralPath $zip -DestinationPath $unpacked -Force
    $src = Get-ChildItem -LiteralPath $unpacked -Directory | Select-Object -First 1
    Get-ChildItem -LiteralPath $src.FullName -Recurse -Force -File | ForEach-Object {
        $relative = $_.FullName.Substring($src.FullName.Length).TrimStart('\', '/')
        $target = Join-Path $Dir $relative
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
        Copy-Item -LiteralPath $_.FullName -Destination $target -Force
    }

    # 2. The engine (uv) that brings its own private copy of Python.
    if (Test-Path -LiteralPath $Uv) {
        Write-Host '  2/4  Engine already installed'
    } else {
        Write-Host '  2/4  Getting the engine that runs it'
        $uvZip = Join-Path $Tmp 'uv.zip'
        Get-File $UvSource $uvZip
        Expand-Archive -LiteralPath $uvZip -DestinationPath $Tools -Force
    }

    # 3. Download Python and the components now, so the first start is quick.
    Write-Host '  3/4  Downloading Python and components (1-2 minutes the first time)'
    Push-Location -LiteralPath $Dir
    try { & $Uv run --quiet --python 3.12 --with-requirements requirements.txt python -c "print('ok')" | Out-Null }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw 'Downloading Python or its components failed.' }

    # Skip the web framework's one-time "enter your email" question.
    $streamlitDir = Join-Path $HomeDir '.streamlit'
    New-Item -ItemType Directory -Force -Path $streamlitDir | Out-Null
    $credentials = Join-Path $streamlitDir 'credentials.toml'
    if (-not (Test-Path -LiteralPath $credentials)) {
        Set-Content -LiteralPath $credentials -Value @('[general]', 'email = ""') -Encoding Ascii
    }

    # 4. Shortcuts on the desktop and in the Start menu. They open the app with its window minimized.
    Write-Host '  4/4  Adding Batch Mailer to the desktop and Start menu'
    try {
        $shell = New-Object -ComObject WScript.Shell
        foreach ($folder in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) {
            $link = $shell.CreateShortcut((Join-Path $folder 'Batch Mailer.lnk'))
            $link.TargetPath = $Uv
            $link.Arguments = $RunArgs
            $link.WorkingDirectory = $Dir
            $link.WindowStyle = 7
            $link.IconLocation = (Join-Path $Dir 'icon.ico')
            $link.Description = 'Send personal emails to a list'
            $link.Save()
        }
    } catch {
        Write-Host "       Couldn't add the shortcuts: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host 'Done. Batch Mailer is opening in your browser.' -ForegroundColor Green
    Write-Host 'Next time, open it from the Batch Mailer shortcut on the desktop.'
    if (-not $env:BATCH_MAILER_NO_START) {
        Start-Process -FilePath $Uv -ArgumentList $RunArgs -WorkingDirectory $Dir -WindowStyle Minimized
    }
}
catch {
    Write-Host ''
    Write-Host "Batch Mailer couldn't be installed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Check the internet connection. On a work laptop, IT may be blocking downloads from github.com.'
}
finally {
    Remove-Item -LiteralPath $Tmp -Recurse -Force -ErrorAction SilentlyContinue
}
}
