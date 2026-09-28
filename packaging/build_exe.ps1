# Build dist\AutoCalendar.exe: the sheet-to-.ics converter as one file that
# runs without Python. Needs: pip install pyinstaller openpyxl
#
#   powershell -ExecutionPolicy Bypass -File packaging\build_exe.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    python -m unittest discover -s tests
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed; not building.' }

    # The Outlook automation (pywin32) is left out on purpose: the .exe only
    # writes .ics files, which cannot send anything.
    python -m PyInstaller --noconfirm --clean --onefile --console `
        --name AutoCalendar `
        --paths $root `
        --exclude-module win32com --exclude-module pythoncom --exclude-module pywintypes `
        --distpath "$root\dist" --workpath "$root\build" --specpath "$root\build" `
        "$root\packaging\autocalendar_exe.py"
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed.' }
    Get-Item "$root\dist\AutoCalendar.exe" | Select-Object Name, Length, LastWriteTime
}
finally {
    Pop-Location
}
