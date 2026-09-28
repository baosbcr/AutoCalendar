"""Entry point for the double-click AutoCalendar.exe (built by build_exe.ps1).

PyInstaller needs a script rather than ``python -m autocalendar``. Double-click
it and a file picker opens; drop a sheet on it and that sheet is converted.
"""

from autocalendar.cli import main

raise SystemExit(main())
