# PM Plan Auto Schedule

Desktop app for generating Jan-Dec PM plan files from:

`C:\Users\sthit\Desktop\ALL BACKLINE PM PLAN - JAN - 2026.xls`

## Project layout

- `pm_plan_auto_schedule/` - app package
- `assets/` - icon assets
- `tools/create_icon.py` - icon generator
- `tools/build_exe.ps1` - one-file Windows build script
- `main.py` - app entrypoint

## Run in Python

```powershell
.\venv\Scripts\python.exe main.py
```

## Build the EXE

```powershell
.\tools\build_exe.ps1
```

Output:

```text
dist\PMPlanAutoSchedule.exe
```

## Notes

- Microsoft Excel must be installed.
- The app keeps the Excel formatting and updates the schedule for the selected year.
- `DE-DROSS` anchored from `PM PLAN` applies to `BT01-BT09`, `A12`, and `A13`.
