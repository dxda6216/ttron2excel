# Taylortron TRACES → Excel Converter (Windows EXE build)

`ttron2excel_app.py` reads a Taylortron `TRACES.nnn` file and writes Excel workbooks,
a ZIP of per-channel `.dat` files (LumiCycle) and a multi-page PDF of plots, with a
review window for peaks/troughs, detrending and sine fitting.

This repository builds a stand-alone Windows program, **`TTRON2Excel.exe`**, automatically
with GitHub Actions. Users of the EXE do not need Python.

## Files

| File | Purpose |
|---|---|
| `ttron2excel_app.py` | The application |
| `requirements.txt` | Python packages it needs |
| `ttron2excel.spec` | PyInstaller recipe (one EXE file, no console window) |
| `.github/workflows/build-windows-exe.yml` | GitHub Actions workflow that builds the EXE on Windows |
| `build_windows.bat` | Optional: build the EXE on your own Windows PC |
| `.gitignore` | Keeps build output and data files out of the repository |

## Build on GitHub

1. Create a new repository on GitHub (it can be private).
2. Upload all the files **keeping the folder structure**; the workflow must be at
   `.github/workflows/build-windows-exe.yml`.
   - Web browser: *Add file → Upload files*, then drag the whole unzipped folder
     contents in (including the `.github` folder; on macOS press `Cmd+Shift+.` in
     Finder to see hidden folders).
     Or create the workflow with *Add file → Create new file* and type the path
     `.github/workflows/build-windows-exe.yml`, then paste its contents.
   - Git: `git init`, `git add .`, `git commit -m "Initial"`, `git remote add origin <url>`,
     `git push -u origin main`.
3. The build starts automatically on every push to `main`. It can also be started by hand:
   **Actions** tab → *Build Windows EXE* → **Run workflow**.
4. When the run has a green tick (about 5–10 minutes), open it and download
   **TTRON2Excel-windows** under *Artifacts*. It is a ZIP containing `TTRON2Excel.exe`.
   Artifacts are kept for 30 days.

### Publishing a release (optional)

Push a version tag and the EXE is also attached to a GitHub Release, which does not expire:

```
git tag v1.0.0
git push origin v1.0.0
```

(Or on GitHub: *Releases → Draft a new release → Choose a tag → type `v1.0.0` → Create*.)

### Icon (optional)

Put a Windows icon file named `icon.ico` in the repository root; the next build uses it.

## Build on your own Windows PC (optional)

Install Python 3.10–3.12 from python.org, then double-click `build_windows.bat`.
The EXE appears in the `dist` folder.

## Notes for users of the EXE

- **Windows SmartScreen** may say *"Windows protected your PC"* the first time, because
  the EXE is not code-signed. Click **More info → Run anyway**. Some antivirus programs
  are also cautious with unsigned PyInstaller programs; if needed, ask IT to allow the file.
- **Start-up takes several seconds** (the one-file EXE unpacks itself to a temporary
  folder each time). This is normal.
- If something goes wrong inside the program, a message box shows the error details;
  please copy that text when reporting a problem.

## Run from the Python source instead

```
pip install -r requirements.txt
python ttron2excel_app.py
```
