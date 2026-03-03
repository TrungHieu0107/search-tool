# Excel Search Tool

Production-ready Windows desktop app for high-performance searching in Excel files, with automated GitHub Actions CI/CD and downloadable installer artifacts.

## Features

- Modern WinForms desktop UI with smooth DataGridView rendering.
- Search modes:
  - Content search (cell values)
  - File name search
  - Regex search + case sensitivity
- Instant search with 400ms debounce.
- Cancels previous search when a new query starts.
- Search history persisted with auto-suggest via combo box.
- Multi-folder search support.
- Parallelized file scanning with ZIP pre-scan to skip unlikely files.
- EPPlus-powered workbook parsing for accurate content matches.
- Progress bar and status counters.
- Corrupted files are skipped and errors are logged.

## Architecture

- `SearchService`: orchestration, parallel processing, progress reporting.
- `FileNameSearchService`: file-name match logic.
- `ContentSearchService`: ZIP pre-scan + EPPlus content scanning.
- `HistoryService`: query history persistence.
- `FolderManager`: add/remove/clear folder sources.
- `ErrorLogger`: async file logging under `%LOCALAPPDATA%\ExcelSearchTool\errors.log`.

## Local Build & Run

```bash
dotnet restore ./ExcelSearchTool/ExcelSearchTool.csproj
dotnet build ./ExcelSearchTool/ExcelSearchTool.csproj -c Release
dotnet run --project ./ExcelSearchTool/ExcelSearchTool.csproj -c Release
```

## Local Publish (single EXE)

```bash
dotnet publish ./ExcelSearchTool/ExcelSearchTool.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true -o publish
```

## Local Installer Build (Inno Setup)

1. Install Inno Setup.
2. Ensure publish output exists in `publish/`.
3. Compile installer:

```bash
iscc setup.iss
```

Output installer:
- `installer-out/ExcelSearchTool-Setup.exe`

## CI/CD (GitHub Actions)

Workflow file: `.github/workflows/build.yml`.

On push to `main`:
1. Restore
2. Build
3. Publish single-file self-contained EXE
4. Build installer via Inno Setup
5. Upload artifacts (EXE + installer)

On tag push matching `v*` (example: `v1.0.0`):
- Same build pipeline + auto GitHub Release with attached binaries.

## End-to-End Release Flow

1. Commit and push code to `main`.
2. Create and push a version tag:

```bash
git tag v1.0.0
git push origin v1.0.0
```

3. Open GitHub **Releases**.
4. Download either:
   - `ExcelSearchTool.exe` (portable single-file)
   - `ExcelSearchTool-Setup.exe` (installer)
5. Run installer (`Next -> Next -> Finish`) and launch from desktop/start menu shortcut.

No separate .NET runtime installation is required for end users.
