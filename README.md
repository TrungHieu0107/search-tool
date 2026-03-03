# Universal File Search Tool

Production-ready Windows desktop app for high-performance searching across Excel, text, and source-code files, with automated GitHub Actions CI/CD and downloadable installer artifacts.

## Supported file types

- Excel: `.xlsx`
- Text: `.txt`, `.log`, `.csv`
- Code/config/web: `.cs`, `.java`, `.js`, `.ts`, `.json`, `.xml`, `.html`, `.css`

## Features

- Modern WinForms desktop UI with responsive layout and smooth DataGridView rendering.
- Search modes:
  - Content + File Name
  - Content Only
  - File Name Only
- Regex + case sensitivity across all file types.
- File-type filters (Excel/Text/Code).
- Instant search with 400ms debounce and cancel-in-flight support.
- Search history persisted with auto-suggest.
- Multi-folder search support.
- Parallel processing with per-extension handler dispatch.
- Excel ZIP pre-scan + EPPlus parsing for accurate cell-level location results.
- Text/code streaming line-by-line processing (memory-efficient).
- Large file skipping via configurable max-size threshold.
- Binary file skip detection for text/code handlers.
- Progress bar/status counters and CSV export.
- Double-click result to open file with default system app/editor.

## Architecture

- `SearchService`: dispatcher/orchestrator, parallel scanning, progress reporting, regex compilation, file-type routing.
- `ISearchHandler`: extension point for file search handlers.
- `ExcelSearchHandler`: ZIP pre-scan + EPPlus workbook/cell search.
- `TextSearchHandler`: streaming line search for txt/log/csv.
- `CodeSearchHandler`: streaming line search for source/config/web files.
- `HistoryService`: query history persistence.
- `FolderManager`: add/remove/clear folders.
- `ErrorLogger`: async log writing under `%LOCALAPPDATA%\ExcelSearchTool\errors.log`.

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
