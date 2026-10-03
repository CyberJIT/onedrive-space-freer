# OneDrive Space Freer (Python & PowerShell)

A cross-platform utility (Python & PowerShell) to scan files in a parametrized OneDrive folder, ensure they are marked for synchronization, and convert them to cloud-only "on-demand" placeholders (`Recall on Data Access` attribute / `attrib +U -P`), instantly freeing up local disk space without deleting files from OneDrive cloud storage.

---

## Features

- **OneDrive Files On-Demand Dehydration**:
  - Unpins local files (`-P` / clears `FILE_ATTRIBUTE_PINNED`).
  - Sets Recall on Data Access (`+U` / sets `FILE_ATTRIBUTE_UNPINNED` and `FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS`).
  - Converts local files to 0-byte cloud placeholders while maintaining full directory view and seamless rehydration on access.
- **Selective Traversal & Recursion**:
  - Optional recursive scanning across nested subfolders (`-r` / `-Recurse`).
  - Flat folder processing by default.
- **Granular Filtering**:
  - **File Name Include/Exclude**: Match specific patterns or substrings (e.g. `*Report*`, exclude `*Draft*`).
  - **Extension Include/Exclude**: Target specific extensions (e.g. `pbix`, `xlsx`, `zip`, `mp4`) or exclude temporary files (`tmp`, `log`).
  - **Folder Include/Exclude**: Target or prune specific directories during recursive traversal (e.g. skip `*Archive*` or `*Backup*`).
- **Resilient Path Handling**:
  - Automatic Windows extended-length path prefix (`\\?\`) to safely bypass `MAX_PATH` (260 characters).
  - Win32 8.3 short path fallback (`GetShortPathNameW`) if extended path fails.
- **Graceful Error Recovery**:
  - File-level exceptions (locked files, permissions, network drops) are captured, logged, and bypassed so execution continues uninterrupted.
- **Actionable Execution Summary**:
  - Final metrics detailing total discovered, newly marked, already online-only, failed items, candidate local disk space freed, and error details.
- **Dry-Run Simulation**:
  - Preview candidate files and space impact without modifying file attributes (`--dry-run` / `-DryRun`).

---

## 1. Python Tool (`onedrive_space_freer.py`)

Requires Python 3.7+ (pure standard library, zero external dependencies). Works on Windows with native `ctypes` Win32 API calls (`GetFileAttributesW`, `SetFileAttributesW`, `GetShortPathNameW`) and `attrib.exe` fallback.

### Usage

```bash
# Basic scan: free space for all files in a single folder
python onedrive_space_freer.py -p "C:\Users\Username\OneDrive - Company\Reports"

# Recursive scan with dry-run preview
python onedrive_space_freer.py -p "C:\Users\Username\OneDrive" -r --dry-run

# Target large file types only (e.g. .pbix, .zip, .mp4, .csv)
python onedrive_space_freer.py -p "C:\Users\Username\OneDrive" -r \
  --ext-include pbix zip mp4 csv

# Filter by file name and exclude temporary files
python onedrive_space_freer.py -p "C:\Users\Username\OneDrive" -r \
  --file-include "*Finance*" "*Sales*" \
  --file-exclude "*Draft*" \
  --ext-exclude tmp log bak

# Exclude specific folders during recursion
python onedrive_space_freer.py -p "C:\Users\Username\OneDrive" -r \
  --folder-exclude "*Archive*" "*Old*"
```

### CLI Parameters

| Flag | Full Option | Description |
| :--- | :--- | :--- |
| `-p` | `--path` | **(Required)** Path to target OneDrive directory. |
| `-r` | `--recursive` | Recursively process subdirectories. |
| | `--file-include` | Substring(s) or wildcard(s) to include file names. |
| | `--file-exclude` | Substring(s) or wildcard(s) to exclude file names. |
| | `--ext-include` | Extension(s) to include (e.g. `pbix`, `pdf`, `docx`). |
| | `--ext-exclude` | Extension(s) to exclude (e.g. `tmp`, `log`). |
| | `--folder-include` | Substring(s) or wildcard(s) to include directories. |
| | `--folder-exclude` | Substring(s) or wildcard(s) to exclude directories. |
| `-n` | `--dry-run` | Preview actions without changing file attributes. |

---

## 2. PowerShell Tool (`Free-OneDriveSpace.ps1`)

Compatible with Windows PowerShell 5.1+ and PowerShell Core 7+. Uses inline C# Win32 P/Invoke for kernel attribute manipulation and `attrib.exe +U -P` fallback.

### Usage

```powershell
# Basic scan: free space in a single folder
.\Free-OneDriveSpace.ps1 -Path "C:\Users\Username\OneDrive - Company\Reports"

# Recursive scan with dry-run preview
.\Free-OneDriveSpace.ps1 -Path "C:\Users\Username\OneDrive" -Recurse -DryRun

# Target specific extensions
.\Free-OneDriveSpace.ps1 -Path "C:\Users\Username\OneDrive" -Recurse `
  -ExtInclude "pbix", "zip", "mp4"

# Filter by file name and exclude folders
.\Free-OneDriveSpace.ps1 -Path "C:\Users\Username\OneDrive" -Recurse `
  -FileInclude "*Sales*", "*Finance*" `
  -FileExclude "*Draft*" `
  -FolderExclude "*Archive*", "*Old*"
```

### Script Parameters

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `-Path` | `String` | **(Required)** Path to directory containing OneDrive files. |
| `-Recurse` | `Switch` | Search and process subdirectories recursively. |
| `-FileInclude` | `String[]` | Substring(s) or wildcard(s) to include file names. |
| `-FileExclude` | `String[]` | Substring(s) or wildcard(s) to skip file names. |
| `-ExtInclude` | `String[]` | Extension(s) to include (e.g. `pbix`, `pdf`, `xlsx`). |
| `-ExtExclude` | `String[]` | Extension(s) to exclude (e.g. `tmp`, `log`). |
| `-FolderInclude` | `String[]` | Substring(s) or wildcard(s) of folders to traverse. |
| `-FolderExclude` | `String[]` | Substring(s) or wildcard(s) of folders to skip. |
| `-DryRun` | `Switch` | Preview operations without touching disk files. |

---

## How It Works (Technical Details)

OneDrive on Windows uses NTFS Reparse Points and Cloud Filter API attributes:
1. `FILE_ATTRIBUTE_PINNED` (`0x00080000`): Marks a file as "Always keep on this device".
2. `FILE_ATTRIBUTE_UNPINNED` (`0x00100000`): Informs OneDrive the user requested freeing up local space.
3. `FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS` (`0x00400000`): Cloud placeholder state. The file's data payload is offloaded to the cloud and released from physical disk blocks, converting it to an on-demand hydration pointer.
4. Path workarounds: Paths exceeding 260 characters automatically receive the `\\?\` prefix or fallback to Win32 8.3 short paths via `GetShortPathNameW`.

---

## Running Tests

Execute the unit test suite with:

```bash
python3 -m unittest test_onedrive_space_freer.py
```
