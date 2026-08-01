# BRC Disk Usage Calculator

A single-file Windows PowerShell GUI that measures a folder or UNC share and
exports **one CSV containing a row for every directory and a row for every file**.

Built to match the look, feel and CSV style of the other BRC tools
(`BRCFolderPermissionIndexer.ps1`, `BRCStaleDirectoryFinder.ps1`).

- Target: Windows Server 2019 / Windows PowerShell 5.1 (also runs on PowerShell 7 on Windows)
- Read-only: the tool never modifies files or folders

## Running it

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\BRCDiskUsageCalculator.ps1
```

1. Type or **Browse...** to a local folder or a UNC path (`\\server\share\folder`).
2. Pick the output CSV with **Save As...** (defaults to your Desktop).
3. Set the options, then **Start Scan**. **Cancel** stops a long scan and still
   writes whatever was measured.

### Options

| Option | Meaning |
| --- | --- |
| Only report files at least *n* MB | Filters **file rows** only. Folder totals always include every file, so the sizes stay accurate. `0` reports every file. |
| Recurse subfolders | Off = only the selected folder and the files directly in it. |
| Include selected root folder in output | Whether the scan root gets its own row. |
| Follow reparse points | Off by default so junctions/symlinks are listed but not walked, which stops target data being counted twice. |
| Include directory rows / Include file rows | Choose one or both. At least one is required. |
| Open CSV when complete | Launches the CSV in the default handler after the scan. |

## How the scan works

1. **Measure pass** – walks the tree once and rolls up sizes and counts from the
   bottom up, so every folder knows its own total.
2. **Write pass** – walks the tree again breadth-first and streams the CSV out
   through a `StreamWriter`, so rows appear parent-then-children and memory
   stays proportional to the folder count rather than the file count.

Permission failures and unreadable items do not stop the scan. They are recorded
in the `Errors` column of the affected folder's row, or as an `ItemType=Error`
row when the failure happens against a single file.

## CSV columns

| Column | Directory rows | File rows |
| --- | --- | --- |
| `Path` | Full folder path | Full file path |
| `ParentPath` | Containing folder | Containing folder |
| `Name` | Folder name | File name |
| `ItemType` | `Directory` | `File` (or `Error`) |
| `Extension` | empty | `.docx`, `.pst`, ... |
| `Depth` | Levels below the scan root (root = 0) | Levels below the scan root |
| `SizeBytes` / `SizeMB` / `SizeGB` | **Total** size of the folder and everything under it | File length |
| `SizeFriendly` | e.g. `4.62 GB` | e.g. `12.40 MB` |
| `DirectSizeBytes` / `DirectSizeMB` | Files sitting directly in that folder only | empty |
| `TotalSizeBytes` / `TotalSizeMB` / `TotalSizeGB` | Folder plus all descendants | File length |
| `PercentOfScanTotal` | Share of the whole scan | Share of the whole scan |
| `DirectFileCount` / `DirectSubfolderCount` | Immediate children only | empty |
| `TotalFileCount` / `TotalSubfolderCount` | Recursive counts | empty |
| `CreationTime` / `LastWriteTime` / `LastAccessTime` | `yyyy-MM-dd HH:mm:ss` | `yyyy-MM-dd HH:mm:ss` |
| `Attributes` | Folder attributes | File attributes |
| `IsReparsePoint` | `True` for junctions/symlinks | `True` for symlinked files |
| `ScanRoot` | The path that was scanned | The path that was scanned |
| `ErrorCount` / `Errors` | Problems hit reading that folder | Problems hit reading that file |

`SizeBytes` is deliberately populated for both row types, so sorting the whole
CSV by `SizeBytes` descending puts the biggest folders and the biggest files in
one ranked list.

### Handy Excel/PowerShell follow-ups

```powershell
# Top 25 space consumers of any type
Import-Csv .\DiskUsage_20260801_101500.csv |
    Sort-Object { [long]$_.SizeBytes } -Descending |
    Select-Object -First 25 ItemType, SizeFriendly, PercentOfScanTotal, Path

# Biggest folders by their own content, ignoring what is nested below them
Import-Csv .\DiskUsage_20260801_101500.csv |
    Where-Object { $_.ItemType -eq 'Directory' } |
    Sort-Object { [long]$_.DirectSizeBytes } -Descending |
    Select-Object -First 25 Path, DirectSizeBytes, DirectFileCount

# Space by file type
Import-Csv .\DiskUsage_20260801_101500.csv |
    Where-Object { $_.ItemType -eq 'File' } |
    Group-Object Extension |
    Select-Object Name, Count, @{ n = 'GB'; e = { [math]::Round((($_.Group | Measure-Object { [long]$_.SizeBytes } -Sum).Sum) / 1GB, 2) } } |
    Sort-Object GB -Descending
```

## Notes and limits

- Sizes are logical file lengths (what the files contain), not size-on-disk, so
  they do not account for NTFS compression, deduplication or cluster slack.
- Paths longer than 260 characters can fail to enumerate on Windows PowerShell
  5.1; those failures are captured in the `Errors` column rather than aborting
  the scan.
- A share with millions of files produces a very large CSV. Use the minimum file
  size filter, or clear **Include file rows**, when you only need folder-level
  numbers.
