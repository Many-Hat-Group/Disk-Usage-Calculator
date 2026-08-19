# BRC storage tools

| Tool | What it does | Docs |
| --- | --- | --- |
| `BRCDiskUsageCalculator.ps1` | Measures a folder or share, one CSV row per directory and per file | this file |
| `BRC-StaleDirectoryFinder.ps1` | Finds directories nothing has modified in *n* days/months/years | in-script header |
| `BRC-FolderPermissionIndexer.ps1` | Indexes NTFS permissions across a tree | in-script header |
| `BRC-FolderArchiver.ps1` | **Moves** stale folders to an archive, leaving a `.txt` signpost with a 24 character reference code, and reports it all as CSV | [BRC-FolderArchiver.md](BRC-FolderArchiver.md) |

Everything here is read-only **except `BRC-FolderArchiver.ps1`**, which moves and
deletes data. Read its documentation before running it.

---

# BRC Disk Usage Calculator

A single-file Windows PowerShell GUI that measures a folder or UNC share and
exports **one CSV containing a row for every directory and a row for every file**.

Built to match the look, feel and CSV style of the other BRC tools
(`BRCFolderPermissionIndexer.ps1`, `BRCStaleDirectoryFinder.ps1`).

- Target: Windows Server 2019 / Windows PowerShell 5.1 (also runs on PowerShell 7 on Windows)
- Read-only: the tool never modifies files or folders under the scan root
- Saves as it goes, and can pick up where it left off after a cancel, a crash or
  a reboot

## Running it

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\BRCDiskUsageCalculator.ps1
```

1. Type or **Browse...** to a local folder or a UNC path (`\\server\share\folder`).
2. Pick the output CSV with **Save As...** (defaults to your Desktop).
3. Set the options, then **Start Scan**. **Cancel** stops a long scan and keeps
   everything measured so far.

### Options

| Option | Meaning |
| --- | --- |
| Only report files at least *n* MB | Filters **file rows** only. Folder totals always include every file, so the sizes stay accurate. `0` reports every file. |
| Recurse subfolders | Off = only the selected folder and the files directly in it. |
| Include selected root folder in output | Whether the scan root gets its own row. |
| Follow reparse points | Off by default so junctions/symlinks are listed but not walked, which stops target data being counted twice. |
| Include directory rows / Include file rows | Choose one or both. At least one is required. |
| Open CSV when complete | Launches the CSV in the default handler after the scan. |
| Resume from checkpoint if one exists | On by default. See below. |
| Fill PercentOfScanTotal when the scan finishes | Rewrites the finished CSV once to fill the percentage column. |
| Delete checkpoint file when complete | On by default. Turn it off if you want the checkpoint kept as a record. |
| Save a checkpoint at least every *n* seconds | How often progress is committed. Lower = less rework after a crash, slightly more IO. |

## Incremental save and resume

Rows are appended to the CSV while the scan runs, so the file grows as it works
and a scan that dies part way through still leaves usable data. Progress is
recorded in a sidecar file next to the CSV:

```
C:\Reports\DiskUsage_20260806_101500.csv          <- the data
C:\Reports\DiskUsage_20260806_101500.csv.resume   <- the checkpoint
```

**To resume:** run the tool again with the *same folder, same CSV path and the
same options*, leave **Resume from checkpoint if one exists** ticked, and press
**Start Scan**. Completed folders are skipped and the scan carries on from where
it stopped. Repeat as many times as needed — a scan can be stopped and resumed
any number of times.

The checkpoint is deleted automatically when a scan finishes the whole tree
(unless you untick that option).

### How it stays consistent

The checkpoint is an append-only, tab-separated ledger. Each line names a folder
that is finished and the **byte offset in the CSV** that was flushed to disk at
that moment:

- `F` — the file rows for this folder are written.
- `D` — this folder and everything under it is written, including its own
  directory row.

The rule that makes it safe: *everything in the CSV before the last recorded
offset belongs to a folder that has a ledger line*. Anything after that offset is
unrecorded work, so a resume truncates the CSV back to the offset and redoes it.
That is what prevents both duplicated rows and missing rows when a scan is killed
mid-folder. The CSV is flushed to disk *before* the ledger claims those bytes, so
a crash in the gap costs a little rework and never loses data.

The ledger also gets compacted on each resume — once a parent folder is complete,
its children's lines are redundant and get dropped — so it does not grow without
bound across repeated resumes of a big share.

A checkpoint is **refused** (and the tool offers to start over) when:

- the scan root or any option that changes which rows get written has changed
  since the checkpoint was taken;
- the CSV is smaller than the checkpoint expects, meaning it was replaced,
  truncated or edited between runs;
- the CSV named by the checkpoint no longer exists.

A half-written final ledger line, which is what a power cut leaves behind, is
detected and discarded along with the work it referred to.

## How the scan works

The walk is depth first, and a folder's own directory row is written only once
its whole subtree is finished — a folder cannot know its total size before its
children are counted. Two consequences worth knowing:

- **Row order.** A folder's file rows come first, then its subfolders' blocks,
  then the folder's own directory row. The scan root's row is the last row in the
  file. Sort the CSV however you like in Excel; the order is not meaningful.
- **PercentOfScanTotal.** The grand total is not known until the root finishes,
  so the column is written empty and filled in by a single rewrite pass at the
  end. It stays empty if you cancel, or if you untick the finalise option. The
  rewrite goes to a temporary file and only swaps in once it is complete, so the
  scan data is never at risk.

Memory use is proportional to the *depth* of the tree, not the number of files,
so a share with millions of files scans in a flat memory footprint.

Permission failures and unreadable items do not stop the scan. They are recorded
in the `Errors` column of the affected folder's row, or as an `ItemType=Error`
row when the failure happens against a single file. A folder that failed to
enumerate is deliberately left out of the checkpoint so a resume retries it.

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
| `PercentOfScanTotal` | Share of the whole scan (filled at the end) | Share of the whole scan (filled at the end) |
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
- A resumed scan reports sizes as they were when each folder was measured, which
  may have been on an earlier day. For an exact point-in-time picture of a share
  that is actively changing, run a scan that completes in one go.
- Do not edit the CSV between runs while a checkpoint exists for it. The
  checkpoint tracks byte offsets into that file; the tool detects a shortened
  file and refuses to resume, but an edit that keeps the length the same cannot
  be detected. Copy the CSV elsewhere if you want to work on it mid-scan.
