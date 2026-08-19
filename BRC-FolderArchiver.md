# BRC Folder Archiver

A single-file Windows PowerShell GUI that moves stale folders out of live
storage into an archive location, leaves a `.txt` signpost behind in place of
every folder it moves, and reports the lot as a CSV.

Built to match the look, feel and CSV style of the other BRC tools
(`BRCDiskUsageCalculator.ps1`, `BRC-StaleDirectoryFinder.ps1`,
`BRC-FolderPermissionIndexer.ps1`).

- Target: Windows Server 2019 / Windows PowerShell 5.1 (also runs on
  PowerShell 7 on Windows)
- **This tool changes data.** Everything else in this repository is read-only.
  Read [Safety model](#safety-model) before the first live run.

> **Use this document as the as-built record.** Upload it to the team wiki
> alongside the CSV from your first production run, so the archive is
> documented independently of this repository.

## What it produces

Take a folder that has not been touched in years:

```
\\srv-files\data\2020\01\
```

After a run it becomes:

```
\\srv-files\data\2020\01_ARCHIVED_9TFLLADHMUFD4GEGCBKT8Q6C.txt   <- the signpost
```

and on the archive side:

```
\\srv-archive\archive\2020\01\                                   <- the data
\\srv-archive\archive\2020\9TFLLADHMUFD4GEGCBKT8Q6C.txt          <- the code file
\\srv-archive\archive\2020\9TFLLADHMUFD4GEGCBKT8Q6C_manifest.txt <- copy of the signpost
\\srv-archive\archive\_BRC_ArchiveIndex.csv                      <- every archival, ever
```

plus a CSV report wherever you asked for it.

### The signpost left behind

```
==============================================================================
  THIS CONTENT HAS BEEN ARCHIVED
==============================================================================

The folder that used to be in this location has been moved to long-term
archive storage because nothing inside it had been modified for at least
3 Years.

Nothing has been deleted. Every folder and file listed further down this
file still exists in the archive and can be restored on request.

If you need anything from this content, please contact I.T. and quote the
archive reference code below. Please do not delete this file - it is the
only pointer back to where the data went.

  ARCHIVE REFERENCE CODE:  9TFLLADHMUFD4GEGCBKT8Q6C

  Contact:  I.T. Service Desk
  Contact:  servicedesk@example.local
  Contact:  x4500

------------------------------------------------------------------------------
SUMMARY
------------------------------------------------------------------------------
Original location  : \\srv-files\data\2020\01
Archived to        : (held by I.T. - quote the reference code above)
Archived on        : 2026-08-19 14:03:22
Archived by        : EXAMPLE\p.smith on SRV-ADMIN01
Age rule applied   : not modified in the last 3 Years (cut-off 2023-08-19)
Age measured by    : Recursive newest item timestamp
Last modified      : 2021-02-08 09:12:44
Folders archived   : 3
Files archived     : 4
Total size         : 7.04 KB (7,212 bytes)

------------------------------------------------------------------------------
EVERY FOLDER, SUBFOLDER AND FILE THAT WAS ARCHIVED
------------------------------------------------------------------------------
Paths are shown relative to the archived folder itself.
Top level: \\srv-files\data\2020\01
[DIR ] \   (2 subfolders, 4 files, 7.04 KB)
  [FILE] summary.txt                                     1001 bytes 2021-02-08 09:12:44
  [DIR ] \invoices   (1 subfolder, 3 files, 6.07 KB)
    [FILE] inv-001.pdf                                      2.00 KB 2021-02-08 09:12:44
    [DIR ] \invoices\scans   (0 subfolders, 1 file, 4.00 KB)
      [FILE] scan01.tif                                     4.00 KB 2021-02-08 09:12:44
```

Every folder, subfolder and file that went to the archive is named, with its
size and last-modified date, so a user can tell from the signpost alone whether
what they want is in there.

### The reference code

24 characters, drawn from a cryptographic RNG over a 32 character alphabet with
`I`, `O`, `0` and `1` left out so nobody misreads one over the phone. That is
120 bits, so codes do not collide in practice; the tool also checks each new
code against the codes already in `_BRC_ArchiveIndex.csv` and against the codes
issued earlier in the same run.

The code appears in four places on purpose:

| Where | Why |
| --- | --- |
| In the signpost text | The user reads it off the screen and quotes it |
| In the signpost file name | Findable by a file-name search of the live share |
| As the archive-side file name `<CODE>.txt` | Findable by a file-name search of the archive |
| As the entire contents of `<CODE>.txt` | Findable by a content search of the archive |

So whichever way I.T. searches, and whichever side they search, the code lands
them in the right folder.

## Running it

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\BRC-FolderArchiver.ps1
```

1. Set the **source data folder** and the **archive destination**.
2. Set the age threshold and pick how folders are chosen.
3. Put your real service-desk details in the **contact** box.
4. Leave **Preview only** ticked and press **Run Preview**. Read the CSV.
5. Untick **Preview only**, press **Archive Now**, type `ARCHIVE` when asked.

**Cancel** stops after the folder in progress finishes, so a folder is never
left half-copied with its original already gone.

## Options

### What to archive

| Option | Meaning |
| --- | --- |
| Not modified in the last *n* Days / Weeks / Months / Years | The age threshold. Anything whose last modification is older than the resulting cut-off date is in scope. |
| Age measured by | **Recursive newest item timestamp** treats a folder as active if anything anywhere inside it has changed recently — the right choice for archiving. **Folder timestamp only** uses the folder's own `LastWriteTime`, which is faster but misses recent activity in subfolders. |
| Folder selection | Which folders are treated as one archivable unit. See below. |
| Depth | Only used by **Fixed depth below the source root**. |
| Never archive folders named | Semicolon separated, wildcards allowed. Excluded folders are not archived and are not descended into. |
| Stop after this many folders | A cap for the first live runs. `0` means no cap. Run again to continue. |

**Folder selection** decides what a single archive unit is:

| Mode | Behaviour |
| --- | --- |
| **Top-most stale folder** (default) | Walks down and archives the highest folder that is entirely stale. If nothing in `\data\2020` has changed for years, the whole of `\data\2020` goes as one unit with one signpost. If `\data\2020\12` is still active, only its stale siblings go. Fewest signposts, least clutter. |
| **Fixed depth below the source root** | Only folders exactly *n* levels below the source root are considered. With source `\\srv\data` and depth `2`, the units are `\\srv\data\2020\01`, `\\srv\data\2020\02` and so on — one signpost per month, regardless of whether the whole year is stale. |
| **Immediate subfolders of the source root only** | Depth 1. Nothing deeper is looked at. |

The source root itself is never archived, and loose files sitting directly in a
folder that is not being archived are left where they are.

### How to archive

| Option | Meaning |
| --- | --- |
| Copy engine | **Robocopy** (default) handles long paths and retries, and is the better choice for large trees. **PowerShell** uses .NET file copies and needs no external binary. |
| Verify every file with SHA-256 | Off by default. Off still verifies every file by relative path and byte length. On also hashes both sides — much slower, and the strongest guarantee before a delete. |
| Copy NTFS permissions | Robocopy only (`/COPY:DATSO`). Off by default so archived data simply inherits the archive share's permissions. See [Security](#security). |
| Mirror the source folder structure | On by default: `\data\2020\01` lands at `<archive>\2020\01`. Off puts every archived folder directly under the archive root by its own name. |
| Group this run in a dated subfolder | Puts everything from one run under `<archive>\ArchiveRun_yyyyMMdd_HHmmss\`. |
| Send originals to the Recycle Bin | A second safety net on a local volume. Has no effect on most UNC shares, which have no Recycle Bin — do not rely on it there. |
| Clear read-only flags before removing | On by default. Without it a read-only file blocks the delete after the copy has already verified. |
| Skip folders holding junctions/symlinks | On by default. A junction, symlink or mount point can point anywhere, including outside the tree or back into it, so those folders are reported and left alone. See [Reparse points](#reparse-points-dedup-junctions-and-tiered-storage). |
| Skip cloud-tiered, HSM and unrecognised stub files | On by default. Files whose data lives elsewhere and is recalled on read. Copying them pulls every file back, which can be slow and can cost money. |
| Skip folders that contain no files | On by default. |
| Remove the part-copy if a folder fails | On by default. The source is untouched either way; this just stops half-copies accumulating in the archive. |
| List skipped folders in the CSV too | On by default, so the CSV is a full audit of the decision, not just the successes. |

### The .txt files

| Option | Meaning |
| --- | --- |
| I.T. contact details | Free text, printed as a `Contact:` line per line you type. Put your real service desk details here. |
| Placeholder file name | Tokens `{FolderName}`, `{Ref}`, `{Date}`. Default `{FolderName}_ARCHIVED_{Ref}.txt`. `.txt` is appended if you leave it off. |
| Also write `<Ref>_manifest.txt` in the archive | On by default. Puts a copy of the signpost next to the archived data, so the archive is self-describing even if the live share is later rebuilt. |
| Print the archive path in the placeholder | Off by default. Off, the signpost says to quote the reference code instead. On, it names the archive path — convenient for I.T., but it also tells every user where the archive lives. |
| Append to `_BRC_ArchiveIndex.csv` | On by default. A cumulative index of every archival across every run, kept in the archive root. |

## Reparse points: dedup, junctions and tiered storage

"Reparse point" is one NTFS mechanism covering two very different things, and
the archiver treats them differently because only one of them is dangerous.

| The item is | Examples | Archiver behaviour |
| --- | --- | --- |
| **A link** — stands in for something elsewhere | Junction, symbolic link, volume mount point, NFS special file | **Blocks the folder.** Copying it would pull in data from outside, and removing it could reach through to the target. |
| **A stub** — the data really is here, stored differently | **Data Deduplication**, Windows Overlay compression, Single Instance Storage | **Ignored.** These are ordinary files. Reads return the real bytes at local speed. |
| **A tiered stub** — the data is real but lives elsewhere | HSM, Azure File Sync cloud tiering, OneDrive placeholders | **Blocks the folder**, under its own switch, because copying recalls every file. |

**This matters on any server running Data Deduplication.** Dedup stores every
optimized file as a reparse point, and those files show up as
`SparseFile, ReparsePoint`. A tool that treats "has a reparse point" as "do not
touch" will refuse to archive essentially the whole volume. The archiver reads
the actual reparse *tag* rather than just the attribute, so deduplicated files
are archived normally.

The tag is read with `FindFirstFileW`, which returns it without opening the
file. If that is unavailable — Constrained Language Mode, `Add-Type` blocked by
policy, a path that will not enumerate — every reparse point is treated as a
link, which is the conservative answer. You will see this in the CSV as
`Reparse point of an unreadable type`.

Unknown tags are handled by rule rather than guesswork: a tag carrying the
documented **name-surrogate bit** is a link, because that bit means the item
stands in for another named object. Anything else unrecognised is grouped with
the tiered stubs and blocked under that switch.

Skip messages name the tag and an example path, so the CSV tells you what was
found rather than just that something was:

```
Skipped because this folder contains a junction, symbolic link or mount point:
'D:\BRCDATA1\PROJECTS\oldlink' is Junction or volume mount point (tag 0xA0000003)
and 2 other such items. Nothing was copied or deleted, because it points
somewhere else, so copying it would pull in data from outside this folder and
removing it could reach through to the target.
```

### Finding reparse points yourself

```powershell
# What in this folder is not an ordinary file?
Get-ChildItem 'D:\BRCDATA1' -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint } |
    Select-Object FullName, Attributes, LinkType, Target

# The exact tag for one item (needs an elevated session)
fsutil reparsepoint query 'D:\BRCDATA1\some\item'
```

A `LinkType` of `Junction` or `SymbolicLink` with a `Target` is a real link.
Ordinary documents showing `SparseFile, ReparsePoint` with no `LinkType` are
dedup stubs.

## ⚠️ Capacity planning on a deduplicated source

**Copying a deduplicated file rehydrates it.** The archive receives full-size
files, so the archive volume must hold the *logical* size of the data, not the
smaller figure the source volume appears to use.

Check the gap before choosing a destination:

```powershell
# Elevated session required
Get-DedupVolume | Select-Object Volume, SavedSpace, SavingsRate
```

A volume showing a 42 % savings rate holds roughly 1.7 times more logical data
than it occupies. Archiving from it to a plain volume needs that larger figure.

The CSV's `TotalSizeBytes` and `SizeFriendly` columns already report logical
sizes, so the preview CSV gives you the correct number for the target:

```powershell
$Rows = Import-Csv .\FolderArchive_20260819_140322.csv |
    Where-Object { $_.Status -eq 'WouldArchive' }

'{0:N2} GB will land in the archive' -f ((($Rows | Measure-Object -Property TotalSizeBytes -Sum).Sum) / 1GB)
```

Two ways to close the gap:

- **Enable Data Deduplication on the archive volume.** Archived data dedupes
  very well, and it is cold on arrival, so set
  `Set-DedupVolume -Volume X: -MinimumFileAgeDays 1` rather than leaving the
  three-day default.
- **Use cloud tiering on the archive** (Azure File Sync). Cheaper for cold
  data, but note that the archive then becomes a tiered volume itself, so a
  future restore recalls from the cloud.

Reading deduplicated files also costs I/O on the source while the dedup filter
rehydrates them, so a large archive run is heavier on the source server than
the file sizes suggest. Run it outside business hours.

## Safety model

The tool is destructive by exception, not by default:

1. **Preview only is on when the GUI opens.** A preview reads the tree, picks
   the folders, and writes the full CSV — and creates nothing at all, not even
   the archive root.
2. **A live run needs the word `ARCHIVE` typed in**, then a second
   confirmation. There is no single button that starts deleting.
3. **Copy, then verify, then delete — in that order.** Every file is checked by
   relative path and byte length (and SHA-256 if you asked for it) before
   anything is removed.
4. **Any failure leaves the source alone.** A copy error, a verification
   mismatch, or an unreadable file during inventory all end the same way: the
   folder is reported in the CSV with the reason, and the original stays where
   it is.
5. **A folder that cannot be read completely is never archived.** A permissions
   error during inventory is a skip, not a partial archive.
6. **The archive root may not be inside the source root, or vice versa.** Both
   are refused before anything runs.
7. **If the copy verifies but the delete fails**, the run says so explicitly and
   marks the row `Failed` — both copies exist, and a human decides.

What the tool does **not** protect you from:

- Archiving a folder that someone still needs. The age threshold is a proxy for
  "unwanted", not a fact. Preview, then circulate the CSV, then run.
- Loss of the archive itself. Once a folder is archived, the archive copy is the
  only copy. **The archive must be backed up.**

## Security

Read this before pointing the tool at a production share.

- **The archive holds the only copy.** Anyone with write access there can delete
  or alter data that no longer exists anywhere else. Restrict the archive share
  and its NTFS permissions to I.T. with Modify; give end users read-only, or no
  access at all. The tool checks the archive root when a run starts and warns if
  `Everyone`, `Authenticated Users`, `Domain Users` or `Users` hold write,
  delete, change-permissions or take-ownership rights.
- **The signposts are a map of your archive.** Every placeholder file names the
  data that used to be there. Leaving **Print the archive path in the
  placeholder** off keeps the archive's location out of that map.
- **Back the archive up.** Archiving is not a backup. Moving data out of a
  backed-up share and into one that is not backed up converts stale data into
  data with no recovery path at all. Confirm the archive location is in your
  backup scope before the first live run.
- **Permission inheritance changes on the move.** By default the archived copy
  inherits the archive location's permissions rather than carrying the source
  ACLs with it. That is usually what you want. If you need the original ACLs
  preserved for an audit or a legal hold, use the robocopy engine with **Copy
  NTFS permissions**, and run as an account with `SeBackupPrivilege` /
  `SeRestorePrivilege`.
- **Run as a service account with only the access it needs** — Modify on the
  source, Modify on the archive. Not Domain Admin.
- **The CSV and the index name real paths and, indirectly, real business
  content.** Store them where the archive itself is stored, not on a general
  share.

## Searching for archived content

A user rings up quoting `9TFLLADHMUFD4GEGCBKT8Q6C`. Any of these finds it:

```powershell
# 1. The master index - fastest, and shows both paths at once
Import-Csv \\srv-archive\archive\_BRC_ArchiveIndex.csv |
    Where-Object { $_.ReferenceCode -eq '9TFLLADHMUFD4GEGCBKT8Q6C' } |
    Select-Object OriginalPath, ArchivePath, ArchivedLocal, ArchivedBy, FileCount, SizeFriendly

# 2. The archive itself, by file name
Get-ChildItem \\srv-archive\archive -Recurse -Filter '9TFLLADHMUFD4GEGCBKT8Q6C*'

# 3. The archive itself, by file contents
Get-ChildItem \\srv-archive\archive -Recurse -Filter '*.txt' |
    Select-String -SimpleMatch '9TFLLADHMUFD4GEGCBKT8Q6C' |
    Select-Object -ExpandProperty Path

# 4. Which signpost on the live share points at it
Get-ChildItem \\srv-files\data -Recurse -Filter '*_ARCHIVED_*.txt' |
    Select-String -SimpleMatch '9TFLLADHMUFD4GEGCBKT8Q6C'
```

Windows Search finds it too, from either side, as long as the share is indexed —
that is what the `<CODE>.txt` file is for: its name and its entire contents are
the code.

### Restoring

```powershell
$Row = Import-Csv \\srv-archive\archive\_BRC_ArchiveIndex.csv |
    Where-Object { $_.ReferenceCode -eq '9TFLLADHMUFD4GEGCBKT8Q6C' }

robocopy $Row.ArchivePath $Row.OriginalPath /E /COPY:DAT /DCOPY:DAT /R:2 /W:2

# then remove the signpost, once the restore is confirmed
Remove-Item -LiteralPath $Row.PlaceholderFile
```

Restore back to the original path unless there is a reason not to — the
`OriginalPath` column is exactly where it came from. Leave the archive copy in
place until the user confirms they have what they need.

## CSV columns

Both the per-run CSV and `_BRC_ArchiveIndex.csv` use these columns.

| Column | Meaning |
| --- | --- |
| `Status` | `Archived`, `ArchivedWithWarnings`, `WouldArchive` (preview), `Skipped`, `Failed` |
| `ReferenceCode` | The 24 character code |
| `OriginalPath` | Where the folder was |
| `ArchivePath` | Where it is now (in preview, where it would go) |
| `PlaceholderFile` | Full path of the `.txt` left behind |
| `ArchiveCodeFile` | Full path of the archive-side `<CODE>.txt` |
| `RelativePath` / `FolderName` / `OriginalParentPath` / `DepthFromSourceRoot` | Position within the source tree |
| `LastModifiedUsed` / `DaysSinceModified` | The timestamp the decision was made on |
| `CutoffDate` / `ThresholdValue` / `ThresholdUnit` / `AgeBasis` | The rule that was applied |
| `FolderCount` | Folders archived, including the archived folder itself |
| `FileCount` | Files archived |
| `TotalSizeBytes` / `TotalSizeMB` / `TotalSizeGB` / `SizeFriendly` | Size of the unit |
| `CopyMethod` / `VerifyResult` / `SourceRemoved` | What was done and how it went |
| `ArchivedLocal` / `ArchivedUtc` / `ArchivedBy` / `RunFromHost` | Audit trail |
| `SourceRoot` / `ArchiveRoot` | The roots the run used |
| `ErrorCount` / `Errors` | Why a folder was skipped or failed |

Sizes use the same binary convention as the other BRC tools: 1 MB =
1,048,576 bytes, 1 GB = 1,073,741,824 bytes. `TotalSizeBytes` is the exact
figure if you need to do your own arithmetic.

## Suggested rollout

1. Run **BRC-StaleDirectoryFinder** first to size the problem and agree a
   threshold with the business.
2. Preview with the agreed threshold. Circulate the CSV to data owners and give
   them a deadline to object.
3. Confirm the archive location is locked down and in the backup scope.
4. First live run: set **Stop after this many folders** to something small, like
   10, against one department's folder. Check the signposts, check the archive,
   check a restore actually works.
5. Remove the cap and work through the rest, one share at a time.
6. Keep every run's CSV. It is the only complete record of what moved and when.

## Notes and limits

- Sizes are logical file lengths, not size-on-disk, so NTFS compression,
  deduplication and cluster slack are not accounted for.
- Paths longer than 260 characters can fail on Windows PowerShell 5.1. The
  robocopy engine handles them; the PowerShell engine may not, and a failure
  there is reported and the source left alone.
- The tool is single-threaded and updates the GUI between folders. A very large
  tree takes a while; the window may look busy while a single big folder copies.
- Recursion depth follows the folder tree, so an extremely deep tree
  (thousands of levels) can exhaust the PowerShell call stack.
- On a deduplicated volume the reparse tag is read once per file during
  inventory. That is a fast call, but across hundreds of thousands of files it
  is not free — expect inventory of a large folder to take noticeably longer
  there than on a plain volume.
- A run does not resume. Use the folder cap to work through a big share in
  sittings; folders already archived are simply no longer there to match.
- Archived folder timestamps are preserved. The archive copy's own folder
  creation dates reflect the originals, not the archival date — the archival
  date is in the CSV, the signpost and the manifest.
