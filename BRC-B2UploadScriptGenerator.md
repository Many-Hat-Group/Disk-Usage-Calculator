# BRC B2 Upload Script Generator

A single Windows batch (`.bat`) file that turns the CSV from
`BRC-StaleDirectoryFinder.ps1` into a ready-to-run upload job: a WinSCP script
that copies every stale folder listed in the CSV into a Backblaze B2 bucket,
plus a runner `.bat` that calls it.

Unlike the other BRC tools, the generator itself is plain `cmd.exe` batch, not
PowerShell - it needs nothing installed beyond Windows and WinSCP.

- Target: Windows Server 2019 / any Windows with `cmd.exe` and WinSCP installed
- **Read-only against the source.** The generator only reads the CSV and
  checks whether each listed folder still exists. It never touches source
  data.
- **The generated upload is additive by default.** It only adds/updates files
  in the bucket. See [MIRROR_DELETE](#options-in-b2-credentialsbat) before
  turning on the one setting that can delete something remotely.

## What it produces

```
BRC-B2UploadScriptGenerator.bat "C:\Reports\StaleDirectories_20260819_140322.csv"
```

writes, next to the CSV:

```
StaleDirectories_20260819_140322_B2Upload_folders.txt   <- what will be uploaded, and where
StaleDirectories_20260819_140322_B2Upload_winscp.txt    <- the WinSCP script
StaleDirectories_20260819_140322_B2Upload.bat            <- run this one
```

Running the last file uploads everything and writes a WinSCP log next to it.

## One-time setup

### 1. Install WinSCP

Get it from the official site and make sure `winscp.com` (the command-line
build) is either on `PATH` or you know its full install path.

### 2. Create a Backblaze B2 Application Key scoped to one bucket

Do **not** use your B2 master account key in this or any script.

1. B2 Console > **Application Keys** > **Add a New Application Key**.
2. Name it something identifiable (e.g. `brc-upload-<bucket>`).
3. **Allow access to Bucket(s):** pick the one bucket this will upload to -
   not "All".
4. **Type of Access:** Read and Write.
5. Save the **Application Key ID** and the **Application Key** (the secret
   is shown once, at creation time only).
6. Note the bucket's **S3-compatible endpoint** from its page in the console,
   e.g. `s3.us-west-004.backblazeb2.com` (varies by region).

If this key is ever exposed, revoke it in the console immediately and issue a
new one.

### 3. Set up `b2-credentials.bat`

Copy `b2-credentials.bat.example` (next to the generator) to
`b2-credentials.bat` in the same folder, and fill it in. It supports two ways
of holding the credential:

| Option | How | Where the secret lives |
| --- | --- | --- |
| **A - WinSCP saved site (recommended)** | Save a WinSCP site (protocol: Amazon S3, host: the bucket's S3 endpoint, key ID/secret pasted in), then put that site's name in `WINSCP_SITE=`. | Inside WinSCP's own encrypted configuration store. Never in this repo, any generated file, or a plain environment variable. |
| **B - inline Application Key** | Set `B2_KEY_ID`, `B2_APP_KEY`, `B2_ENDPOINT` directly. | In `b2-credentials.bat`, in plain text. Use only if a saved site isn't practical (e.g. a shared machine with no local WinSCP profile), and restrict that file's NTFS permissions. |

Either way, also set `B2_BUCKET`. Full details and every optional setting are
in the comments inside `b2-credentials.bat.example`.

**`b2-credentials.bat` is listed in `.gitignore`.** Never commit your filled-in
copy, paste it into chat, or email it. Only `b2-credentials.bat.example` (the
blank template) is tracked in this repository.

## Running it

```bat
BRC-B2UploadScriptGenerator.bat "C:\Reports\StaleDirectories_20260819_140322.csv"
```

An optional second argument sets the output file base name (defaults to the
CSV's own name):

```bat
BRC-B2UploadScriptGenerator.bat "C:\Reports\StaleDirectories_20260819_140322.csv" MarketingArchive
```

The generator prints how many folders it queued and how many it skipped (not
stale, no longer exists, blank path). **Nothing is uploaded by the generator
itself** - it only writes the three output files.

Before running the generated `*_B2Upload.bat`:

1. Open `*_B2Upload_folders.txt` and check the local -> bucket mapping looks
   right.
2. Recommended: point `B2_REMOTE_PREFIX` at a throwaway test prefix, or trim
   the folder list down to one folder, and run it once to confirm the layout
   in the bucket is what you expect before doing a full run.
3. Run `*_B2Upload.bat`. Progress and the exit code print to the console; the
   full transfer log is written to `*_B2Upload.log` next to it.

## Options (set in `b2-credentials.bat`)

| Setting | Default | Meaning |
| --- | --- | --- |
| `ONLY_STALE` | `1` | Only upload rows where the CSV's `IsStale` column is `True`. Set to `0` to upload every row regardless. |
| `B2_REMOTE_PREFIX` | *(none)* | A folder inside the bucket to upload under, instead of the bucket root. |
| `MIRROR_DELETE` | `0` | **Destructive if turned on.** `1` also deletes files in the bucket that no longer exist in the corresponding local folder. Leave this off unless you specifically want the bucket to mirror local deletions. |
| `WINSCP_EXE` | *(PATH)* | Full path to `WinSCP.com`, only needed if it isn't on `PATH`. |

Change a setting, then re-run the generator to pick it up - it isn't read
again by files already generated.

## How the remote layout works

A local path becomes a bucket path by dropping the drive letter or UNC server
name and keeping the rest, the same convention `BRC-StaleDataArchiver.ps1`
uses for its destination layout:

```
D:\Data\2020\01              ->  /<bucket>/[<prefix>/]D/Data/2020/01
\\srv-files\data\2020\01     ->  /<bucket>/[<prefix>/]srv-files/data/2020/01
```

So two stale folders with the same folder name from different drives or
servers never collide in the bucket.

## Safety model

- **Read-only against the source, always.** The generator and the generated
  upload script only read local folders; neither ever deletes, moves, or
  modifies anything on the machine being scanned.
- **Additive by default on the bucket side.** `synchronize remote` (without
  `-delete`) only adds and updates objects in B2. Nothing already in the
  bucket is removed unless you explicitly set `MIRROR_DELETE=1`.
- **Credentials never land in a committed or shareable file.** Generated
  scripts hold only `%PLACEHOLDER%` tokens; the real key stays in
  `b2-credentials.bat`, which is gitignored, or in WinSCP's own encrypted
  site store. The runner `.bat` resolves the placeholder and passes it to
  `winscp.com` as a command-line argument (WinSCP script files can't read
  environment variables themselves) - with Option B (inline key) that means
  the key is briefly visible in that process's own command line while it
  runs, to anything with local access to inspect running processes. Option A
  (a saved site) has no such exposure, since a site name isn't a secret.
- **Review before you run.** The generator writes the folder list and the
  WinSCP script but performs no transfer itself - you always get a chance to
  read `*_B2Upload_folders.txt` first.

## Known limitations

Batch has no real CSV parser, so this generator does simple comma-splitting
rather than full CSV-quote-aware parsing:

- A folder path containing a literal **comma** will misread that row's
  columns. Windows paths essentially never contain commas; if
  `BRC-StaleDirectoryFinder.ps1` ever reports one, it will show up wrong in
  `*_B2Upload_folders.txt` - check that file before running the upload.
- If the CSV lists both a stale parent folder and a stale folder inside it,
  both are uploaded independently, so the inner one's files transfer twice.
  This wastes time and bandwidth but isn't harmful - WinSCP's synchronize
  just re-verifies files that already match and skips them.

If any of this matters for your data (paths with commas or percent signs at
scale, or you need it to run unattended on a schedule), the PowerShell BRC
tools are the better fit; this generator trades that robustness for having no
dependency beyond `cmd.exe` and WinSCP.

## Troubleshooting

| Symptom | Likely cause |
| --- | --- |
| `ERROR: neither WINSCP_SITE ... nor B2_KEY_ID is set` | `b2-credentials.bat` exists but neither option is filled in / uncommented. |
| `ERROR: WinSCP command-line executable not found` | WinSCP isn't installed, or `winscp.com` isn't on `PATH` and `WINSCP_EXE` isn't set. |
| `Folders queued for upload : 0` | Either the CSV has no `IsStale=True` rows, or every listed folder has since been moved/archived. Check `SKIPPED_MISSING` in the console output. |
| WinSCP exits non-zero | Check `*_B2Upload.log` - usually an authentication failure (bad key, wrong endpoint/region) or a folder that vanished between generation and upload. |
| Console spams "No session", log shows `Host "%WINSCP_SITE%"` (or `%B2_KEY_ID%`) `does not exist` | Files were generated by an older copy of this generator that wrote `open %WINSCP_SITE%` directly into the `.txt` script - WinSCP does not expand `%VAR%` inside script files, so it tried to connect to a literal host named that. Fixed in the current version, which passes the session as a command-line argument instead. Re-run the generator against the same CSV to regenerate the three output files. |
