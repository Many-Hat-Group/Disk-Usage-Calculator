@echo off
setlocal EnableDelayedExpansion
:: ============================================================================
:: BRC B2 Upload Script Generator
:: Windows batch (cmd.exe) - no PowerShell required
:: ============================================================================
::
:: What it does:
:: - Takes the CSV that BRC-StaleDirectoryFinder.ps1 produced
:: - Reads the Path and IsStale columns
:: - Writes a WinSCP script that uploads every stale folder to a Backblaze B2
::   bucket, mirroring each folder's structure under the bucket
:: - Writes a runner .bat that loads your B2 credentials and calls WinSCP with
::   that script
::
:: What it does NOT do:
:: - It never deletes or modifies anything at the source. It only reads the CSV
::   and the existence of the folders it lists.
:: - The generated upload only ADDS/UPDATES files in the bucket. It does not
::   delete remote files unless you explicitly set MIRROR_DELETE=1 in your
::   credentials file (see b2-credentials.bat.example). Leave that off unless
::   you specifically want the bucket to mirror deletions too.
::
:: Requirements:
:: - WinSCP installed (winscp.com on PATH, or WINSCP_EXE set in your
::   credentials file)
:: - A Backblaze B2 bucket and an Application Key SCOPED TO THAT BUCKET ONLY
::   with read+write access - not your master account key. Create one at:
::   B2 Console > Application Keys > Add a New Application Key.
::
:: Usage:
::   BRC-B2UploadScriptGenerator.bat "C:\Reports\StaleDirectories_2026...csv"
::   BRC-B2UploadScriptGenerator.bat "C:\Reports\StaleDirectories_2026...csv" MyUpload
::
:: Output (next to the CSV, unless a second argument names the base):
::   <name>_B2Upload.bat           <- run this to upload
::   <name>_B2Upload_winscp.txt    <- the WinSCP script it calls
::   <name>_B2Upload_folders.txt   <- the local/remote folder pairs, for review
::
:: Credentials:
:: - Kept in b2-credentials.bat, next to THIS generator, never inside the
::   generated files. See b2-credentials.bat.example for setup (a WinSCP saved
::   site is the recommended, no-secrets-on-disk-here option).
:: - b2-credentials.bat is listed in .gitignore. Never commit your filled-in
::   copy.
:: - The WinSCP script (.txt) file never contains an "open" line - WinSCP does
::   NOT substitute %VAR% environment references inside a script file, so
::   that would send WinSCP a literal, unresolved "%VAR%" as a hostname.
::   Instead, the runner .bat passes the resolved site name / session URL as a
::   plain argument on the winscp.com command line, which cmd.exe (not
::   WinSCP) expands. That argument briefly appears in the winscp.com
::   process's own command line while it runs - visible to Task Manager /
::   Process Explorer / command-line auditing on this machine, but never
::   written to any file. This only matters for Option B (inline key); a
::   WinSCP saved site name is not a secret at all.
::
:: Known limitations:
:: - This is plain batch CSV parsing (no quoted-field support). If a folder
::   path ever contains a literal comma, its row will be misread. Windows
::   paths essentially never contain commas, but if BRC-StaleDirectoryFinder
::   reports something odd, check <name>_B2Upload_folders.txt before running
::   the upload.
:: - If the CSV lists both a stale parent folder and a stale child folder
::   inside it, both are uploaded independently, so the child's files are
::   transferred twice. This wastes time and bandwidth but is not harmful
::   (WinSCP synchronize simply re-verifies files that already match).
:: ============================================================================

set "GEN_DIR=%~dp0"
set "CSV_PATH=%~1"
set "OUT_BASE=%~2"

if "%CSV_PATH%"=="" (
    set /p "CSV_PATH=Path to the BRC-StaleDirectoryFinder.ps1 CSV: "
)
set CSV_PATH=%CSV_PATH:"=%

if "%CSV_PATH%"=="" (
    echo ERROR: no CSV path given.
    exit /b 1
)

if not exist "%CSV_PATH%" (
    echo ERROR: CSV not found: %CSV_PATH%
    exit /b 1
)

if "%OUT_BASE%"=="" (
    for %%F in ("%CSV_PATH%") do set "OUT_BASE=%%~nF"
)

set "OUT_DIR="
for %%F in ("%CSV_PATH%") do set "OUT_DIR=%%~dpF"

set "OUT_BAT=%OUT_DIR%%OUT_BASE%_B2Upload.bat"
set "OUT_WINSCP=%OUT_DIR%%OUT_BASE%_B2Upload_winscp.txt"
set "OUT_LIST=%OUT_DIR%%OUT_BASE%_B2Upload_folders.txt"
set "OUT_LOG=%OUT_DIR%%OUT_BASE%_B2Upload.log"

set "CREDS_EXAMPLE=%GEN_DIR%b2-credentials.bat.example"
set "CREDS_FILE=%GEN_DIR%b2-credentials.bat"

:: --- Make sure a credentials template exists, and stop if the real file isn't set up yet ---
if not exist "%CREDS_EXAMPLE%" (
    echo ERROR: %CREDS_EXAMPLE% is missing from this generator's folder.
    echo Restore it from source control, or see the repository README.
    exit /b 1
)

if not exist "%CREDS_FILE%" (
    echo.
    echo No b2-credentials.bat found next to this generator:
    echo   %GEN_DIR%
    echo.
    echo 1. Copy b2-credentials.bat.example to b2-credentials.bat in that folder.
    echo 2. Fill in either Option A ^(recommended: a WinSCP saved site^) or
    echo    Option B ^(inline Application Key^) - read the comments in the file.
    echo 3. Re-run this generator.
    echo.
    echo b2-credentials.bat holds access to your B2 bucket. Do not commit it,
    echo email it, or store it anywhere outside this machine's local disk.
    exit /b 1
)

:: --- Load the non-secret settings from the credentials file (bucket name,
::     prefix, filters). We only read these into THIS process; nothing from
::     b2-credentials.bat is ever written into a generated file. ---
call "%CREDS_FILE%"

if not defined B2_BUCKET (
    echo ERROR: B2_BUCKET is not set in %CREDS_FILE%.
    exit /b 1
)

if not defined WINSCP_SITE if not defined B2_KEY_ID (
    echo ERROR: neither WINSCP_SITE ^(Option A^) nor B2_KEY_ID ^(Option B^) is set
    echo in %CREDS_FILE%. Fill in one of them - see the comments in that file.
    exit /b 1
)

if not defined ONLY_STALE set "ONLY_STALE=1"
if not defined MIRROR_DELETE set "MIRROR_DELETE=0"

echo Reading %CSV_PATH% ...

:: --- Filter the CSV down to the folders to upload ---
set "COUNT=0"
set "SKIPPED_NOTSTALE=0"
set "SKIPPED_MISSING=0"
set "SKIPPED_BLANK=0"

if exist "%OUT_LIST%" del "%OUT_LIST%"

> "%OUT_LIST%" (
    for /f "usebackq skip=1 eol=| tokens=1,12 delims=," %%A in ("%CSV_PATH%") do (
        call :ProcessRow "%%~A" "%%~B"
    )
)

echo.
echo Folders queued for upload : %COUNT%
echo Skipped, not stale        : %SKIPPED_NOTSTALE%
echo Skipped, path missing now : %SKIPPED_MISSING%
echo Skipped, blank path       : %SKIPPED_BLANK%
echo.

if "%COUNT%"=="0" (
    echo Nothing to upload. No upload scripts were generated.
    del "%OUT_LIST%" >nul 2>nul
    exit /b 1
)

:: --- Work out how the runner .bat should tell WinSCP which session to open.
::     IMPORTANT: WinSCP does NOT substitute %VAR% environment references
::     inside a script (.txt) file - a line like "open %WINSCP_SITE%" in the
::     script is sent to WinSCP literally and it will try to connect to a
::     host actually named "%WINSCP_SITE%". Instead, the session is passed as
::     a plain argument on the winscp.com command line in the runner .bat, so
::     it is cmd.exe (not WinSCP) that expands %%PLACEHOLDER%% - and cmd.exe
::     genuinely does expand it, at the moment the runner .bat runs. That
::     keeps the actual secret value out of every generated file; it only
::     ever exists as a resolved environment variable, briefly, in the
::     winscp.com process's own command line. ---
set "SESSION_ARG="
if defined WINSCP_SITE (
    set "SESSION_ARG=%%WINSCP_SITE%%"
) else (
    set "SESSION_ARG=s3://%%B2_KEY_ID%%:%%B2_APP_KEY%%@%%B2_ENDPOINT%%/"
)

set "SYNC_FLAGS=remote"
if "%MIRROR_DELETE%"=="1" set "SYNC_FLAGS=remote -delete"

echo Writing %OUT_WINSCP% ...

> "%OUT_WINSCP%" (
    echo # Generated by BRC-B2UploadScriptGenerator.bat on %DATE% %TIME%
    echo # Source CSV: %CSV_PATH%
    echo # Credentials come from environment variables set by %OUT_BASE%_B2Upload.bat
    echo # at run time - no secret is stored in this file.
    echo option batch on
    echo option confirm off
    echo option transfer binary
    for /f "usebackq tokens=1,2 delims=|" %%L in ("%OUT_LIST%") do (
        echo synchronize %SYNC_FLAGS% "%%~L" "%%~M"
    )
    echo close
    echo exit
)

echo Writing %OUT_BAT% ...

> "%OUT_BAT%" (
    echo @echo off
    echo :: Generated by BRC-B2UploadScriptGenerator.bat on %DATE% %TIME%
    echo :: Source CSV: %CSV_PATH%
    echo :: Uploads %COUNT% stale folder^(s^) to Backblaze B2 bucket "%B2_BUCKET%".
    echo :: This only adds/updates files remotely. MIRROR_DELETE was %MIRROR_DELETE%
    echo :: at generation time ^(1 = also deletes remote files no longer present
    echo :: locally - re-run the generator after changing it^).
    echo.
    echo set "CREDS=%CREDS_FILE%"
    echo.
    echo if not exist "%%CREDS%%" ^(
    echo     echo ERROR: credentials file not found: %%CREDS%%
    echo     exit /b 1
    echo ^)
    echo.
    echo call "%%CREDS%%"
    echo.
    echo if not defined WINSCP_SITE if not defined B2_KEY_ID ^(
    echo     echo ERROR: neither WINSCP_SITE nor B2_KEY_ID is set in %%CREDS%%.
    echo     exit /b 1
    echo ^)
    echo.
    echo set "WINSCP=%%WINSCP_EXE%%"
    echo if "%%WINSCP%%"=="" set "WINSCP=winscp.com"
    echo.
    echo where "%%WINSCP%%" ^>nul 2^>nul
    echo if errorlevel 1 if not exist "%%WINSCP%%" ^(
    echo     echo ERROR: WinSCP command-line executable not found: %%WINSCP%%
    echo     echo Install WinSCP, or set WINSCP_EXE in %%CREDS%% to its full path.
    echo     exit /b 1
    echo ^)
    echo.
    echo echo Uploading %COUNT% stale folder^(s^) to Backblaze B2, bucket "%%B2_BUCKET%%"...
    echo echo Log: %OUT_LOG%
    echo.
    echo "%%WINSCP%%" "!SESSION_ARG!" /ini=nul /log="%OUT_LOG%" /script="%OUT_WINSCP%"
    echo set "RC=%%ERRORLEVEL%%"
    echo.
    echo if "%%RC%%"=="0" ^(
    echo     echo Done. All transfers reported success. See the log for details:
    echo     echo   %OUT_LOG%
    echo ^) else ^(
    echo     echo WinSCP exited with code %%RC%%. Check the log for what failed:
    echo     echo   %OUT_LOG%
    echo ^)
    echo.
    echo exit /b %%RC%%
)

echo.
echo Done. Review before running:
echo   %OUT_LIST%    - the folders that will be uploaded, and where
echo   %OUT_WINSCP%  - the WinSCP script that will run
echo   %OUT_BAT%     - run this one to actually upload
echo.
echo Nothing has been uploaded yet. Recommended: run it once against a test
echo prefix or a single folder first, and check the bucket before doing a full run.

exit /b 0

:: ============================================================================
:: Subroutines
:: ============================================================================

:ProcessRow
set "ROWPATH=%~1"
set "ROWSTALE=%~2"

if "%ROWPATH%"=="" (
    set /a SKIPPED_BLANK+=1
    exit /b
)

if /i "%ONLY_STALE%"=="1" if /i not "%ROWSTALE%"=="True" (
    set /a SKIPPED_NOTSTALE+=1
    exit /b
)

if not exist "%ROWPATH%\" (
    set /a SKIPPED_MISSING+=1
    exit /b
)

call :BuildRemotePath "%ROWPATH%" REMOTEPATH

set /a COUNT+=1
echo "%ROWPATH%"^|"%REMOTEPATH%"
exit /b

:BuildRemotePath
setlocal
set "P=%~1"

if "%P:~0,2%"=="\\" (
    set "P=%P:~2%"
) else if "%P:~1,1%"==":" (
    set "DRIVE=%P:~0,1%"
    set "REST=%P:~2%"
    set "P=%DRIVE%%REST%"
)

set "P=%P:\=/%"

set "REMOTE=/%B2_BUCKET%"
if defined B2_REMOTE_PREFIX set "REMOTE=%REMOTE%/%B2_REMOTE_PREFIX%"
set "REMOTE=%REMOTE%/%P%"

endlocal & set "%~2=%REMOTE%"
exit /b
