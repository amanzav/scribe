Param(
    [string]$MonitorFolder,   # optional override; if not set, use config
    [switch]$Batch,           # batch mode: process existing files once
    [switch]$DryRun           # don't move anything, just log what would happen
)

# ------------------------------
# Header
# ------------------------------
Write-Host "Starting Scribe..." -ForegroundColor Cyan
if ($Batch) {
    if ($DryRun) {
        Write-Host "Mode: BATCH (dry run)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Mode: BATCH" -ForegroundColor Yellow
    }
}
else {
    if ($DryRun) {
        Write-Host "Mode: WATCHER (dry run)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Mode: WATCHER" -ForegroundColor Yellow
    }
}
Write-Host ""

# ------------------------------
# 1) Load config (config.json or config.example.json)
# ------------------------------
$scriptRoot = $PSScriptRoot
$configPath = Join-Path $scriptRoot "config.json"
$examplePath = Join-Path $scriptRoot "config.example.json"

if (Test-Path $configPath) {
    $config = Get-Content -Path $configPath | ConvertFrom-Json
    Write-Host "Loaded config: config.json" -ForegroundColor Green
}
elseif (Test-Path $examplePath) {
    $config = Get-Content -Path $examplePath | ConvertFrom-Json
    Write-Host "WARNING: config.json not found, using config.example.json" -ForegroundColor Yellow
    Write-Host "Edit config.example.json and save as config.json for your own setup." -ForegroundColor Yellow
}
else {
    Write-Host "ERROR: No config.json or config.example.json found next to the script." -ForegroundColor Red
    exit 1
}

# MonitorFolder: param overrides config, config overrides nothing
if (-not $MonitorFolder) {
    $MonitorFolder = $config.MonitorFolder
}

if (-not $MonitorFolder) {
    Write-Host "ERROR: MonitorFolder is not set in param or config." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $MonitorFolder -PathType Container)) {
    Write-Host "ERROR: MonitorFolder path does not exist: $MonitorFolder" -ForegroundColor Red
    exit 1
}

# Normalize MonitorFolder to a full canonical path
$MonitorFolder = (Resolve-Path $MonitorFolder).ProviderPath
Write-Host "Watching folder: $MonitorFolder"
Write-Host ""

# ------------------------------
# 2) Load rules from config
# ------------------------------
$global:MonitorFolder = $MonitorFolder
$global:SiteRules = @{}

if ($config.Rules) {
    $config.Rules.PSObject.Properties | ForEach-Object {
        $global:SiteRules[$_.Name] = $_.Value
    }
}

if ($global:SiteRules.Count -eq 0) {
    Write-Host "WARNING: No rules found in config. Files will not be sorted by course." -ForegroundColor Yellow
}

# ------------------------------
# 3) Duplicate policy
# ------------------------------
$global:DuplicatePolicy = $config.DuplicatePolicy
if (-not $global:DuplicatePolicy) {
    $global:DuplicatePolicy = "rename"
}
$global:DuplicatePolicy = $global:DuplicatePolicy.ToLowerInvariant()
if ($global:DuplicatePolicy -notin @("rename", "skip", "overwrite")) {
    Write-Host "WARNING: Unknown DuplicatePolicy '$($global:DuplicatePolicy)'. Defaulting to 'rename'." -ForegroundColor Yellow
    $global:DuplicatePolicy = "rename"
}

# Dry-run flag
$global:DryRun = [bool]$DryRun

# ------------------------------
# 4) Smart filename-based categories (within each course folder)
# ------------------------------
$global:SmartCategories = @(
    @{ Name = "Solutions"; Patterns = @("solution", "solutions", "soln", "sols") }
    @{ Name = "Assignments"; Patterns = @("assignment", "homework", "hw", "problem set", "pset") }
    @{ Name = "Labs"; Patterns = @("lab", "laboratory") }
    @{ Name = "Worksheets"; Patterns = @("worksheet", "worksheets", "practice", "exercise") }
    @{ Name = "Lectures"; Patterns = @("lecture", "lectures", "slides", "lec", "week") }
    @{ Name = "Tutorials"; Patterns = @("tutorial", "tutorials", "tut") }
    @{ Name = "Quizzes"; Patterns = @("quiz", "quizzes") }
    @{ Name = "Exams"; Patterns = @("midterm", "mid-term", "exam", "final") }
)

# ------------------------------
# 5) Helper: resolve course by URL or filename
#     Returns a PSCustomObject with:
#       BaseRel    = "University/MTE-252"
#       Source     = "Url" or "Filename"
#       CourseCode = "MTE-252" or $null
# ------------------------------
function Resolve-CourseFolder {
    param(
        [string]$Url,
        [string]$FileName
    )

    # 1) Try URL-based rule match
    if ($Url -and $global:SiteRules.Count -gt 0) {
        foreach ($key in $global:SiteRules.Keys) {
            if ($Url -like $key) {
                return [PSCustomObject]@{
                    BaseRel    = $global:SiteRules[$key]
                    Source     = "Url"
                    CourseCode = $null
                }
            }
        }
    }

    # 2) Fallback: try filename-based course detection if URL failed
    if ($FileName) {
        $lower = $FileName.ToLowerInvariant()

        # Look for MTE220, MTE-220, MTE 220 pattern
        if ($lower -match 'mte[\s\-]?(\d{3})') {
            $num = $matches[1]
            $courseCode = "MTE-$num"  # normalized
            # Try to match any rule value that contains this course code (e.g. "University/MTE-220")
            foreach ($val in $global:SiteRules.Values) {
                if ($val -like ("*" + $courseCode + "*")) {
                    return [PSCustomObject]@{
                        BaseRel    = $val
                        Source     = "Filename"
                        CourseCode = $courseCode
                    }
                }
            }
        }
    }

    return $null
}

# ------------------------------
# 6) Core processor (used by watcher & batch)
# ------------------------------
function Process-ScribeFile {
    param(
        [string]$Path,
        [string]$ChangeType
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    $fileName = [System.IO.Path]::GetFileName($Path)

    # Skip temp download files
    if ($fileName -match '\.(crdownload|tmp|opdownload)$') { return }

    # Skip folders
    if (Test-Path -LiteralPath $Path -PathType Container) { return }

    # Only watch root of MonitorFolder
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    $dirFull = [System.IO.Path]::GetFullPath($dir)
    if ($dirFull -ne $global:MonitorFolder) { return }

    Start-Sleep -Milliseconds 200

    # Extract URL (if any)
    $url = $null
    try {
        $zone = Get-Content -Path $Path -Stream Zone.Identifier -ErrorAction Stop
        $meta = $zone | Where-Object { $_ -match '^HostUrl=' -or $_ -match '^ReferrerUrl=' } | Select-Object -First 1
        if ($meta) { $url = $meta.Split('=')[1].Trim() }
    }
    catch {}

    # No rules at all? Just silently ignore.
    if (-not $global:SiteRules -or $global:SiteRules.Count -eq 0) {
        return
    }

    # Resolve course folder (by URL first, then filename)
    $courseInfo = Resolve-CourseFolder -Url $url -FileName $fileName
    if (-not $courseInfo) {
        # No course matched: skip silently
        return
    }

    $baseRel = $courseInfo.BaseRel
    $courseSrc = $courseInfo.Source
    $courseCode = $courseInfo.CourseCode

    # Smart category
    $lowerName = $fileName.ToLowerInvariant()
    $category = $null

    foreach ($cat in $global:SmartCategories) {
        foreach ($p in $cat.Patterns) {
            if ($lowerName -like ("*" + $p + "*")) {
                $category = $cat.Name
                break
            }
        }
        if ($category) { break }
    }

    if ($category) {
        $targetRel = Join-Path $baseRel $category   # e.g. "University/MTE-252/Solutions"
    }
    else {
        $targetRel = $baseRel                       # no smart match, just course root
    }

    $targetDir = Join-Path $global:MonitorFolder $targetRel

    # Build "reason" text for dry-run summary
    $reasonParts = @()
    if ($courseSrc -eq "Url") {
        $reasonParts += "course via URL"
    }
    elseif ($courseSrc -eq "Filename") {
        if ($courseCode) {
            $reasonParts += "course via filename $courseCode"
        }
        else {
            $reasonParts += "course via filename"
        }
    }
    if ($category) {
        $reasonParts += "category $category"
    }
    $reasonText = ""
    if ($reasonParts.Count -gt 0) {
        $reasonText = " (" + ($reasonParts -join "; ") + ")"
    }

    # In WATCH mode (non-dry), show detailed info
    if (-not $global:DryRun) {
        Write-Host ""
        Write-Host "[$ChangeType] $fileName" -ForegroundColor White
        Write-Host "  Course folder: $baseRel"
        if ($courseSrc -eq "Url") {
            Write-Host "  Course match: URL"
        }
        elseif ($courseSrc -eq "Filename") {
            if ($courseCode) {
                Write-Host "  Course match: filename ($courseCode)"
            }
            else {
                Write-Host "  Course match: filename"
            }
        }
        if ($category) {
            Write-Host "  Category: $category"
        }
    }

    # --------------------------
    # Duplicate handling with Batch "(1)" cleanup
    # --------------------------
    $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $ext = [System.IO.Path]::GetExtension($fileName)

    # In BATCH mode, normalize "Name (1).pdf" -> "Name.pdf" as canonical
    $canonicalBase = $base
    if ($ChangeType -eq "Batch" -and $base -match '^(.*)\s+\((\d+)\)$') {
        $canonicalBase = $matches[1]
    }
    $canonicalFileName = $canonicalBase + $ext

    $targetPath = Join-Path $targetDir $canonicalFileName

    if (Test-Path $targetPath) {
        try {
            $existingHash = Get-FileHash -Path $targetPath -Algorithm SHA256
            $newHash = Get-FileHash -Path $Path       -Algorithm SHA256

            if ($existingHash.Hash -eq $newHash.Hash) {
                # Same content: skip silently (no log, as requested)
                if (-not $global:DryRun -and $global:DuplicatePolicy -eq "overwrite") {
                    # If someone explicitly set overwrite, respect it quietly
                    Move-Item -LiteralPath $Path -Destination $targetPath -Force
                    Write-Host "  Replaced identical file at: $targetRel" -ForegroundColor Green
                }
                return
            }
            else {
                # Different content, same canonical name
                switch ($global:DuplicatePolicy) {
                    "overwrite" {
                        if ($global:DryRun) {
                            $relative = $targetPath.Replace($global:MonitorFolder + '\', '')
                            Write-Host "MOVE: $fileName -> $relative$reasonText" -ForegroundColor Green
                        }
                        else {
                            Move-Item -LiteralPath $Path -Destination $targetPath -Force
                            Write-Host "  Moved (overwrite) to: $targetRel" -ForegroundColor Green
                        }
                        return
                    }
                    "skip" {
                        # Skip silently
                        return
                    }
                    default {
                        # rename mode -> fall through to versioned name logic
                    }
                }
            }
        }
        catch {
            if (-not $global:DryRun) {
                Write-Host "  Warning: couldn't hash files, falling back to rename behavior." -ForegroundColor Yellow
            }
        }

        # Fallback or rename mode: create "Name (1).pdf", "Name (2).pdf", etc.
        $i = 1
        while (Test-Path $targetPath) {
            $targetPath = Join-Path $targetDir ("{0} ({1}){2}" -f $canonicalBase, $i, $ext)
            $i++
        }
    }

    # --------------------------
    # Final MOVE / DRY RUN output
    # --------------------------
    if ($global:DryRun) {
        # Show clean move line with path relative to MonitorFolder
        $relative = $targetPath.Replace($global:MonitorFolder + '\', '')
        Write-Host "MOVE: $fileName -> $relative$reasonText" -ForegroundColor Green
    }
    else {
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        Move-Item -LiteralPath $Path -Destination $targetPath
        Write-Host "  Moved to: $targetRel" -ForegroundColor Green
    }
}

# ------------------------------
# 7) Batch mode vs watcher mode
# ------------------------------
if ($Batch) {
    # Batch mode: process all existing files in MonitorFolder (top level)
    $files = Get-ChildItem -Path $MonitorFolder -File

    if (-not $files -or $files.Count -eq 0) {
        Write-Host "No files found in $MonitorFolder to process." -ForegroundColor DarkGray
        exit 0
    }

    foreach ($f in $files) {
        Process-ScribeFile -Path $f.FullName -ChangeType "Batch"
    }

    Write-Host ""
    Write-Host "Batch processing complete." -ForegroundColor Cyan
    exit 0
}

# ------------------------------
# 8) Watcher mode
# ------------------------------
Unregister-Event -SourceIdentifier "DL_Created" -ErrorAction SilentlyContinue
Unregister-Event -SourceIdentifier "DL_Renamed" -ErrorAction SilentlyContinue

$action = {
    $path = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    Process-ScribeFile -Path $path -ChangeType $changeType
}

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $MonitorFolder
$watcher.Filter = "*.*"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

Register-ObjectEvent $watcher Created -SourceIdentifier "DL_Created" -Action $action | Out-Null
Register-ObjectEvent $watcher Renamed -SourceIdentifier "DL_Renamed" -Action $action | Out-Null

Write-Host "Watcher running. Press Ctrl+C to stop." -ForegroundColor Cyan

while ($true) { Start-Sleep -Seconds 1 }
