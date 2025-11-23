Param(
    [string]$MonitorFolder  # optional override; if not set, use config
)

Write-Host "Starting Scribe..."
Write-Host ""

# ------------------------------
# 1) Load config (config.json or config.example.json)
# ------------------------------
$scriptRoot  = $PSScriptRoot
$configPath  = Join-Path $scriptRoot "config.json"
$examplePath = Join-Path $scriptRoot "config.example.json"

if (Test-Path $configPath) {
    $config = Get-Content -Path $configPath | ConvertFrom-Json
    Write-Host "Loaded config: config.json"
}
elseif (Test-Path $examplePath) {
    $config = Get-Content -Path $examplePath | ConvertFrom-Json
    Write-Host "WARNING: config.json not found, using config.example.json"
    Write-Host "Edit config.example.json and save as config.json for your own setup."
}
else {
    Write-Host "ERROR: No config.json or config.example.json found next to the script."
    exit 1
}

# MonitorFolder: param overrides config, config overrides nothing
if (-not $MonitorFolder) {
    $MonitorFolder = $config.MonitorFolder
}

if (-not $MonitorFolder) {
    Write-Host "ERROR: MonitorFolder is not set in param or config."
    exit 1
}

if (-not (Test-Path $MonitorFolder -PathType Container)) {
    Write-Host "ERROR: MonitorFolder path does not exist: $MonitorFolder"
    exit 1
}

Write-Host "Watching folder: $MonitorFolder"
Write-Host "Press Ctrl+C to stop."
Write-Host ""

# ------------------------------
# 2) Load rules from config
# ------------------------------
$global:MonitorFolder = $MonitorFolder
$global:SiteRules     = @{}

if ($config.Rules) {
    $config.Rules.PSObject.Properties | ForEach-Object {
        $global:SiteRules[$_.Name] = $_.Value
    }
}

if ($global:SiteRules.Count -eq 0) {
    Write-Host "WARNING: No rules found in config. Files will not be sorted by course."
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
    Write-Host "WARNING: Unknown DuplicatePolicy '$($global:DuplicatePolicy)'. Defaulting to 'rename'."
    $global:DuplicatePolicy = "rename"
}

# ------------------------------
# 4) Smart filename-based categories (within each course folder)
# ------------------------------
$global:SmartCategories = @(
    @{ Name = "Solutions";   Patterns = @("solution", "solutions", "soln", "sols") }
    @{ Name = "Assignments"; Patterns = @("assignment", "homework", "hw", "problem set", "pset") }
    @{ Name = "Labs";        Patterns = @("lab", "laboratory") }
    @{ Name = "Worksheets";  Patterns = @("worksheet", "worksheets", "practice", "exercise") }
    @{ Name = "Lectures";    Patterns = @("lecture", "lectures", "slides") }
    @{ Name = "Tutorials";   Patterns = @("tutorial", "tutorials", "tut") }
    @{ Name = "Quizzes";     Patterns = @("quiz", "quizzes") }
    @{ Name = "Exams";       Patterns = @("midterm", "mid-term", "exam", "final") }
)

# ------------------------------
# 5) Clean old events on rerun
# ------------------------------
Unregister-Event -SourceIdentifier "DL_Created" -ErrorAction SilentlyContinue
Unregister-Event -SourceIdentifier "DL_Renamed" -ErrorAction SilentlyContinue

# ------------------------------
# 6) Event action
# ------------------------------
$action = {
    $path       = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    $fileName   = [System.IO.Path]::GetFileName($path)

    # Skip temp download files
    if ($fileName -match '\.(crdownload|tmp|opdownload)$') { return }

    # Skip folders
    if (Test-Path -LiteralPath $path -PathType Container) { return }

    # Only watch root of MonitorFolder
    $dir = [System.IO.Path]::GetDirectoryName($path)
    if ($dir -ne $global:MonitorFolder) { return }

    Start-Sleep -Milliseconds 200

    Write-Host ""
    Write-Host "[$changeType] $fileName"

    # Extract URL (if any)
    $url = $null
    try {
        $zone = Get-Content -Path $path -Stream Zone.Identifier -ErrorAction Stop
        $meta = $zone | Where-Object { $_ -match '^HostUrl=' -or $_ -match '^ReferrerUrl=' } | Select-Object -First 1
        if ($meta) { $url = $meta.Split('=')[1].Trim() }
    } catch {}

    if ($url) {
        Write-Host "  URL: $url"
    } else {
        Write-Host "  URL: <none>"
        return
    }

    if (-not $global:SiteRules -or $global:SiteRules.Count -eq 0) {
        Write-Host "  No rules configured."
        return
    }

    # Match course rule
    $matchedKey = $null
    foreach ($key in $global:SiteRules.Keys) {
        if ($url -like $key) {
            $matchedKey = $key
            break
        }
    }

    if (-not $matchedKey) {
        Write-Host "  No course rule matched."
        return
    }

    $baseRel = $global:SiteRules[$matchedKey]  # e.g. "University/MTE-252"

    # --------------------------
    # Smart filename categorization
    # --------------------------
    $lowerName = $fileName.ToLowerInvariant()
    $category  = $null

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
    } else {
        $targetRel = $baseRel                       # no smart match, just course root
    }

    $targetDir = Join-Path $global:MonitorFolder $targetRel

    Write-Host "  Matched course: $baseRel"
    if ($category) {
        Write-Host "  Category: $category"
    }

    # Create folder if missing
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        Write-Host "  Created folder: $targetRel"
    }

    # --------------------------
    # Duplicate handling
    # --------------------------
    $targetPath = Join-Path $targetDir $fileName
    $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $ext  = [System.IO.Path]::GetExtension($fileName)

    if (Test-Path $targetPath) {
        try {
            $existingHash = Get-FileHash -Path $targetPath -Algorithm SHA256
            $newHash      = Get-FileHash -Path $path       -Algorithm SHA256

            if ($existingHash.Hash -eq $newHash.Hash) {
                Write-Host "  Duplicate content detected (same name, same file)."

                switch ($global:DuplicatePolicy) {
                    "overwrite" {
                        Write-Host "  DuplicatePolicy=overwrite -> replacing existing file."
                        Move-Item -LiteralPath $path -Destination $targetPath -Force
                        return
                    }
                    "skip" {
                        Write-Host "  DuplicatePolicy=skip -> keeping existing, ignoring new file."
                        return
                    }
                    default { # rename
                        Write-Host "  DuplicatePolicy=rename -> identical content, ignoring new file."
                        return
                    }
                }
            }
            else {
                Write-Host "  Different content with same name detected."

                switch ($global:DuplicatePolicy) {
                    "overwrite" {
                        Write-Host "  DuplicatePolicy=overwrite -> replacing existing file."
                        Move-Item -LiteralPath $path -Destination $targetPath -Force
                        return
                    }
                    "skip" {
                        Write-Host "  DuplicatePolicy=skip -> keeping existing, ignoring new version."
                        return
                    }
                    default { # rename
                        Write-Host "  DuplicatePolicy=rename -> saving as new version."
                    }
                }
            }
        }
        catch {
            Write-Host "  Warning: couldn't hash files, falling back to rename behavior."
        }

        # Fallback or rename mode: create " (1)", " (2)", etc.
        $i = 1
        while (Test-Path $targetPath) {
            $targetPath = Join-Path $targetDir ("{0} ({1}){2}" -f $base, $i, $ext)
            $i++
        }
    }

    Move-Item -LiteralPath $path -Destination $targetPath
    Write-Host "  Moved to: $targetRel"
}

# ------------------------------
# 7) Watcher registration
# ------------------------------
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $MonitorFolder
$watcher.Filter = "*.*"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

Register-ObjectEvent $watcher Created -SourceIdentifier "DL_Created" -Action $action | Out-Null
Register-ObjectEvent $watcher Renamed -SourceIdentifier "DL_Renamed" -Action $action | Out-Null

# ------------------------------
# 8) Keep alive
# ------------------------------
while ($true) { Start-Sleep -Seconds 1 }
