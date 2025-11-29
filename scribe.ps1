Param(
    [string]$MonitorFolder,   # optional override; if not set, use config
    [switch]$Batch,           # kept for compatibility, but script always runs in batch mode
    [switch]$DryRun,          # don't move anything, just log what would happen
    [switch]$All              # process ALL files, ignoring LastRunUtc
)

# ------------------------------
# 0) Setup logging
# ------------------------------
$scriptRoot = $PSScriptRoot
$global:LogPath = Join-Path $scriptRoot "scribe.log"
$statePath = Join-Path $scriptRoot "scribe_state.json"

function Write-Log {
    param(
        [string]$Message,
        [System.ConsoleColor]$Color
    )

    if ($Message -eq "") {
        Write-Host ""
        Add-Content -Path $global:LogPath -Value ""
        return
    }

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$timestamp] $Message"

    if ($PSBoundParameters.ContainsKey('Color')) {
        Write-Host $line -ForegroundColor $Color
    }
    else {
        Write-Host $line
    }

    Add-Content -Path $global:LogPath -Value $line
}

function Log-Info { param([string]$m) Write-Log $m }
function Log-Success { param([string]$m) Write-Log $m ([System.ConsoleColor]::Green) }
function Log-Warn { param([string]$m) Write-Log $m ([System.ConsoleColor]::Yellow) }
function Log-Error { param([string]$m) Write-Log $m ([System.ConsoleColor]::Red) }
function Log-Dim { param([string]$m) Write-Log $m ([System.ConsoleColor]::DarkGray) }

Log-Info "Starting Scribe..."

# ------------------------------
# 1) Mode header
# ------------------------------
$mode =
if ($All -and $DryRun) { "BATCH-ALL (dry run)" }
elseif ($All) { "BATCH-ALL" }
elseif ($DryRun) { "BATCH (dry run)" }
else { "BATCH" }

Log-Info  "Mode: $mode"
Log-Info  "Log file: $global:LogPath"
Log-Info  ""

# ------------------------------
# 2) Load config (config.json or config.example.json)
# ------------------------------
$configPath = Join-Path $scriptRoot "config.json"
$examplePath = Join-Path $scriptRoot "config.example.json"
$config = $null

if (Test-Path $configPath) {
    $config = Get-Content -Path $configPath | ConvertFrom-Json
    Log-Success "Loaded config: config.json"
}
elseif (Test-Path $examplePath) {
    $config = Get-Content -Path $examplePath | ConvertFrom-Json
    Log-Warn "config.json not found, using config.example.json"
    Log-Warn "Edit config.example.json and save as config.json for your own setup."
}
else {
    Log-Error "No config.json or config.example.json found next to the script."
    exit 1
}

# MonitorFolder: param overrides config
if (-not $MonitorFolder) {
    $MonitorFolder = $config.MonitorFolder
}
if (-not $MonitorFolder) {
    Log-Error "MonitorFolder is not set in param or config."
    exit 1
}
if (-not (Test-Path $MonitorFolder -PathType Container)) {
    Log-Error "MonitorFolder path does not exist: $MonitorFolder"
    exit 1
}

$MonitorFolder = (Resolve-Path $MonitorFolder).ProviderPath
$global:MonitorFolder = $MonitorFolder

Log-Info "Downloads folder: $MonitorFolder"
Log-Info ""

# ------------------------------
# 3) Load rules from config
# ------------------------------
$global:SiteRules = @{}
if ($config.Rules) {
    $config.Rules.PSObject.Properties | ForEach-Object {
        $global:SiteRules[$_.Name] = $_.Value
    }
}

if ($global:SiteRules.Count -eq 0) {
    Log-Warn "No rules found in config. Files will not be sorted by course."
}

# ------------------------------
# 4) Duplicate policy
# ------------------------------
$global:DuplicatePolicy = $config.DuplicatePolicy
if (-not $global:DuplicatePolicy) {
    $global:DuplicatePolicy = "rename"
}
$global:DuplicatePolicy = $global:DuplicatePolicy.ToLowerInvariant()

if ($global:DuplicatePolicy -notin @("rename", "skip", "overwrite")) {
    Log-Warn "Unknown DuplicatePolicy '$($global:DuplicatePolicy)'. Defaulting to 'rename'."
    $global:DuplicatePolicy = "rename"
}

$global:DryRun = [bool]$DryRun

# ------------------------------
# Image handling
# ------------------------------
$global:ImageExtensions = @(
    ".png", ".jpg", ".jpeg", ".gif",
    ".bmp", ".tif", ".tiff", ".webp", ".heic"
)

# ------------------------------
# 5) AI config (Groq for category classification)
# ------------------------------
$global:AIConfig = $null
$global:UseAIForCategory = $false
$global:AllowedCategories = @(
    "Solutions", "Assignments", "Labs", "Projects",
    "Worksheets", "Lectures", "Tutorials", "Quizzes", "Exams", "Misc"
)

if ($config.AI) {
    $global:AIConfig = $config.AI
    if ($global:AIConfig.UseCategory -and
        $global:AIConfig.Endpoint -and
        $global:AIConfig.Model -and
        $global:AIConfig.ApiKeyEnvVar) {

        $global:UseAIForCategory = $true
        Log-Success "AI category classification: ENABLED (provider: $($global:AIConfig.Provider), model: $($global:AIConfig.Model))"
    }
    else {
        Log-Warn "AI config present but incomplete; AI category classification DISABLED."
    }
}
else {
    Log-Warn "No AI config found; AI category classification DISABLED."
}

Log-Info ""

# ------------------------------
# 6) Smart filename-based categories (fallback)
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
# 7) Load state: LastRunUtc
# ------------------------------
$global:LastRunUtc = $null

if (-not $All -and (Test-Path $statePath)) {
    try {
        $stateObj = Get-Content -Path $statePath | ConvertFrom-Json
        if ($stateObj.LastRunUtc) {
            $global:LastRunUtc = [DateTime]::Parse($stateObj.LastRunUtc).ToUniversalTime()
            Log-Info "Loaded last run time from state: $($global:LastRunUtc.ToString('u'))"
        }
    }
    catch {
        Log-Warn "Failed to read state file; defaulting to start of today."
    }
}

if (-not $All -and -not $global:LastRunUtc) {
    $todayLocal = [DateTime]::Today
    $global:LastRunUtc = $todayLocal.ToUniversalTime()
    Log-Info "No previous state found. Using start of today as LastRunUtc: $todayLocal"
}

Log-Info ""

# ------------------------------
# 8) Helper: resolve course by URL or filename
# ------------------------------
function Resolve-CourseFolder {
    param(
        [string]$Url,
        [string]$FileName
    )

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

    if ($FileName) {
        $lower = $FileName.ToLowerInvariant()

        if ($lower -match 'mte[\s\-]?(\d{3})') {
            $num = $matches[1]
            $courseCode = "MTE-$num"

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
# 9) Helper: call Groq to get AI category
# ------------------------------
function Get-AICategory {
    param(
        [string]$FileName,
        [string]$Url,
        [string]$CourseCode
    )

    if (-not $global:UseAIForCategory) {
        return $null
    }

    $apiKeyEnv = $global:AIConfig.ApiKeyEnvVar
    $apiKey = [System.Environment]::GetEnvironmentVariable($apiKeyEnv, "User")
    if (-not $apiKey) {
        Log-Warn "AI: Missing API key in env var '$apiKeyEnv'. Skipping AI category."
        return $null
    }

    $endpoint = $global:AIConfig.Endpoint
    $model = $global:AIConfig.Model
    $categories = $global:AllowedCategories -join ", "

    $systemMsg = @"
You classify university course files into one of a fixed set of categories based on their filename (and optionally source URL or course code).
Return ONLY a compact JSON object with a single field "category".
Allowed categories are:
$categories

Choose the best category based on the filename. If it doesn't clearly fit, use "Misc".
"@

    $userMsg = @"
Filename: $FileName
URL: $Url
CourseCode: $CourseCode
"@

    $bodyObj = @{
        model           = $model
        messages        = @(
            @{ role = "system"; content = $systemMsg },
            @{ role = "user"; content = $userMsg }
        )
        temperature     = 0.2
        response_format = @{ type = "json_object" }
    }

    $jsonBody = $bodyObj | ConvertTo-Json -Depth 6

    try {
        $headers = @{
            "Authorization" = "Bearer $apiKey"
            "Content-Type"  = "application/json"
        }

        $response = Invoke-RestMethod -Method Post -Uri $endpoint -Headers $headers -Body $jsonBody
        $content = $response.choices[0].message.content
        $parsed = $content | ConvertFrom-Json

        $rawCat = $parsed.category
        if (-not $rawCat) {
            return $null
        }

        $catNorm = ($rawCat.Trim()).ToLowerInvariant()
        foreach ($allowed in $global:AllowedCategories) {
            if ($allowed.ToLowerInvariant() -eq $catNorm) {
                return $allowed
            }
        }

        return "Misc"
    }
    catch [System.Net.WebException] {
        $resp = $_.Exception.Response
        if ($resp -and $resp.GetResponseStream()) {
            $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
            $responseBody = $reader.ReadToEnd()

            # Truncate the body so it doesn't nuke your log
            $snippet = $responseBody
            if ($snippet.Length -gt 400) {
                $snippet = $snippet.Substring(0, 400) + " ..."
            }

            Write-Log "AI: HTTP error for '$FileName': $($_.Exception.Message) | Response body: $snippet" `
            ([System.ConsoleColor]::Yellow)
        }
        else {
            Write-Log "AI: HTTP error for '$FileName': $($_.Exception.Message)" `
            ([System.ConsoleColor]::Yellow)
        }
        return $null
    }

    catch {
        Log-Warn "AI: Error calling Groq for '$FileName': $($_.Exception.Message)"
        return $null
    }
}

# ------------------------------
# 10) Core processor (batch only)
# ------------------------------
function Process-ScribeFile {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    $fileName = [System.IO.Path]::GetFileName($Path)

    # Skip temp download files
    if ($fileName -match '\.(crdownload|tmp|opdownload)$') { return }
    # Skip folders
    if (Test-Path -LiteralPath $Path -PathType Container) { return }

    $dir = [System.IO.Path]::GetDirectoryName($Path)
    $dirFull = [System.IO.Path]::GetFullPath($dir)
    if ($dirFull -ne $global:MonitorFolder) { return }

    # --------------------------
    # Image special-case
    # --------------------------
    $ext = [System.IO.Path]::GetExtension($fileName)
    $extLower = $ext.ToLowerInvariant()
    $isImage = $global:ImageExtensions -contains $extLower

    if ($isImage) {
        # All images -> Downloads\Images
        $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)

        # Strip " (1)" style suffixes for canonical name
        $canonicalBase = $base
        if ($base -match '^(.*)\s+\((\d+)\)$') {
            $canonicalBase = $matches[1]
        }
        $canonicalFileName = $canonicalBase + $ext

        $targetRel = "Images"
        $targetDir = Join-Path $global:MonitorFolder $targetRel
        $targetPath = Join-Path $targetDir $canonicalFileName

        # Duplicate handling for images too
        if (Test-Path $targetPath) {
            try {
                $existingHash = Get-FileHash -Path $targetPath -Algorithm SHA256
                $newHash = Get-FileHash -Path $Path       -Algorithm SHA256

                if ($existingHash.Hash -eq $newHash.Hash) {
                    # Same content → skip silently
                    return
                }
                else {
                    switch ($global:DuplicatePolicy) {
                        "overwrite" {
                            if ($global:DryRun) {
                                $relative = $targetPath.Replace($global:MonitorFolder + '\', '')
                                Log-Info "MOVE: $fileName -> $relative (image; overwrite)"
                            }
                            else {
                                if (-not (Test-Path $targetDir)) {
                                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                                }
                                Move-Item -LiteralPath $Path -Destination $targetPath -Force
                                Log-Info "Moved: $fileName -> $targetRel (image; overwrite)"
                            }
                            return
                        }
                        "skip" {
                            return
                        }
                        default {
                            # fall through to rename logic
                        }
                    }
                }
            }
            catch {
                Log-Warn "Warning: couldn't hash image '$fileName', falling back to rename behavior."
            }

            # rename mode: Name (1).ext, Name (2).ext, ...
            $i = 1
            while (Test-Path $targetPath) {
                $targetPath = Join-Path $targetDir ("{0} ({1}){2}" -f $canonicalBase, $i, $ext)
                $i++
            }
        }

        if ($global:DryRun) {
            $relative = $targetPath.Replace($global:MonitorFolder + '\', '')
            Log-Info "MOVE: $fileName -> $relative (image)"
        }
        else {
            if (-not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }

            Move-Item -LiteralPath $Path -Destination $targetPath
            Log-Info "Moved: $fileName -> $targetRel (image)"
        }

        return
    }

    # --------------------------
    # NON-IMAGE FILES: existing behavior (AI + course rules)
    # --------------------------

    # URL from Zone.Identifier
    $url = $null
    try {
        $zone = Get-Content -Path $Path -Stream Zone.Identifier -ErrorAction Stop
        $meta = $zone | Where-Object { $_ -match '^HostUrl=' -or $_ -match '^ReferrerUrl=' } | Select-Object -First 1
        if ($meta) { $url = $meta.Split('=')[1].Trim() }
    }
    catch {}

    if (-not $global:SiteRules -or $global:SiteRules.Count -eq 0) {
        return
    }

    $courseInfo = Resolve-CourseFolder -Url $url -FileName $fileName
    if (-not $courseInfo) { return }

    $baseRel = $courseInfo.BaseRel
    $courseSrc = $courseInfo.Source
    $courseCode = $courseInfo.CourseCode

    # Category: AI first, fallback to regex
    $category = $null
    $categorySource = "none"

    $aiCat = Get-AICategory -FileName $fileName -Url $url -CourseCode $courseCode
    if ($aiCat) {
        if ($aiCat -ne "Misc") {
            $category = $aiCat
            $categorySource = "AI"
        }
        else {
            $categorySource = "AI-Misc"
        }
    }

    if (-not $category) {
        $lowerName = $fileName.ToLowerInvariant()
        foreach ($cat in $global:SmartCategories) {
            foreach ($p in $cat.Patterns) {
                if ($lowerName -like ("*" + $p + "*")) {
                    $category = $cat.Name
                    $categorySource = "rules"
                    break
                }
            }
            if ($category) { break }
        }
    }

    if ($category) {
        $targetRel = Join-Path $baseRel $category
    }
    else {
        $targetRel = $baseRel
    }

    $targetDir = Join-Path $global:MonitorFolder $targetRel

    # Reason text for log
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
        if ($categorySource -eq "AI") {
            $reasonParts += "category $category (AI)"
        }
        elseif ($categorySource -eq "rules") {
            $reasonParts += "category $category (rules)"
        }
        else {
            $reasonParts += "category $category"
        }
    }
    elseif ($categorySource -eq "AI-Misc") {
        $reasonParts += "AI category Misc -> course root"
    }

    $reasonText = ""
    if ($reasonParts.Count -gt 0) {
        $reasonText = " (" + ($reasonParts -join "; ") + ")"
    }

    # Duplicate handling with "(1)" cleanup
    $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)

    $canonicalBase = $base
    if ($base -match '^(.*)\s+\((\d+)\)$') {
        $canonicalBase = $matches[1]
    }
    $canonicalFileName = $canonicalBase + $ext

    $targetPath = Join-Path $targetDir $canonicalFileName

    if (Test-Path $targetPath) {
        try {
            $existingHash = Get-FileHash -Path $targetPath -Algorithm SHA256
            $newHash = Get-FileHash -Path $Path       -Algorithm SHA256

            if ($existingHash.Hash -eq $newHash.Hash) {
                return
            }
            else {
                switch ($global:DuplicatePolicy) {
                    "overwrite" {
                        if ($global:DryRun) {
                            $relative = $targetPath.Replace($global:MonitorFolder + '\', '')
                            Log-Info "MOVE: $fileName -> $relative$reasonText"
                        }
                        else {
                            if (-not (Test-Path $targetDir)) {
                                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                            }
                            Move-Item -LiteralPath $Path -Destination $targetPath -Force
                            Log-Info "Moved (overwrite): $fileName -> $targetRel$reasonText"
                        }
                        return
                    }
                    "skip" {
                        return
                    }
                    default {
                        # fall through to rename
                    }
                }
            }
        }
        catch {
            Log-Warn "Warning: couldn't hash files for '$fileName', falling back to rename behavior."
        }

        $i = 1
        while (Test-Path $targetPath) {
            $targetPath = Join-Path $targetDir ("{0} ({1}){2}" -f $canonicalBase, $i, $ext)
            $i++
        }
    }

    if ($global:DryRun) {
        $relative = $targetPath.Replace($global:MonitorFolder + '\', '')
        Log-Info "MOVE: $fileName -> $relative$reasonText"
    }
    else {
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        Move-Item -LiteralPath $Path -Destination $targetPath
        Log-Info "Moved: $fileName -> $targetRel$reasonText"
    }
}


# ------------------------------
# 11) Batch mode (new files vs ALL)
# ------------------------------
if (-not $All) {
    Log-Info "Scanning for new files since: $($global:LastRunUtc.ToString('u'))"
}
else {
    Log-Info "Scanning ALL files in root (ignoring LastRunUtc)."
}

$files = Get-ChildItem -Path $MonitorFolder -File

if (-not $files -or $files.Count -eq 0) {
    Log-Dim "No files found in $MonitorFolder to process."
    exit 0
}

if ($All) {
    $filesToProcess = $files
    Log-Info "Found $($filesToProcess.Count) files to process (ALL mode)."
}
else {
    $filesToProcess = $files | Where-Object { $_.LastWriteTimeUtc -gt $global:LastRunUtc }
    Log-Info "Found $($files.Count) files in root; $($filesToProcess.Count) new since last run."
}

Log-Info ""

foreach ($f in $filesToProcess) {
    Process-ScribeFile -Path $f.FullName
}

Log-Info ""
Log-Success "Batch processing complete."

if (-not $global:DryRun -and -not $All) {
    $nowUtc = (Get-Date).ToUniversalTime()
    $state = @{ LastRunUtc = $nowUtc.ToString("o") }
    $state | ConvertTo-Json | Set-Content -Path $statePath
    Log-Info "Updated LastRunUtc in state: $($nowUtc.ToString('u'))"
}
