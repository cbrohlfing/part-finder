# ==========================================================
# Part Finder (Multi-folder Search)
# Version: v1.3 (Layout fix + sortable columns)
#
# Search engine:
# - Search runs inside Start-Job (separate powershell.exe process) - stable on PS 5.1
# - UI polls job output via WinForms Timer and updates progress/results
#
# Features:
# - No default folders; user adds folders (typed/pasted or Browse)
# - Search mode selectable (default: Filename contains)
# - Include subfolders checkbox
# - Max results limit
# - Ignore patterns (semicolon-separated wildcards), remembered
# - Remembers last query in settings JSON
# - Clear "actively searching" indicators:
#     * Marquee progress bar
#     * Status line + current folder
#     * Results placeholder row "Searching..."
# - Stop cancels/terminates the job
# - Sort columns by clicking column headers (toggle asc/desc)
# - Debug log: part_finder_debug.log
# - Settings:  part_finder_settings.json
# ==========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# -------------------------
# Version-local paths
# -------------------------
$script:Version = "v1.3"
$script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsPath = Join-Path $script:ScriptDir "part_finder_settings.json"
$script:DebugLog = Join-Path $script:ScriptDir "part_finder_debug.log"



# -------------------------
# Debug logging
# -------------------------
$script:EnableDebugLog = $true

function DLog([string]$msg) {
    if (-not $script:EnableDebugLog) { return }
    try {
        Add-Content -LiteralPath $script:DebugLog -Value ("{0} {1}" -f (Get-Date -Format s), $msg) -Encoding UTF8
    } catch { }
}

DLog ("SCRIPT START ({0})" -f $script:Version)

trap {
    try { DLog ("TRAP: {0}`r`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace) } catch {}
    continue
}

[System.Windows.Forms.Application]::add_ThreadException({
        try { DLog ("UI THREAD EXCEPTION: {0}`r`n{1}" -f $_.Exception.Message, $_.Exception.StackTrace) } catch {}
    })

[System.AppDomain]::CurrentDomain.add_UnhandledException({
        try {
            $ex = $_.ExceptionObject
            DLog ("UNHANDLED EXCEPTION: {0}" -f $ex.ToString())
        } catch {}
    })

# -------------------------
# Settings load/save
# -------------------------
$script:DefaultIgnore = "*.log;*.bak"
$script:DefaultSettings = [PSCustomObject]@{
    folders           = @()                    # intentionally empty
    searchMode        = "Filename contains"
    includeSubfolders = $true
    maxResults        = 200
    ignorePatterns    = $script:DefaultIgnore
    lastQuery         = ""
}

function Load-Settings([string]$path, $fallback) {
    try {
        if (Test-Path -LiteralPath $path) {
            $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
            $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
            return [PSCustomObject]@{
                folders           = if ($cfg.folders) { @($cfg.folders | ForEach-Object { [string]$_ }) } else { @($fallback.folders) }
                searchMode        = if ($cfg.searchMode) { [string]$cfg.searchMode } else { [string]$fallback.searchMode }
                includeSubfolders = if ($null -ne $cfg.includeSubfolders) { [bool]$cfg.includeSubfolders } else { [bool]$fallback.includeSubfolders }
                maxResults        = if ($null -ne $cfg.maxResults) { [int]$cfg.maxResults } else { [int]$fallback.maxResults }
                ignorePatterns    = if ($cfg.ignorePatterns) { [string]$cfg.ignorePatterns } else { [string]$fallback.ignorePatterns }
                lastQuery         = if ($cfg.lastQuery) { [string]$cfg.lastQuery } else { [string]$fallback.lastQuery }
            }
        }
    } catch {
        DLog ("Load-Settings error: {0}" -f $_.Exception.Message)
    }
    return $fallback
}

function Save-Settings([string]$path, $settingsObject) {
    try {
        $json = $settingsObject | ConvertTo-Json -Depth 6
        Set-Content -LiteralPath $path -Value $json -Encoding UTF8
    } catch {
        DLog ("Save-Settings error: {0}" -f $_.Exception.Message)
    }
}

$script:cfg = Load-Settings -path $script:SettingsPath -fallback $script:DefaultSettings

# -------------------------
# UI helpers
# -------------------------
function Ui([System.Windows.Forms.Control]$ctl, [scriptblock]$sb) {
    try {
        if ($null -eq $ctl -or $ctl.IsDisposed) { return }
        if (-not $ctl.IsHandleCreated) { $null = $ctl.Handle }
        if ($ctl.InvokeRequired) {
            $null = $ctl.BeginInvoke([Action] { try { & $sb } catch {} })
        } else {
            & $sb
        }
    } catch {
        DLog ("UI invoke error: {0}" -f $_.Exception.Message)
    }
}

function Add-FolderToList([System.Windows.Forms.ListBox]$list, [string]$folder) {
    if ([string]::IsNullOrWhiteSpace($folder)) { return }
    $folder = $folder.Trim().Trim('"').TrimEnd('\')

    foreach ($item in $list.Items) {
        if ([string]::Equals([string]$item, $folder, [System.StringComparison]::OrdinalIgnoreCase)) { return }
    }
    [void]$list.Items.Add($folder)
}

function Get-FoldersFromList([System.Windows.Forms.ListBox]$list) {
    $out = @()
    foreach ($item in $list.Items) { $out += [string]$item }
    return $out
}

function Safe-OpenFile([string]$path) { try { Start-Process -FilePath $path | Out-Null } catch {} }
function Safe-OpenFolderAndSelect([string]$filePath) { try { Start-Process explorer.exe ("/select,`"$filePath`"") | Out-Null } catch {} }
function Copy-ToClipboard([string]$text) { try { [System.Windows.Forms.Clipboard]::SetText($text) } catch {} }

function Parse-IgnorePatterns([string]$raw) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    $raw.Split(';') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne "" }
}

# -------------------------
# UI Layout
# -------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = ("Part Finder ({0})" -f $script:Version)
$form.StartPosition = "CenterScreen"
$form.Width = 1100
$form.Height = 740

# Left group: folders
$grpFolders = New-Object System.Windows.Forms.GroupBox
$grpFolders.Text = "Folders to Search"
$grpFolders.Left = 10
$grpFolders.Top = 10
$grpFolders.Width = 420
$grpFolders.Height = 680

$listFolders = New-Object System.Windows.Forms.ListBox
$listFolders.Left = 12
$listFolders.Top = 25
$listFolders.Width = 395
$listFolders.Height = 500
$listFolders.HorizontalScrollbar = $true

$lblAdd = New-Object System.Windows.Forms.Label
$lblAdd.Left = 12
$lblAdd.Top = 535
$lblAdd.Width = 395
$lblAdd.Text = "Type/paste a folder path (UNC or mapped), or Browse:"

$txtAddFolder = New-Object System.Windows.Forms.TextBox
$txtAddFolder.Left = 12
$txtAddFolder.Top = 555
$txtAddFolder.Width = 395

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Left = 12
$btnBrowse.Top = 585
$btnBrowse.Width = 90

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add"
$btnAdd.Left = 110
$btnAdd.Top = 585
$btnAdd.Width = 70

$btnRemoveFolder = New-Object System.Windows.Forms.Button
$btnRemoveFolder.Text = "Remove Selected"
$btnRemoveFolder.Left = 190
$btnRemoveFolder.Top = 585
$btnRemoveFolder.Width = 130

$btnClearFolders = New-Object System.Windows.Forms.Button
$btnClearFolders.Text = "Clear All"
$btnClearFolders.Left = 330
$btnClearFolders.Top = 585
$btnClearFolders.Width = 77

$lblFolderHint = New-Object System.Windows.Forms.Label
$lblFolderHint.Left = 12
$lblFolderHint.Top = 620
$lblFolderHint.Width = 395
$lblFolderHint.Height = 45
$lblFolderHint.Text = "Tip: Add multiple folders. Search runs across ALL folders. UNC paths like \\server\share\folder are supported."

$grpFolders.Controls.AddRange(@(
        $listFolders,
        $lblAdd, $txtAddFolder,
        $btnBrowse, $btnAdd, $btnRemoveFolder, $btnClearFolders,
        $lblFolderHint
    ))

# Right group: search/results
$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = "Search"
$grpSearch.Left = 440
$grpSearch.Top = 10
$grpSearch.Width = 635
$grpSearch.Height = 680

# --- Row 1: Query + Mode
$lblQuery = New-Object System.Windows.Forms.Label
$lblQuery.Text = "Search:"
$lblQuery.Left = 12
$lblQuery.Top = 30
$lblQuery.Width = 60

$txtQuery = New-Object System.Windows.Forms.TextBox
$txtQuery.Left = 75
$txtQuery.Top = 26
$txtQuery.Width = 370
$txtQuery.Text = [string]$script:cfg.lastQuery

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Text = "Mode:"
$lblMode.Left = 455
$lblMode.Top = 30
$lblMode.Width = 45

$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.Left = 505
$cmbMode.Top = 26
$cmbMode.Width = 115
$cmbMode.DropDownStyle = "DropDownList"
$cmbMode.Items.AddRange(@(
        "Filename contains",
        "Filename starts with",
        "Filename ends with",
        "Exact filename",
        "Wildcard (* and ?)",
        "Regex"
    ))
$cmbMode.SelectedItem = $script:cfg.searchMode
if (-not $cmbMode.SelectedItem) { $cmbMode.SelectedItem = "Filename contains" }

# --- Row 2: Include subfolders + buttons (moved UP so it never collides)
$chkSub = New-Object System.Windows.Forms.CheckBox
$chkSub.Left = 75
$chkSub.Top = 55
$chkSub.Width = 200
$chkSub.Text = "Include subfolders"
$chkSub.Checked = [bool]$script:cfg.includeSubfolders

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Search"
$btnSearch.Left = 285
$btnSearch.Top = 50
$btnSearch.Width = 80

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "Stop"
$btnStop.Left = 375
$btnStop.Top = 50
$btnStop.Width = 70
$btnStop.Enabled = $false

# --- Row 3: Max results + Ignore (dedicated row, no overlap)
$lblMax = New-Object System.Windows.Forms.Label
$lblMax.Text = "Max results:"
$lblMax.Left = 12
$lblMax.Top = 82
$lblMax.Width = 70

$numMax = New-Object System.Windows.Forms.NumericUpDown
$numMax.Left = 85
$numMax.Top = 78
$numMax.Width = 70
$numMax.Minimum = 1
$numMax.Maximum = 10000
$numMax.Value = [decimal][int]$script:cfg.maxResults

$lblIgnore = New-Object System.Windows.Forms.Label
$lblIgnore.Text = "Ignore:"
$lblIgnore.Left = 175
$lblIgnore.Top = 82
$lblIgnore.Width = 50

$txtIgnore = New-Object System.Windows.Forms.TextBox
$txtIgnore.Left = 230
$txtIgnore.Top = 78
$txtIgnore.Width = 392
$txtIgnore.Text = [string]$script:cfg.ignorePatterns

$lblIgnoreDefault = New-Object System.Windows.Forms.Label
$lblIgnoreDefault.Left = 230
$lblIgnoreDefault.Top = 100
$lblIgnoreDefault.Width = 392
$lblIgnoreDefault.Height = 16
$lblIgnoreDefault.ForeColor = [System.Drawing.Color]::DimGray
$lblIgnoreDefault.Text = "Default: *.log;*.bak"

# --- Row 4: progress
$prg = New-Object System.Windows.Forms.ProgressBar
$prg.Left = 12
$prg.Top = 125
$prg.Width = 610
$prg.Height = 14
$prg.Style = 'Blocks'
$prg.Visible = $false

# --- Row 5/6: status + current folder
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Left = 12
$lblStatus.Top = 145
$lblStatus.Width = 610
$lblStatus.Height = 18
$lblStatus.Text = "Ready."

$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Left = 12
$lblCurrent.Top = 165
$lblCurrent.Width = 610
$lblCurrent.Height = 18
$lblCurrent.Text = ""

# --- Results list
$listResults = New-Object System.Windows.Forms.ListView
$listResults.Left = 12
$listResults.Top = 190
$listResults.Width = 610
$listResults.Height = 475
$listResults.View = "Details"
$listResults.FullRowSelect = $true
$listResults.GridLines = $true
$listResults.Font = New-Object System.Drawing.Font("Consolas", 10)
$listResults.HideSelection = $false

$null = $listResults.Columns.Add("Name", 220)
$null = $listResults.Columns.Add("Folder", 300)
$null = $listResults.Columns.Add("Modified", 80)

# Context menu
$resultsMenu = New-Object System.Windows.Forms.ContextMenuStrip
$miOpen = $resultsMenu.Items.Add("Open")
$miOpenFolder = $resultsMenu.Items.Add("Open Containing Folder")
$resultsMenu.Items.Add("-") | Out-Null
$miCopyFull = $resultsMenu.Items.Add("Copy Full Path")
$miCopyFolder = $resultsMenu.Items.Add("Copy Folder Path")
$listResults.ContextMenuStrip = $resultsMenu

$grpSearch.Controls.AddRange(@(
        $lblQuery, $txtQuery, $lblMode, $cmbMode,
        $chkSub, $btnSearch, $btnStop,
        $lblMax, $numMax, $lblIgnore, $txtIgnore, $lblIgnoreDefault,
        $prg, $lblStatus, $lblCurrent, $listResults
    ))

$form.Controls.AddRange(@($grpFolders, $grpSearch))

# Load folders from settings
foreach ($f in @($script:cfg.folders)) { Add-FolderToList $listFolders $f }

# -------------------------
# Persist settings
# -------------------------
function Persist-UiSettings {
    $script:cfg = [PSCustomObject]@{
        folders           = @(Get-FoldersFromList $listFolders)
        searchMode        = [string]$cmbMode.SelectedItem
        includeSubfolders = [bool]$chkSub.Checked
        maxResults        = [int]$numMax.Value
        ignorePatterns    = [string]$txtIgnore.Text
        lastQuery         = [string]$txtQuery.Text
    }
    Save-Settings -path $script:SettingsPath -settingsObject $script:cfg
    DLog ("Settings saved. folders={0} mode='{1}' sub={2} max={3} ignore='{4}' lastQuery='{5}'" -f `
            $script:cfg.folders.Count, $script:cfg.searchMode, $script:cfg.includeSubfolders, $script:cfg.maxResults, $script:cfg.ignorePatterns, $script:cfg.lastQuery)
}

function Set-UiSearching([bool]$isSearching) {
    Ui $form {
        $btnSearch.Enabled = -not $isSearching
        $btnStop.Enabled = $isSearching
        $prg.Visible = $isSearching
        if ($isSearching) {
            $prg.Style = 'Marquee'
            $prg.MarqueeAnimationSpeed = 30
        } else {
            $prg.Style = 'Blocks'
            $prg.MarqueeAnimationSpeed = 0
        }
    }
}

function Set-Status([string]$text) { Ui $form { $lblStatus.Text = $text } }
function Set-Current([string]$text) { Ui $form { $lblCurrent.Text = $text } }

function Add-ResultRow([string]$name, [string]$folder, [datetime]$modified, [string]$full) {
    Ui $form {
        $item = New-Object System.Windows.Forms.ListViewItem($name)
        [void]$item.SubItems.Add($folder)
        [void]$item.SubItems.Add(($modified.ToString("yyyy-MM-dd")))
        # Tag holds richer info for sorting + actions
        $item.Tag = [pscustomobject]@{ FullPath = $full; Modified = $modified; Name = $name; Folder = $folder }
        [void]$listResults.Items.Add($item)
    }
}

# -------------------------
# Sorting support (ListView ColumnClick)
# -------------------------
$script:SortColumn = -1
$script:SortAsc = $true

function Compare-Text([string]$a, [string]$b) {
    return [string]::Compare($a, $b, $true)  # ignore case
}

$listResults.Add_ColumnClick({
        param($sender, $e)

        try {
            $col = [int]$e.Column
            if ($script:SortColumn -eq $col) {
                $script:SortAsc = -not $script:SortAsc
            } else {
                $script:SortColumn = $col
                $script:SortAsc = $true
            }

            # Capture items, sort, re-add
            $items = @()
            foreach ($it in $listResults.Items) { $items += $it }

            $sorted = $items | Sort-Object -Stable -Property @{
                Expression = {
                    $tag = $_.Tag
                    switch ($script:SortColumn) {
                        0 { [string]$tag.Name }
                        1 { [string]$tag.Folder }
                        2 { [datetime]$tag.Modified }
                        default { [string]$tag.Name }
                    }
                }
                Ascending  = $script:SortAsc
            }

            Ui $form {
                $listResults.BeginUpdate()
                try {
                    $listResults.Items.Clear()
                    foreach ($it in $sorted) { [void]$listResults.Items.Add($it) }
                } finally {
                    $listResults.EndUpdate()
                }
            }
        } catch {
            DLog ("SORT ERROR: {0}" -f $_.Exception.Message)
        }
    })

# -------------------------
# Job-based search engine
# -------------------------
$script:SearchJob = $null
$script:PollTimer = $null
$script:Found = 0
$script:Scanned = 0
$script:JobStart = $null
$script:SeenPaths = $null   # UI-side dedup

function Stop-Search {
    try {
        DLog "Stop-Search: requested"
        if ($script:PollTimer) {
            try { $script:PollTimer.Stop(); $script:PollTimer.Dispose() } catch {}
            $script:PollTimer = $null
        }
        if ($script:SearchJob) {
            try { Stop-Job -Job $script:SearchJob -Force -ErrorAction SilentlyContinue } catch {}
            try { Remove-Job -Job $script:SearchJob -Force -ErrorAction SilentlyContinue } catch {}
            $script:SearchJob = $null
        }
    } catch {
        DLog ("Stop-Search error: {0}" -f $_.Exception.Message)
    } finally {
        Set-Current ""
        Set-UiSearching $false
        if ($listResults.Items.Count -ge 1 -and $listResults.Items[0].Text -eq "Searching...") {
            Ui $form { $listResults.Items.RemoveAt(0) }
        }
        if ($lblStatus.Text -like "Starting*") { Set-Status "Stopped." }
    }
}

function Start-Search([string[]]$folders, [string]$query, [string]$mode, [bool]$recurse, [int]$maxResults, [string[]]$ignorePatterns) {
    DLog ("Start-Search: folders={0} query='{1}' mode='{2}' recurse={3} max={4} ignore='{5}'" -f `
            $folders.Count, $query, $mode, $recurse, $maxResults, ($ignorePatterns -join ';'))

    Stop-Search

    # UI pre-state
    Ui $form {
        $listResults.Items.Clear()
        $ph = New-Object System.Windows.Forms.ListViewItem("Searching...")
        [void]$ph.SubItems.Add("")
        [void]$ph.SubItems.Add("")
        $ph.ForeColor = [System.Drawing.Color]::Gray
        [void]$listResults.Items.Add($ph)
    }

    $script:Found = 0
    $script:Scanned = 0
    $script:JobStart = Get-Date
    $script:SeenPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    Set-UiSearching $true
    Set-Status "Starting..."
    Set-Current ""

    $jobScript = {
        param($folders, $query, $mode, $recurse, $maxResults, $ignorePatterns)

        function IsMatch([string]$name, [string]$q, [string]$mode) {
            switch ($mode) {
                "Filename contains" { return ($name.IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) }
                "Filename starts with" { return $name.StartsWith($q, [System.StringComparison]::OrdinalIgnoreCase) }
                "Filename ends with" { return $name.EndsWith($q, [System.StringComparison]::OrdinalIgnoreCase) }
                "Exact filename" { return [string]::Equals($name, $q, [System.StringComparison]::OrdinalIgnoreCase) }
                "Wildcard (* and ?)" { return ($name -like $q) }
                "Regex" {
                    try {
                        $rx = New-Object System.Text.RegularExpressions.Regex($q, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        return $rx.IsMatch($name)
                    } catch {
                        return $false
                    }
                }
                default { return ($name.IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) }
            }
        }

        function ShouldIgnore([string]$name, [string[]]$patterns) {
            if (-not $patterns -or $patterns.Count -lt 1) { return $false }
            foreach ($p in $patterns) {
                if ([string]::IsNullOrWhiteSpace($p)) { continue }
                if ($name -like $p) { return $true }
            }
            return $false
        }

        $scanned = 0
        $found = 0
        $folderIndex = 0
        $folderCount = $folders.Count
        $lastEmit = Get-Date

        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($root in $folders) {
            $folderIndex++
            if ($found -ge $maxResults) { break }

            if (-not (Test-Path -LiteralPath $root)) {
                [pscustomobject]@{ Type = 'progress'; Current = ("[{0}/{1}] Unreachable: {2}" -f $folderIndex, $folderCount, $root); Scanned = $scanned; Found = $found } | Write-Output
                continue
            }

            [pscustomobject]@{ Type = 'progress'; Current = ("[{0}/{1}] {2}" -f $folderIndex, $folderCount, $root); Scanned = $scanned; Found = $found } | Write-Output

            try {
                $opt = if ($recurse) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }

                foreach ($filePath in [System.IO.Directory]::EnumerateFiles($root, "*", $opt)) {
                    $scanned++

                    $now = Get-Date
                    if (($now - $lastEmit).TotalMilliseconds -ge 300) {
                        $lastEmit = $now
                        [pscustomobject]@{ Type = 'progress'; Current = ("[{0}/{1}] {2}" -f $folderIndex, $folderCount, $root); Scanned = $scanned; Found = $found } | Write-Output
                    }

                    $name = [System.IO.Path]::GetFileName($filePath)

                    if (ShouldIgnore $name $ignorePatterns) { continue }

                    if (IsMatch $name $query $mode) {
                        if ($seen.Add($filePath)) {
                            $found++
                            try {
                                $fi = New-Object System.IO.FileInfo($filePath)
                                [pscustomobject]@{
                                    Type     = 'match'
                                    Name     = $fi.Name
                                    Folder   = $fi.DirectoryName
                                    Modified = $fi.LastWriteTime
                                    FullPath = $fi.FullName
                                } | Write-Output
                            } catch { }

                            if ($found -ge $maxResults) { break }
                        }
                    }
                }
            } catch {
                [pscustomobject]@{
                    Type    = 'progress'
                    Current = ("[{0}/{1}] Error: {2}" -f $folderIndex, $folderCount, $_.Exception.Message)
                    Scanned = $scanned
                    Found   = $found
                } | Write-Output
            }
        }

        [pscustomobject]@{ Type = 'done'; Scanned = $scanned; Found = $found } | Write-Output
    }

    $script:SearchJob = Start-Job -ScriptBlock $jobScript -ArgumentList @($folders, $query, $mode, $recurse, $maxResults, $ignorePatterns)
    DLog ("Start-Search: job started id={0}" -f $script:SearchJob.Id)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 200
    $timer.Add_Tick({
            try {
                if (-not $script:SearchJob) { return }

                $items = @()
                try { $items = Receive-Job -Job $script:SearchJob -Keep -ErrorAction SilentlyContinue } catch {}

                foreach ($o in $items) {
                    if ($null -eq $o) { continue }
                    if (-not $o.PSObject.Properties['Type']) { continue }

                    switch ($o.Type) {
                        'progress' {
                            $script:Scanned = [int]$o.Scanned
                            $script:Found = [int]$o.Found
                            Set-Current ([string]$o.Current)
                            Set-Status ("Scanning... Files: {0} | Matches: {1}" -f $script:Scanned, $script:Found)
                        }
                        'match' {
                            # UI dedup (just in case)
                            $fp = [string]$o.FullPath
                            if ($script:SeenPaths -and $fp -and $script:SeenPaths.Add($fp)) {
                                Add-ResultRow ([string]$o.Name) ([string]$o.Folder) ([datetime]$o.Modified) $fp
                                Ui $form {
                                    if ($listResults.Items.Count -ge 1 -and $listResults.Items[0].Text -eq "Searching...") {
                                        $listResults.Items.RemoveAt(0)
                                    }
                                }
                            }
                        }
                        'done' {
                            $script:Scanned = [int]$o.Scanned
                            $script:Found = [int]$o.Found
                        }
                    }
                }

                if ($script:SearchJob.State -in @('Completed', 'Failed', 'Stopped')) {
                    $state = $script:SearchJob.State
                    DLog ("Job finished. state={0}" -f $state)

                    try { $script:PollTimer.Stop(); $script:PollTimer.Dispose() } catch {}
                    $script:PollTimer = $null

                    try { $null = Receive-Job -Job $script:SearchJob -ErrorAction SilentlyContinue } catch {}

                    if ($state -eq 'Completed') {
                        Set-Status ("Done. Matches: {0} | Files scanned: {1}" -f $script:Found, $script:Scanned)
                    } elseif ($state -eq 'Stopped') {
                        Set-Status "Stopped."
                    } else {
                        $err = ""
                        try {
                            $errs = $script:SearchJob.ChildJobs[0].Error
                            if ($errs -and $errs.Count -gt 0) { $err = $errs[0].ToString() }
                        } catch {}
                        if ($err) { Set-Status ("Error: {0}" -f $err) } else { Set-Status "Error: search job failed." }
                    }

                    Set-Current ""
                    Set-UiSearching $false

                    Ui $form {
                        if ($listResults.Items.Count -ge 1 -and $listResults.Items[0].Text -eq "Searching...") {
                            $listResults.Items.RemoveAt(0)
                        }
                    }

                    try { Remove-Job -Job $script:SearchJob -Force -ErrorAction SilentlyContinue } catch {}
                    $script:SearchJob = $null
                }
            } catch {
                DLog ("POLL TIMER ERROR: {0}" -f $_.Exception.Message)
                Set-Status ("Error: {0}" -f $_.Exception.Message)
                Stop-Search
            }
        })

    $script:PollTimer = $timer
    $timer.Start()
}

# -------------------------
# Events - folder controls
# -------------------------
$btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select a folder to include in searches"
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtAddFolder.Text = $dlg.SelectedPath
        }
    })

$btnAdd.Add_Click({
        $p = $txtAddFolder.Text
        if ([string]::IsNullOrWhiteSpace($p)) { return }
        Add-FolderToList $listFolders $p
        $txtAddFolder.Text = ""
        Persist-UiSettings
    })

$txtAddFolder.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnAdd.PerformClick()
            $_.SuppressKeyPress = $true
        }
    })

$btnRemoveFolder.Add_Click({
        $idx = $listFolders.SelectedIndex
        if ($idx -ge 0) {
            $listFolders.Items.RemoveAt($idx)
            Persist-UiSettings
        }
    })

$btnClearFolders.Add_Click({
        $listFolders.Items.Clear()
        Persist-UiSettings
    })

# -------------------------
# Events - search controls
# -------------------------
$btnSearch.Add_Click({
        try {
            DLog "Search button clicked"

            $folders = @(Get-FoldersFromList $listFolders)
            DLog ("Folders count: {0}" -f $folders.Count)

            $query = ($txtQuery.Text).Trim()
            DLog ("Query: '{0}'" -f $query)

            if ($folders.Count -lt 1) { Set-Status "Error: Add at least one folder."; return }
            if ([string]::IsNullOrWhiteSpace($query)) { Set-Status "Error: Enter a search value."; return }

            $mode = [string]$cmbMode.SelectedItem
            $recurse = [bool]$chkSub.Checked
            $max = [int]$numMax.Value

            $ignoreRaw = [string]$txtIgnore.Text
            if ([string]::IsNullOrWhiteSpace($ignoreRaw)) { $ignoreRaw = $script:DefaultIgnore }
            $ignore = Parse-IgnorePatterns $ignoreRaw

            Persist-UiSettings
            DLog ("Mode: '{0}' recurse={1} max={2} ignore='{3}'" -f $mode, $recurse, $max, ($ignore -join ';'))

            Start-Search -folders $folders -query $query -mode $mode -recurse $recurse -maxResults $max -ignorePatterns $ignore
            DLog "Search click handler finished normally"
        } catch {
            DLog ("SEARCH CLICK ERROR: {0}`r`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace)
            try { Set-UiSearching $false } catch {}
            try { Set-Status ("Error: {0}" -f $_.Exception.Message) } catch {}
        }
    })

$btnStop.Add_Click({ Stop-Search })

$txtQuery.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnSearch.PerformClick()
            $_.SuppressKeyPress = $true
        }
    })

# -------------------------
# Results actions
# -------------------------
$listResults.Add_DoubleClick({
        if ($listResults.SelectedItems.Count -lt 1) { return }
        $tag = $listResults.SelectedItems[0].Tag
        $full = if ($tag -and $tag.FullPath) { [string]$tag.FullPath } else { "" }
        if ($full) { Safe-OpenFile $full }
    })

$miOpen.Add_Click({
        if ($listResults.SelectedItems.Count -lt 1) { return }
        $tag = $listResults.SelectedItems[0].Tag
        $full = if ($tag -and $tag.FullPath) { [string]$tag.FullPath } else { "" }
        if ($full) { Safe-OpenFile $full }
    })

$miOpenFolder.Add_Click({
        if ($listResults.SelectedItems.Count -lt 1) { return }
        $tag = $listResults.SelectedItems[0].Tag
        $full = if ($tag -and $tag.FullPath) { [string]$tag.FullPath } else { "" }
        if ($full) { Safe-OpenFolderAndSelect $full }
    })

$miCopyFull.Add_Click({
        if ($listResults.SelectedItems.Count -lt 1) { return }
        $tag = $listResults.SelectedItems[0].Tag
        $full = if ($tag -and $tag.FullPath) { [string]$tag.FullPath } else { "" }
        if ($full) { Copy-ToClipboard $full }
    })

$miCopyFolder.Add_Click({
        if ($listResults.SelectedItems.Count -lt 1) { return }
        $tag = $listResults.SelectedItems[0].Tag
        $full = if ($tag -and $tag.FullPath) { [string]$tag.FullPath } else { "" }
        if ($full) { Copy-ToClipboard (Split-Path $full -Parent) }
    })

$form.Add_FormClosing({
        try { Stop-Search } catch {}
        try { Persist-UiSettings } catch {}
        DLog "FORM CLOSING"
    })

# Show
DLog "Showing UI"










[void]$form.ShowDialog()
DLog "SCRIPT END"




