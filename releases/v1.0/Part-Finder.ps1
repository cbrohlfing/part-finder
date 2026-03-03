# ==========================================================
# Part Finder (Multi-folder Search)
# Version: v1.0.9 (PS5.1-safe enumerator + no-runspace worker)
#
# - No default folders; user adds folders in UI
# - Search style selectable in UI (default: Filename contains)
# - Clear "actively searching" indicators:
#     * Marquee progress bar
#     * Results placeholder row: "Searching..."
#     * Status line + current folder line
# - Background search runs on dedicated Thread
# - Worker never executes PowerShell ScriptBlocks (avoids "no Runspace")
# - UI updates are posted via Control.BeginInvoke(Action)
# - Debug log file beside script: part_finder_debug.log
# - Settings saved beside script: part_finder_settings.json
# ==========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Core

# -------------------------
# Version-local paths
# -------------------------
$script:ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsPath  = Join-Path $script:ScriptDir "part_finder_settings.json"
$script:DebugLog      = Join-Path $script:ScriptDir "part_finder_debug.log"

# -------------------------
# Debug
# -------------------------
$script:EnableDebugLog = $true

function DLog([string]$msg) {
    if (-not $script:EnableDebugLog) { return }
    try {
        Add-Content -LiteralPath $script:DebugLog -Value ("{0} {1}" -f (Get-Date -Format s), $msg) -Encoding UTF8
    } catch {}
}

DLog "SCRIPT START (v1.0.9)"

# -------------------------
# Safety nets
# -------------------------
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
$script:DefaultSettings = [PSCustomObject]@{
    folders = @()                  # intentionally empty
    searchMode = "Filename contains"
    includeSubfolders = $true
}

function Load-Settings([string]$path, $fallback) {
    try {
        if (Test-Path -LiteralPath $path) {
            $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
            $cfg = $raw | ConvertFrom-Json -ErrorAction Stop

            return [PSCustomObject]@{
                folders = if ($cfg.folders) { @($cfg.folders | ForEach-Object { [string]$_ }) } else { @($fallback.folders) }
                searchMode = if ($cfg.searchMode) { [string]$cfg.searchMode } else { [string]$fallback.searchMode }
                includeSubfolders = if ($null -ne $cfg.includeSubfolders) { [bool]$cfg.includeSubfolders } else { [bool]$fallback.includeSubfolders }
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
# UI helper: post to UI thread safely
# -------------------------
function UiPost([System.Windows.Forms.Control]$ctl, [scriptblock]$sb) {
    try {
        if ($null -eq $ctl -or $ctl.IsDisposed) { return }
        if (-not $ctl.IsHandleCreated) { $null = $ctl.Handle }

        if ($ctl.InvokeRequired) {
            $null = $ctl.BeginInvoke([Action]{
                try { & $sb } catch {}
            })
        } else {
            & $sb
        }
    } catch {
        DLog ("UiPost error: {0}" -f $_.Exception.Message)
    }
}

# -------------------------
# Helpers
# -------------------------
function Normalize-Folder([string]$p) {
    if ([string]::IsNullOrWhiteSpace($p)) { return $null }
    $p = $p.Trim().Trim('"').Trim()
    if ($p.EndsWith('\')) { $p = $p.TrimEnd('\') }
    return $p
}

function Add-FolderToList([System.Windows.Forms.ListBox]$list, [string]$folder) {
    $folder = Normalize-Folder $folder
    if (-not $folder) { return }

    foreach ($item in $list.Items) {
        if ([string]::Equals([string]$item, $folder, [System.StringComparison]::OrdinalIgnoreCase)) {
            return
        }
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

# -------------------------
# Matching (NO scriptblocks in worker)
# -------------------------
function New-Matcher([string]$mode, [string]$query) {
    switch ($mode) {
        "Filename contains"    { return @{ Mode="contains"; Query=$query } }
        "Filename starts with" { return @{ Mode="starts";   Query=$query } }
        "Filename ends with"   { return @{ Mode="ends";     Query=$query } }
        "Exact filename"       { return @{ Mode="exact";    Query=$query } }
        "Wildcard (* and ?)"   {
            $wp = New-Object System.Management.Automation.WildcardPattern($query, [System.Management.Automation.WildcardOptions]::IgnoreCase)
            return @{ Mode="wildcard"; Wild=$wp }
        }
        "Regex" {
            $rx = New-Object System.Text.RegularExpressions.Regex($query, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            return @{ Mode="regex"; Rx=$rx }
        }
        default { return @{ Mode="contains"; Query=$query } }
    }
}

function Is-Match([hashtable]$matcher, [string]$name) {
    try {
        switch ($matcher.Mode) {
            "contains" { return ($name.IndexOf($matcher.Query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) }
            "starts"   { return $name.StartsWith($matcher.Query, [System.StringComparison]::OrdinalIgnoreCase) }
            "ends"     { return $name.EndsWith($matcher.Query, [System.StringComparison]::OrdinalIgnoreCase) }
            "exact"    { return [string]::Equals($name, $matcher.Query, [System.StringComparison]::OrdinalIgnoreCase) }
            "wildcard" { return $matcher.Wild.IsMatch($name) }
            "regex"    { return $matcher.Rx.IsMatch($name) }
            default    { return ($name.IndexOf($matcher.Query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) }
        }
    } catch {
        return $false
    }
}

# -------------------------
# PS5.1-safe enumeration (iterative DFS, skips errors)
# -------------------------
function Enumerate-FilesSafe {
    param(
        [Parameter(Mandatory=$true)] [string]$Root,
        [Parameter(Mandatory=$true)] [bool]$Recurse,
        [Parameter(Mandatory=$true)] $Token   # CancellationToken
    )

    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Root)

    while ($stack.Count -gt 0) {
        if ($Token.IsCancellationRequested) { break }

        $dir = $stack.Pop()

        # Enumerate files
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($dir)) {
                if ($Token.IsCancellationRequested) { break }
                $f
            }
        } catch {}

        if (-not $Recurse) { continue }

        # Enumerate subdirs
        try {
            foreach ($d in [System.IO.Directory]::EnumerateDirectories($dir)) {
                if ($Token.IsCancellationRequested) { break }
                $stack.Push($d)
            }
        } catch {}
    }
}

# -------------------------
# UI Layout
# -------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Part Finder (v1.0.9)"
$form.StartPosition = "CenterScreen"
$form.Width = 1100
$form.Height = 720

# Left group: folders
$grpFolders = New-Object System.Windows.Forms.GroupBox
$grpFolders.Text = "Folders to Search"
$grpFolders.Left = 10
$grpFolders.Top = 10
$grpFolders.Width = 420
$grpFolders.Height = 660

$listFolders = New-Object System.Windows.Forms.ListBox
$listFolders.Left = 12
$listFolders.Top = 25
$listFolders.Width = 395
$listFolders.Height = 500
$listFolders.HorizontalScrollbar = $true

$lblPaste = New-Object System.Windows.Forms.Label
$lblPaste.Left = 12
$lblPaste.Top = 530
$lblPaste.Width = 395
$lblPaste.Height = 18
$lblPaste.Text = "Type/paste a folder path (UNC or mapped), or Browse:"

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Left = 12
$txtFolder.Top = 552
$txtFolder.Width = 395

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Left = 12
$btnBrowse.Top = 585
$btnBrowse.Width = 90

$btnAddTyped = New-Object System.Windows.Forms.Button
$btnAddTyped.Text = "Add"
$btnAddTyped.Left = 110
$btnAddTyped.Top = 585
$btnAddTyped.Width = 70

$btnRemoveFolder = New-Object System.Windows.Forms.Button
$btnRemoveFolder.Text = "Remove Selected"
$btnRemoveFolder.Left = 188
$btnRemoveFolder.Top = 585
$btnRemoveFolder.Width = 130

$btnClearFolders = New-Object System.Windows.Forms.Button
$btnClearFolders.Text = "Clear All"
$btnClearFolders.Left = 326
$btnClearFolders.Top = 585
$btnClearFolders.Width = 81

$lblFolderHint = New-Object System.Windows.Forms.Label
$lblFolderHint.Left = 12
$lblFolderHint.Top = 620
$lblFolderHint.Width = 395
$lblFolderHint.Height = 40
$lblFolderHint.Text = "Tip: Add multiple folders. Search runs across ALL folders. UNC paths like \\server\share\folder are supported."

$grpFolders.Controls.AddRange(@(
    $listFolders, $lblPaste, $txtFolder,
    $btnBrowse, $btnAddTyped, $btnRemoveFolder, $btnClearFolders,
    $lblFolderHint
))

# Right group: search/results
$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = "Search"
$grpSearch.Left = 440
$grpSearch.Top = 10
$grpSearch.Width = 635
$grpSearch.Height = 660

$lblQuery = New-Object System.Windows.Forms.Label
$lblQuery.Text = "Search:"
$lblQuery.Left = 12
$lblQuery.Top = 30
$lblQuery.Width = 60

$txtQuery = New-Object System.Windows.Forms.TextBox
$txtQuery.Left = 75
$txtQuery.Top = 26
$txtQuery.Width = 370

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

$prg = New-Object System.Windows.Forms.ProgressBar
$prg.Left = 12
$prg.Top = 85
$prg.Width = 610
$prg.Height = 14
$prg.Style = 'Marquee'
$prg.MarqueeAnimationSpeed = 30
$prg.Visible = $false

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Left = 12
$lblStatus.Top = 105
$lblStatus.Width = 610
$lblStatus.Height = 18
$lblStatus.Text = "Ready."

$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Left = 12
$lblCurrent.Top = 125
$lblCurrent.Width = 610
$lblCurrent.Height = 18
$lblCurrent.Text = ""

$listResults = New-Object System.Windows.Forms.ListView
$listResults.Left = 12
$listResults.Top = 150
$listResults.Width = 610
$listResults.Height = 500
$listResults.View = "Details"
$listResults.FullRowSelect = $true
$listResults.GridLines = $true
$listResults.Font = New-Object System.Drawing.Font("Consolas", 10)

$null = $listResults.Columns.Add("Name", 220)
$null = $listResults.Columns.Add("Folder", 290)
$null = $listResults.Columns.Add("Modified", 90)

# Results context menu
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
    $prg, $lblStatus, $lblCurrent, $listResults
))

$form.Controls.AddRange(@($grpFolders, $grpSearch))

# Load folders from settings
foreach ($f in @($script:cfg.folders)) { Add-FolderToList $listFolders $f }

# -------------------------
# Persist UI settings
# -------------------------
function Persist-UiSettings {
    $script:cfg = [PSCustomObject]@{
        folders = @(Get-FoldersFromList $listFolders)
        searchMode = [string]$cmbMode.SelectedItem
        includeSubfolders = [bool]$chkSub.Checked
    }
    Save-Settings -path $script:SettingsPath -settingsObject $script:cfg
    DLog ("Settings saved. folders={0} mode='{1}' sub={2}" -f $script:cfg.folders.Count, $script:cfg.searchMode, $script:cfg.includeSubfolders)
}

# -------------------------
# Search control
# -------------------------
$script:cts = $null
$script:workerThread = $null

function Set-UiSearching([bool]$isSearching) {
    UiPost $form {
        $btnSearch.Enabled = -not $isSearching
        $btnStop.Enabled   = $isSearching
        $prg.Visible       = $isSearching
        if ($isSearching) {
            $prg.Style = 'Marquee'
            $prg.MarqueeAnimationSpeed = 30
        } else {
            $prg.Style = 'Blocks'
            $prg.MarqueeAnimationSpeed = 0
        }
    }
}

function Set-Status([string]$text)  { UiPost $form { $lblStatus.Text  = $text } }
function Set-Current([string]$text) { UiPost $form { $lblCurrent.Text = $text } }

function Clear-ResultsAndShowSearching {
    UiPost $form {
        $listResults.Items.Clear()
        $ph = New-Object System.Windows.Forms.ListViewItem("Searching...")
        [void]$ph.SubItems.Add("")
        [void]$ph.SubItems.Add("")
        $ph.ForeColor = [System.Drawing.Color]::Gray
        [void]$listResults.Items.Add($ph)
    }
}

function Remove-SearchingPlaceholder {
    UiPost $form {
        if ($listResults.Items.Count -ge 1 -and $listResults.Items[0].Text -eq "Searching...") {
            $listResults.Items.RemoveAt(0)
        }
    }
}

function Add-ResultRow([string]$name, [string]$folder, [datetime]$modified, [string]$full) {
    UiPost $form {
        $item = New-Object System.Windows.Forms.ListViewItem($name)
        [void]$item.SubItems.Add($folder)
        [void]$item.SubItems.Add(($modified.ToString("yyyy-MM-dd")))
        $item.Tag = $full
        [void]$listResults.Items.Add($item)
    }
}

function Stop-Search {
    try {
        if ($script:cts) {
            DLog "Stop-Search: requested"
            $script:cts.Cancel()
        }
    } catch {
        DLog ("Stop-Search error: {0}" -f $_.Exception.Message)
    }
}

function Start-Search([string[]]$folders, [string]$query, [string]$mode, [bool]$recurse) {
    DLog ("Start-Search: folders={0} query='{1}' mode='{2}' recurse={3}" -f $folders.Count, $query, $mode, $recurse)

    Stop-Search

    $script:cts = New-Object System.Threading.CancellationTokenSource
    $token = $script:cts.Token

    Clear-ResultsAndShowSearching
    Set-UiSearching $true
    Set-Status "Starting..."
    Set-Current ""

    $foldersLocal = @($folders)
    $queryLocal   = $query
    $modeLocal    = $mode
    $recurseLocal = $recurse

    # Build matcher once (PS objects only used on UI thread; worker uses .NET hashtable + .NET types)
    $matcher = New-Matcher -mode $modeLocal -query $queryLocal

    $threadBody = {
        try {
            DLog "Worker thread started"
            $found = 0
            $scanned = 0
            $folderIndex = 0
            $folderCount = $foldersLocal.Count
            $lastUi = [DateTime]::UtcNow

            foreach ($root in $foldersLocal) {
                $folderIndex++
                if ($token.IsCancellationRequested) { break }

                if (-not [System.IO.Directory]::Exists($root)) {
                    DLog ("Folder unreachable: {0}" -f $root)
                    Set-Current ("[{0}/{1}] Unreachable: {2}" -f $folderIndex, $folderCount, $root)
                    continue
                }

                Set-Current ("[{0}/{1}] {2}" -f $folderIndex, $folderCount, $root)
                DLog ("Searching folder: {0}" -f $root)

                foreach ($filePath in (Enumerate-FilesSafe -Root $root -Recurse $recurseLocal -Token $token)) {
                    if ($token.IsCancellationRequested) { break }
                    $scanned++

                    $nowUtc = [DateTime]::UtcNow
                    if (($nowUtc - $lastUi).TotalMilliseconds -ge 250) {
                        $lastUi = $nowUtc
                        Set-Status ("Scanning... Files: {0} | Matches: {1}" -f $scanned, $found)
                    }

                    $name = [System.IO.Path]::GetFileName($filePath)
                    if (Is-Match -matcher $matcher -name $name) {
                        $found++
                        try {
                            $fi = New-Object System.IO.FileInfo($filePath)
                            Add-ResultRow -name $fi.Name -folder $fi.DirectoryName -modified $fi.LastWriteTime -full $fi.FullName
                        } catch {}
                    }
                }
            }

            if ($token.IsCancellationRequested) {
                DLog "Search stopped by user"
                Set-Status "Stopped."
            } else {
                DLog ("Search complete. matches={0} scanned={1}" -f $found, $scanned)
                Set-Status ("Done. Matches: {0} | Files scanned: {1}" -f $found, $scanned)
            }
        }
        catch {
            DLog ("WORKER FATAL: {0}" -f $_.Exception.Message)
            Set-Status ("Fatal error: {0}" -f $_.Exception.Message)
        }
        finally {
            Remove-SearchingPlaceholder
            Set-Current ""
            Set-UiSearching $false
            DLog "Worker thread finished"
        }
    }.GetNewClosure()

    $script:workerThread = New-Object System.Threading.Thread([System.Threading.ThreadStart]$threadBody)
    $script:workerThread.IsBackground = $true
    $script:workerThread.Start()

    DLog "Start-Search: thread started"
}

# -------------------------
# Events: folders
# -------------------------
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select a folder to include in searches"
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolder.Text = $dlg.SelectedPath
    }
})

$btnAddTyped.Add_Click({
    $p = Normalize-Folder $txtFolder.Text
    if (-not $p) { return }

    Add-FolderToList $listFolders $p
    $txtFolder.Text = ""
    Persist-UiSettings
})

$txtFolder.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $btnAddTyped.PerformClick()
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
# Events: search
# -------------------------
$btnSearch.Add_Click({
    try {
        DLog "Search button clicked"

        $folders = @(Get-FoldersFromList $listFolders)
        DLog ("Folders count: {0}" -f $folders.Count)

        $query = [string]$txtQuery.Text
        $query = $query.Trim()
        DLog ("Query: '{0}'" -f $query)

        if ($folders.Count -lt 1) {
            Set-Status "Error: Add at least one folder."
            return
        }
        if ([string]::IsNullOrWhiteSpace($query)) {
            Set-Status "Error: Enter a search value."
            return
        }

        $mode = [string]$cmbMode.SelectedItem
        $recurse = [bool]$chkSub.Checked
        DLog ("Mode: '{0}' recurse={1}" -f $mode, $recurse)

        Persist-UiSettings

        Start-Search -folders $folders -query $query -mode $mode -recurse $recurse
        DLog "Search click handler finished normally"
    }
    catch {
        DLog ("SEARCH CLICK ERROR: {0}`r`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace)
        try { Set-UiSearching $false } catch {}
        try { Set-Status ("Error: {0}" -f $_.Exception.Message) } catch {}
    }
})

$btnStop.Add_Click({
    Stop-Search
    Set-Status "Stopping..."
})

$txtQuery.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $btnSearch.PerformClick()
        $_.SuppressKeyPress = $true
    }
})

# Results interactions
$listResults.Add_DoubleClick({
    if ($listResults.SelectedItems.Count -lt 1) { return }
    $full = [string]$listResults.SelectedItems[0].Tag
    if ($full) { Safe-OpenFile $full }
})

$miOpen.Add_Click({
    if ($listResults.SelectedItems.Count -lt 1) { return }
    $full = [string]$listResults.SelectedItems[0].Tag
    if ($full) { Safe-OpenFile $full }
})

$miOpenFolder.Add_Click({
    if ($listResults.SelectedItems.Count -lt 1) { return }
    $full = [string]$listResults.SelectedItems[0].Tag
    if ($full) { Safe-OpenFolderAndSelect $full }
})

$miCopyFull.Add_Click({
    if ($listResults.SelectedItems.Count -lt 1) { return }
    $full = [string]$listResults.SelectedItems[0].Tag
    if ($full) { Copy-ToClipboard $full }
})

$miCopyFolder.Add_Click({
    if ($listResults.SelectedItems.Count -lt 1) { return }
    $full = [string]$listResults.SelectedItems[0].Tag
    if ($full) { Copy-ToClipboard (Split-Path $full -Parent) }
})

$form.Add_FormClosing({
    try { Stop-Search } catch {}
    try { Persist-UiSettings } catch {}
    DLog "FORM CLOSING"
})

# -------------------------
# Show UI
# -------------------------
DLog "Showing UI"
[void]$form.ShowDialog()
DLog "SCRIPT END"
