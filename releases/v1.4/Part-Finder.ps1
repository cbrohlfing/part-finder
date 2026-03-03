# ==========================================================
# Part Finder (Multi-folder Search)
# Version: v1.4 (Layout fix + sortable columns)
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
$script:Version = "v1.4"
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
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
# Add a little breathing room so controls don't touch the window edge
$form.Padding = New-Object System.Windows.Forms.Padding(8)

# Main split: Left (Folders/Metadata) | Right (Search/Results)
$mainSplit = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.Orientation = 'Vertical'
$mainSplit.SplitterWidth = 6
$mainSplit.SplitterDistance = 420
$mainSplit.Panel1MinSize = 300
$mainSplit.Panel2MinSize = 450

# Left split: Top (Folders) / Bottom (Metadata)
$leftSplit = New-Object System.Windows.Forms.SplitContainer
$leftSplit.Dock = 'Fill'
$leftSplit.Orientation = 'Horizontal'
$leftSplit.SplitterWidth = 6
$leftSplit.SplitterDistance = 520
$leftSplit.Panel1MinSize = 260
$leftSplit.Panel2MinSize = 140

# -------------------------
# Left - Folders group (table layout so it resizes cleanly)
# -------------------------
$grpFolders = New-Object System.Windows.Forms.GroupBox
$grpFolders.Text = "Folders to Search"
$grpFolders.Dock = 'Fill'
$grpFolders.Padding = New-Object System.Windows.Forms.Padding(10)

$tblFolders = New-Object System.Windows.Forms.TableLayoutPanel
$tblFolders.Dock = 'Fill'
$tblFolders.ColumnCount = 1
$tblFolders.RowCount = 2
$tblFolders.Padding = New-Object System.Windows.Forms.Padding(0)
$tblFolders.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$tblFolders.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

$listFolders = New-Object System.Windows.Forms.ListBox
$listFolders.Dock = 'Fill'
$listFolders.HorizontalScrollbar = $true

# Bottom controls panel (table layout so wrapping labels and button wrapping never clip)
$pnlFolderBottom = New-Object System.Windows.Forms.Panel
$pnlFolderBottom.Dock = 'Fill'
$pnlFolderBottom.AutoScroll = $true

$tblFolderBottom = New-Object System.Windows.Forms.TableLayoutPanel
$tblFolderBottom.Dock = 'Top'
$tblFolderBottom.AutoSize = $true
$tblFolderBottom.AutoSizeMode = 'GrowAndShrink'
$tblFolderBottom.ColumnCount = 1
$tblFolderBottom.RowCount = 4
$tblFolderBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null  # label
$tblFolderBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null  # textbox
$tblFolderBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null  # buttons
$tblFolderBottom.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null  # tip

$lblAdd = New-Object System.Windows.Forms.Label
$lblAdd.AutoSize = $true
$lblAdd.Text = "Type/paste a folder path (UNC or mapped), or Browse:"
$lblAdd.MaximumSize = New-Object System.Drawing.Size(1000, 0)   # adjusted on resize for word-wrap

$txtAddFolder = New-Object System.Windows.Forms.TextBox
$txtAddFolder.Dock = 'Top'
$txtAddFolder.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$txtAddFolder.Height = 22

$flowFolderBtns = New-Object System.Windows.Forms.FlowLayoutPanel
$flowFolderBtns.FlowDirection = 'LeftToRight'
$flowFolderBtns.WrapContents = $true   # wrap buttons when pane is narrow
$flowFolderBtns.AutoSize = $true
$flowFolderBtns.AutoSizeMode = 'GrowAndShrink'
$flowFolderBtns.Dock = 'Top'
$flowFolderBtns.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Width = 90

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add"
$btnAdd.Width = 70

$btnRemoveFolder = New-Object System.Windows.Forms.Button
$btnRemoveFolder.Text = "Remove Selected"
$btnRemoveFolder.Width = 130

$btnClearFolders = New-Object System.Windows.Forms.Button
$btnClearFolders.Text = "Clear All"
$btnClearFolders.Width = 80

$flowFolderBtns.Controls.AddRange(@($btnBrowse, $btnAdd, $btnRemoveFolder, $btnClearFolders))

$lblFolderHint = New-Object System.Windows.Forms.Label
$lblFolderHint.AutoSize = $true
$lblFolderHint.MaximumSize = New-Object System.Drawing.Size(1000, 0)  # adjusted on resize for word-wrap
$lblFolderHint.MinimumSize = New-Object System.Drawing.Size(0, 34)
$lblFolderHint.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$lblFolderHint.Text = "Tip: Add multiple folders. Search runs across ALL folders. UNC paths like \\server\share\folder are supported."

$tblFolderBottom.Controls.Add($lblAdd, 0, 0)
$tblFolderBottom.Controls.Add($txtAddFolder, 0, 1)
$tblFolderBottom.Controls.Add($flowFolderBtns, 0, 2)
$tblFolderBottom.Controls.Add($lblFolderHint, 0, 3)

$pnlFolderBottom.Controls.Add($tblFolderBottom)

# Keep the "Type/paste..." label and Tip label word-wrapped as the left pane is resized
function Update-FolderBottomWrap {
    try {
        $w = [Math]::Max(120, $grpFolders.ClientSize.Width - 24)
        $lblAdd.MaximumSize = New-Object System.Drawing.Size($w, 0)
        $lblFolderHint.MaximumSize = New-Object System.Drawing.Size($w, 0)

        # Ensure the bottom row in the folders table is tall enough so the Tip isn't clipped.
        $needed =
        $lblAdd.PreferredSize.Height +
        $txtAddFolder.Height +
        $flowFolderBtns.PreferredSize.Height +
        $lblFolderHint.PreferredSize.Height +
        28

        $minH = 165
        $maxH = 260
        $h = [Math]::Max($minH, [Math]::Min($maxH, $needed))

        if ($tblFolders.RowStyles.Count -ge 2) {
            $tblFolders.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Absolute
            $tblFolders.RowStyles[1].Height = $h
        }
    } catch {}
}
Update-FolderBottomWrap
$grpFolders.Add_Resize({ Update-FolderBottomWrap })


# Add to folders table
$tblFolders.Controls.Add($listFolders, 0, 0)
$tblFolders.Controls.Add($pnlFolderBottom, 0, 1)

$grpFolders.Controls.Add($tblFolders)

# -------------------------
# Left - Metadata group
# -------------------------
$grpMeta = New-Object System.Windows.Forms.GroupBox
$grpMeta.Text = "Selected File Metadata"
$grpMeta.Dock = 'Fill'
$grpMeta.Padding = New-Object System.Windows.Forms.Padding(10)

$tblMeta = New-Object System.Windows.Forms.TableLayoutPanel
$tblMeta.Dock = 'Fill'
$tblMeta.ColumnCount = 2
$tblMeta.RowCount = 6
$tblMeta.Padding = New-Object System.Windows.Forms.Padding(8)
$tblMeta.AutoSize = $false
$tblMeta.AutoScroll = $true
$tblMeta.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 80))) | Out-Null
$tblMeta.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

function New-MetaRow([string]$labelText) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $labelText
    $l.AutoSize = $true
    $l.Anchor = 'Left'
    $v = New-Object System.Windows.Forms.Label
    $v.Text = ""
    $v.AutoEllipsis = $true
    $v.Dock = 'Fill'
    $v.Padding = New-Object System.Windows.Forms.Padding(2)
    $v.BorderStyle = 'Fixed3D'
    return @($l, $v)
}

$metaNameRow = New-MetaRow "Name"
$metaPathRow = New-MetaRow "Path"
$metaTypeRow = New-MetaRow "Type"
$metaSizeRow = New-MetaRow "Size"
$metaModRow = New-MetaRow "Modified"
$metaCreRow = New-MetaRow "Created"

$lblMetaName = $metaNameRow[0]; $valMetaName = $metaNameRow[1]
$lblMetaPath = $metaPathRow[0]; $valMetaPath = $metaPathRow[1]
$lblMetaType = $metaTypeRow[0]; $valMetaType = $metaTypeRow[1]
$lblMetaSize = $metaSizeRow[0]; $valMetaSize = $metaSizeRow[1]
$lblMetaMod = $metaModRow[0]; $valMetaMod = $metaModRow[1]
$lblMetaCre = $metaCreRow[0]; $valMetaCre = $metaCreRow[1]

$tblMeta.Controls.Add($lblMetaName, 0, 0); $tblMeta.Controls.Add($valMetaName, 1, 0)
$tblMeta.Controls.Add($lblMetaPath, 0, 1); $tblMeta.Controls.Add($valMetaPath, 1, 1)
$tblMeta.Controls.Add($lblMetaType, 0, 2); $tblMeta.Controls.Add($valMetaType, 1, 2)
$tblMeta.Controls.Add($lblMetaSize, 0, 3); $tblMeta.Controls.Add($valMetaSize, 1, 3)
$tblMeta.Controls.Add($lblMetaMod, 0, 4); $tblMeta.Controls.Add($valMetaMod, 1, 4)
$tblMeta.Controls.Add($lblMetaCre, 0, 5); $tblMeta.Controls.Add($valMetaCre, 1, 5)

$tblMeta.RowStyles.Clear()
for ($i = 0; $i -lt 6; $i++) {
    $tblMeta.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26))) | Out-Null
}

$grpMeta.Controls.Add($tblMeta)

# -------------------------
# Right group: search/results (table layout so it doesn't cut off when narrow)
# -------------------------
$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = "Search"
$grpSearch.Dock = 'Fill'
$grpSearch.Padding = New-Object System.Windows.Forms.Padding(10)

$tblSearch = New-Object System.Windows.Forms.TableLayoutPanel
$tblSearch.Dock = 'Fill'
$tblSearch.ColumnCount = 1
$tblSearch.RowCount = 3
$tblSearch.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$tblSearch.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$tblSearch.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

# Top controls grid
$tblTop = New-Object System.Windows.Forms.TableLayoutPanel
$tblTop.Dock = 'Top'
$tblTop.AutoSize = $true
$tblTop.ColumnCount = 4
$tblTop.RowCount = 4
$tblTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null  # labels
$tblTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null # stretch
$tblTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$tblTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

# Row 1: Query + Mode
$lblQuery = New-Object System.Windows.Forms.Label
$lblQuery.Text = "Search:"
$lblQuery.AutoSize = $true
$lblQuery.Anchor = 'Left'

$txtQuery = New-Object System.Windows.Forms.TextBox
$txtQuery.Dock = 'Fill'
$txtQuery.Text = [string]$script:cfg.lastQuery

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Text = "Mode:"
$lblMode.AutoSize = $true
$lblMode.Anchor = 'Left'

$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.Width = 160
$cmbMode.DropDownStyle = "DropDownList"
$cmbMode.Anchor = 'Left'
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

$tblTop.Controls.Add($lblQuery, 0, 0)
$tblTop.Controls.Add($txtQuery, 1, 0)
$tblTop.Controls.Add($lblMode, 2, 0)
$tblTop.Controls.Add($cmbMode, 3, 0)

# Row 2: Include subfolders + buttons
$chkSub = New-Object System.Windows.Forms.CheckBox
$chkSub.Text = "Include subfolders"
$chkSub.AutoSize = $true
$chkSub.Checked = [bool]$script:cfg.includeSubfolders
$chkSub.Anchor = 'Left'

$flowSearchBtns = New-Object System.Windows.Forms.FlowLayoutPanel
$flowSearchBtns.FlowDirection = 'LeftToRight'
$flowSearchBtns.WrapContents = $false
$flowSearchBtns.AutoSize = $true
$flowSearchBtns.Anchor = 'Right'
$flowSearchBtns.Dock = 'Fill'

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Search"
$btnSearch.Width = 80

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "Stop"
$btnStop.Width = 70
$btnStop.Enabled = $false

$flowSearchBtns.Controls.AddRange(@($btnSearch, $btnStop))

$tblTop.Controls.Add($chkSub, 0, 1)
$tblTop.SetColumnSpan($chkSub, 2)
$tblTop.Controls.Add($flowSearchBtns, 2, 1)
$tblTop.SetColumnSpan($flowSearchBtns, 2)

# Row 3: Max results + Ignore
$lblMax = New-Object System.Windows.Forms.Label
$lblMax.Text = "Max results:"
$lblMax.AutoSize = $true
$lblMax.Anchor = 'Left'

$numMax = New-Object System.Windows.Forms.NumericUpDown
$numMax.Width = 70
$numMax.Minimum = 1
$numMax.Maximum = 10000
$numMax.Value = [decimal][int]$script:cfg.maxResults
$numMax.Anchor = 'Left'

$lblIgnore = New-Object System.Windows.Forms.Label
$lblIgnore.Text = "Ignore:"
$lblIgnore.AutoSize = $true
$lblIgnore.Anchor = 'Left'

$txtIgnore = New-Object System.Windows.Forms.TextBox
$txtIgnore.Dock = 'Fill'
$txtIgnore.Text = [string]$script:cfg.ignorePatterns

$tblTop.Controls.Add($lblMax, 0, 2)
$tblTop.Controls.Add($numMax, 1, 2)
$tblTop.Controls.Add($lblIgnore, 2, 2)
$tblTop.Controls.Add($txtIgnore, 3, 2)

# Row 4: ignore default label (spans)
$lblIgnoreDefault = New-Object System.Windows.Forms.Label
$lblIgnoreDefault.AutoSize = $true
$lblIgnoreDefault.ForeColor = [System.Drawing.Color]::DimGray
$lblIgnoreDefault.Text = "Default: *.log;*.bak"
$lblIgnoreDefault.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)

$tblTop.Controls.Add($lblIgnoreDefault, 2, 3)
$tblTop.SetColumnSpan($lblIgnoreDefault, 2)

# Progress/status panel
$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Dock = 'Top'
$pnlStatus.AutoSize = $true
$pnlStatus.AutoSizeMode = 'GrowAndShrink'
$pnlStatus.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)

$prg = New-Object System.Windows.Forms.ProgressBar
$prg.Dock = 'Top'
$prg.Height = 14
$prg.Style = 'Blocks'
$prg.Visible = $false

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Dock = 'Top'
$lblStatus.Height = 18
$lblStatus.Text = "Ready."

$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Dock = 'Top'
$lblCurrent.Height = 18
$lblCurrent.Text = ""

$pnlStatus.Controls.Add($lblCurrent)
$pnlStatus.Controls.Add($lblStatus)
$pnlStatus.Controls.Add($prg)

# Results list (fill)
$listResults = New-Object System.Windows.Forms.ListView
$listResults.Dock = 'Fill'
$listResults.View = "Details"
$listResults.FullRowSelect = $true
$listResults.GridLines = $true
$listResults.Font = New-Object System.Drawing.Font("Consolas", 10)
$listResults.HideSelection = $false

$null = $listResults.Columns.Add("Name", 220)
$null = $listResults.Columns.Add("Folder", 300)
$null = $listResults.Columns.Add("Modified", 120)

# Context menu
$resultsMenu = New-Object System.Windows.Forms.ContextMenuStrip
$miOpen = $resultsMenu.Items.Add("Open")
$miOpenFolder = $resultsMenu.Items.Add("Open Containing Folder")
$resultsMenu.Items.Add("-") | Out-Null
$miCopyFull = $resultsMenu.Items.Add("Copy Full Path")
$miCopyFolder = $resultsMenu.Items.Add("Copy Folder Path")
$listResults.ContextMenuStrip = $resultsMenu

# Assemble right side
$tblSearch.Controls.Add($tblTop, 0, 0)
$tblSearch.Controls.Add($pnlStatus, 0, 1)
$tblSearch.Controls.Add($listResults, 0, 2)
$grpSearch.Controls.Add($tblSearch)

# Assemble split panes
$leftSplit.Panel1.Controls.Add($grpFolders)
$leftSplit.Panel2.Controls.Add($grpMeta)

$mainSplit.Panel1.Controls.Add($leftSplit)
$mainSplit.Panel2.Controls.Add($grpSearch)

$form.Controls.Add($mainSplit)

# Metadata updater (called when selection changes)
function Set-Metadata([string]$fullPath) {
    Ui $form {
        if ([string]::IsNullOrWhiteSpace($fullPath) -or -not (Test-Path -LiteralPath $fullPath)) {
            $valMetaName.Text = ""
            $valMetaPath.Text = ""
            $valMetaType.Text = ""
            $valMetaSize.Text = ""
            $valMetaMod.Text = ""
            $valMetaCre.Text = ""
            return
        }

        try {
            $fi = New-Object System.IO.FileInfo($fullPath)
            $valMetaName.Text = $fi.Name
            $valMetaPath.Text = $fi.FullName
            $valMetaType.Text = $fi.Extension
            $valMetaSize.Text = ("{0:n0} KB" -f [math]::Ceiling($fi.Length / 1KB))
            $valMetaMod.Text = $fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
            $valMetaCre.Text = $fi.CreationTime.ToString("yyyy-MM-dd HH:mm")
        } catch {
            $valMetaName.Text = ""
            $valMetaPath.Text = $fullPath
            $valMetaType.Text = ""
            $valMetaSize.Text = ""
            $valMetaMod.Text = ""
            $valMetaCre.Text = ""
        }
    }
}

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

    # Clear metadata
    try { Set-Metadata "" } catch {}

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

# Update metadata panel when selection changes
$listResults.Add_SelectedIndexChanged({
        try {
            if ($listResults.SelectedItems.Count -lt 1) { Set-Metadata ""; return }
            $tag = $listResults.SelectedItems[0].Tag
            $full = if ($tag -and $tag.FullPath) { [string]$tag.FullPath } else { "" }
            Set-Metadata $full
        } catch { }
    })

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





