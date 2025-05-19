Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import ExtractIconEx function
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Shell32 {
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern int ExtractIconEx(string szFileName, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);
}
"@

function Get-ShellIcon {
    param (
        [string]$dll = "$env:SystemRoot\System32\shell32.dll",
        [int]$index
    )
    $iconSmall = New-Object IntPtr[] 1
    [Shell32]::ExtractIconEx($dll, $index, $null, $iconSmall, 1) | Out-Null
    if ($iconSmall[0] -ne [IntPtr]::Zero) {
        $icon = [System.Drawing.Icon]::FromHandle($iconSmall[0])
        return $icon
    }
    return $null
}

function Get-ShellIconBitmap {
    param (
        [int]$index
    )
    $icon = Get-ShellIcon -index $index
    return $icon
}

function Set-LastColumnAutoSize {
    if ($listView.Columns.Count -gt 0) {
        $listView.Columns[$listView.Columns.Count - 1].Width = -2
    }
}

# File paths
$hostsFolderPath = "hosts"
$categoryCsvPath = Join-Path $PSScriptRoot "categories.csv"
$preferencesPath = Join-Path $PSScriptRoot "preferences.json"
$themePath = Join-Path $PSScriptRoot "themes"


# Ensure paths exist
if (-not (Test-Path $themePath)) {
    New-Item -ItemType Directory -Path $themePath | Out-Null
}

# Default theme JSON
if (-not (Test-Path $themePath)) {
    New-Item -ItemType Directory -Path $themePath | Out-Null
}

$defaultThemePath = Join-Path $themePath "default.json"
if (-not (Test-Path $defaultThemePath)) {
@'
{
  "Light": {
    "FormBack": "#FFFFFF",
    "FormFore": "#000000",
    "ListBack": "#FFFFFF",
    "ListFore": "#000000",
    "SearchBack": "#FFFFFF",
    "SearchFore": "#000000",
    "ButtonBack": "#E0E0E0",
    "RowFore": "#000000",
    "RowBack1": "#FFFFFF",
    "RowBack2": "#F5F5F5",
    "TextColor": "#000000",
    "HeaderBack": "#DDDDDD",
    "HeaderFore": "#000000",
    "SelectedBack": "#3399FF",
    "SelectedFore": "#FFFFFF",
    "Border": "#AAAAAA",
    "FontName": "Consolas",
    "FontSize": 10
  },
  "Dark": {
    "FormBack": "#1E1E1E",
    "FormFore": "#CCCCCC",
    "ListBack": "#2D2D2D",
    "ListFore": "#CCCCCC",
    "SearchBack": "#333333",
    "SearchFore": "#FFFFFF",
    "ButtonBack": "#444444",
    "RowFore": "#FFFFFF",
    "RowBack1": "#3A3A3A",
    "RowBack2": "#2A2A2A",
    "TextColor": "#FFFFFF",
    "HeaderBack": "#FFFFFF",
    "HeaderFore": "#000000",
    "SelectedBack": "#3399FF",
    "SelectedFore": "#FFFFFF",
    "Border": "#000000",
    "FontName": "Segoe UI",
    "FontSize": 10
  }
}
'@ | Set-Content -Encoding UTF8 $defaultThemePath
}

# Default preferences.json
if (-not (Test-Path $preferencesPath)) {
@'
{
    "DarkMode": false,
    "WindowWidth": 700,
    "WindowHeight": 800,
    "Theme": "default.json",
    "Hosts": "hosts.csv"
}
'@ | Set-Content -Encoding UTF8 $preferencesPath
}

# Default categories.csv
if (-not (Test-Path $categoryCsvPath)) {
@"
Category,Description
Uncategorized,Default fallback category
"@ | Set-Content -Encoding UTF8 $categoryCsvPath
}

# Load preferences
function Get-Preferences {
    if (Test-Path $preferencesPath) {
        return Get-Content $preferencesPath | ConvertFrom-Json
    }
}

# Save preferences
function Save-Preferences($prefs) {
    $prefs | ConvertTo-Json | Set-Content $preferencesPath
}

$prefs = Get-Preferences
$hostsCsvFile = Join-Path $hostsFolderPath $prefs.Hosts

# Load category list
function Get-Categories {
    if (Test-Path $categoryCsvPath) {
        try {
            $data = Import-Csv $categoryCsvPath | Where-Object { $_.Category -and $_.Category.Trim() -ne "" }
            if ($data.Count -gt 0) { return $data }
        } catch {
            Write-Warning "Failed to load categories.csv"
        }
    }
    return @([PSCustomObject]@{ Category = "Uncategorized"; Description = "Default" })
}

#Create hosts template file
function New-TemplateHostsFile {
    param (
        [string]$hostsDefaultCsvPath
    )
    if (-not (Test-Path $hostsFolderPath)) {
        New-Item -ItemType Directory -Path $hostsFolderPath | Out-Null
    }
    $template = @"
Name,Hostname,Username,Category
"@
    $template | Set-Content -Path $hostsDefaultCsvPath -Encoding UTF8
    Write-Host "Created template hosts.csv file at $hostsDefaultCsvPath"
}

# Load hosts list
function Get-Hosts {  
    if (Test-Path $hostsCsvFile) {
        try {
            Import-Csv $hostsCsvFile | ForEach-Object {
                [PSCustomObject]@{
                    HostIndex = $_.HostIndex
                    Name     = $_.Name
                    Hostname = $_.Hostname
                    Username = $_.Username
                    Category = $_.Category
                }
            }
        } catch {
            Write-Warning "Failed to load hosts.csv. Ensure the file is properly formatted."
            return @()
        }
    } else {
        Write-Warning "hosts.csv not found. Creating template file."
        New-TemplateHostsFile $hostsCsvFile
        Get-Hosts
    }
}

function Connect-Host {
    param (
        [string]$hostName
    )
    $hostConnect = $allHosts | Where-Object { $_.Name -eq $hostName }
    if ($hostConnect) {
        $ssh = "ssh $($hostConnect.Username)@$($hostConnect.Hostname)"
        $serverName = $hostConnect.Name
        Start-Process -NoNewWindow -FilePath "wt.exe" -ArgumentList @("new-tab", "--title", "`"$serverName`"", $ssh)
    } else {
        Write-Warning "Host '$hostName' not found."
    }
}

# Global host data
$allHosts = Get-Hosts
$script:displayedHosts = @()
$script:sortColumn = 0
$script:sortOrder = 'Ascending'

# GUI setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "TerminalMagic"
$form.Size = New-Object System.Drawing.Size($prefs.WindowWidth, $prefs.WindowHeight)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Icon = Get-ShellIcon -Index 111
$form.KeyPreview = $true

$form.Add_KeyDown({
    param($sender, $e)

    if ($e.Control -and $e.KeyCode -eq "A") {
        Show-HostForm
        $e.Handled = $true
    }
    elseif ($e.Control -and $e.KeyCode -eq "C") {
        $selectedIndex = $listview.SelectedIndices
        if ($selectedIndex.Count -gt 0) {
            $row = $script:displayedHosts[$selectedIndex[0]]
            if ($row.Type -eq 'Host' -and $row.Host) {
                Connect-Host -hostName $row.Host.Hostname
            }
        }
        $e.Handled = $true
    }
    elseif ($e.Control -and $e.KeyCode -eq "E") {
        $selectedIndex = $listview.SelectedIndices
        if ($selectedIndex.Count -gt 0) {
            $row = $script:displayedHosts[$selectedIndex[0]]
            if ($row.Type -eq 'Host' -and $row.Host) {
                Show-HostForm -hostObj $row.Host
            }
        }
        $e.Handled = $true
    }
    elseif ($e.Control -and $e.KeyCode -eq "D") {
        $selectedIndex = $listview.SelectedIndices
        if ($selectedIndex.Count -gt 0) {
            $row = $script:displayedHosts[$selectedIndex[0]]
            if ($row.Type -eq 'Host' -and $row.Host) {
                $hostObj = $row.Host
                $confirm = [System.Windows.Forms.MessageBox]::Show("Delete host '$($hostObj.Name)'?", "Confirm", "YesNo")
                if ($confirm -eq "Yes") {
                    $allHosts = $allHosts | Where-Object { $_.HostIndex -ne $hostObj.HostIndex }
                    $allHosts | Export-Csv $hostsCsvFile -NoTypeInformation
                    $script:allHosts = Get-Hosts
                    Update-List
                }
            }
        }
        $e.Handled = $true
    }
})


# Dark mode checkbox for ToolStrip
$darkModeBox = New-Object System.Windows.Forms.CheckBox
$darkModeBox.Text = "Dark Mode"
$darkModeBox.Checked = $prefs.DarkMode
$darkModeBox.Margin = '10,0,0,0'

$darkModeHost = New-Object System.Windows.Forms.ToolStripControlHost($darkModeBox)

$darkModeBox.Add_CheckedChanged({
    $prefs.DarkMode = $darkModeBox.Checked
    Save-Preferences $prefs
    Set-Theme
    Update-List
})

# Search box
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Dock = "Top"
$searchBox.Font = 'Segoe UI,10'
$searchBox.Height = 30
$searchBox.Margin = '10,10,10,10'

$placeholder = "Enter search term..."
$searchBox.Text = $placeholder

$searchBox.Add_GotFocus({
    if ($searchBox.Text -eq $placeholder) {
        $searchBox.Select(0, 0)  # Optional: put cursor at beginning
    }
})

$searchBox.Add_KeyDown({
    if ($searchBox.Text -eq $placeholder) {
        $searchBox.Clear()
    }
})

$searchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
        $searchBox.Text = $placeholder
    }
})
# Button strip (top)
$buttonStrip = New-Object System.Windows.Forms.ToolStrip
$buttonStrip.Dock = "Top"

# Use Windows system icons
$connectIcon = Get-ShellIconBitmap -index 92
$addIcon    = Get-ShellIconBitmap -index 79
$editIcon   = Get-ShellIconBitmap -index 269
$deleteIcon = Get-ShellIconBitmap -index 131
$themeIcon = Get-ShellIconBitmap -index 303
$aboutIcon = Get-ShellIconBitmap -index 23

$connectButton = New-Object System.Windows.Forms.ToolStripButton("", $connectIcon)
$addButton = New-Object System.Windows.Forms.ToolStripButton("", $addIcon)
$editButton = New-Object System.Windows.Forms.ToolStripButton("", $editIcon)
$deleteButton = New-Object System.Windows.Forms.ToolStripButton("", $deleteIcon)
$themeButton = New-Object System.Windows.Forms.ToolStripButton("", $themeIcon)
$aboutButton = New-Object System.Windows.Forms.ToolStripButton("", $aboutIcon)

$connectButton.ToolTipText = "Connect to selected host - Ctrl+C"
$addButton.ToolTipText    = "Add a new host - Ctrl+A"
$editButton.ToolTipText   = "Edit the selected host - Ctrl+E"
$deleteButton.ToolTipText = "Delete the selected host - Ctrl+D"
$themeButton.ToolTipText = "Change theme"
$aboutButton.ToolTipText = "Support"

$buttonStrip.Items.AddRange(@($connectButton, $addButton, $editButton, $deleteButton, $themeButton, $aboutButton, $darkModeHost))

$connectButton.Add_Click({
    $selectedIndex = $listview.SelectedIndices
    if ($selectedIndex.Count -gt 0) {
        $row = $script:displayedHosts[$selectedIndex[0]]
        if ($row.Type -eq 'Host' -and $row.Host) {
            Connect-Host -hostName $row.Host.Hostname
        }
    }
})

$editButton.Add_Click({
    $selectedIndex = $listview.SelectedIndices
    if ($selectedIndex.Count -gt 0) {
        $row = $script:displayedHosts[$selectedIndex[0]]
        if ($row.Type -eq 'Host' -and $row.Host) {
            Show-HostForm -hostObj $row.Host
        }
    }
})

$deleteButton.Add_Click({
    $selectedIndex = $listview.SelectedIndices
    if ($selectedIndex.Count -gt 0) {
        $row = $script:displayedHosts[$selectedIndex[0]]
        if ($row.Type -eq 'Host' -and $row.Host) {
            $hostObj = $row.Host
            $confirm = [System.Windows.Forms.MessageBox]::Show("Delete host '$($hostObj.Name)'?", "Confirm", "YesNo")
            if ($confirm -eq "Yes") {
                $allHosts = $allHosts | Where-Object { $_.HostIndex -ne $hostObj.HostIndex }
                $allHosts | Export-Csv $hostsCsvFile -NoTypeInformation
                $script:allHosts = Get-Hosts
                Update-List
            }
        }
    }
})

function Show-ThemePicker {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = $themePath
    $dialog.Filter = "JSON Theme Files (*.json)|*.json"
    $dialog.Title = "Select a Theme File"
    $dialog.Multiselect = $false
    $dialog.CheckFileExists = $true
    $dialog.RestoreDirectory = $true
    $dialog.ShowHelp = $false

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $fileName = [System.IO.Path]::GetFileName($dialog.FileName)
        $prefs.Theme = $fileName
        Save-Preferences $prefs
        Set-Theme
        Update-List
    }
}

$themeButton.Add_Click({ Show-ThemePicker })

# Create the About form
$aboutForm = New-Object System.Windows.Forms.Form
$aboutForm.Text = "About"
$aboutForm.Size = New-Object System.Drawing.Size(300, 180)
$aboutForm.StartPosition = "CenterParent"
$aboutForm.Icon = Get-ShellIcon -Index 23

# Add a label with company information
$aboutLabel = New-Object System.Windows.Forms.Label
$aboutLabel.Text = "TerminalMagic v1.0`nFor support, email:"
$aboutLabel.AutoSize = $true
$aboutLabel.Location = New-Object System.Drawing.Point(20, 10)
$aboutForm.Controls.Add($aboutLabel)

# Add a LinkLabel for the email
$aboutLinkLabel = New-Object System.Windows.Forms.LinkLabel
$aboutLinkLabel.Text = "support@tek.tools"
$aboutLinkLabel.Location = New-Object System.Drawing.Point(20, 40)
$aboutLinkLabel.AutoSize = $true
$aboutLinkLabel.Add_LinkClicked({
    [System.Diagnostics.Process]::Start("mailto:support@tek.tools")
})
$aboutForm.Controls.Add($aboutLinkLabel)

# Add a label with company information
$aboutLabel2 = New-Object System.Windows.Forms.Label
$aboutLabel2.Text = "Or visit:"
$aboutLabel2.AutoSize = $true
$aboutLabel2.Location = New-Object System.Drawing.Point(20, 55)
$aboutForm.Controls.Add($aboutLabel2)

# Add a LinkLabel for the website
$aboutLinkLabel = New-Object System.Windows.Forms.LinkLabel
$aboutLinkLabel.Text = "https://tek.tools"
$aboutLinkLabel.Location = New-Object System.Drawing.Point(20, 70)
$aboutLinkLabel.AutoSize = $true
$aboutLinkLabel.Add_LinkClicked({
    [System.Diagnostics.Process]::Start("https://tek.tools")
})
$aboutForm.Controls.Add($aboutLinkLabel)

# Add an OK button to close the dialog
$aboutOkButton = New-Object System.Windows.Forms.Button
$aboutOkButton.Text = "OK"
$aboutOkButton.Location = New-Object System.Drawing.Point(110, 100)
$aboutOkButton.Add_Click({ $aboutForm.Close() })
$aboutForm.Controls.Add($aboutOkButton)

$aboutButton.Add_Click({
    $aboutForm.ShowDialog()
})

# ListView setup
$listview = New-Object System.Windows.Forms.ListView
$listview.Dock = "Fill"
$listview.View = 'Details'
$listview.FullRowSelect = $true
$listview.GridLines = $true
$listview.Font = 'Segoe UI, 9'
$listview.MultiSelect = $false
$listview.HideSelection = $true
$listview.HoverSelection = $false
$listview.OwnerDraw = $true

# Columns
$columns = @("Name", "Hostname", "Username", "Category")
$widths = @(150, 180, 140, 140)
for ($i = 0; $i -lt $columns.Count; $i++) {
    $col = New-Object System.Windows.Forms.ColumnHeader
    $col.Text = $columns[$i]
    $col.Width = $widths[$i]
    [void]$listview.Columns.Add($col)
}

# Properly typed event handlers to preserve default row rendering
$listview.add_DrawItem({
    param([object]$sender, [System.Windows.Forms.DrawListViewItemEventArgs]$e)
    $e.DrawDefault = $true
})

$listview.add_DrawSubItem({
    param([object]$sender, [System.Windows.Forms.DrawListViewSubItemEventArgs]$e)
    $e.DrawDefault = $true
})

$listview.add_DrawColumnHeader({
    param([object]$sender, [System.Windows.Forms.DrawListViewColumnHeaderEventArgs]$e)

    $dark = $darkModeBox.Checked
    $colors = if ($dark) { $theme.Dark } else { $theme.Light }

    $backColor = Convert-HexToColor $colors.HeaderBack
    $foreColor = Convert-HexToColor $colors.HeaderFore
    $borderColor = Convert-HexToColor $colors.Border

    $backBrush  = New-Object System.Drawing.SolidBrush($backColor)
    $foreBrush  = New-Object System.Drawing.SolidBrush($foreColor)
    $pen        = New-Object System.Drawing.Pen($borderColor)

    # Fill background
    $e.Graphics.FillRectangle($backBrush, $e.Bounds)

    # Draw column text
    $e.Graphics.DrawString($e.Header.Text, $listview.Font, $foreBrush, $e.Bounds.X + 4, $e.Bounds.Y + 2)

    # Draw simulated gridlines
    $e.Graphics.DrawLine($pen, $e.Bounds.Left, $e.Bounds.Bottom - 1, $e.Bounds.Right, $e.Bounds.Bottom - 1)  # bottom line
    $e.Graphics.DrawLine($pen, $e.Bounds.Right - 1, $e.Bounds.Top, $e.Bounds.Right - 1, $e.Bounds.Bottom)    # right line

    $backBrush.Dispose()
    $foreBrush.Dispose()
    $pen.Dispose()
})

Set-LastColumnAutoSize

# Create the ContextMenuStrip
$contextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip

# Create "Edit" and "Delete" menu items
$connectMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$connectMenuItem.Text = "Connect"
$editMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$editMenuItem.Text = "Edit"
$deleteMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$deleteMenuItem.Text = "Delete"

# Add menu items to the ContextMenuStrip
$contextMenuStrip.Items.AddRange(@($connectMenuItem, $editMenuItem, $deleteMenuItem))

# Attach the ContextMenuStrip to the ListView
$listview.ContextMenuStrip = $contextMenuStrip

# Handle "Connect" menu item click
$connectMenuItem.Add_Click({
    $selectedIndex = $listview.SelectedIndices
    if ($selectedIndex.Count -gt 0) {
        $row = $script:displayedHosts[$selectedIndex[0]]
        if ($row.Type -eq 'Host' -and $row.Host) {
            Connect-Host -hostName $row.Host.Hostname
        }
    }
})

# Handle "Edit" menu item click
$editMenuItem.Add_Click({
    $selectedIndex = $listview.SelectedIndices
    if ($selectedIndex.Count -gt 0) {
        $row = $script:displayedHosts[$selectedIndex[0]]
        if ($row.Type -eq 'Host' -and $row.Host) {
            Show-HostForm -hostObj $row.Host
        }
    }
})

# Handle "Delete" menu item click
$deleteMenuItem.Add_Click({
    $selectedIndex = $listview.SelectedIndices
    if ($selectedIndex.Count -gt 0) {
        $row = $script:displayedHosts[$selectedIndex[0]]
        if ($row.Type -eq 'Host' -and $row.Host) {
            $hostObj = $row.Host
            $confirm = [System.Windows.Forms.MessageBox]::Show("Delete host '$($hostObj.Name)'?", "Confirm", "YesNo")
            if ($confirm -eq "Yes") {
                $allHosts = $allHosts | Where-Object { $_.HostIndex -ne $hostObj.HostIndex }
                $allHosts | Export-Csv $hostsCsvFile -NoTypeInformation
                $script:allHosts = Get-Hosts
                Update-List
            }
        }
    }
})

function Convert-HexToColor {
    param([string]$hex)
    return [System.Drawing.ColorTranslator]::FromHtml($hex)
}

function Set-ControlProperties {
    param (
        [System.Windows.Forms.Control]$control,
        [System.Drawing.Color]$backColor = $null,
        [System.Drawing.Color]$foreColor = $null,
        [System.Drawing.Font]$font = $null
    )
    if ($null -ne $backColor) {
        $control.BackColor = $backColor
    }
    if ($null -ne $foreColor) {
        $control.ForeColor = $foreColor
    }
    if ($null -ne $font) {
        $control.Font = $font
    }
}

function Set-Theme {
    $themeFile = Join-Path $themePath $prefs.Theme
    if (!(Test-Path $themeFile)) {
        $themeFile = Join-Path $themePath "default.json"
    }
    $global:theme = Get-Content $themeFile | ConvertFrom-Json
    
    $dark = $darkModeBox.Checked
    $colors = if ($dark) { $theme.Dark } else { $theme.Light }

    $fontName = $colors.FontName
    $fontSize = [float]$colors.FontSize
    $customFont = New-Object System.Drawing.Font($fontName, $fontSize)

    # Ensure valid colors are passed
    $formBackColor = if ($colors.FormBack) { Convert-HexToColor $colors.FormBack } else { $null }
    $formForeColor = if ($colors.FormFore) { Convert-HexToColor $colors.FormFore } else { $null }
    $rowBackColor = if ($colors.RowBack1) { Convert-HexToColor $colors.RowBack1 } else { $null }
    $rowForeColor = if ($colors.RowFore) { Convert-HexToColor $colors.RowFore } else { $null }
    $searchBackColor = if ($colors.SearchBack) { Convert-HexToColor $colors.SearchBack } else { $null }
    $searchForeColor = if ($colors.SearchFore) { Convert-HexToColor $colors.SearchFore } else { $null }
    $textColor = if ($colors.TextColor) { Convert-HexToColor $colors.TextColor } else { $null }

    # Apply colors and fonts to controls
    Set-ControlProperties -control $form -backColor $formBackColor -foreColor $formForeColor -font $customFont
    Set-ControlProperties -control $listview -backColor $rowBackColor -foreColor $rowForeColor -font $customFont
    Set-ControlProperties -control $searchBox -backColor $searchBackColor -foreColor $searchForeColor -font $customFont
    Set-ControlProperties -control $darkModeBox -foreColor $textColor -backColor $rowBackColor -font $customFont

    # Customize the ToolStrip
    $buttonStrip.BackColor = $rowBackColor
    foreach ($btn in $buttonStrip.Items) {
        if ($btn -is [System.Windows.Forms.ToolStripButton]) {
            $btn.ForeColor = $textColor
            $btn.Font = $customFont
        }
    }

    Update-List
}

function Update-List {
    $listview.Items.Clear()
    $script:displayedHosts = @()
    $sortedHosts = $allHosts | Sort-Object -Property $columns[$script:sortColumn] -Descending:($script:sortOrder -eq 'Descending')
    $filtered = $sortedHosts | Where-Object {
        $text = $searchBox.Text.ToLower()
        $text -eq "" -or
        $text -eq $placeholder -or
        $_.Name.ToLower().Contains($text) -or
        $_.Hostname.ToLower().Contains($text) -or
        $_.Username.ToLower().Contains($text) -or
        ($_.Category -and $_.Category.ToLower().Contains($text))
    }

    $dark = $darkModeBox.Checked
    $colors = if ($dark) { $theme.Dark } else { $theme.Light }
    $foreColor = Convert-HexToColor $colors.RowFore
    $backColor1 = Convert-HexToColor $colors.RowBack1
    $backColor2 = Convert-HexToColor $colors.RowBack2

    $index = 0
    foreach ($entry in $filtered) {
        $item = New-Object System.Windows.Forms.ListViewItem($entry.Name)
        $item.UseItemStyleForSubItems = $false

        # Add subitems
        $item.SubItems.Add($entry.Hostname) | Out-Null
        $item.SubItems.Add($entry.Username) | Out-Null
        $item.SubItems.Add($entry.Category) | Out-Null

        # Alternating background
        $bgColor = if ($index % 2 -eq 0) { $backColor1 } else { $backColor2 }
        $item.BackColor = $bgColor

        # Apply fore/back color to each subitem
        foreach ($sub in $item.SubItems) {
            $sub.ForeColor = $foreColor
            $sub.BackColor = $bgColor
        }

        $listview.Items.Add($item) | Out-Null
        $script:displayedHosts += [PSCustomObject]@{ Type = 'Host'; Host = $entry }
        $index++
    }
    $listview.Refresh()
}

$listview.Add_ColumnClick({
    param($sender, $e)
    if ($script:sortColumn -eq $e.Column) {
        $script:sortOrder = if ($script:sortOrder -eq 'Ascending') { 'Descending' } else { 'Ascending' }
    } else {
        $script:sortColumn = $e.Column
        $script:sortOrder = 'Ascending'
    }
    Update-List
})

$searchBox.Add_TextChanged({ Update-List })

$form.Controls.AddRange(@($listview, $buttonStrip, $searchBox, $darkModeBox))
$form.Add_Shown({ $searchBox.Focus(); Set-Theme })
$form.Add_Resize({
    Set-LastColumnAutoSize
    $prefs.WindowWidth = $form.Width
    $prefs.WindowHeight = $form.Height
    Save-Preferences $prefs
})

$listview.Add_DoubleClick({
    $i = $listview.SelectedIndices
    if ($i.Count -gt 0) {
        $row = $script:displayedHosts[$i[0]]
        if ($row.Type -eq 'Host' -and $row.Host) {
            Connect-Host -hostName $row.Host.Name
        }
    }
})

# Add/Edit form
function Show-HostForm {
    param (
        [Parameter(Mandatory = $false)]
        [object]$hostObj
    )
    $isEdit = $null -ne $hostObj
    $formTitle = if ($isEdit) { "Edit Host - $($hostObj.Name)" } else { "Add New Host" }

    $f = New-Object System.Windows.Forms.Form -Property @{
        Text = $formTitle
        Size = '300,330'
        StartPosition = 'CenterParent'
    }

    $lbl = @{ Size = '260,20' }
    $txt = @{ Size = '260,20' }

    $nameLabel = New-Object System.Windows.Forms.Label -Property ($lbl + @{ Text = "Name:"; Location = "10,20" })
    $nameBox   = New-Object System.Windows.Forms.TextBox -Property ($txt + @{
        Location = "10,40"
        Text = if ($isEdit) { $hostObj.Name } else { "" }
    })

    $hostLabel = New-Object System.Windows.Forms.Label -Property ($lbl + @{ Text = "Hostname:"; Location = "10,70" })
    $hostBox   = New-Object System.Windows.Forms.TextBox -Property ($txt + @{
        Location = "10,90"
        Text = if ($isEdit) { $hostObj.Hostname } else { "" }
    })

    $userLabel = New-Object System.Windows.Forms.Label -Property ($lbl + @{ Text = "Username:"; Location = "10,120" })
    $userBox   = New-Object System.Windows.Forms.TextBox -Property ($txt + @{
        Location = "10,140"
        Text = if ($isEdit) { $hostObj.Username } else { "" }
    })

    $catLabel = New-Object System.Windows.Forms.Label -Property ($lbl + @{ Text = "Category:"; Location = "10,170" })
    $catBox   = New-Object System.Windows.Forms.ComboBox -Property @{ Location = "10,190"; Size = '260,20' }

    $allCats = Get-Categories
    foreach ($c in $allCats) { $catBox.Items.Add($c.Category) }

    if ($isEdit -and $hostObj.Category -and $catBox.Items.Contains($hostObj.Category)) {
        $catBox.SelectedItem = $hostObj.Category
    } elseif (-not $isEdit -and $catBox.Items.Count -gt 0) {
        $catBox.SelectedIndex = 0
    }

    $okButton = New-Object System.Windows.Forms.Button -Property @{
        Text = if ($isEdit) { "Save" } else { "Add" }
        Location = "10,230"
        Size = '260,30'
    }

    $f.AcceptButton = $okButton

    $okButton.Add_Click({
        if ($nameBox.Text -and $hostBox.Text -and $userBox.Text) {
            $updated = [PSCustomObject]@{
                HostIndex = if ($isEdit) { $hostObj.HostIndex } else {
                    if ($script:allHosts.Count -eq 0) { 0 } else {
                        ($script:allHosts | Measure-Object -Property HostIndex -Maximum).Maximum + 1
                    }
                }
                Name     = $nameBox.Text
                Hostname = $hostBox.Text
                Username = $userBox.Text
                Category = if ($catBox.SelectedItem) { $catBox.SelectedItem } else { "Uncategorized" }
            }

            if ($isEdit) {
                $allHosts = $allHosts | ForEach-Object {
                    if ($_.HostIndex -eq $hostObj.HostIndex) {
                        $updated
                    } else { $_ }
                }
            } else {
                $script:allHosts = @(Get-Hosts)
                $script:allHosts += $updated
            }

            $allHosts | Export-Csv $hostsCsvFile -NoTypeInformation
            $script:allHosts = Get-Hosts
            Update-List
            $f.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please fill in Name, Hostname, and Username.", "Missing Info")
        }
    })

    $f.Controls.AddRange(@(
        $nameLabel, $nameBox,
        $hostLabel, $hostBox,
        $userLabel, $userBox,
        $catLabel, $catBox,
        $okButton
    ))

    $f.Icon = Get-ShellIcon -Index 17
    $f.ShowDialog($form) | Out-Null
}

# Button click handlers

$addButton.Add_Click({
    Show-HostForm
})

[void]$form.ShowDialog()