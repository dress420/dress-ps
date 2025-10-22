# CCL Service Checker - Made by Dress
# Universal Version (Centered Red ASCII, Works in CMD, PowerShell, Windows Terminal)

# --- FIX: Enable UTF-8 and ANSI color in CMD-hosted PowerShell ---
if (-not $Host.UI.SupportsVirtualTerminal) {
    try { $Host.UI.SupportsVirtualTerminal = $true } catch {}
}
$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

Clear-Host

# === ASCII HEADER ===
$ascii = @"
  /$$$$$$   /$$$$$$  /$$              /$$$$$$                                 /$$                            /$$$$$$   /$$$$$$  /$$                           /$$                          
 /$$__  $$ /$$__  $$| $$             /$$__  $$                               |__/                           /$$__  $$ /$$__  $$| $$                          | $$                          
| $$  \__/| $$  \__/| $$            | $$  \__/  /$$$$$$   /$$$$$$  /$$    /$$ /$$  /$$$$$$$  /$$$$$$       | $$  \__/| $$  \ $$| $$   /$$$$$$   /$$$$$$$     | $$   /$$  /$$$$$$   /$$$$$$ 
| $$      | $$      | $$            |  $$$$$$  /$$__  $$ /$$__  $$|  $$  /$$/| $$ /$$_____/ /$$__  $$      | $$      | $$$$$$$$| $$  /$$__  $$ /$$_____/     | $$  /$$/ /$$__  $$ /$$__  $$
| $$      | $$      | $$             \____  $$| $$$$$$$$| $$  \__/ \  $$/$$/ | $$| $$      | $$$$$$$$      | $$      | $$__  $$| $$ | $$$$$$$$| $$           | $$$$$$/ | $$$$$$$$| $$  \__/
| $$    $$| $$    $$| $$             /$$  \ $$| $$_____/| $$        \  $$$/  | $$| $$      | $$_____/      | $$    $$| $$  | $$| $$ | $$_____/| $$           | $$_  $$ | $$_____/| $$      
|  $$$$$$/|  $$$$$$/| $$$$$$$$      |  $$$$$$/|  $$$$$$$| $$         \  $/   | $$|  $$$$$$$|  $$$$$$$      |  $$$$$$/| $$  | $$| $$ |  $$$$$$$|  $$$$$$$      | $$ \  $$|  $$$$$$$| $$      
 \______/  \______/ |________/       \______/  \_______/|__/          \_/    |__/ \_______/ \_______/       \______/ |__/  |__/|__/  \_______/ \_______/      |__/  \__/ \_______/|__/      
                                                                                                                                                                                             
                                                                                   CCL SERVICE CHECKER — Made by Dress                                                                      
"@

# === CENTER ASCII ===
function Show-CenteredAscii {
    param ([string]$Text, [ConsoleColor]$Color = [ConsoleColor]::Red)
    
    $lines = $Text -split "`n"
    $width = [console]::WindowWidth
    foreach ($line in $lines) {
        $trimmed = $line.TrimEnd()
        $padding = [Math]::Max(0, [Math]::Floor(($width - $trimmed.Length) / 2))
        Write-Host (" " * $padding + $trimmed) -ForegroundColor $Color
    }
}

Show-CenteredAscii -Text $ascii -Color Red

# === GUI SECTION ===

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$servicesToCheck = @(
    "DusmSvc", "AppInfo", "PcaSvc", "DPS", "DiagTrack", 
    "SysMain", "Dnscache", "EventLog", "Bam", "Schedule", "WSearch"
)

$form = New-Object System.Windows.Forms.Form
$form.Text = "CCL Service Checker"
$form.Size = New-Object System.Drawing.Size(600,450)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.FormBorderStyle = "FixedDialog"

$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Text = "Select services to enable and set to auto-start:"
$headerLabel.Size = New-Object System.Drawing.Size(560,20)
$headerLabel.Location = New-Object System.Drawing.Point(20,20)
$headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($headerLabel)

$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(560,280)
$listView.Location = New-Object System.Drawing.Point(20,50)
$listView.View = "Details"
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.MultiSelect = $true
$listView.CheckBoxes = $true

$listView.Columns.Add("Enable", 60) | Out-Null
$listView.Columns.Add("Service Name", 150) | Out-Null
$listView.Columns.Add("Status", 80) | Out-Null
$listView.Columns.Add("Startup Type", 100) | Out-Null
$listView.Columns.Add("Description", 150) | Out-Null

$form.Controls.Add($listView)

$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Size = New-Object System.Drawing.Size(560,40)
$buttonPanel.Location = New-Object System.Drawing.Point(20,340)

$selectAllBox = New-Object System.Windows.Forms.CheckBox
$selectAllBox.Text = "Select All"
$selectAllBox.Size = New-Object System.Drawing.Size(100,20)
$selectAllBox.Location = New-Object System.Drawing.Point(10,10)
$selectAllBox.Add_CheckedChanged({
    foreach ($item in $listView.Items) {
        $item.Checked = $selectAllBox.Checked
    }
})
$buttonPanel.Controls.Add($selectAllBox)

$buttonEnable = New-Object System.Windows.Forms.Button
$buttonEnable.Text = "Enable Selected Services"
$buttonEnable.Size = New-Object System.Drawing.Size(180,30)
$buttonEnable.Location = New-Object System.Drawing.Point(120,5)
$buttonEnable.BackColor = [System.Drawing.Color]::LightBlue
$buttonPanel.Controls.Add($buttonEnable)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Text = "Refresh List"
$buttonRefresh.Size = New-Object System.Drawing.Size(100,30)
$buttonRefresh.Location = New-Object System.Drawing.Point(310,5)
$buttonPanel.Controls.Add($buttonRefresh)

$form.Controls.Add($buttonPanel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready - Select services and click 'Enable Selected Services'"
$statusLabel.Size = New-Object System.Drawing.Size(560,20)
$statusLabel.Location = New-Object System.Drawing.Point(20,390)
$statusLabel.TextAlign = "MiddleLeft"
$form.Controls.Add($statusLabel)

$serviceDescriptions = @{
    "DusmSvc"   = "Data Usage Service"
    "AppInfo"   = "Application Information"
    "PcaSvc"    = "Program Compatibility Assistant"
    "DPS"       = "Diagnostic Policy Service"
    "DiagTrack" = "Connected User Experience and Telemetry"
    "SysMain"   = "System Maintenance (Prefetcher)"
    "Dnscache"  = "DNS Client Cache"
    "EventLog"  = "Windows Event Log"
    "Bam"       = "Background Activity Moderator"
    "Schedule"  = "Task Scheduler"
    "WSearch"   = "Windows Search Service"
}

function Refresh-Services {
    $listView.Items.Clear()
    
    foreach ($serviceName in $servicesToCheck) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        $description = $serviceDescriptions[$serviceName]
        
        if ($service) {
            $status = $service.Status.ToString()
            $startupObj = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue
            $startup = if ($startupObj) { $startupObj.StartMode } else { "Unknown" }
            
            $item = New-Object System.Windows.Forms.ListViewItem("")
            $item.Checked = $false
            $item.SubItems.Add($serviceName) | Out-Null
            $item.SubItems.Add($status) | Out-Null
            $item.SubItems.Add($startup) | Out-Null
            $item.SubItems.Add($description) | Out-Null
            
            if ($status -eq "Running") {
                $item.BackColor = [System.Drawing.Color]::FromArgb(240, 255, 240)
            } else {
                $item.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 240)
            }
            
            $listView.Items.Add($item) | Out-Null
        } else {
            $item = New-Object System.Windows.Forms.ListViewItem("")
            $item.Checked = $false
            $item.SubItems.Add($serviceName) | Out-Null
            $item.SubItems.Add("Not Found") | Out-Null
            $item.SubItems.Add("N/A") | Out-Null
            $item.SubItems.Add($description) | Out-Null
            $item.BackColor = [System.Drawing.Color]::LightYellow
            $listView.Items.Add($item) | Out-Null
        }
    }
}

$buttonEnable.Add_Click({
    $selectedCount = 0
    $successCount = 0
    
    foreach ($item in $listView.Items) {
        if ($item.Checked) {
            $selectedCount++
            $serviceName = $item.SubItems[1].Text
            
            try {
                Set-Service -Name $serviceName -StartupType Automatic -ErrorAction SilentlyContinue
                $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
                if ($service -and $service.Status -ne 'Running') {
                    Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                }
                $successCount++
            } catch {}
        }
    }
    
    Refresh-Services
    
    if ($selectedCount -eq 0) {
        $statusLabel.Text = "Please select at least one service to enable"
    } else {
        $statusLabel.Text = "Completed: $successCount of $selectedCount services enabled successfully"
    }
})

$buttonRefresh.Add_Click({
    Refresh-Services
    $statusLabel.Text = "Service list refreshed"
})

Refresh-Services
[void]$form.ShowDialog()
