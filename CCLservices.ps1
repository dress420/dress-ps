# Full GUI Service Manager with custom red ASCII header
Clear-Host

$ascii = @"
  /$$$$$$   /$$$$$$  /$$              /$$$$$$                                 /$$                            /$$$$$$  /$$                           /$$                          
 /$$__  $$ /$$__  $$| $$             /$$__  $$                               |__/                           /$$__  $$| $$                          | $$                          
| $$  \__/| $$  \__/| $$            | $$  \__/  /$$$$$$   /$$$$$$  /$$    /$$ /$$  /$$$$$$$  /$$$$$$       | $$  \__/| $$$$$$$   /$$$$$$   /$$$$$$$| $$   /$$  /$$$$$$   /$$$$$$ 
| $$      | $$      | $$            |  $$$$$$  /$$__  $$ /$$__  $$|  $$  /$$/| $$ /$$_____/ /$$__  $$      | $$      | $$__  $$ /$$__  $$ /$$_____/| $$  /$$/ /$$__  $$ /$$__  $$
| $$      | $$      | $$             \____  $$| $$$$$$$$| $$  \__/ \  $$/$$/ | $$| $$      | $$$$$$$$      | $$      | $$  \ $$| $$$$$$$$| $$      | $$$$$$/ | $$$$$$$$| $$  \__/
| $$    $$| $$    $$| $$             /$$  \ $$| $$_____/| $$        \  $$$/  | $$| $$      | $$_____/      | $$    $$| $$  | $$| $$_____/| $$      | $$_  $$ | $$_____/| $$      
|  $$$$$$/|  $$$$$$/| $$$$$$$$      |  $$$$$$/|  $$$$$$$| $$         \  $/   | $$|  $$$$$$$|  $$$$$$$      |  $$$$$$/| $$  | $$|  $$$$$$$|  $$$$$$$| $$ \  $$|  $$$$$$$| $$      
 \______/  \______/ |________/       \______/  \_______/|__/          \_/    |__/ \_______/ \_______/       \______/ |__/  |__/ \_______/ \_______/|__/  \__/ \_______/|__/      
                                                                                                                                                                                 
                                                                                                                                                                                 
                                                                                                                                                                                 
                                                             /$$                 /$$                       /$$$$$$$                                                              
                                                            | $$                | $$                      | $$__  $$                                                             
                               /$$$$$$/$$$$   /$$$$$$   /$$$$$$$  /$$$$$$       | $$$$$$$  /$$   /$$      | $$  \ $$  /$$$$$$   /$$$$$$   /$$$$$$$ /$$$$$$$                      
                              | $$_  $$_  $$ |____  $$ /$$__  $$ /$$__  $$      | $$__  $$| $$  | $$      | $$  | $$ /$$__  $$ /$$__  $$ /$$_____//$$_____/                      
                              | $$ \ $$ \ $$  /$$$$$$$| $$  | $$| $$$$$$$$      | $$  \ $$| $$  | $$      | $$  | $$| $$  \__/| $$$$$$$$|  $$$$$$|  $$$$$$                       
                              | $$ | $$ | $$ /$$__  $$| $$  | $$| $$_____/      | $$  | $$| $$  | $$      | $$  | $$| $$      | $$_____/ \____  $$\____  $$                      
                              | $$ | $$ | $$|  $$$$$$$|  $$$$$$$|  $$$$$$$      | $$$$$$$/|  $$$$$$$      | $$$$$$$/| $$      |  $$$$$$$ /$$$$$$$//$$$$$$$/                      
                              |__/ |__/ |__/ \_______/ \_______/ \_______/      |_______/  \____  $$      |_______/ |__/       \_______/|_______/|_______/                       
                                                                                           /$$  | $$                                                                             
                                                                                          |  $$$$$$/                                                                             
                                                                                           \______/                                                                              
"@

# Print ASCII header in red
Write-Host $ascii -ForegroundColor Red

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Services to include
$servicesToCheck = @(
    "DusmSvc", "Appinfo", "PcaSvc", "DPS",
    "DiagTrack", "SysMain", "Dnscache", "EventLog",
    "Bam", "Schedule", "WSearch"
)

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Service Manager"
$form.Size = New-Object System.Drawing.Size(720,520)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.FormBorderStyle = "FixedDialog"

# Header label
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Text = "Select services to enable and set to auto-start:"
$headerLabel.Size = New-Object System.Drawing.Size(660,20)
$headerLabel.Location = New-Object System.Drawing.Point(20,20)
$headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($headerLabel)

# ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(660,320)
$listView.Location = New-Object System.Drawing.Point(20,50)
$listView.View = "Details"
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.MultiSelect = $true
$listView.CheckBoxes = $true

$listView.Columns.Add("Enable", 60) | Out-Null
$listView.Columns.Add("Service Name", 180) | Out-Null
$listView.Columns.Add("Status", 100) | Out-Null
$listView.Columns.Add("Startup Type", 120) | Out-Null
$listView.Columns.Add("Description", 200) | Out-Null

$form.Controls.Add($listView)

# Button panel
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Size = New-Object System.Drawing.Size(660,50)
$buttonPanel.Location = New-Object System.Drawing.Point(20,380)

$selectAllBox = New-Object System.Windows.Forms.CheckBox
$selectAllBox.Text = "Select All"
$selectAllBox.Size = New-Object System.Drawing.Size(100,20)
$selectAllBox.Location = New-Object System.Drawing.Point(10,15)
$selectAllBox.Add_CheckedChanged({
    foreach ($item in $listView.Items) {
        $item.Checked = $selectAllBox.Checked
    }
})
$buttonPanel.Controls.Add($selectAllBox)

$buttonEnable = New-Object System.Windows.Forms.Button
$buttonEnable.Text = "Enable Selected Services"
$buttonEnable.Size = New-Object System.Drawing.Size(200,30)
$buttonEnable.Location = New-Object System.Drawing.Point(130,10)
$buttonEnable.BackColor = [System.Drawing.Color]::LightBlue
$buttonPanel.Controls.Add($buttonEnable)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Text = "Refresh List"
$buttonRefresh.Size = New-Object System.Drawing.Size(120,30)
$buttonRefresh.Location = New-Object System.Drawing.Point(350,10)
$buttonPanel.Controls.Add($buttonRefresh)

$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "Close"
$buttonClose.Size = New-Object System.Drawing.Size(120,30)
$buttonClose.Location = New-Object System.Drawing.Point(490,10)
$buttonClose.Add_Click({ $form.Close() })
$buttonPanel.Controls.Add($buttonClose)

$form.Controls.Add($buttonPanel)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready - Select services and click 'Enable Selected Services'"
$statusLabel.Size = New-Object System.Drawing.Size(660,20)
$statusLabel.Location = New-Object System.Drawing.Point(20,440)
$statusLabel.TextAlign = "MiddleLeft"
$form.Controls.Add($statusLabel)

# Service descriptions
$serviceDescriptions = @{
    "DusmSvc"   = "Data Usage Service (tracks data usage and network stats)"
    "Appinfo"   = "Application Information Service (manages elevation for apps)"
    "PcaSvc"    = "Program Compatibility Assistant Service"
    "DPS"       = "Diagnostic Policy Service"
    "DiagTrack" = "Connected User Experiences and Telemetry"
    "SysMain"   = "Prefetcher for faster app loading (Superfetch)"
    "Dnscache"  = "DNS Client Service (resolves and caches DNS queries)"
    "EventLog"  = "Windows Event Log Service"
    "Bam"       = "Background Activity Moderator (controls background app activity)"
    "Schedule"  = "Windows Task Scheduler (runs scheduled tasks and maintenance)"
    "WSearch"   = "Windows Search Service (indexes files and content for fast search)"
}

# Function to refresh list
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

# Button click behaviors
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
            } catch {
                # silent catch
            }
        }
    }
    
    # Refresh automatically after enabling
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

# Initial population
Refresh-Services

# Show the form
[void]$form.ShowDialog()
