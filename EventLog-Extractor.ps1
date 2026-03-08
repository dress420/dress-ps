# ============================================================
#  Dress EventLog Extractor
#  GUI Launcher + HTML Report Generator + Forensic Mode
# ============================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Dress EventLog Extractor"
    Width="640" Height="600"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    Background="#08090D"
    WindowStyle="None"
    AllowsTransparency="True">

    <Window.Resources>

        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#00D4FF"/>
            <Setter Property="Foreground" Value="#000000"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#33DDFF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#0099BB"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="ForensicButton" TargetType="Button">
            <Setter Property="Background" Value="#1a0a2e"/>
            <Setter Property="Foreground" Value="#bf5af2"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#bf5af2"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#2d0f4e"/>
                                <Setter Property="BorderBrush" Value="#d97fff"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#0f0620"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="OutlineButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#4A5568"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#1E2229"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="#00D4FF"/>
                                <Setter Property="BorderBrush" Value="#00D4FF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="ModernCombo" TargetType="ComboBox">
            <Setter Property="Background" Value="#111318"/>
            <Setter Property="Foreground" Value="#00D4FF"/>
            <Setter Property="BorderBrush" Value="#1E2229"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Focusable="False" IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" Background="#111318" BorderBrush="#1E2229" BorderThickness="1">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="20"/>
                                                </Grid.ColumnDefinitions>
                                                <ContentPresenter Grid.Column="0" Margin="10,6" VerticalAlignment="Center"/>
                                                <Path Grid.Column="1" Data="M0,0 L4,4 L8,0" Stroke="#00D4FF" StrokeThickness="1.5" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                            </Grid>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter IsHitTestVisible="False" Margin="10,6,26,6" VerticalAlignment="Center"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
                            <Popup IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom" AllowsTransparency="True" Focusable="False">
                                <Border Background="#111318" BorderBrush="#1E2229" BorderThickness="1" CornerRadius="4" MaxHeight="200">
                                    <ScrollViewer>
                                        <ItemsPresenter/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Resources>
                <Style TargetType="ComboBoxItem">
                    <Setter Property="Background" Value="#111318"/>
                    <Setter Property="Foreground" Value="#00D4FF"/>
                    <Setter Property="FontFamily" Value="Consolas"/>
                    <Setter Property="FontSize" Value="12"/>
                    <Setter Property="Padding" Value="10,6"/>
                    <Style.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="#1a2030"/>
                        </Trigger>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter Property="Background" Value="#0d1a2a"/>
                            <Setter Property="Foreground" Value="#00D4FF"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
            </Style.Resources>
        </Style>

        <Style x:Key="ModernCheck" TargetType="CheckBox">
            <Setter Property="Foreground" Value="#C9D1D9"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,4,0,4"/>
        </Style>

    </Window.Resources>

    <Border BorderBrush="#1E2229" BorderThickness="1" CornerRadius="12">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="36"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="#0D1117" CornerRadius="12,12,0,0" Name="TitleBar">
                <Grid Margin="16,0">
                    <TextBlock Text="DRESS EVENTLOG EXTRACTOR" Foreground="#2A3040"
                               FontFamily="Consolas" FontSize="10" FontWeight="Bold"
                               VerticalAlignment="Center"/>
                    <Button Content="x" HorizontalAlignment="Right" VerticalAlignment="Center"
                            Width="24" Height="24" Name="CloseBtn"
                            Style="{StaticResource OutlineButton}" Padding="0"/>
                </Grid>
            </Border>

            <StackPanel Grid.Row="1" Margin="48,22,48,22">

                <!-- Title -->
                <StackPanel HorizontalAlignment="Center" Margin="0,0,0,24">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,12">
                        <Rectangle Width="40" Height="1" Fill="#1E2229" VerticalAlignment="Center"/>
                        <TextBlock Text="   *   " Foreground="#00D4FF" FontSize="10" VerticalAlignment="Center"/>
                        <Rectangle Width="40" Height="1" Fill="#1E2229" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock Text="DRESS" HorizontalAlignment="Center"
                               FontFamily="Consolas" FontSize="11" FontWeight="Bold"
                               Foreground="#00D4FF" Margin="0,0,0,4"/>
                    <TextBlock Text="EventLog Extractor" HorizontalAlignment="Center"
                               FontFamily="Georgia" FontSize="32" FontWeight="Bold"
                               Foreground="#FFFFFF" Margin="0,0,0,6"/>
                    <TextBlock Text="Windows  Event  Intelligence" HorizontalAlignment="Center"
                               FontFamily="Consolas" FontSize="11" Foreground="#2A3A4A"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,12,0,0">
                        <Rectangle Width="60" Height="1" Fill="#1E2229" VerticalAlignment="Center"/>
                        <Rectangle Width="7" Height="7" Fill="#00D4FF" Margin="8,0" VerticalAlignment="Center">
                            <Rectangle.RenderTransform><RotateTransform Angle="45"/></Rectangle.RenderTransform>
                        </Rectangle>
                        <Rectangle Width="60" Height="1" Fill="#1E2229" VerticalAlignment="Center"/>
                    </StackPanel>
                </StackPanel>

                <!-- Settings Grid -->
                <Grid Margin="0,0,0,16">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="16"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="14"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Column="0" Grid.Row="0">
                        <TextBlock Text="TIME RANGE" Foreground="#2A3A4A" FontFamily="Consolas"
                                   FontSize="9" FontWeight="Bold" Margin="0,0,0,6"/>
                        <ComboBox Name="DaysCombo" Style="{StaticResource ModernCombo}">
                            <ComboBoxItem Content="Last 1 Day" Tag="1"/>
                            <ComboBoxItem Content="Last 3 Days" Tag="3"/>
                            <ComboBoxItem Content="Last 7 Days" Tag="7" IsSelected="True"/>
                            <ComboBoxItem Content="Last 14 Days" Tag="14"/>
                            <ComboBoxItem Content="Last 30 Days" Tag="30"/>
                        </ComboBox>
                    </StackPanel>

                    <StackPanel Grid.Column="2" Grid.Row="0">
                        <TextBlock Text="MIN LEVEL" Foreground="#2A3A4A" FontFamily="Consolas"
                                   FontSize="9" FontWeight="Bold" Margin="0,0,0,6"/>
                        <ComboBox Name="LevelCombo" Style="{StaticResource ModernCombo}">
                            <ComboBoxItem Content="Critical" Tag="Critical"/>
                            <ComboBoxItem Content="Error" Tag="Error"/>
                            <ComboBoxItem Content="Warning" Tag="Warning" IsSelected="True"/>
                            <ComboBoxItem Content="Information" Tag="Information"/>
                            <ComboBoxItem Content="All" Tag="All"/>
                        </ComboBox>
                    </StackPanel>

                    <StackPanel Grid.Column="0" Grid.Row="2" Grid.ColumnSpan="3">
                        <TextBlock Text="CHANNELS" Foreground="#2A3A4A" FontFamily="Consolas"
                                   FontSize="9" FontWeight="Bold" Margin="0,0,0,8"/>
                        <StackPanel Orientation="Horizontal">
                            <CheckBox Name="ChkSystem" Content="System" IsChecked="True"
                                      Style="{StaticResource ModernCheck}" Margin="0,0,28,0"/>
                            <CheckBox Name="ChkApplication" Content="Application" IsChecked="True"
                                      Style="{StaticResource ModernCheck}" Margin="0,0,28,0"/>
                            <CheckBox Name="ChkSecurity" Content="Security" IsChecked="True"
                                      Style="{StaticResource ModernCheck}"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>

                <!-- Forensic Mode Toggle -->
                <Border Background="#0d0a1a" CornerRadius="8" Padding="14,10" Margin="0,0,0,14"
                        BorderBrush="#2a1a4a" BorderThickness="1">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0">
                            <TextBlock Text="FORENSIC MODE" Foreground="#bf5af2" FontFamily="Consolas"
                                       FontSize="10" FontWeight="Bold" Margin="0,0,0,3"/>
                            <TextBlock Text="Scans for: USN deletion, audit clears, shutdowns, permission changes, script execution, time changes and more"
                                       Foreground="#5a3a7a" FontFamily="Consolas" FontSize="10"
                                       TextWrapping="Wrap"/>
                        </StackPanel>
                        <CheckBox Name="ChkForensic" Grid.Column="1" VerticalAlignment="Center"
                                  Margin="12,0,0,0" IsChecked="False">
                            <CheckBox.Style>
                                <Style TargetType="CheckBox">
                                    <Setter Property="Foreground" Value="#bf5af2"/>
                                    <Setter Property="FontFamily" Value="Consolas"/>
                                    <Setter Property="FontSize" Value="11"/>
                                </Style>
                            </CheckBox.Style>
                        </CheckBox>
                    </Grid>
                </Border>

                <!-- Status -->
                <Border Background="#0D1117" CornerRadius="6" Padding="14,10" Margin="0,0,0,12"
                        BorderBrush="#1E2229" BorderThickness="1">
                    <TextBlock Name="StatusText" Text="Ready. Configure settings and start extraction."
                               Foreground="#4A5568" FontFamily="Consolas" FontSize="11"
                               TextWrapping="Wrap"/>
                </Border>

                <!-- Progress -->
                <ProgressBar Name="ProgressBar" Height="3" Background="#1E2229"
                             Foreground="#00D4FF" BorderThickness="0"
                             Minimum="0" Maximum="100" Value="0" Margin="0,0,0,16"/>

                <!-- Buttons -->
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="12"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="12"/>
                        <ColumnDefinition Width="110"/>
                    </Grid.ColumnDefinitions>
                    <Button Grid.Column="0" Name="RunBtn" Content="EXTRACT LOGS"
                            Style="{StaticResource ModernButton}" Height="44" Padding="0"/>
                    <Button Grid.Column="2" Name="ForensicBtn" Content="FORENSIC EXTRACT"
                            Style="{StaticResource ForensicButton}" Height="44" Padding="0"/>
                    <Button Grid.Column="4" Name="CancelBtn" Content="CANCEL"
                            Style="{StaticResource OutlineButton}" Height="44" Padding="0"/>
                </Grid>

            </StackPanel>
        </Grid>
    </Border>
</Window>
"@

# ============================================================
#  FORENSIC EVENT IDs (with descriptions)
# ============================================================
$ForensicIDs = @{
    3079 = "USN Journal Deleted"
    1102 = "Audit Log Cleared"
    7    = "Unexpected System Shutdown"
    7036 = "Service State Modified (started/stopped)"
    1001 = "Service Crashed"
    1002 = "Service Crash (repeat)"
    104  = "User Logon"
    1074 = "System Shut Down by User"
    4663 = "Folder Permissions Changed"
    4688 = "Script / Process Executed"
    4660 = "File Permissions Menu Accessed"
    4616 = "System Time Changed"
    4720 = "User Account Created"
    4726 = "User Account Deleted"
    4724 = "Password Reset Attempt"
    4723 = "Password Change Attempt"
    4719 = "Audit Policy Changed"
    4672 = "Special Privileges Assigned"
    4625 = "Failed Login Attempt"
    4624 = "Successful Login"
    4634 = "User Logoff"
    4728 = "User Added to Security Group"
    4732 = "User Added to Local Group"
    4756 = "User Added to Universal Group"
    4698 = "Scheduled Task Created"
    4702 = "Scheduled Task Updated"
    4699 = "Scheduled Task Deleted"
    4657 = "Registry Value Modified"
    4648 = "Login with Explicit Credentials"
    4768 = "Kerberos Ticket Requested"
    4771 = "Kerberos Pre-Auth Failed"
}

# Load XAML
$Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$XAML)
try {
    $Window = [Windows.Markup.XamlReader]::Load($Reader)
} catch {
    Write-Host "GUI Error: $_" -ForegroundColor Red
    exit
}

$CloseBtn    = $Window.FindName("CloseBtn")
$RunBtn      = $Window.FindName("RunBtn")
$ForensicBtn = $Window.FindName("ForensicBtn")
$CancelBtn   = $Window.FindName("CancelBtn")
$StatusText  = $Window.FindName("StatusText")
$ProgressBar = $Window.FindName("ProgressBar")
$DaysCombo   = $Window.FindName("DaysCombo")
$LevelCombo  = $Window.FindName("LevelCombo")
$ChkSystem   = $Window.FindName("ChkSystem")
$ChkApp      = $Window.FindName("ChkApplication")
$ChkSecurity = $Window.FindName("ChkSecurity")
$ChkForensic = $Window.FindName("ChkForensic")
$TitleBar    = $Window.FindName("TitleBar")

$TitleBar.Add_MouseLeftButtonDown({ $Window.DragMove() })
$CloseBtn.Add_Click({ $Window.Close() })
$CancelBtn.Add_Click({ $Window.Close() })

# ============================================================
#  SHARED EXTRACTION FUNCTION
# ============================================================
function Run-Extraction {
    param([bool]$ForensicMode)

    $RunBtn.IsEnabled      = $false
    $ForensicBtn.IsEnabled = $false

    if ($ForensicMode) {
        $StatusText.Foreground = [Windows.Media.Brushes]::Violet
        $ProgressBar.Foreground = [Windows.Media.Brushes]::Violet
    } else {
        $StatusText.Foreground = [Windows.Media.Brushes]::LightSkyBlue
        $ProgressBar.Foreground = [Windows.Media.Brushes]::DeepSkyBlue
    }

    $SelectedDays  = [int]($DaysCombo.SelectedItem.Tag)
    $SelectedLevel = $LevelCombo.SelectedItem.Tag

    $SelectedChannels = @()
    if ($ChkSystem.IsChecked)   { $SelectedChannels += "System" }
    if ($ChkApp.IsChecked)      { $SelectedChannels += "Application" }
    if ($ChkSecurity.IsChecked) { $SelectedChannels += "Security" }

    if ($SelectedChannels.Count -eq 0) {
        $StatusText.Text = "Please select at least one channel."
        $StatusText.Foreground = [Windows.Media.Brushes]::OrangeRed
        $RunBtn.IsEnabled = $true; $ForensicBtn.IsEnabled = $true
        return
    }

    $StatusText.Text = if ($ForensicMode) { "Forensic scan starting..." } else { "Starting extraction..." }
    $ProgressBar.Value = 5
    $Window.Dispatcher.Invoke([action]{}, "Render")

    $StartZeit = (Get-Date).AddDays(-$SelectedDays)
    $JetztZeit = Get-Date

    # Forensic mode forces All levels to catch Info events like 3079, 104 etc
    $LevelFilter = if ($ForensicMode) {
        @(1, 2, 3, 4, 5)
    } else {
        switch ($SelectedLevel) {
            "Critical"    { @(1) }
            "Error"       { @(1, 2) }
            "Warning"     { @(1, 2, 3) }
            "Information" { @(1, 2, 3, 4) }
            "All"         { @(1, 2, 3, 4, 5) }
        }
    }

    $AlleEvents = @()
    $Step = 0

    foreach ($Kanal in $SelectedChannels) {
        $Step++
        $Pct = [int](($Step / ($SelectedChannels.Count + 1)) * 75)
        $ProgressBar.Value = $Pct
        $StatusText.Text = "Reading $Kanal..."
        $Window.Dispatcher.Invoke([action]{}, "Render")

        try {
            $FilterHash = @{
                LogName   = $Kanal
                StartTime = $StartZeit
                Level     = $LevelFilter
            }

            # In forensic mode also explicitly pull by ID to catch any that slip through level filter
            $Events = Get-WinEvent -FilterHashtable $FilterHash -ErrorAction SilentlyContinue

            if ($Events) {
                foreach ($E in $Events) {
                    # In forensic mode, skip non-forensic IDs
                    if ($ForensicMode -and -not $ForensicIDs.ContainsKey([int]$E.Id)) { continue }

                    $Msg = $E.Message
                    if ($Msg) {
                        $Msg = $Msg -replace '"', '&quot;'
                        $Msg = $Msg -replace '<', '&lt;'
                        $Msg = $Msg -replace '>', '&gt;'
                        $Msg = $Msg -replace "`r`n|`n", '<br>'
                    } else { $Msg = "" }

                    $ForensicLabel = ""
                    if ($ForensicIDs.ContainsKey([int]$E.Id)) {
                        $ForensicLabel = $ForensicIDs[[int]$E.Id]
                    }

                    $AlleEvents += [PSCustomObject]@{
                        EventID       = $E.Id
                        Zeit          = $E.TimeCreated
                        Stufe         = $E.LevelDisplayName
                        Kanal         = $E.LogName
                        Quelle        = $E.ProviderName
                        Maschine      = $E.MachineName
                        Nachricht     = $Msg
                        StufeNum      = $E.Level
                        ForensicLabel = $ForensicLabel
                        IsForensic    = ($ForensicLabel -ne "")
                    }
                }
                $StatusText.Text = "$Kanal - $($Events.Count) events read"
            } else {
                $StatusText.Text = "$Kanal - No events in range"
            }
        } catch {
            $StatusText.Text = "Error reading $Kanal"
        }
        $Window.Dispatcher.Invoke([action]{}, "Render")
    }

    # In forensic mode also scan Application log specifically for 3079 if not already included
    if ($ForensicMode -and $SelectedChannels -notcontains "Application") {
        $ProgressBar.Value = 80
        $StatusText.Text = "Scanning Application log for forensic IDs..."
        $Window.Dispatcher.Invoke([action]{}, "Render")
        try {
            $ExtraEvents = Get-WinEvent -FilterHashtable @{
                LogName   = "Application"
                StartTime = $StartZeit
                Level     = @(1,2,3,4,5)
            } -ErrorAction SilentlyContinue
            if ($ExtraEvents) {
                foreach ($E in $ExtraEvents) {
                    if (-not $ForensicIDs.ContainsKey([int]$E.Id)) { continue }
                    $Msg = $E.Message
                    if ($Msg) {
                        $Msg = $Msg -replace '"','&quot;' -replace '<','&lt;' -replace '>','&gt;'
                        $Msg = $Msg -replace "`r`n|`n",'<br>'
                    } else { $Msg = "" }
                    $AlleEvents += [PSCustomObject]@{
                        EventID       = $E.Id
                        Zeit          = $E.TimeCreated
                        Stufe         = $E.LevelDisplayName
                        Kanal         = $E.LogName
                        Quelle        = $E.ProviderName
                        Maschine      = $E.MachineName
                        Nachricht     = $Msg
                        StufeNum      = $E.Level
                        ForensicLabel = $ForensicIDs[[int]$E.Id]
                        IsForensic    = $true
                    }
                }
            }
        } catch {}
    }

    $ProgressBar.Value = 85
    $StatusText.Text = "Building HTML report ($($AlleEvents.Count) events)..."
    $Window.Dispatcher.Invoke([action]{}, "Render")

    $AlleEvents = $AlleEvents | Sort-Object Zeit -Descending

    $AnzahlKritisch = ($AlleEvents | Where-Object { $_.StufeNum -eq 1 }).Count
    $AnzahlFehler   = ($AlleEvents | Where-Object { $_.StufeNum -eq 2 }).Count
    $AnzahlWarnung  = ($AlleEvents | Where-Object { $_.StufeNum -eq 3 }).Count
    $AnzahlInfo     = ($AlleEvents | Where-Object { $_.StufeNum -eq 4 }).Count
    $AnzahlForensic = ($AlleEvents | Where-Object { $_.IsForensic }).Count

    $TopIDs = $AlleEvents | Group-Object EventID | Sort-Object Count -Descending | Select-Object -First 10

    # Build forensic summary
    $ForensicSummaryHtml = ""
    if ($ForensicMode -or $AnzahlForensic -gt 0) {
        $ForensicFound = $AlleEvents | Where-Object { $_.IsForensic } | Group-Object EventID | Sort-Object Count -Descending
        foreach ($F in $ForensicFound) {
            $Desc = $ForensicIDs[[int]$F.Name]
            $ForensicSummaryHtml += "<div class='fitem'><span class='fid'>#$($F.Name)</span><span class='fdesc'>$Desc</span><span class='fcount'>$($F.Count)x</span></div>"
        }
    }

    $HtmlZeilen = ""
    foreach ($E in $AlleEvents) {
        $FarbeKlasse = switch ($E.StufeNum) {
            1 { "critical" } 2 { "error" } 3 { "warning" } 4 { "info" } default { "verbose" }
        }
        $Symbol = switch ($E.StufeNum) {
            1 { "CRITICAL" } 2 { "ERROR" } 3 { "WARNING" } 4 { "INFO" } default { "VERBOSE" }
        }
        $ZeitFormatiert = $E.Zeit.ToString("dd.MM.yyyy HH:mm:ss")
        $MsgShort = ""; $MsgFull = ""
        if ($E.Nachricht.Length -gt 120) {
            $MsgShort = $E.Nachricht.Substring(0, 120) + "..."
            $MsgFull  = $E.Nachricht
        } else { $MsgShort = $E.Nachricht }
        $expandBtn = ""
        if ($MsgFull -ne "") {
            $expandBtn = "<button class='xbtn' onclick='toggleMsg(this)'>Show more</button><div class='mfull hidden'>$MsgFull</div>"
        }
        $forensicBadge = ""
        $rowClass = "r$($E.StufeNum)"
        if ($E.IsForensic) {
            $forensicBadge = "<span class='fbadge'>&#9888; $($E.ForensicLabel)</span>"
            $rowClass += " forensic-row"
        }
        $HtmlZeilen += "<tr class='$rowClass' data-level='$($E.StufeNum)' data-kanal='$($E.Kanal)' data-forensic='$($E.IsForensic.ToString().ToLower())'>"
        $HtmlZeilen += "<td><span class='badge b$FarbeKlasse'>$Symbol</span></td>"
        $HtmlZeilen += "<td class='mono'>$ZeitFormatiert</td>"
        $HtmlZeilen += "<td class='mono bold'>$($E.EventID)</td>"
        $HtmlZeilen += "<td>$($E.Kanal)</td>"
        $HtmlZeilen += "<td class='src'>$($E.Quelle)</td>"
        $HtmlZeilen += "<td class='mc'><div class='mp'>$forensicBadge$MsgShort</div>$expandBtn</td></tr>"
    }

    $TopIDsHtml = ""
    foreach ($T in $TopIDs) {
        $Pct2 = if ($AlleEvents.Count -gt 0) { [Math]::Round($T.Count / $AlleEvents.Count * 100) } else { 0 }
        $isF = $ForensicIDs.ContainsKey([int]$T.Name)
        $fMark = if ($isF) { " <span style='color:#bf5af2'>&#9888;</span>" } else { "" }
        $TopIDsHtml += "<tr><td class='mono bold'>#$($T.Name)$fMark</td><td><div class='bw'><div class='bb' style='width:$Pct2%'></div></div></td><td class='mono'>$($T.Count)x</td></tr>"
    }

    $ModeLabel  = if ($ForensicMode) { "FORENSIC SCAN" } else { "Standard Extract" }
    $StartStr   = $StartZeit.ToString('dd.MM.yyyy')
    $JetztStr   = $JetztZeit.ToString('dd.MM.yyyy HH:mm:ss')
    $Computer   = $env:COMPUTERNAME
    $AusgabePfad = "$env:USERPROFILE\Desktop\EventLog-Report.html"

    $HTML = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'><title>Dress EventLog Report</title><style>"
    $HTML += ":root{--bg:#08090D;--sf:#0F1116;--bd:#1a1f28;--tx:#c9d1d9;--mt:#3a4455;--ac:#00D4FF;--cr:#ff2d55;--er:#ff6b35;--wn:#ffd60a;--in:#30d158;--vb:#636e7b;--fo:#bf5af2}"
    $HTML += "*{margin:0;padding:0;box-sizing:border-box}body{background:var(--bg);color:var(--tx);font-family:'Segoe UI',Arial,sans-serif;font-size:14px}"
    $HTML += "header{background:#0D1117;border-bottom:1px solid var(--bd);padding:22px 36px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}"
    $HTML += ".lt{display:flex;flex-direction:column;gap:2px}.ltag{font-family:Consolas,monospace;font-size:9px;font-weight:700;letter-spacing:4px;color:var(--ac)}"
    $HTML += ".ltitle{font-family:Georgia,serif;font-size:22px;font-weight:700;color:#fff}.lsub{font-family:Consolas,monospace;font-size:10px;letter-spacing:2px;color:var(--mt);margin-top:2px}"
    $HTML += ".mode-badge{display:inline-block;padding:3px 10px;border-radius:4px;font-size:10px;font-weight:700;font-family:Consolas,monospace;margin-top:6px}"
    $HTML += ".mode-forensic{background:rgba(191,90,242,0.15);color:var(--fo);border:1px solid rgba(191,90,242,0.3)}"
    $HTML += ".mode-standard{background:rgba(0,212,255,0.1);color:var(--ac);border:1px solid rgba(0,212,255,0.2)}"
    $HTML += ".hm{text-align:right;color:var(--mt);font-size:11px;font-family:Consolas,monospace;line-height:1.8}.hm span{color:var(--ac)}"
    $HTML += ".stats{display:grid;grid-template-columns:repeat(6,1fr);background:var(--bd)}"
    $HTML += ".sc{background:var(--sf);padding:16px 20px;cursor:pointer;border-bottom:2px solid transparent;transition:all .15s}.sc:hover{background:#13171f}.sc.active{border-bottom-color:var(--ac)}"
    $HTML += ".sn{font-family:Consolas,monospace;font-size:24px;font-weight:700;line-height:1;margin-bottom:4px}.sl{font-size:9px;text-transform:uppercase;letter-spacing:1.5px;color:var(--mt)}"
    $HTML += ".sc.all .sn{color:var(--ac)}.sc.crit .sn{color:var(--cr)}.sc.err .sn{color:var(--er)}.sc.warn .sn{color:var(--wn)}.sc.info .sn{color:var(--in)}.sc.fo .sn{color:var(--fo)}"
    $HTML += ".sc.fo.active{border-bottom-color:var(--fo)}"
    $HTML += ".layout{display:flex}.sidebar{width:260px;min-width:260px;background:var(--sf);border-right:1px solid var(--bd);padding:20px;position:sticky;top:99px;height:calc(100vh - 99px);overflow-y:auto}"
    $HTML += ".sidebar h3{font-size:9px;text-transform:uppercase;letter-spacing:2px;color:var(--mt);margin-bottom:9px;margin-top:18px;font-family:Consolas,monospace}.sidebar h3:first-child{margin-top:0}"
    $HTML += ".fg{display:flex;flex-direction:column;gap:5px}"
    $HTML += ".fb{background:transparent;border:1px solid var(--bd);border-radius:5px;color:var(--tx);padding:7px 11px;text-align:left;cursor:pointer;font-size:12px;font-family:Consolas,monospace;transition:all .15s;display:flex;align-items:center;gap:8px}"
    $HTML += ".fb:hover{border-color:var(--ac);color:var(--ac)}.fb.active{background:rgba(0,212,255,.08);border-color:var(--ac);color:var(--ac)}"
    $HTML += ".fb.fo-btn:hover{border-color:var(--fo);color:var(--fo)}.fb.fo-btn.active{background:rgba(191,90,242,.08);border-color:var(--fo);color:var(--fo)}"
    $HTML += ".fd{width:7px;height:7px;border-radius:50%;flex-shrink:0}.d0{background:var(--ac)}.d1{background:var(--in)}.d2{background:var(--wn)}.d3{background:var(--cr)}.d4{background:var(--fo)}"
    $HTML += ".ti table{width:100%;border-collapse:collapse}.ti td{padding:4px 3px;font-size:11px;font-family:Consolas,monospace}"
    $HTML += ".bw{height:5px;background:var(--bd);border-radius:3px;overflow:hidden}.bb{height:100%;background:var(--ac);border-radius:3px;opacity:.6}"
    $HTML += ".fitem{display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid var(--bd)}"
    $HTML += ".fid{font-family:Consolas,monospace;font-size:11px;font-weight:700;color:var(--fo);min-width:48px}"
    $HTML += ".fdesc{font-family:Consolas,monospace;font-size:11px;color:var(--tx);flex:1}"
    $HTML += ".fcount{font-family:Consolas,monospace;font-size:11px;color:var(--mt)}"
    $HTML += ".ma{flex:1;overflow:hidden}.sw{padding:13px 22px;background:var(--sf);border-bottom:1px solid var(--bd);position:relative}"
    $HTML += ".sw input{width:100%;background:var(--bg);border:1px solid var(--bd);border-radius:6px;color:var(--tx);padding:9px 13px 9px 36px;font-family:Consolas,monospace;font-size:12px;outline:none;transition:border .15s}"
    $HTML += ".sw input:focus{border-color:var(--ac)}.si{position:absolute;left:36px;top:50%;transform:translateY(-50%);color:var(--mt);pointer-events:none}"
    $HTML += ".tw{overflow-x:auto}table.mt{width:100%;border-collapse:collapse}"
    $HTML += "table.mt thead th{background:#0a0c12;padding:10px 15px;text-align:left;font-size:9px;text-transform:uppercase;letter-spacing:1.5px;color:var(--mt);border-bottom:1px solid var(--bd);white-space:nowrap;font-family:Consolas,monospace}"
    $HTML += "table.mt tbody tr{border-bottom:1px solid var(--bd);transition:background .1s}table.mt tbody tr:hover{background:rgba(255,255,255,.025)}"
    $HTML += "table.mt tbody td{padding:9px 15px;vertical-align:top}"
    $HTML += ".forensic-row{background:rgba(191,90,242,0.04)}"
    $HTML += ".forensic-row td:first-child{border-left:2px solid var(--fo) !important}"
    $HTML += ".r1 td:first-child{border-left:2px solid var(--cr)}.r2 td:first-child{border-left:2px solid var(--er)}.r3 td:first-child{border-left:2px solid var(--wn)}.r4 td:first-child{border-left:2px solid var(--in)}"
    $HTML += ".badge{display:inline-block;padding:3px 7px;border-radius:3px;font-size:9px;font-weight:700;white-space:nowrap;font-family:Consolas,monospace}"
    $HTML += ".bcritical{background:rgba(255,45,85,.12);color:var(--cr)}.berror{background:rgba(255,107,53,.12);color:var(--er)}.bwarning{background:rgba(255,214,10,.12);color:var(--wn)}.binfo{background:rgba(48,209,88,.12);color:var(--in)}.bverbose{background:rgba(99,110,123,.12);color:var(--vb)}"
    $HTML += ".fbadge{display:inline-block;background:rgba(191,90,242,0.12);color:var(--fo);border:1px solid rgba(191,90,242,0.25);padding:2px 7px;border-radius:3px;font-size:9px;font-weight:700;font-family:Consolas,monospace;margin-right:8px;margin-bottom:4px}"
    $HTML += ".mono{font-family:Consolas,monospace;font-size:11px}.bold{font-weight:700}"
    $HTML += ".src{color:var(--mt);font-size:11px;font-family:Consolas,monospace;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}"
    $HTML += ".mc{max-width:450px}.mp{color:var(--tx);font-size:12px;line-height:1.5}"
    $HTML += ".mfull{color:var(--mt);font-size:11px;line-height:1.6;margin-top:6px;padding:9px;background:var(--bg);border-radius:4px;border:1px solid var(--bd);font-family:Consolas,monospace}"
    $HTML += ".xbtn{background:none;border:none;color:var(--ac);cursor:pointer;font-size:10px;margin-top:4px;padding:0;font-family:Consolas,monospace}"
    $HTML += ".hidden{display:none}.nr{padding:60px;text-align:center;color:var(--mt)}"
    $HTML += "::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:var(--sf)}::-webkit-scrollbar-thumb{background:var(--bd);border-radius:3px}"
    $HTML += "</style></head><body>"

    $ModeBadgeClass = if ($ForensicMode) { "mode-forensic" } else { "mode-standard" }
    $HTML += "<header><div class='lt'><div class='ltag'>DRESS</div><div class='ltitle'>EventLog Extractor</div><span class='mode-badge $ModeBadgeClass'>$ModeLabel</span></div>"
    $HTML += "<div class='hm'>Generated: <span>$JetztStr</span><br>Period: <span>$StartStr to $JetztStr</span><br>Machine: <span>$Computer</span></div></header>"
    $HTML += "<div class='stats'>"
    $HTML += "<div class='sc all active' onclick='fL(`"all`",this)'><div class='sn'>$($AlleEvents.Count)</div><div class='sl'>Total</div></div>"
    $HTML += "<div class='sc crit' onclick='fL(`"1`",this)'><div class='sn'>$AnzahlKritisch</div><div class='sl'>Critical</div></div>"
    $HTML += "<div class='sc err' onclick='fL(`"2`",this)'><div class='sn'>$AnzahlFehler</div><div class='sl'>Error</div></div>"
    $HTML += "<div class='sc warn' onclick='fL(`"3`",this)'><div class='sn'>$AnzahlWarnung</div><div class='sl'>Warning</div></div>"
    $HTML += "<div class='sc info' onclick='fL(`"4`",this)'><div class='sn'>$AnzahlInfo</div><div class='sl'>Info</div></div>"
    $HTML += "<div class='sc fo' onclick='fL(`"forensic`",this)'><div class='sn'>$AnzahlForensic</div><div class='sl'>&#9888; Forensic</div></div>"
    $HTML += "</div>"
    $HTML += "<div class='layout'><aside class='sidebar'>"
    $HTML += "<h3>Channel</h3><div class='fg' id='kf'>"
    $HTML += "<button class='fb active' onclick='fK(`"all`",this)'><span class='fd d0'></span>All Channels</button>"
    $HTML += "<button class='fb' onclick='fK(`"System`",this)'><span class='fd d1'></span>System</button>"
    $HTML += "<button class='fb' onclick='fK(`"Application`",this)'><span class='fd d2'></span>Application</button>"
    $HTML += "<button class='fb' onclick='fK(`"Security`",this)'><span class='fd d3'></span>Security</button>"
    $HTML += "</div>"
    if ($ForensicSummaryHtml -ne "") {
        $HTML += "<h3>&#9888; Forensic Hits</h3><div>$ForensicSummaryHtml</div>"
    }
    $HTML += "<h3>Top Event IDs</h3><div class='ti'><table>$TopIDsHtml</table></div></aside>"
    $HTML += "<div class='ma'><div class='sw'><span class='si'>&#128269;</span>"
    $HTML += "<input type='text' id='sb' placeholder='Search by ID, source, message...' oninput='af()'></div>"
    $HTML += "<div class='tw'><table class='mt'><thead><tr><th>Level</th><th>Timestamp</th><th>Event ID</th><th>Channel</th><th>Source</th><th>Message</th></tr></thead>"
    $HTML += "<tbody id='tb'>$HtmlZeilen</tbody></table>"
    $HTML += "<div class='nr hidden' id='nr'>No results found.</div></div></div></div>"
    $HTML += "<script>var al='all',ak='all';"
    $HTML += "function fL(l,el){al=l;document.querySelectorAll('.sc').forEach(function(c){c.classList.remove('active')});el.classList.add('active');af();}"
    $HTML += "function fK(k,b){ak=k;document.querySelectorAll('#kf .fb').forEach(function(x){x.classList.remove('active')});b.classList.add('active');af();}"
    $HTML += "function af(){var s=document.getElementById('sb').value.toLowerCase();var rows=document.querySelectorAll('#tb tr');var v=0;rows.forEach(function(r){"
    $HTML += "var lok=al==='all'||(al==='forensic'?r.dataset.forensic==='true':r.dataset.level===al);"
    $HTML += "var kok=ak==='all'||r.dataset.kanal===ak;var sok=!s||r.textContent.toLowerCase().indexOf(s)>=0;"
    $HTML += "r.style.display=(lok&&kok&&sok)?'':'none';if(lok&&kok&&sok)v++;});document.getElementById('nr').classList.toggle('hidden',v>0);}"
    $HTML += "function toggleMsg(btn){var f=btn.nextElementSibling;btn.textContent=f.classList.toggle('hidden')?'Show more':'Show less';}"
    $HTML += "</script></body></html>"

    $HTML | Out-File -FilePath $AusgabePfad -Encoding UTF8

    $ProgressBar.Value = 100
    $StatusText.Foreground = if ($ForensicMode) { [Windows.Media.Brushes]::Violet } else { [Windows.Media.Brushes]::LightGreen }
    $StatusText.Text = "Done! $($AlleEvents.Count) events$(if($AnzahlForensic -gt 0){" | $AnzahlForensic FORENSIC HITS"}). Opening report..."
    $Window.Dispatcher.Invoke([action]{}, "Render")

    Start-Sleep -Milliseconds 800
    Start-Process $AusgabePfad
    Start-Sleep -Milliseconds 500
    $Window.Close()
}

$RunBtn.Add_Click({ Run-Extraction -ForensicMode $false })
$ForensicBtn.Add_Click({ Run-Extraction -ForensicMode $true })

$Window.ShowDialog() | Out-Null
