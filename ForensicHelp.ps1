#    Comprehensive Forensics Script
#    This script retrieves SRUM, DPS, BAM, and .jar file information for tracking executions and provides detailed analysis.
#    Made with love by Dress <3

Write-Host @"
   _____              .___       ___.           ________                               
  /     \ _____     __| _/____   \_ |__ ___.__. \______ \_______   ____   ______ ______
 /  \ /  \\__  \   / __ |/ __ \   | __ <   |  |  |    |  \_  __ \_/ __ \ /  ___//  ___/
/    Y    \/ __ \_/ /_/ \  ___/   | \_\ \___  |  |    `   \  | \/\  ___/ \___ \ \___ \
\____|__  (____  /\____ |\___  >  |___  / ____| /_______  /__|    \___  >____  >____  >
        \/     \/      \/    \/       \/\/              \/            \/     \/     \/

"@ -ForegroundColor Cyan

$OutputDirectory = "C:\Screenshare"
$OutputFile = Join-Path $OutputDirectory "comprehensive_forensics.txt"
$Artifacts = @()

if (-not (Test-Path $OutputDirectory)) {
    try {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-Host "Created output directory: $OutputDirectory" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create output directory: $OutputDirectory"
        Write-Error "Error: $($_.Exception.Message)"
        exit 1
    }
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Write-Host $logEntry
    Add-Content -Path $OutputFile -Value $logEntry
}

function Get-DigitalSignature {
    param([string]$FilePath)

    try {
        $sig = Get-AuthenticodeSignature -FilePath $FilePath -ErrorAction SilentlyContinue
        if ($sig -and $sig.Status -eq "Valid") {
            return "Signed - $($sig.SignerCertificate.Subject)"
        } elseif ($sig -and $sig.Status -ne "Valid") {
            return "Invalid - $($sig.Status)"
        } else {
            return "Not Signed"
        }
    }
    catch {
        return "Error checking signature"
    }
}

function Get-SRUMInformation {
    Write-Log "Retrieving SRUM information..."
    $results = @()

    try {
        $srumPath = "$env:SystemRoot\System32\srumbin.exe"
        if (Test-Path $srumPath) {
            $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processStartInfo.FileName = $srumPath
            $processStartInfo.Arguments = "/dump"
            $processStartInfo.RedirectStandardOutput = $true
            $processStartInfo.RedirectStandardError = $true
            $processStartInfo.UseShellExecute = $false
            $processStartInfo.CreateNoWindow = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processStartInfo
            $process.Start() | Out-Null
            $output = $process.StandardOutput.ReadToEnd()
            $errorOutput = $process.StandardError.ReadToEnd()
            $process.WaitForExit()

            if (-not [string]::IsNullOrEmpty($errorOutput)) {
                Write-Log "Error retrieving SRUM information: $errorOutput"
                return $results
            }

            $lines = $output -split "`r?`n"
            foreach ($line in $lines) {
                if ($line -match "Process:\s*(.+)\s+PID:\s*(\d+)\s+Time:\s*(.+)") {
                    $processName = $matches[1].Trim()
                    $pid = $matches[2].Trim()
                    $time = $matches[3].Trim()

                    $result = [PSCustomObject]@{
                        Source = "SRUM"
                        ProcessName = $processName
                        PID = $pid
                        Time = $time
                        Signature = "N/A"
                        FileExists = $false
                        ArtifactFile = $srumPath
                        SuspiciousActivity = "N/A"
                    }
                    $results += $result
                }
            }
        } else {
            Write-Log "SRUM binary not found at $srumPath"
        }
    }
    catch {
        Write-Log "Error retrieving SRUM information: $($_.Exception.Message)"
    }

    return $results
}

function Get-DPSInformation {
    Write-Log "Retrieving DPS information..."
    $results = @()

    try {
        $dpsPath = "$env:SystemRoot\System32\dps.exe"
        if (Test-Path $dpsPath) {
            $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processStartInfo.FileName = $dpsPath
            $processStartInfo.Arguments = "/dump"
            $processStartInfo.RedirectStandardOutput = $true
            $processStartInfo.RedirectStandardError = $true
            $processStartInfo.UseShellExecute = $false
            $processStartInfo.CreateNoWindow = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processStartInfo
            $process.Start() | Out-Null
            $output = $process.StandardOutput.ReadToEnd()
            $errorOutput = $process.StandardError.ReadToEnd()
            $process.WaitForExit()

            if (-not [string]::IsNullOrEmpty($errorOutput)) {
                Write-Log "Error retrieving DPS information: $errorOutput"
                return $results
            }

            $lines = $output -split "`r?`n"
            foreach ($line in $lines) {
                if ($line -match "Process:\s*(.+)\s+PID:\s*(\d+)\s+Time:\s*(.+)") {
                    $processName = $matches[1].Trim()
                    $pid = $matches[2].Trim()
                    $time = $matches[3].Trim()

                    $result = [PSCustomObject]@{
                        Source = "DPS"
                        ProcessName = $processName
                        PID = $pid
                        Time = $time
                        Signature = "N/A"
                        FileExists = $false
                        ArtifactFile = $dpsPath
                        SuspiciousActivity = "N/A"
                    }
                    $results += $result
                }
            }
        } else {
            Write-Log "DPS binary not found at $dpsPath"
        }
    }
    catch {
        Write-Log "Error retrieving DPS information: $($_.Exception.Message)"
    }

    return $results
}

function Get-BAMInformation {
    Write-Log "Retrieving BAM information..."
    $results = @()

    try {
        $bamPath = "$env:SystemRoot\System32\bam.exe"
        if (Test-Path $bamPath) {
            $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processStartInfo.FileName = $bamPath
            $processStartInfo.Arguments = "/dump"
            $processStartInfo.RedirectStandardOutput = $true
            $processStartInfo.RedirectStandardError = $true
            $processStartInfo.UseShellExecute = $false
            $processStartInfo.CreateNoWindow = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processStartInfo
            $process.Start() | Out-Null
            $output = $process.StandardOutput.ReadToEnd()
            $errorOutput = $process.StandardError.ReadToEnd()
            $process.WaitForExit()

            if (-not [string]::IsNullOrEmpty($errorOutput)) {
                Write-Log "Error retrieving BAM information: $errorOutput"
                return $results
            }

            $lines = $output -split "`r?`n"
            foreach ($line in $lines) {
                if ($line -match "Process:\s*(.+)\s+PID:\s*(\d+)\s+Time:\s*(.+)") {
                    $processName = $matches[1].Trim()
                    $pid = $matches[2].Trim()
                    $time = $matches[3].Trim()

                    $result = [PSCustomObject]@{
                        Source = "BAM"
                        ProcessName = $processName
                        PID = $pid
                        Time = $time
                        Signature = "N/A"
                        FileExists = $false
                        ArtifactFile = $bamPath
                        SuspiciousActivity = "N/A"
                    }
                    $results += $result
                }
            }
        } else {
            Write-Log "BAM binary not found at $bamPath"
        }
    }
    catch {
        Write-Log "Error retrieving BAM information: $($_.Exception.Message)"
    }

    return $results
}

function Get-JARInformation {
    Write-Log "Retrieving .jar file information..."
    $results = @()

    try {
        $jarPaths = @(
            "$env:ProgramFiles*\*\*.jar",
            "$env:ProgramFiles (x86)*\*\*.jar",
            "$env:LocalAppData\*\*.jar",
            "$env:UserProfile\Downloads\*\*.jar",
            "$env:Temp\*\*.jar"
        )

        foreach ($path in $jarPaths) {
            $files = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt (Get-Date).AddHours(-24) }
            foreach ($file in $files) {
                $result = [PSCustomObject]@{
                    Source = "JAR"
                    FilePath = $file.FullName
                    Time = $file.LastWriteTime
                    Signature = "N/A"
                    FileExists = $true
                    ArtifactFile = $file.FullName
                    SuspiciousActivity = "N/A"
                }
                $results += $result
            }
        }
    }
    catch {
        Write-Log "Error retrieving .jar file information: $($_.Exception.Message)"
    }

    return $results
}

Write-Log "Starting comprehensive forensics scan..."
Write-Log "Output file: $OutputFile"

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Warning "Not running as administrator. Some artifacts may not be accessible."
}

if (Test-Path $OutputFile) {
    try {
        Remove-Item $OutputFile -Force
        Write-Log "Cleared existing output file."
    }
    catch {
        Write-Log "Warning: Could not clear existing output file: $($_.Exception.Message)"
    }
}

Write-Log "Collecting data from SRUM, DPS, BAM, and .jar files..."

$Artifacts += Get-SRUMInformation
$Artifacts += Get-DPSInformation
$Artifacts += Get-BAMInformation
$Artifacts += Get-JARInformation

$uniqueResults = $Artifacts | Sort-Object ProcessName, Source, FilePath | Get-Unique -AsString

Write-Log "Writing results to output file..."
$header = "Source`tProcessName`tPID`tTime`tSignature`tFileExists`tArtifactFile`tSuspiciousActivity`tFilePath"
Add-Content -Path $OutputFile -Value "COMPREHENSIVE FORENSICS REPORT"
Add-Content -Path $OutputFile -Value "Generated: $(Get-Date)"
Add-Content -Path $OutputFile -Value "Scan Target: All drives including C:"
Add-Content -Path $OutputFile -Value ("=" * 80)
Add-Content -Path $OutputFile -Value $header

foreach ($result in $uniqueResults) {
    $filePath = if ($result.Source -eq "JAR") { $result.FilePath } else { "$env:SystemRoot\System32\$($result.ProcessName).exe" }
    $fileExists = Test-Path $filePath
    $signature = if ($fileExists) { Get-DigitalSignature -FilePath $filePath } else { "N/A" }

    $line = "$($result.Source)`t$($result.ProcessName)`t$($result.PID)`t$($result.Time)`t$signature`t$fileExists`t$($result.ArtifactFile)`t$($result.SuspiciousActivity)`t$($result.FilePath)"
    Add-Content -Path $OutputFile -Value $line
}

Write-Log "Scan completed!"
Write-Log "Total unique entries found: $($uniqueResults.Count)"
Write-Log "Results saved to: $OutputFile"

Write-Host "`n> SCAN COMPLETED <" -ForegroundColor Green
Write-Host "Total entries found: $($uniqueResults.Count)" -ForegroundColor Yellow
Write-Host "File also saved to: $OutputFile" -ForegroundColor Cyan