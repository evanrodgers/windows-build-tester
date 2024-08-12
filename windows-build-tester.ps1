# User-configurable parameters
param(
    [int]$CPUThreadCount = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors,
    [int]$TestDuration = 15,  # Duration in seconds for each stress test
    [int]$AppWaitTime = 30,    # Wait time in seconds after launching/closing apps
    [string]$LogFile = "last_run_logs.txt",
    [string]$TempFilePath = "C:\Temp\StressTestFiles",
    [string]$TempFileDestination = "C:\Temp\StressTestFilesCopy",
    [string]$NetworkTestUrl = "https://mirror.us.leaseweb.net/ubuntu-cdimage/xubuntu/releases/24.04/release/xubuntu-24.04-desktop-amd64.iso",
    [string]$NetworkTestDestination = "C:\Temp\network_stress_test.zip"
)

# Global variable to control script execution
$global:continueTests = $true

# Function to handle script termination
function Stop-Script {
    $global:continueTests = $false
    Write-Host "`nStopping all tests and cleaning up..."
    Add-Content -Path $LogFile -Value "Script stopped by user. Cleaning up..."
    
    # Stop all running jobs
    Get-Job | Stop-Job
    Get-Job | Remove-Job
    
    # Close all applications launched by the script
    $apps | ForEach-Object {
        Get-Process -Name $_.Executable -ErrorAction SilentlyContinue | Stop-Process -Force
    }
    
    # Delete temporary files and directories
    if (Test-Path $TempFilePath) {
        Remove-Item $TempFilePath -Recurse -Force
    }
    if (Test-Path $TempFileDestination) {
        Remove-Item $TempFileDestination -Recurse -Force
    }
    if (Test-Path $NetworkTestDestination) {
        Remove-Item $NetworkTestDestination -Force
    }
    
    Write-Host "Cleanup completed. Exiting script."
    Add-Content -Path $LogFile -Value "Cleanup completed. Script exited."
    exit
}

# Function to check for Escape key press
function Check-ForEscapeKey {
    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq 'Escape') {
            return $true
        }
    }
    return $false
}

# Define applications to launch and close
$officeRoot = "C:\Program Files\Microsoft Office\root\Office16"
$edgeRoot = "C:\Program Files (x86)\Microsoft\Edge\Application"
$otherRoot = "C:\Windows\System32"
$ntRoot = "C:\Windows\System32"

$apps = @(
    @{Name = "Outlook"; Path = "$officeRoot\OUTLOOK.EXE"; Executable = "OUTLOOK.EXE"},
    @{Name = "Excel"; Path = "$officeRoot\EXCEL.EXE"; Executable = "EXCEL.EXE"},
    @{Name = "PowerPoint"; Path = "$officeRoot\POWERPNT.EXE"; Executable = "POWERPNT.EXE"},
    @{Name = "Microsoft Edge"; Path = "$edgeRoot\msedge.exe"; Executable = "msedge.exe"},
    @{Name = "Notepad"; Path = "$otherRoot\notepad.exe"; Executable = "notepad.exe"},
    @{Name = "WordPad"; Path = "$ntRoot\wordpad.exe"; Executable = "wordpad.exe"}
)

# Get machine name
$machineName = $env:COMPUTERNAME

# Set up logging
if (Test-Path $LogFile) {
    Remove-Item $LogFile
}

# User prompt
function Show-Menu {
    Clear-Host
    Write-Host "This script tests newly-built machines by launching and closing various Windows applications."
    Write-Host "This generates events in Aternity which can be further analyzed."
    Write-Host "`nThe script can also do basic hardware stress testing."
    Write-Host "`nMachine name: $machineName"
    Write-Host "`nCurrent settings:"
    Write-Host "CPU Thread Count: $CPUThreadCount"
    Write-Host "Test Duration: $TestDuration seconds"
    Write-Host "App Wait Time: $AppWaitTime seconds"
    Write-Host "`nSelect what you would like to do:"
    Write-Host "[1] Launch apps"
    Write-Host "[2] Hardware stress test"
    Write-Host "[3] Both"
    Write-Host "[4] Exit"
    Write-Host "`nPress Escape at any time to stop the script."
}

# Function to launch and close applications with error handling
function Launch-Apps {
    while ($global:continueTests) {
        foreach ($app in $apps) {
            if (Check-ForEscapeKey) { 
                Stop-Script
                return
            }
            try {
                Write-Host "Launching $($app.Name)... Press the ESC key to stop"
                Add-Content -Path $LogFile -Value "Launching $($app.Name)..."
                Start-Process $app.Path
                Start-Sleep -Seconds $AppWaitTime
                Write-Host "Waiting $AppWaitTime seconds... Press the ESC key to stop"
                Add-Content -Path $LogFile -Value "Waiting $AppWaitTime seconds..."
            }
            catch {
                Write-Host "Error launching $($app.Name): $_.Exception.Message"
                Add-Content -Path $LogFile -Value "Error launching $($app.Name): $_.Exception.Message"
            }
            
            try {
                Write-Host "Closing $($app.Name)... Press the ESC key to stop"
                Add-Content -Path $LogFile -Value "Closing $($app.Name)..."
                Stop-Process -Name $app.Executable -Force
                Start-Sleep -Seconds $AppWaitTime
                Write-Host "Waiting $AppWaitTime seconds... Press the ESC key to stop"
                Add-Content -Path $LogFile -Value "Waiting $AppWaitTime seconds..."
            }
            catch {
                Write-Host "Error closing $($app.Name): $_.Exception.Message"
                Add-Content -Path $LogFile -Value "Error closing $($app.Name): $_.Exception.Message"
            }
        }
    }
}

# Function to perform hardware stress tests with remaining time
function Run-StressTest {
    param (
        [string]$testName,
        [scriptblock]$testAction
    )

    Write-Host "Starting $testName... Press the ESC key to stop"
    Add-Content -Path $LogFile -Value "Starting $testName..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $testJob = Start-Job -ScriptBlock $testAction

    while ($stopwatch.Elapsed.TotalSeconds -lt $TestDuration -and $global:continueTests) {
        $remainingTime = [math]::Ceiling($TestDuration - $stopwatch.Elapsed.TotalSeconds)
        Write-Host "$remainingTime seconds remaining for $testName... Press the ESC key to stop"
        Add-Content -Path $LogFile -Value "$remainingTime seconds remaining for $testName..."
        Start-Sleep -Seconds 1
        
        if (Check-ForEscapeKey) {
            Stop-Script
            return
        }
    }

    # Stop the job and clean up
    $stopwatch.Stop()
    Stop-Job -Job $testJob
    Remove-Job -Job $testJob

    Write-Host "$testName completed."
    Add-Content -Path $LogFile -Value "$testName completed."
}

# Stress test scriptblocks
$cpuTest = {
    while ($true) {
        1..$using:CPUThreadCount | ForEach-Object { [Math]::Sqrt((Get-Random)) }
    }
}

$memoryTest = {
    $memoryList = @()
    while ($true) {
        $memoryList += ,@(0..1000000)
        Start-Sleep -Milliseconds 100
    }
}

$diskTest = {
    $sourceDir = $using:TempFilePath
    $destDir = $using:TempFileDestination

    # Ensure directories exist
    New-Item -ItemType Directory -Force -Path $sourceDir | Out-Null
    New-Item -ItemType Directory -Force -Path $destDir | Out-Null

    while ($true) {
        # Create 1000 100KB files
        1..1000 | ForEach-Object {
            $fileName = "file_$_.txt"
            $filePath = Join-Path $sourceDir $fileName
            $content = "A" * 102400  # 100KB of content
            [System.IO.File]::WriteAllText($filePath, $content)
        }

        # Copy files to destination
        Copy-Item "$sourceDir\*" -Destination $destDir -Force

        # Delete files from both locations
        Remove-Item "$sourceDir\*" -Force
        Remove-Item "$destDir\*" -Force
    }
}

$networkTest = {
    while ($true) {
        Invoke-WebRequest -Uri $using:NetworkTestUrl -OutFile $using:NetworkTestDestination
    }
}

# Main loop
while ($true) {
    Show-Menu
    $selection = Read-Host "Enter your choice (1-4)"
    
    switch ($selection) {
        1 { Launch-Apps }
        2 {
            while ($global:continueTests) {
                Run-StressTest -testName "CPU stress test" -testAction $cpuTest
                if (-not $global:continueTests) { break }
                Run-StressTest -testName "Memory stress test" -testAction $memoryTest
                if (-not $global:continueTests) { break }
                Run-StressTest -testName "Disk stress test" -testAction $diskTest
                if (-not $global:continueTests) { break }
                Run-StressTest -testName "Network stress test" -testAction $networkTest
            }
        }
        3 {
            while ($global:continueTests) {
                Launch-Apps
                if (-not $global:continueTests) { break }
                Run-StressTest -testName "CPU stress test" -testAction $cpuTest
                if (-not $global:continueTests) { break }
                Run-StressTest -testName "Memory stress test" -testAction $memoryTest
                if (-not $global:continueTests) { break }
                Run-StressTest -testName "Disk stress test" -testAction $diskTest
                if (-not $global:continueTests) { break }
                Run-StressTest -testName "Network stress test" -testAction $networkTest
            }
        }
        4 { Stop-Script }
        default { 
            Write-Host "Invalid selection, please try again." 
            Add-Content -Path $LogFile -Value "Invalid selection, please try again."
        }
    }
}