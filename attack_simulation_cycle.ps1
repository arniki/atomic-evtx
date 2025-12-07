# Start the script block
$startTime = Get-Date
Set-Location "D:\Virtual Box\"

# Define the array of TTP IDs
$ttpArray = @("T1548.002-21")

# VM management functions
function Start-TargetVM {
    param (
        [int]$Timeout = 300,
        [int]$Interval = 5
    )

    .\VBoxManage.exe startvm VM-WIN10-TRGT --type headless
    $elapsedTime = 0

    while (-not (Test-VMReady)) {
        if ($elapsedTime -ge $Timeout) {
            Write-Host "Timed out waiting for the VM to boot." -ForegroundColor Red
            exit 1
        }
        Write-Host "   [?] Waiting for the VM to boot..." -ForegroundColor Yellow
        Start-Sleep -Seconds $Interval
        $elapsedTime += $Interval
    }
    
    Write-Host "[+] Target Machine has fully booted" -ForegroundColor Green
}

function Test-VMReady {
    $result = .\VBoxManage.exe guestproperty get "VM-WIN10-TRGT" "/VirtualBox/GuestInfo/OS/LoggedInUsers"
    return ($result -notmatch "No value set!")
}

function Stop-TargetVM {
    $vmState = Get-VMState

    switch -Regex ($vmState) {
        'running|paused|starting' {
            Write-Host "[+] VM is $vmState. Powering off..." -ForegroundColor Yellow
            .\VBoxManage.exe controlvm VM-WIN10-TRGT poweroff
        }
        'saved' {
            Write-Host "[+] VM is in saved state. Discarding..." -ForegroundColor Yellow
            .\VBoxManage.exe discardstate VM-WIN10-TRGT
        }
    }

    Wait-For-VMState -TargetState 'poweroff'
    Write-Host "[+] VM is powered off and ready to restore the snapshot." -ForegroundColor Green
}

function Get-VMState {
    return (.\VBoxManage.exe showvminfo "VM-WIN10-TRGT" --machinereadable | Select-String "VMState=").ToString().Split('=')[1].Trim('"')
}

function Wait-For-VMState {
    param (
        [string]$TargetState,
        [int]$Timeout = 60,
        [int]$Interval = 3
    )

    $elapsedTime = 0
    do {
        Start-Sleep -Seconds $Interval
        $currentState = Get-VMState
        $elapsedTime += $Interval

        if ($elapsedTime -ge $Timeout) {
            Write-Host "Timed out waiting for VM state $TargetState" -ForegroundColor Red
            exit 1
        }
    } while ($currentState -ne $TargetState)
}

function New-RetryPSSession {
    param (
        [string]$ComputerName,
        [PSCredential]$Credential,
        [int]$MaxAttempts = 5,
        [int]$RetryIntervalSeconds = 10
    )

    Add-Type -AssemblyName System.Windows.Forms

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Write-Host "Attempt $attempt to establish remote session..." -ForegroundColor Yellow
            $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
            Write-Host "Remote session established successfully." -ForegroundColor Green
            return $session
        }
        catch {
            Write-Host "Failed to establish remote session (Attempt $attempt of $MaxAttempts)" -ForegroundColor Red
            Write-Host "Error: $_" -ForegroundColor Red
            
            if ($attempt -lt $MaxAttempts) {
                Write-Host "Retrying in $RetryIntervalSeconds seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RetryIntervalSeconds
            }
            else {
                Write-Host "Max attempts reached. Unable to establish remote session." -ForegroundColor Red
                
                # Create and show Windows Notification
                $notification = New-Object System.Windows.Forms.NotifyIcon
                $notification.Icon = [System.Drawing.SystemIcons]::Error
                $notification.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Error
                $notification.BalloonTipTitle = "Remote Session Connection Failed"
                $notification.BalloonTipText = "Failed to establish a remote session with $ComputerName after $MaxAttempts attempts."
                $notification.Visible = $true
                $notification.ShowBalloonTip(5000)

                # Exit the script
                exit
            }
        }
    }
}

# Preliminary configurations
Enable-PSRemoting -SkipNetworkProfileCheck -Force
winrm set winrm/config/client/auth '@{Basic="true"}'
winrm set winrm/config/service/auth '@{Basic="true"}'
winrm set winrm/config/client '@{AllowUnencrypted="true"}'
winrm set winrm/config/service '@{AllowUnencrypted="true"}'
Set-NetFirewallRule -Name 'WINRM-HTTP-In-TCP' -RemoteAddress Any
Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force

Import-Module "C:\AtomicRedTeam\invoke-atomicredteam\Invoke-AtomicRedTeam.psd1" -Force
$PSDefaultParameterValues = @{"Invoke-AtomicTest:PathToAtomicsFolder"="C:\AtomicRedTeam\atomics"}

$targetComputer = "Server002"
$credential = New-Object System.Management.Automation.PSCredential ("admin_test", (ConvertTo-SecureString "123123" -AsPlainText -Force))

foreach ($ttp in $ttpArray) {
    $startTime_ttp = Get-Date
    $outputFolder = "D:\atomic_results\$ttp"

    Write-Host "[*] Processing TTP: $ttp" -ForegroundColor Cyan

    Start-TargetVM
    Start-Sleep -Seconds 30

    $session = New-RetryPSSession -ComputerName $targetComputer -Credential $credential
    if ($null -eq $session) {
        Write-Host "Skipping TTP $ttp due to connection failure." -ForegroundColor Red
        continue
    }

    New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
    Write-Host "Output folder created at $outputFolder" -ForegroundColor Green

    Write-Host "[*] Executing Pre-req installation" -ForegroundColor Yellow
    Invoke-AtomicTest $ttp -GetPrereqs -Session $session -ExecutionLogPath "$outputFolder\$ttp.csv" *>&1 | Tee-Object "$outputFolder\$ttp.txt" -Append

    Write-Host "[*] Clearing pre-attack logs" -ForegroundColor Yellow
    Invoke-Command -Session $session -ScriptBlock {
        @("Microsoft-Windows-Sysmon/Operational", "Application", "System", "Security", "Windows PowerShell") | ForEach-Object {
            Write-Host "Clearing $_ logs..."
            wevtutil cl $_
        }
    }

    Write-Host "[*] Executing TTP" -ForegroundColor Yellow
    Invoke-AtomicTest $ttp -Session $session -ExecutionLogPath "$outputFolder\$ttp.csv" *>&1 | Tee-Object "$outputFolder\$ttp.txt" -Append
    Start-Sleep -Seconds 30
    Write-Host "[*] Exporting logs" -ForegroundColor Yellow
    @("Microsoft-Windows-Sysmon/Operational", "Application", "System", "Security", "Windows PowerShell") | ForEach-Object {
        $logName = $_
        $outputFile = Join-Path $outputFolder "$ttp`_$($logName -replace '/', '_').evtx"

        try {
            Invoke-Command -Session $session -ScriptBlock {
                param($LogName)
                $tempPath = Join-Path $env:TEMP "TempLog.evtx"
                wevtutil epl $LogName $tempPath
                Get-Content -Path $tempPath -Raw -Encoding Byte
                Remove-Item -Path $tempPath -Force
            } -ArgumentList $logName | Set-Content -Path $outputFile -Encoding Byte

            Write-Host "Log $logName exported successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "Error exporting $logName : $_" -ForegroundColor Red
        }
    }

    Remove-PSSession $session

    Stop-TargetVM
    Write-Host "[*] Restoring backup" -ForegroundColor Yellow
    .\VBoxManage.exe snapshot VM-WIN10-TRGT restorecurrent

    $executionTime_ttp = (Get-Date) - $startTime_ttp
    Write-Host "TTP $ttp completed. Execution time: $executionTime_ttp" -ForegroundColor Cyan
    Write-Host "--------------------------------------------" -ForegroundColor Cyan
}

$totalExecutionTime = (Get-Date) - $startTime
Write-Host "Total script execution time: $totalExecutionTime" -ForegroundColor Green
