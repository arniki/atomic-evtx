# Configuration
$EvtxECmdPath = ".\EvtxECmd.exe"  # Path to EvtxECmd executable
$RootFolderPath = ".\atomic-evtx\testing"   # Root folder path

# Function to create formatted log messages
function Write-CustomLog {
    param(
        [string]$Message,
        [string]$Type = 'INFO'  # Default type is INFO
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Type) {
        'INFO'    { 'Cyan' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'SCAN'    { 'Yellow' }
        default   { 'White' }
    }
    
    Write-Host "[$timestamp] [$Type] $Message" -ForegroundColor $color
}

# Function to ensure directory exists
function Ensure-DirectoryExists {
    param(
        [string]$Path
    )
    
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-CustomLog "Created directory: $Path" -Type 'INFO'
    }
}

# Function to scan and report EVTX files
function Get-EVTXFilesReport {
    param(
        [string]$RootPath
    )
    
    Write-CustomLog "Starting pre-conversion scanning phase..." -Type 'SCAN'
    Write-CustomLog "Scanning directory: $RootPath" -Type 'SCAN'
    
    # Get all EVTX files recursively
    $evtxFiles = Get-ChildItem -Path $RootPath -Filter "*.evtx" -Recurse
    
    # Group files by directory for reporting
    $filesByDirectory = $evtxFiles | Group-Object DirectoryName
    
    # Print summary header
    Write-CustomLog "=== EVTX Files Scan Report ===" -Type 'SCAN'
    Write-CustomLog "Total EVTX files found: $($evtxFiles.Count)" -Type 'SCAN'
    Write-CustomLog "Files found in $($filesByDirectory.Count) directories" -Type 'SCAN'
    Write-CustomLog "-------------------------" -Type 'SCAN'
    
    # Print detailed directory breakdown
    foreach ($dirGroup in $filesByDirectory) {
        Write-CustomLog "Directory: $($dirGroup.Name)" -Type 'SCAN'
        Write-CustomLog "Files found: $($dirGroup.Count)" -Type 'SCAN'
        
        # List each file in the directory
        foreach ($file in $dirGroup.Group) {
            Write-CustomLog "  - $($file.Name)" -Type 'SCAN'
        }
        Write-CustomLog "-------------------------" -Type 'SCAN'
    }
    
    # Return files for later processing
    return $evtxFiles
}

# Main script execution
try {
    Write-CustomLog "Starting EVTX processing script..." -Type 'INFO'
    Write-CustomLog "EvtxECmd Path: $EvtxECmdPath" -Type 'INFO'
    Write-CustomLog "Root Folder Path: $RootFolderPath" -Type 'INFO'
    
    # Verify EvtxECmd exists
    if (-not (Test-Path -Path $EvtxECmdPath)) {
        Write-CustomLog "EvtxECmd executable not found at: $EvtxECmdPath" -Type 'ERROR'
        exit
    }
    
    # Verify root folder exists
    if (-not (Test-Path -Path $RootFolderPath)) {
        Write-CustomLog "Root folder not found at: $RootFolderPath" -Type 'ERROR'
        exit
    }
    
    # First phase: Scan and report
    $evtxFiles = Get-EVTXFilesReport -RootPath $RootFolderPath
    $totalFiles = $evtxFiles.Count
    
    if ($totalFiles -eq 0) {
        Write-CustomLog "No EVTX files found in the specified directory" -Type 'ERROR'
        exit
    }
    
    # Ask for confirmation to proceed
    Write-CustomLog "Do you want to proceed with converting $totalFiles EVTX files? (Y/N)" -Type 'INFO'
    $confirmation = Read-Host
    
    if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
        Write-CustomLog "Conversion cancelled by user" -Type 'INFO'
        exit
    }
    
    Write-CustomLog "Starting conversion process..." -Type 'INFO'
    
    # Counter for progress tracking
    $processedFiles = 0
    
    foreach ($evtxFile in $evtxFiles) {
        $processedFiles++
        $progress = [math]::Round(($processedFiles / $totalFiles) * 100, 2)
        
        # Get the directory where the EVTX file is located
        $sourceDir = $evtxFile.DirectoryName
        $evtxFileName = $evtxFile.Name
        $jsonFileName = $evtxFile.BaseName + ".json"
        
        # Create 'json' subdirectory in the same folder as the EVTX file
        $jsonDir = Join-Path -Path $sourceDir -ChildPath "json"
        Ensure-DirectoryExists -Path $jsonDir
        
        Write-CustomLog "Processing ($processedFiles/$totalFiles - $progress%) : $evtxFileName" -Type 'INFO'
        
        try {
            # Construct the full command with proper arguments
            $arguments = @(
                "-f",
                "`"$($evtxFile.FullName)`"",
                "--json",
                "`"$jsonDir`"",
                "--fj",
                "--jsonf",
                "`"$jsonFileName`""
            )
            
            # Execute the command and capture output
            $process = Start-Process -FilePath $EvtxECmdPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru
            
            if ($process.ExitCode -eq 0) {
                Write-CustomLog "Successfully converted: $evtxFileName" -Type 'SUCCESS'
            }
            else {
                Write-CustomLog "Conversion failed for $evtxFileName with exit code $($process.ExitCode)" -Type 'ERROR'
            }
        }
        catch {
            Write-CustomLog "Error converting $($evtxFile.Name): $($_.Exception.Message)" -Type 'ERROR'
        }
    }
    
    Write-CustomLog "Conversion process completed. Processed $processedFiles files." -Type 'SUCCESS'
}
catch {
    Write-CustomLog "Fatal error in main script execution: $($_.Exception.Message)" -Type 'ERROR'
}
finally {
    Write-CustomLog "Script execution finished" -Type 'INFO'
}