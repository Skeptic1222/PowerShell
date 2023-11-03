# Logwrite Function
Function LogWrite {
    Param (
        [string]$logstring, # The message to write to the log file and display in the console
        [switch]$overwrite, # If present, clears the log file before writing the message
        [switch]$complete, # If present, sets $LASTEXITCODE to 0
        [switch]$red, # If present, sets the console text color to red
        [switch]$green, # If present, sets the console text color to green
        [string]$StaticLogfileName = $null
    )

    # If $StaticLogfileName is defined, it will override the default behavior of using the script name for the log file
    # To use a custom log file name, uncomment the following line and set the value of $StaticLogfileName
    # Example: $StaticLogfileName = "MyCustomLog" to make the logfile be "c:\Batch\Results\MyCustomLog.log"
    # $StaticLogfileName = "MyCustomLog"

    # Determine the log file name
    $logfileName = if ($StaticLogfileName) {
        $StaticLogfileName 
    }
    else {
        [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath) 
    }
    $logfile = "C:\Batch\Results\$logfileName.log"

    if ($overwrite) {
        if (Test-Path $logfile) {
            Remove-Item $logfile
        }
        New-Item -Path $logfile -ItemType File -Force | Out-Null
    }
    elseif (-not [System.IO.File]::Exists($logfile)) {
        New-Item -Path $logfile -ItemType File -Force | Out-Null
    }
    
    $time = ("[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date))
    $formattedLogEntry = "$time $logstring"
    
    Add-Content $Logfile -Value $formattedLogEntry

    $foregroundColor = if ($red) {
        "Red" 
    }
    elseif ($green) {
        "Green" 
    }
    else {
        "White" 
    }
    Write-Host $formattedLogEntry -ForegroundColor $foregroundColor -NoNewline
    Write-Host -ForegroundColor $foregroundColor -NoNewline "`n" # To apply color and bold to the newline character as well
    
    if ($complete) {
        $LASTEXITCODE = 0
    }

    if ($env:CORESERVER) {
        $returnCode = if ($null -eq $LASTEXITCODE) {
            0 
        }
        else {
            $LASTEXITCODE 
        }
        $arguments = "-core=$($env:CORESERVER) -taskId=$($env:task_id) -retcode=$returnCode -message=`"$logstring`" -log=$logfile"
        Start-Process -FilePath "sendtaskstatus.exe" -ArgumentList $arguments
    }
    
    Start-Sleep -Seconds 1
}
$global:officeShortcutsPresent = $false
# Functions to get Office Info
function Get-OfficeInfoFromOSPP {
    $osppFolders = @(
        "${env:ProgramFiles}\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "${env:ProgramFiles}\Microsoft Office\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
    )

    foreach ($osppFolder in $osppFolders) {
        $osppScript = Join-Path $osppFolder "ospp.vbs"
        if (Test-Path $osppScript) {
            $output = & cscript.exe //NoLogo $osppScript /dstatus
            if ($output -match "LICENSE NAME: Office 16") {
                $version = "Microsoft Office 2016"
            }
            elseif ($output -match "LICENSE NAME: Office 21") {
                $version = "Microsoft Office 2021 LTSC"
            }
            else {
                continue
            }

            $skuId = "Unknown"
            if ($output -match "SKU ID: (\w+)") {
                if ($Matches.Count -gt 1) {
                    $skuId = $Matches[1]
                }
            }

            if ($skuId -match "ProPlus") {
                $edition = "Professional"
            }
            else {
                $edition = "Standard"
            }

            if ($osppFolder -match "x86") {
                $architecture = "32-bit"
            }
            else {
                $architecture = "64-bit"
            }

            return @{
                Version      = $version
                Edition      = $edition
                Architecture = $architecture
            }
        }
    }

    return $null
}

# Function to place the Outlook shortcut on desktop
function OutlookDesktopIcon {
    $sourceShortcutPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Outlook 2016.lnk"
    $destinationShortcutPath = "$env:Public\Desktop\Outlook 2016.lnk"
    if (Test-Path $sourceShortcutPath) {
        Copy-Item -Path $sourceShortcutPath -Destination $destinationShortcutPath -Force
    }
    else {
        Write-Warning "Source shortcut $sourceShortcutPath does not exist."
    }
}


# Function to Check if Office 2016 shortcuts exist
function CheckOfficeShortcuts {
    $officeApps = @("Word", "Excel", "PowerPoint", "Outlook")
    $global:officeShortcutsPresent = $true
    foreach ($app in $officeApps) {
        $shortcutFiles = Get-ChildItem -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\" -Include "*$app 2016.lnk" -File -Recurse -ErrorAction SilentlyContinue
        if ($null -eq $shortcutFiles -or $shortcutFiles.Count -eq 0) {
            $global:officeShortcutsPresent = $false
            break
        }
    }
}

# Function to remove Access Database Engine
function Remove-AccessDatabaseEngine {
    LogWrite "Checking for Access Database Engine"
    $dbEngine = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE (Name LIKE 'Microsoft Access database engine%')"
    if ($null -ne $dbEngine) {
        LogWrite "Found Access Database Engine. Starting uninstallation..."
        foreach ($product in $dbEngine) {
            $product.Uninstall()
            LogWrite "Uninstalled $($product.Name)"
        }
    }
    else {
        LogWrite "No Access Database Engine found."
    }
}

# Function to install Office 2016
function Install-Office2016 {
    LogWrite "Installing Office 2016 from local cache"

    # Terminate all Microsoft Office processes
    Get-Process WINWORD, EXCEL, POWERPNT, MSACCESS, OUTLOOK, ONENOTE, VISIO, PROJECT -ErrorAction SilentlyContinue | Stop-Process -Force

    LogWrite "All Microsoft Office processes terminated."

    Remove-AccessDatabaseEngine

    # Define the setup file path
    $setupFile = Join-Path -Path $PSScriptRoot -ChildPath "setup.exe"

    # Define the admin file path
    $adminFile = Join-Path -Path $PSScriptRoot -ChildPath "EMC_Office_2016_Standard.MSP"

    LogWrite "Running command: Start-Process $setupFile"

    try {
        $process = Start-Process -FilePath $setupFile -ArgumentList "/adminfile `"$adminFile`"" -PassThru -NoNewWindow -RedirectStandardError error.txt -RedirectStandardOutput output.txt
        $process | Wait-Process -Timeout 600
        OutlookDesktopIcon
        Start-Sleep -Seconds 30
        CheckOfficeShortcuts
    }
    catch {
        LogWrite $_.Exception.Message
        LogWrite "Error during installation from local cache. Trying UNC path installation."
        Install-Office2016-UNC
    }
}

function Install-Office2016-UNC {
    LogWrite "Installing Office 2016 from UNC path"

    # Terminate all Microsoft Office processes
    Get-Process WINWORD, EXCEL, POWERPNT, MSACCESS, OUTLOOK, ONENOTE, VISIO, PROJECT -ErrorAction SilentlyContinue | Stop-Process -Force

    LogWrite "All Microsoft Office processes terminated."

    Remove-AccessDatabaseEngine

    LogWrite "Running command: Start-Process $setupUNC"

    try {
        $process = Start-Process -FilePath $setupUNC -ArgumentList "/adminfile `"$adminUNC`"" -PassThru -NoNewWindow -RedirectStandardError error.txt -RedirectStandardOutput output.txt
        $process | Wait-Process -Timeout 600
        OutlookDesktopIcon
        Start-Sleep -Seconds 30
        CheckOfficeShortcuts
    }
    catch {
        LogWrite $_.Exception.Message
        LogWrite "Error during installation from UNC path. Aborting."
        exit 1
    }
}

# Function to Detect any remaining Office 2013 files
function DetectOffice2013Files {
    $office2013Files = Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office\Office15" -Recurse -ErrorAction SilentlyContinue

    if ($null -ne $office2013Files -and $office2013Files.Count -gt 0) {
        LogWrite "Found $($office2013Files.Count) Office 2013 files. Removing..."
        foreach ($file in $office2013Files) {
            Remove-Item -Path $file.FullName -Force
            LogWrite "Removed $file"
        }
    }
    else {
        LogWrite "No Office 2013 files found."
    }
}
# Log file path
$logfile = Join-Path -Path $PSScriptRoot -ChildPath "install.log"

# Get Office version
$officeInfo = Get-OfficeInfoFromOSPP
$version = $officeInfo.Version

CheckOfficeShortcuts

# If the version is not 2016 or shortcuts are missing, initiate installation process
if (($version -ne 'Microsoft Office 2016') -or ($officeShortcutsPresent) -eq ($false)) {
    LogWrite "Office 2016 not detected or shortcuts missing. Starting Installation."
    Install-Office2016
}

# If the version is 2016 and shortcuts are present, log the success
if (($version -eq 'Microsoft Office 2016') -and ($officeShortcutsPresent)) {
    LogWrite "Office 2016 detected along with shortcuts"
}
else {
    LogWrite "Office 2016 shortcuts not found, installation failed."
    exit 1
}

# Call the DetectOffice2013Files function to ensure no leftover 2013 files
LogWrite "Checking for any remaining Office 2013 files"
DetectOffice2013Files

# Re-check office version
$officeInfo = Get-OfficeInfoFromOSPP
$version = $officeInfo.Version
if ($version -match 'Microsoft Office 2016') {
    logwrite "Office 2016 detected, exiting"
    exit 0
}
else {
    logwrite "Office 2016 installation failed, please try again"
    exit 1
}
