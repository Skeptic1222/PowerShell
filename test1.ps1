# LogWrite Function
Function LogWrite {
    Param (
        [string]$logstring, # The message to write to the log file and display in the console
        [switch]$overwrite, # If present, clears the log file before writing the message
        [switch]$complete, # If present, sets $LASTEXITCODE to 0
        [switch]$red, # If present, sets the console text color to red
        [switch]$green, # If present, sets the console text color to green
        [string]$StaticLogfileName = $null # If $StaticLogfileName is defined, it will override the default behavior of using the script name for the log file
    )

    # Determine the log file name
    $logfileName = if ($StaticLogfileName) { $StaticLogfileName  } else { [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)  }
    $logfile = "C:\Batch\Results\$logfileName.log"
    
    if ($overwrite) {
        if (Test-Path $logfile) { Remove-Item $logfile }
        New-Item -Path $logfile -ItemType File -Force | Out-Null
    }
    elseif (-not [System.IO.File]::Exists($logfile)) {
        New-Item -Path $logfile -ItemType File -Force | Out-Null
    }

    $time = ("[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date))
    $formattedLogEntry = "$time $logstring"
    Add-Content $Logfile -Value $formattedLogEntry

    $foregroundColor = if ($red) { "Red"  } elseif ($green) { "Green"  } else { "White"  }
    Write-Host $formattedLogEntry -ForegroundColor $foregroundColor -NoNewline
    Write-Host -ForegroundColor $foregroundColor -NoNewline "`n" # To apply color and bold to the newline character as well

    if ($complete) { $LASTEXITCODE = 0 }
}

# Function to install Office 2016 with retry logic
function Install-Office2016 {
    Param (
        [int]$retryCount = 3  # Number of retries
    )

    LogWrite "Installing Office 2016 from local cache"

    # Terminate all Microsoft Office processes
    Get-Process WINWORD, EXCEL, POWERPNT, MSACCESS, OUTLOOK, ONENOTE, VISIO, PROJECT -ErrorAction SilentlyContinue | Stop-Process -Force
    LogWrite "All Microsoft Office processes terminated."

    Remove-AccessDatabaseEngine

    # Define the setup file path
    $setupFile = Join-Path -Path $PSScriptRoot -ChildPath "setup.exe"

    # Define the admin file path
    $adminFile = Join-Path -Path $PSScriptRoot -ChildPath "EMC_Office_2016_Standard.MSP"

    for ($i=0; $i -lt $retryCount; $i++) {
        LogWrite "Attempt $($i + 1) of $retryCount to install Office 2016"
        try {
            LogWrite "Running command: Start-Process $setupFile"
            $process = Start-Process -FilePath $setupFile -ArgumentList "/adminfile `"$adminFile`"" -PassThru -NoNewWindow -RedirectStandardError error.txt -RedirectStandardOutput output.txt
            $process | Wait-Process -Timeout 800
            CheckOfficeInstallation
            Start-Sleep -Seconds 30
            

            # Verify if installation was successful
            if (($global:officeInstallationSuccessful) -eq ($true)) {
                LogWrite "Office 2016 installation successful on attempt $($i + 1)"
                $global:officeInstallationSuccessful = $true
                break  # Exit loop if installation was successful
            }
            else {
                LogWrite "Office 2016 installation failed on attempt $($i + 1)"
            }
        }
        catch {
            LogWrite $_.Exception.Message
            LogWrite "Error during installation from local cache. Retrying..."
        }
    }

    if (($global:officeInstallationSuccessful) -eq ($false)) {
        LogWrite "Office 2016 installation failed after $retryCount attempts."
        Install-Office2016-UNC
    }
}
# Function to install Office 2016 from UNC path with retry logic
function Install-Office2016-UNC {
    Param (
        [int]$retryCount = 3  # Number of retries
    )

    LogWrite "Installing Office 2016 from UNC path"
    $setupUNC = "\\isilon\Ivanti\Packages\Microsoft\Office\Office 2016 Standard\setup.exe"
    $adminUNC = "\\isilon\Ivanti\Packages\Microsoft\Office\Office 2016 Standard\Eisenhower_Standard_HideProgress_Testing.MSP"

    for ($i=0; $i -lt $retryCount; $i++) {
        LogWrite "Attempt $($i + 1) of $retryCount to install Office 2016 from UNC path"
        try {
            # Terminate all Microsoft Office processes
            Get-Process WINWORD, EXCEL, POWERPNT, MSACCESS, OUTLOOK, ONENOTE, VISIO, PROJECT -ErrorAction SilentlyContinue | Stop-Process -Force
            LogWrite "All Microsoft Office processes terminated."
            Remove-AccessDatabaseEngine
            LogWrite "Running command: Start-Process $setupUNC"
            $process = Start-Process -FilePath $setupUNC -ArgumentList "/adminfile `"$adminUNC`"" -PassThru -NoNewWindow -RedirectStandardError error.txt -RedirectStandardOutput output.txt
            $process | Wait-Process -Timeout 600
            CheckOfficeInstallation
            Start-Sleep -Seconds 30
            

            # Verify if installation was successful
            if (($global:officeInstallationSuccessful) -eq ($true)) {
                LogWrite "Office 2016 installation successful on attempt $($i + 1)"
                $global:officeInstallationSuccessful = $true
                break  # Exit loop if installation was successful
                
            }
            else {
                LogWrite "Office 2016 installation failed on attempt $($i + 1)"
            }
        }
        catch {
            LogWrite $_.Exception.Message
            LogWrite "Error during installation from UNC path. Retrying..."
        }
    }

    if (($global:officeInstallationSuccessful) -eq ($false)) {
        LogWrite "Office 2016 installation failed after $retryCount attempts."
        exit 1
    }
}
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

# Revised CheckOfficeInstallation function to check for Office 2016 installation and ensure Outlook desktop icon is present
function CheckOfficeInstallation {
    LogWrite "Checking for Office 2016 installation and ensuring Outlook desktop icon is present"
    $Global:officeInfo = Get-OfficeInfoFromOSPP
    $Global:version = $officeInfo.Version

    if (($version -ne 'Microsoft Office 2016') -or ($global:officeInstallationSuccessful -eq $false)) {
    LogWrite "Office 2016 not detected or shortcuts missing. "
    $global:officeInstallationSuccessful = $false
}

if (($version -eq 'Microsoft Office 2016') -and ($global:officeInstallationSuccessful)) {
    LogWrite "Office 2016 detected along with shortcuts"
    LogWrite "Office 2016 is installed"
    $global:officeInstallationSuccessful = $true
}
    
    $officeApps = @("Word", "Excel", "PowerPoint", "Outlook")
    $global:officeInstallationSuccessful = $false

    foreach ($app in $officeApps) {
        $shortcutFiles = Get-ChildItem -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\" -Include "*$app 2016.lnk" -File -Recurse -ErrorAction SilentlyContinue
        if ($null -eq $shortcutFiles -or $shortcutFiles.Count -eq 0) {
            $global:officeInstallationSuccessful = $false
            LogWrite "Shortcut for $app not found."
            break
        }
        else {
            LogWrite "Shortcut for $app found."
            $global:officeInstallationSuccessful = $true
        }
    }

    # Ensure Outlook desktop icon is present
    $sourceShortcutPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Outlook 2016.lnk"
    $destinationShortcutPath = "$env:Public\Desktop\Outlook 2016.lnk"
    if (Test-Path $sourceShortcutPath) {
        Copy-Item -Path $sourceShortcutPath -Destination $destinationShortcutPath -Force
        LogWrite "Outlook desktop icon is present."
        $global:officeInstallationSuccessful = $true
    }
    else {
        LogWrite "Source shortcut $sourceShortcutPath does not exist."
        $global:officeInstallationSuccessful = $false
    }
}

# Function to remove Access Database Engine
function Remove-AccessDatabaseEngine {
    LogWrite "Checking for Access Database Engine"
    $dbEngine = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE (Name LIKE 'Microsoft Access database engine%')"
    if ($null -ne $dbEngine) {
        LogWrite "Found Access Database Engine. Starting uninstallation..."
        foreach ($product in $dbEngine) {
            try {
                $product.Uninstall()
                LogWrite "Uninstalled $($product.Name)"
            }
            catch {
                LogWrite "Error uninstalling $($product.Name): $_.Exception.Message"
            }
        }
    }
    else {
        LogWrite "No Access Database Engine found."
    }
}

CheckOfficeInstallation

if (($global:officeInstallationSuccessful) -eq ($false)) {
Install-Office2016
} else {
LogWrite "Office 2016 is installed"
}
LogWrite "Checking for any remaining Office 2013 files"
DetectOffice2013Files
