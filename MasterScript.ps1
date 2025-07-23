# ___ MASTER SCRIPT 126 ___

# Requires PowerShell version 5.1

param (
    [switch]$skipAD,
    [switch]$fullexcel,
    [switch]$xlsonly
)

$StartTime = Get-Date
Write-Host " "
Write-Host "Script started at: $StartTime"
Write-Host " "

# Check if the script is running with administrative privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Get domain credentials
$cred = Get-Credential -Message "Enter credentials for domain admin"

if (-not $isAdmin) {
    try {
        # Relaunch the script with elevated privileges using provided credentials in an interactive session
        Start-Process powershell -Credential $cred -ArgumentList "-NoProfile -ExecutionPolicy Bypass -NoExit -File `"$PSCommandPath`"" -Verb RunAs
        # Exit the current non-elevated script
        exit
    }
    catch {
        Write-Host "Failed to elevate privileges. Please ensure the provided credentials have administrative rights."
        exit
    }
}

# ------------------- Initialize Variables -------------------
$DataRootPath = "C:\data\PCscript\Data"
$PclistRootPath = "C:\data\PCscript\pclists"
$ExcelFilePath = "C:\data\PCscript\pc-overview.xlsx"
$FileAgeThresholdDays = 14

# Ensure directories and files exist
$null = New-Item -ItemType Directory -Path $DataRootPath -Force
$null = New-Item -ItemType Directory -Path $PclistRootPath -Force

# Create files only if they don't exist, do not empty existing ones
if (-not (Test-Path "$PclistRootPath\pclist-noWinRM.txt")) {
    $null = New-Item -ItemType File -Path "$PclistRootPath\pclist-noWinRM.txt"
}
if (-not (Test-Path "$PclistRootPath\pclist-offline.txt")) {
    $null = New-Item -ItemType File -Path "$PclistRootPath\pclist-offline.txt"
}
if (-not (Test-Path "$PclistRootPath\pclist-to-skip.txt")) {
    $null = New-Item -ItemType File -Path "$PclistRootPath\pclist-to-skip.txt"
}

# ------------------- Script Start -------------------


# Check for ActiveDirectory module
Write-Host " - Checking AD module"
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
# was :    Write-Warning "ActiveDirectory module not installed. Run Install-Module -Name ActiveDirectory -Scope CurrentUser"
	Write-Warning "ActiveDirectory module not installed."
	Install-Module -Name ActiveDirectory -Scope CurrentUser
}

# Check for ImportExcel module
Write-Host " - Checking ImportExcel module"
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
# was :    Write-Warning "ImportExcel module not installed. Run Install-Module -Name ImportExcel -Scope CurrentUser"
    Write-Warning "ImportExcel module not installed."
    Install-Module -Name ImportExcel -Scope CurrentUser
}

# Check if the Excel file exists
if (-not (Test-Path $ExcelFilePath)) {
    # Create an empty worksheet (single sheet, no data)
    $null | Export-Excel -Path $ExcelFilePath -WorksheetName "everything"
}

if (-not $xlsonly) {
# Get computers from AD unless -skipAD is specified
$computers = if (-not $skipAD) {
    Write-Host " - Getting computers from AD"
    Get-ADComputer -Filter {Enabled -eq $true} -Credential $cred | Select-Object -ExpandProperty Name | Sort-Object
} else {
    Write-Host " - Reading computers from file"
    Get-Content "$PclistRootPath\pclist-AD.txt" -ErrorAction SilentlyContinue
}

# Read computers to skip
Write-Host " - Reading computers to skip from $PclistRootPath\pclist-to-skip.txt"
$computersToSkip = Get-Content "$PclistRootPath\pclist-to-skip.txt" -ErrorAction SilentlyContinue | Where-Object { $_ -ne "" }

# Filter out computers to skip
if ($computersToSkip) {
    Write-Host " - Filtering out $($computersToSkip.Count) computers from the list."
    $computers = $computers | Where-Object { $_ -notin $computersToSkip }
}

if ($computers) {
    $computers | Out-File -FilePath "$PclistRootPath\pclist-AD.txt" -Force
} else {
    Write-Host "No computers found after filtering. Exiting."
    exit
}

# Create subfolders for each command
Write-Host " - Checking if subfolders exist"
$CommandNames = @(
    "ComputerInfo", "LocalAdmins", "FileSystemDrives", "Services", "Logs",
    "Processes", "FirstRun", "NetworkAdapters", "DiskTotalSize",
    "DiskFreeSize", "InstalledModules", "LocalUsers", "BinSize", "SWList", "Temps", "TeamviewerID" 
)
foreach ($subfolder in $CommandNames) {
    $folderPath = "$DataRootPath\$subfolder"
    $null = New-Item -ItemType Directory -Path $folderPath -Force
    try {
        $testFile = "$folderPath\test-write.txt"
        "Test" | Out-File -FilePath $testFile -Force -ErrorAction Stop
        Remove-Item -Path $testFile -Force
    } catch {
        Write-Warning "Failed to write to $folderPath. Check permissions."
    }
}

# --- Core Logic: Loop through computers and collect data sequentially ---

# Define a simple function to check file age (can be reused)
function Test-FileAge {
    param (
        [string]$FilePath,
        [int]$Days
    )
    if (-not (Test-Path $FilePath)) {
        return $true
    }
    $file = Get-Item $FilePath
    $age = (Get-Date) - $file.LastWriteTime
    return $age.Days -gt $Days
}

function Temps {
	param(
		$ComputerName,
		$Cred
	)
	$paths = @(
		"C:\Windows\Temp",
        "$env:SystemDrive\Users\*\AppData\Local\Temp"
	)
    $fileCount = 0
    foreach ($path in $paths) {
		$count = Get-ChildItem -Path $path -File -Recurse -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
        $fileCount += $count
    }
    [PSCustomObject]@{
		TempFileCount = $fileCount
	}
}

# Define commands hashtable for remote execution
$Commands = @{
    "ComputerInfo"     = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-ComputerInfo -ErrorAction SilentlyContinue } -ErrorAction SilentlyContinue }
    "LocalAdmins"      = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name } -ErrorAction SilentlyContinue }
    "FileSystemDrives" = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Select-Object Name, Used, Free, Description, Root } -ErrorAction SilentlyContinue }
    "Services"         = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-Service -ErrorAction SilentlyContinue | Select-Object Name, DisplayName, Status, StartType } -ErrorAction SilentlyContinue }
    "Logs"             = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-EventLog -LogName System -EntryType Error -Newest 100 -ErrorAction SilentlyContinue | Select-Object TimeGenerated, EntryType, Source, Message } -ErrorAction SilentlyContinue }
    "Processes"        = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-Process -ErrorAction SilentlyContinue | Select-Object Name, CPU } -ErrorAction SilentlyContinue }
    "FirstRun"         = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { (Get-ChildItem 'C:\Windows\Logs\DirectX.log' -ErrorAction SilentlyContinue).LastWriteTime } -ErrorAction SilentlyContinue }
    "NetworkAdapters"  = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-NetAdapter -ErrorAction SilentlyContinue | Select-Object Name, Status, LinkSpeed, MacAddress } -ErrorAction SilentlyContinue }
    "DiskTotalSize"    = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue | ForEach-Object { [math]::Round($_.Size / 1GB, 1) } } -ErrorAction SilentlyContinue }
    "DiskFreeSize"     = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue | ForEach-Object { [math]::Round($_.FreeSpace / 1GB, 1) } } -ErrorAction SilentlyContinue }
    "InstalledModules" = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-InstalledModule -ErrorAction SilentlyContinue | Select-Object Name, Version } -ErrorAction SilentlyContinue }
    "LocalUsers"       = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { Get-LocalUser -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name } -ErrorAction SilentlyContinue }
    "BinSize"          = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock {(Get-ChildItem -LiteralPath 'C:\$Recycle.Bin' -File -Force -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum/1000000; " mb" } -ErrorAction SilentlyContinue}
    "SWList"           = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock { $apps1 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*
                                $apps2 = Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*
                                $allApps = $apps1 + $apps2
                                $allApps | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Where-Object { $_.DisplayName }
                                } -ErrorAction SilentlyContinue }
    "Temps"            = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock {
							$paths = @(
								"C:\Windows\Temp",
								"$env:SystemDrive\Users\*\AppData\Local\Temp"
							)
							$fileCount = 0
							foreach ($path in $paths) {
								$count = Get-ChildItem -Path $path -File -Recurse -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
								$fileCount += $count
							}
							[PSCustomObject]@{
								TempFileCount = $fileCount
							}
						}
						}
    "TeamviewerID"     = { param($ComputerName, $Cred) Invoke-Command -ComputerName $ComputerName -Credential $Cred -ScriptBlock {(Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\TeamViewer' -Name ClientID).ClientID} -ErrorAction SilentlyContinue}
}

$offlineComputers = @()
$noWinRMComputers = @()

foreach ($computer in $computers) {
    Write-Host "Processing computer: $computer"
    $computerStatus = "success"
    $computerMessage = ""

    # Test if computer is online and WinRM is available
    try {
        $pingResult = Test-Connection -ComputerName $computer -Count 1 -Quiet -ErrorAction Stop
        if (-not $pingResult) {
            $computerStatus = "offline"
            $computerMessage = "${computer}: is offline"
        } else {
            # Attempt a simple WinRM command to check connectivity
            Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock { $null } -ErrorAction Stop
        }
    } catch {
        if ($_.Exception.Message -like "*WinRM*" -or $_.Exception.Message -like "*access is denied*") {
            $computerStatus = "noWinRM"
            $computerMessage = "${computer}: WinRM not available or access denied. $($_.Exception.Message)"
        } else {
            $computerStatus = "error"
            $computerMessage = "${computer}: An unexpected error occurred during connectivity check. $($_.Exception.Message)"
        }
    }

    if ($computerStatus -ne "success") {
        Write-Warning $computerMessage
        if ($computerStatus -eq "offline") {
            $offlineComputers += $computer
        } elseif ($computerStatus -eq "noWinRM") {
            $noWinRMComputers += $computer
        }
        continue # Skip to next computer if not online or WinRM issue
    }

    foreach ($commandName in $CommandNames) {
        $filePath = "$DataRootPath\$commandName\$computer.txt"
        if (Test-FileAge -FilePath $filePath -Days $FileAgeThresholdDays) {
            try {
                $scriptBlock = $Commands[$commandName]
                $data = & $scriptBlock $computer $cred # Pass computer name AND credentials to the scriptblock
                if ($data) {
                    # Convert objects to a formatted string for consistent text file output
                    if ($commandName -in @("FileSystemDrives", "Services", "Logs", "Processes", "NetworkAdapters", "InstalledModules", "SWList")) {
                        $data | Format-Table -AutoSize | Out-String | Out-File -FilePath $filePath -Force
                    } else {
                        $data | Out-File -FilePath $filePath -Force
                    }
                    $computerMessage += "Collected $commandName; "
                } else {
                    $computerMessage += "No $commandName data returned; "
                }
            } catch {
                $computerMessage += "Failed to collect $commandName $($_.Exception.Message); "
            }
        }
    }
    Write-Host "$computer $computerMessage"
}

# Log results and update pclist files
if ($offlineComputers) {
    $offlineComputers | Out-File -FilePath "$PclistRootPath\pclist-offline.txt" -Force
    Write-Host "Offline computers logged to $PclistRootPath\pclist-offline.txt"
}
if ($noWinRMComputers) {
    $noWinRMComputers | Out-File -FilePath "$PclistRootPath\pclist-noWinRM.txt" -Force
    Write-Host "No WinRM computers logged to $PclistRootPath\pclist-noWinRM.txt"
}
}  # this is the end of the xlsonly switch


# ____________________________________ ____ EXCEL Export Section ____ ____________________________________

Write-Host " "
Write-Host "______ Starting Excel Export ______"

# Export pclist files to separate sheets
$pclistFiles = @(
    "$PclistRootPath\pclist-AD.txt",
    "$PclistRootPath\pclist-offline.txt",
    "$PclistRootPath\pclist-noWinRM.txt",
    "$PclistRootPath\pclist-to-skip.txt"
)

foreach ($file in $pclistFiles) {
    if (Test-Path $file) {
        $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        Write-Host "Exporting $sheetName to Excel..."
        try {
            Get-Content $file | Export-Excel -Path $excelFilePath -WorksheetName $sheetName -ClearSheet -AutoSize -ErrorAction Stop
        } catch {
            Write-Warning "Failed to export $sheetName to Excel: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "Pclist file not found: $file. Skipping export."
    }
}

# Update 'everything' sheet with collected data
Write-Host "Updating 'everything' sheet with collected data..."

# Define column headers for the 'everything' sheet
$headers = @(
    "Computer", "PSComputerName", "WindowsEditionId", "WindowsInstallationType", "WindowsInstallDateFromRegistry",
    "WindowsProductName", "WindowsRegisteredOrganization", "WindowsRegisteredOwner", "BiosCaption",
    "BiosFirmwareType", "BiosSeralNumber", "BiosSMBIOSBIOSVersion", "CsCurrentTimeZone", "CsDaylightInEffect",
    "CsDomain", "CsDomainRole", "CsEnableDaylightSavingsTime", "CsManufacturer", "CsModel",
    "CsNumberOfLogicalProcessors", "CsNumberOfProcessors", "CsPCSystemType", "CsPowerOnPasswordStatus",
    "CsSystemType", "CsTotalPhysicalMemory", "CsPhyicallyInstalledMemory", "CsUserName", "OsBuildNumber",
    "OsCountryCode", "OsCurrentTimeZone", "OsLocale", "OsUptime", "OsCodeSet", "OsInstallDate",
    "OsMuiLanguages", "OsNumberOfUsers", "OsOrganization", "OsArchitecture", "OsLanguage", "OsProductType",
    "OsRegisteredUser", "KeyboardLayout", "PowerPlatformRole", "A", "DiskFreeSize", "DiskTotalSize",
    "FirstRun", "LogsErrorsWarnings", "ProcessesCount", "ServicesCount"
)

# Initialize $everythingSheet as an ArrayList
$everythingSheet = New-Object System.Collections.ArrayList

# Read the existing 'everything' sheet and add to ArrayList
$importedData = Import-Excel -Path $excelFilePath -WorksheetName "everything" -ErrorAction SilentlyContinue
if ($importedData) {
    $null = $everythingSheet.AddRange($importedData)
}

# If 'everything' sheet doesn't exist or is empty, create a basic structure
if ($everythingSheet.Count -eq 0) {
    Write-Host "'everything' sheet not found or empty. Creating a new one."
    $newRow = [PSCustomObject]@{
        Computer = ""
        PSComputerName = ""
        WindowsEditionId = ""
        WindowsInstallationType = ""
        WindowsInstallDateFromRegistry = ""
        WindowsProductName = ""
        WindowsRegisteredOrganization = ""
        WindowsRegisteredOwner = ""
        BiosCaption = ""
        BiosFirmwareType = ""
        BiosSeralNumber = ""
        BiosSMBIOSBIOSVersion = ""
        CsCurrentTimeZone = ""
        CsDaylightInEffect = ""
        CsDomain = ""
        CsDomainRole = ""
        CsEnableDaylightSavingsTime = ""
        CsManufacturer = ""
        CsModel = ""
        CsNumberOfLogicalProcessors = ""
        CsNumberOfProcessors = ""
        CsPCSystemType = ""
        CsPowerOnPasswordStatus = ""
        CsSystemType = ""
        CsTotalPhysicalMemory = ""
        CsPhyicallyInstalledMemory = ""
        CsUserName = ""
        OsBuildNumber = ""
        OsCountryCode = ""
        OsCurrentTimeZone = ""
        OsLocale = ""
        OsUptime = ""
        OsCodeSet = ""
        OsInstallDate = ""
        OsMuiLanguages = ""
        OsNumberOfUsers = ""
        OsOrganization = ""
        OsArchitecture = ""
        OsLanguage = ""
        OsProductType = ""
        OsRegisteredUser = ""
        KeyboardLayout = ""
        PowerPlatformRole = ""
        A = ""
        DiskFreeSize = ""
        DiskTotalSize = ""
        FirstRun = ""
        LogsErrorsWarnings = ""
        ProcessesCount = ""
        ServicesCount = ""
		Temps = ""
    }
    $null = $everythingSheet.Add($newRow)
    try {
        $everythingSheet | Export-Excel -Path $excelFilePath -WorksheetName "everything" -ClearSheet -AutoSize -ErrorAction Stop
    } catch {
        Write-Warning "Failed to create 'everything' sheet: $($_.Exception.Message)"
    }
}

foreach ($computer in $computers) {
    Write-Host "Processing data for $computer for 'everything' sheet."
    $computerRow = $everythingSheet | Where-Object { $_.Computer -eq $computer } | Select-Object -First 1

    if (-not $computerRow) {
        Write-Host "Computer $computer not found in 'everything' sheet. Adding new row."
        $newRow = [PSCustomObject]@{
            Computer = $computer
            PSComputerName = ""
            WindowsEditionId = ""
            WindowsInstallationType = ""
            WindowsInstallDateFromRegistry = ""
            WindowsProductName = ""
            WindowsRegisteredOrganization = ""
            WindowsRegisteredOwner = ""
            BiosCaption = ""
            BiosFirmwareType = ""
            BiosSeralNumber = ""
            BiosSMBIOSBIOSVersion = ""
            CsCurrentTimeZone = ""
            CsDaylightInEffect = ""
            CsDomain = ""
            CsDomainRole = ""
            CsEnableDaylightSavingsTime = ""
            CsManufacturer = ""
            CsModel = ""
            CsNumberOfLogicalProcessors = ""
            CsNumberOfProcessors = ""
            CsPCSystemType = ""
            CsPowerOnPasswordStatus = ""
            CsSystemType = ""
            CsTotalPhysicalMemory = ""
            CsPhyicallyInstalledMemory = ""
            CsUserName = ""
            OsBuildNumber = ""
            OsCountryCode = ""
            OsCurrentTimeZone = ""
            OsLocale = ""
            OsUptime = ""
            OsCodeSet = ""
            OsInstallDate = ""
            OsMuiLanguages = ""
            OsNumberOfUsers = ""
            OsOrganization = ""
            OsArchitecture = ""
            OsLanguage = ""
            OsProductType = ""
            OsRegisteredUser = ""
            KeyboardLayout = ""
            PowerPlatformRole = ""
            A = ""
            DiskFreeSize = ""
            DiskTotalSize = ""
            FirstRun = ""
            LogsErrorsWarnings = ""
            ProcessesCount = ""
            ServicesCount = ""
			Temps = ""
        }
        $null = $everythingSheet.Add($newRow)
        $computerRow = $newRow
    }

    # Parse Get-ComputerInfo data from text file
    $computerInfoPath = "$DataRootPath\ComputerInfo\$computer.txt"
    if (Test-Path $computerInfoPath) {
        try {
            $content = Get-Content $computerInfoPath -ErrorAction Stop
            $computerInfo = @{}
            foreach ($line in $content) {
                if ($line -match "^\s*(\S+?)\s*:\s*(.*)$") {
                    $computerInfo[$matches[1]] = $matches[2].Trim()
                }
            }
            # Map Get-ComputerInfo properties to $computerRow
            $computerRow.PSComputerName = $computerInfo['PSComputerName'] 
            $computerRow.WindowsEditionId = $computerInfo['WindowsEditionId'] 
            $computerRow.WindowsInstallationType = $computerInfo['WindowsInstallationType'] 
            $computerRow.WindowsInstallDateFromRegistry = $computerInfo['WindowsInstallDateFromRegistry'] 
            $computerRow.WindowsProductName = $computerInfo['WindowsProductName'] 
            $computerRow.WindowsRegisteredOrganization = $computerInfo['WindowsRegisteredOrganization'] 
            $computerRow.WindowsRegisteredOwner = $computerInfo['WindowsRegisteredOwner'] 
            $computerRow.BiosCaption = $computerInfo['BiosCaption'] 
            $computerRow.BiosFirmwareType = $computerInfo['BiosFirmwareType'] 
            $computerRow.BiosSeralNumber = $computerInfo['BiosSeralNumber'] 
            $computerRow.BiosSMBIOSBIOSVersion = $computerInfo['BiosSMBIOSBIOSVersion'] 
            $computerRow.CsCurrentTimeZone = $computerInfo['CsCurrentTimeZone'] 
            $computerRow.CsDaylightInEffect = $computerInfo['CsDaylightInEffect'] 
            $computerRow.CsDomain = $computerInfo['CsDomain'] 
            $computerRow.CsDomainRole = $computerInfo['CsDomainRole'] 
            $computerRow.CsEnableDaylightSavingsTime = $computerInfo['CsEnableDaylightSavingsTime'] 
            $computerRow.CsManufacturer = $computerInfo['CsManufacturer'] 
            $computerRow.CsModel = $computerInfo['CsModel'] 
            $computerRow.CsNumberOfLogicalProcessors = $computerInfo['CsNumberOfLogicalProcessors'] 
            $computerRow.CsNumberOfProcessors = $computerInfo['CsNumberOfProcessors'] 
            $computerRow.CsPCSystemType = $computerInfo['CsPCSystemType'] 
            $computerRow.CsPowerOnPasswordStatus = $computerInfo['CsPowerOnPasswordStatus'] 
            $computerRow.CsSystemType = $computerInfo['CsSystemType'] 
            $computerRow.CsTotalPhysicalMemory = $computerInfo['CsTotalPhysicalMemory'] 
            $computerRow.CsPhyicallyInstalledMemory = $computerInfo['CsPhyicallyInstalledMemory'] 
            $computerRow.CsUserName = $computerInfo['CsUserName'] 
            $computerRow.OsBuildNumber = $computerInfo['OsBuildNumber'] 
            $computerRow.OsCountryCode = $computerInfo['OsCountryCode'] 
            $computerRow.OsCurrentTimeZone = $computerInfo['OsCurrentTimeZone'] 
            $computerRow.OsLocale = $computerInfo['OsLocale'] 
            $computerRow.OsUptime = $computerInfo['OsUptime'] 
            $computerRow.OsCodeSet = $computerInfo['OsCodeSet'] 
            $computerRow.OsInstallDate = $computerInfo['OsInstallDate'] 
            $computerRow.OsMuiLanguages = $computerInfo['OsMuiLanguages'] 
            $computerRow.OsNumberOfUsers = $computerInfo['OsNumberOfUsers'] 
            $computerRow.OsOrganization = $computerInfo['OsOrganization'] 
            $computerRow.OsArchitecture = $computerInfo['OsArchitecture'] 
            $computerRow.OsLanguage = $computerInfo['OsLanguage'] 
            $computerRow.OsProductType = $computerInfo['OsProductType'] 
            $computerRow.OsRegisteredUser = $computerInfo['OsRegisteredUser'] 
            $computerRow.KeyboardLayout = $computerInfo['KeyboardLayout'] 
            $computerRow.PowerPlatformRole = $computerInfo['PowerPlatformRole'] 
        } catch {
            Write-Warning "Failed to parse ComputerInfo file for $computer : $($_.Exception.Message)"
        }
    } else {
        Write-Warning "ComputerInfo file not found for $computer at $computerInfoPath"
    }

    # DiskFreeSize
    $diskFreeSizePath = "$DataRootPath\DiskFreeSize\$computer.txt"
    if (Test-Path $diskFreeSizePath) {
        $computerRow.DiskFreeSize = (Get-Content $diskFreeSizePath -ErrorAction SilentlyContinue | Select-Object -First 1)
    }

    # DiskTotalSize
    $diskTotalSizePath = "$DataRootPath\DiskTotalSize\$computer.txt"
    if (Test-Path $diskTotalSizePath) {
        $computerRow.DiskTotalSize = (Get-Content $diskTotalSizePath -ErrorAction SilentlyContinue | Select-Object -First 1)
    }

    # FirstRun
    $firstRunPath = "$DataRootPath\FirstRun\$computer.txt"
    if (Test-Path $firstRunPath) {
        $computerRow.FirstRun = (Get-Content $firstRunPath -ErrorAction SilentlyContinue | Select-Object -First 1)
    }

    # Logs (count errors/warnings)
    $logsPath = "$DataRootPath\Logs\$computer.txt"
    if (Test-Path $logsPath) {
        $logContent = Get-Content $logsPath -ErrorAction SilentlyContinue
        $errorWarningCount = ($logContent | Where-Object { $_ -match "error|warning" }).Count
        $computerRow.LogsErrorsWarnings = $errorWarningCount
    }

    # Processes (total lines)
    $processesPath = "$DataRootPath\Processes\$computer.txt"
    if (Test-Path $processesPath) {
        $processCount = (Get-Content $processesPath -ErrorAction SilentlyContinue).Count
        $computerRow.ProcessesCount = $processCount
    }

    # Services (total lines, assuming similar structure to Processes)
    $servicesPath = "$DataRootPath\Services\$computer.txt"
    if (Test-Path $servicesPath) {
        $serviceCount = (Get-Content $servicesPath -ErrorAction SilentlyContinue).Count
        $computerRow.ServicesCount = $serviceCount
    }
	
    # SWList (total lines, assuming similar structure to Processes)
    $SWListPath = "$DataRootPath\SWList\$computer.txt"
    if (Test-Path $SWListPath) {
        $SWListCount = (Get-Content $SWListPath -ErrorAction SilentlyContinue).Count
        $computerRow.SWListCount = $SWListCount
    }
	
    # Temps (TempFileCount)
    $TempsPath = "$DataRootPath\Temps\$computer.txt"
    if (Test-Path $TempsPath) {
        $TempsCount = (Get-Content $TempsPath -ErrorAction SilentlyContinue).Count
        $computerRow.TempFileCount = $TempsCount
    }
}

# Export the updated $everythingSheet to Excel
Write-Host "Exporting updated 'everything' sheet to Excel..."
try {
    $everythingSheet | Export-Excel -Path $excelFilePath -WorksheetName "everything" -ClearSheet -AutoSize -ErrorAction Stop
} catch {
    Write-Warning "Failed to export 'everything' sheet to Excel: $($_.Exception.Message)"
}

# Excel Everything and pclists Export Complete 

# _______ Second excel part _______

# Nu de andere sheets export
# List of required sheets
$requiredSheets = @(
    "ComputerInfo", "LocalAdmins", "FileSystemDrives", "Services", "Logs",
    "Processes", "NetworkAdapters", "InstalledModules", "LocalUsers", 
    "SWList", "Temps", "Summary", "everything"
)

# Function to check and create sheets
function Ensure-ExcelSheet {
    param (
        [string]$SheetName,
        [string]$ExcelFilePath
    )
    # Check if the Excel file exists and get existing sheets
    if (Test-Path $ExcelFilePath) {
        $existingSheets = Get-ExcelSheetInfo -Path $ExcelFilePath | Select-Object -ExpandProperty Name
    } else {
        $existingSheets = @()
    }
    # Create sheet if it doesn't exist
    if ($SheetName -notin $existingSheets) {

        # Use Export-Excel to create a new sheet (requires at least one object to create the sheet)
        $dummyData = [PSCustomObject]@{} # Empty object to create the sheet
        Export-Excel -Path $ExcelFilePath -WorksheetName $SheetName -InputObject $dummyData -NoHeader
        Write-Host "Created sheet: $SheetName"
    }
}

# Create all required sheets
foreach ($sheet in $requiredSheets) {
    Ensure-ExcelSheet -SheetName $sheet -ExcelFilePath $ExcelFilePath
}

# Define $computers if not already defined (fallback: use pclist-AD.txt)
if (-not $computers) {
    $adPclist = "$PclistRootPath\pclist-AD.txt"
    if (Test-Path $adPclist) {
        $computers = Get-Content $adPclist -ErrorAction SilentlyContinue
        Write-Host "Loaded computers from $adPclist."
    } else {
        Write-Error "No computers defined and $adPclist not found. Exiting."
        exit
    }
}


if ($rows.Count -gt 0) {
        $rows | Export-Excel -Path $ExcelFilePath -WorksheetName $folder -AutoSize
}

# End of Logs, localUsers, LocalAdmins, Services and NW adapters saving
# Collect all data in memory before writing to Excel

$OtherSheets = $RequiredSheets | Where-Object { $_ -notin @("Logs", "LocalUsers", "LocalAdmins", "Services", "NetworkAdapters") }

$summarySheet = New-Object System.Collections.ArrayList
$sheetData = @{}
foreach ($sheet in $OtherSheets) {
    if ($sheet -ne "Summary") {
        $sheetData[$sheet] = New-Object System.Collections.ArrayList
    }
}

foreach ($computer in $computers) {
    Write-Host "Processing data for $computer for Excel export."

    # Initialize summary row
    $summaryRow = [PSCustomObject]@{
        ComputerName   = $computer
        TeamviewerID   = ""
        BinSize        = ""
        DiskFreeSize   = ""
        DiskTotalSize  = ""
        FirstRun       = ""
    }

    foreach ($folder in $CommandNames) {
        $filePath = Join-Path -Path "$DataRootPath\$folder" -ChildPath "$computer.txt"
        if (Test-Path $filePath) {
            if (-not $fullexcel -and -not (Test-FileAge -FilePath $filePath -Days $FileAgeThresholdDays)) {
                Write-Host "Skipping $filePath (less than $FileAgeThresholdDays days old)"
                continue
            }
            $content = Get-Content $filePath -Raw -ErrorAction SilentlyContinue
            if ($content) {
                if ($folder -eq "Temps") {
                    $tempFileCount = $null
                    foreach ($line in (Get-Content $filePath)) {
                        if ($line -match "TempFileCount\s+(\d+)") {
                            $tempFileCount = $matches[1]
                            break
                        }
                    }
                    if ($null -ne $tempFileCount) {
                        $row = [PSCustomObject]@{
                            ComputerName = $computer
                            TempFileCount = $tempFileCount
                        }
                        $sheetData["Temps"].Add($row) | Out-Null
                    }
                } elseif ($folder -in @("TeamviewerID", "BinSize", "DiskFreeSize", "DiskTotalSize", "FirstRun")) {
                    $summaryRow.$folder = $content.Trim()
                } elseif ($folder -in @("FileSystemDrives", "Services", "Logs", "Processes")) {
                    # Parse structured data for these sheets
                    try {
                        $lines = Get-Content $filePath -ErrorAction SilentlyContinue
                        $parsedData = @()
                        switch ($folder) {
                            "FileSystemDrives" {
                                foreach ($line in $lines) {
                                    if ($line -match "^\s*(\S+)\s+(\d+)\s+(\d+)\s+(.*)\s+(\S+)\s*$") {
                                        $parsedData += [PSCustomObject]@{
                                            ComputerName = $computer
                                            Name = $matches[1]
                                            Used = $matches[2]
                                            Free = $matches[3]
                                            Description = $matches[4]
                                            Root = $matches[5]
                                        }
                                    }
                                }
                            }
                            "Services" {
                                foreach ($line in $lines) {
                                    if ($line -match "^\s*(\S+)\s+(.+?)\s+(\S+)\s+(\S+)\s*$") {
                                        $parsedData += [PSCustomObject]@{
                                            ComputerName = $computer
                                            Name = $matches[1]
                                            DisplayName = $matches[2]
                                            Status = $matches[3]
                                            StartType = $matches[4]
                                        }
                                    }
                                }
                            }
                            "Logs" {
                                foreach ($line in $lines) {
                                    if ($line -match "^\s*(\S+\s+\S+)\s+(\S+)\s+(\S+)\s+(.+)$") {
                                        $parsedData += [PSCustomObject]@{
                                            ComputerName = $computer
                                            TimeGenerated = $matches[1]
                                            EntryType = $matches[2]
                                            Source = $matches[3]
                                            Message = $matches[4]
                                        }
                                    }
                                }
                            }
							"SWList" {
                                foreach ($line in $lines) {
                                    if ($line -match "^\s*(\S+\s+\S+)\s+(\S+)\s+(\S+)\s+(.+)$") {
                                        $parsedData += [PSCustomObject]@{
                                            ComputerName = $computer
                                            DisplayName = $matches[1]
                                        }
                                    }
                                }
                            }
                            "Processes" {
                                foreach ($line in $lines) {
                                    if ($line -match "^\s*(\S+)\s+(\d+)\s+(\d+)\s+(\d+)\s*$") {
                                        $parsedData += [PSCustomObject]@{
                                            PSComputerName = $computer
                                            Name = $matches[1]
                                            CPU = $matches[2]
                                        }
                                    }
                                }
                            }
                        }
                        if ($parsedData) {
                            $sheetData[$folder].AddRange($parsedData)
                        }
                    } catch {
                        Write-Warning "Failed to parse $folder data for $computer : $($_.Exception.Message)"
                    }
                } else {
                    $row = [PSCustomObject]@{
                        ComputerName = $computer
                        Data = $content
                    }
                    $sheetData[$folder].Add($row) | Out-Null
                }
            }
        }
    }
    $summarySheet.Add($summaryRow) | Out-Null
}

# Export pclist files to separate sheets
$pclistFiles = @(
    "$PclistRootPath\pclist-AD.txt",
    "$PclistRootPath\pclist-offline.txt",
    "$PclistRootPath\pclist-noWinRM.txt",
    "$PclistRootPath\pclist-to-skip.txt"
)

foreach ($file in $pclistFiles) {
    if (Test-Path $file) {
        $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        Write-Host "Exporting $sheetName to Excel..."
        try {
            Get-Content $file | ForEach-Object {
                [PSCustomObject]@{
                    ComputerName = $_
                }
            } | Export-Excel -Path $ExcelFilePath -WorksheetName $sheetName -ClearSheet -AutoSize -ErrorAction Stop
        } catch {
            Write-Warning "Failed to export $sheetName to Excel: $($_.Exception.Message)"
        }
    }
}

# Export all collected data to Excel at once
Write-Host "Exporting all data to Excel..."
foreach ($sheet in $requiredSheets) {
    if ($sheet -eq "Summary") {
        try {
            $summarySheet | Export-Excel -Path $ExcelFilePath -WorksheetName $sheet -ClearSheet -AutoSize -ErrorAction Stop
            Write-Host "Exported data to $sheet sheet"
        } catch {
            Write-Warning "Failed to export $sheet sheet: $($_.Exception.Message)"
        }
    } else {
        if ($sheetData[$sheet].Count -gt 0) {
            try {
                $sheetData[$sheet] | Export-Excel -Path $ExcelFilePath -WorksheetName $sheet -ClearSheet -AutoSize -ErrorAction Stop
                Write-Host "Exported data to $sheet sheet"
            } catch {
                Write-Warning "Failed to export $sheet sheet: $($_.Exception.Message)"
            }
        }
    }
}


# begin van xlsexport 2.5

Write-Host "___ Final Optimized Excel Export ___"

Import-Module ImportExcel

$DataRootPath = "\\d0500\C$\data\PCscript\Data"
$PCListPath = "\\d0500\C$\data\PCscript\pclists\pclist-AD.txt"
$ExcelFilePath = "\\d0500\C$\data\PCscript\pc-overview.xlsx"
$folders = @("Logs", "LocalUsers", "LocalAdmins", "Services", "Processes", "NetworkAdapters")

$PCList = Get-Content $PCListPath | ForEach-Object { $_.Trim() }

foreach ($folder in $folders) {
    Write-Host "Processing folder: $folder"
    $rows = @()

    foreach ($pc in $PCList) {
        $filePath = Join-Path -Path "$DataRootPath\$folder" -ChildPath "$pc.txt"
        if (-not (Test-Path $filePath)) { continue }

        if ($folder -eq "Logs") {
            $lines = Get-Content $filePath | Where-Object { $_ -match '\S' }
            $headerFound = $false

            foreach ($line in $lines) {
                if ($line -match "^\s*Index\s+Time\s+EntryType\s+Source\s+InstanceID\s+Message") {
                    $headerFound = $true
                    continue
                }

                if ($headerFound) {
                    $columns = $line -split '\s{2,}'  # Split by 2+ spaces
                    if ($columns.Count -ge 6) {
                        $rows += [PSCustomObject]@{
                            ComputerName = $pc
                            Time         = $columns[1]
                            EntryType    = $columns[2]
                            Source       = $columns[3]
                            Message      = $columns[5]
                        }
                    }
                }
            }
        }
        elseif ($folder -eq "LocalUsers" -or $folder -eq "LocalAdmins") {
            $lines = Get-Content $filePath 
            foreach ($line in $lines) {
                $rows += [PSCustomObject]@{
                    ComputerName = $pc
                    User         = $line.Trim()
                }
            }
        }
        elseif ($folder -eq "NetworkAdapters") {
            $adapter = @{}
            $lines = Get-Content $filePath

            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) {
                    if ($adapter.Count -gt 0) {
                        $rows += [PSCustomObject]@{
                            ComputerName         = $pc
                            MacAddress           = $adapter["MacAddress"]
                            LinkSpeed            = $adapter["LinkSpeed"]
                            AdminStatus          = $adapter["AdminStatus"]
                            MediaConnectionState = $adapter["MediaConnectionState"]
                            Name                 = $adapter["Name"]
                        }
                        $adapter.Clear()
                    }
                    continue
                }

                if ($line -match "^\s*(.+?)\s*:\s*(.*)$") {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    $adapter[$key] = $value
                }
            }

            # Add last adapter if file doesn't end with blank line
            if ($adapter.Count -gt 0) {
                $rows += [PSCustomObject]@{
                    ComputerName         = $pc
                    MacAddress           = $adapter["MacAddress"]
                    LinkSpeed            = $adapter["LinkSpeed"]
                    AdminStatus          = $adapter["AdminStatus"]
                    MediaConnectionState = $adapter["MediaConnectionState"]
                    Name                 = $adapter["Name"]
                }
            }
        }
        elseif ($folder -eq "Processes") {
            $lines = Get-Content $filePath
            $headerFound = $false

            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                if ($line -match "^\s*Handles\s+NPM\(K\)\s+PM\(K\)\s+WS\(K\)\s+VM\(M\)\s+CPU\s+Id\s+ProcessName") {
                    $headerFound = $true
                    continue
                }

                if ($headerFound) {
                    $columns = $line -split '\s+'
                    if ($columns.Count -ge 8) {
                        $rows += [PSCustomObject]@{
                            ComputerName = $pc
                            ProcessName  = $columns[-1]
                            CPU          = $columns[5]
                        }
                    }
                }
            }
        }
        elseif ($folder -eq "Services") {
            $lines = Get-Content $filePath | Where-Object { $_ -match '\S' }
            foreach ($line in $lines) {
                # Split by 2 or more spaces to separate Status, Name, and DisplayName
                $columns = $line -split '\s{2,}'
                if ($columns.Count -ge 3) {
                    $rows += [PSCustomObject]@{
                        ComputerName = $pc
                        Status       = $columns[0].Trim()
                        Name         = $columns[1].Trim()
                        DisplayName  = $columns[2].Trim()
                    }
                }
            }
        }
        else {
            $lines = Get-Content $filePath | Where-Object { $_ -match '\S' }
            foreach ($line in $lines) {
                $rows += [PSCustomObject]@{
                    ComputerName = $pc
                    Data         = $line.Trim()
                }
            }
        }
    }

    if ($rows.Count -gt 0) {
        $rows | Export-Excel -Path $ExcelFilePath -WorksheetName $folder -AutoSize
    }
} 


# end of xlsexport 2.5

Write-Host "--- Excel Export Complete ---"

# Log completion time
$EndTime = Get-Date
Write-Host "`nScript completed at: $EndTime"
Write-Host "Total duration: $($EndTime - $StartTime)"

# THE END
