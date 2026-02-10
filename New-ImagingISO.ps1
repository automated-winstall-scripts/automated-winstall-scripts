Clear-Host

#region Functions

#Function to get installed programs
function Remove-Apps {
    param([string]$Workspace)

    $WhitelistedApps = @( 
        "Microsoft.549981C3F5F10",
        "Microsoft.DesktopAppInstaller",
        "Microsoft.HEVCVideoExtension",
        "Microsoft.Paint",
        "Microsoft.RawImageExtension",
        "Microsoft.ScreenSketch",
        "Microsoft.SecHealthUI",
        "Microsoft.VCLibs.140.00",
        "Microsoft.VP9VideoExtensions",
        "Microsoft.WebpImageExtension",
        "Microsoft.Windows.Photos",
        "Microsoft.WindowsCalculator",
        "Microsoft.WindowsCamera",
        "Microsoft.WindowsNotepad",
        "Microsoft.WindowsSoundRecorder",
        "Microsoft.WindowsStore",
        "Microsoft.WindowsTerminal",
        "Microsoft.XboxIdentityProvider",
        "MicrosoftWindows.Client.WebExperience",
        "Microsoft.MSPaint",
        "Microsoft.StorePurchaseApp",
        "Microsoft.MicrosoftStickyNotes",
        "Microsoft.Office.OneNote",
        "Microsoft.HEIFImageExtension",
        "Microsoft.WebMediaExtensions",
        "Microsoft.WindowsAlarms",
        "Microsoft.Getstarted",
        "Microsoft.ApplicationCompatibilityEnhancements",
        "Microsoft.AV1VideoExtension",
        "Microsoft.AVCEncoderVideoExtension",
        "Microsoft.MPEG2VideoExtension"
    )
    
    #Get Apps from WIM
    $apps = Get-AppxProvisionedPackage -Path $Workspace

    #Cycle through Apps
    ForEach ($app in $apps) {
        #If app not in whitelist
        if($app.DisplayName -notin $WhitelistedApps) {
            #Remove App
            Write-Host "Removing" $app.DisplayName
            Remove-AppxProvisionedPackage -Path $Workspace -PackageName $app.PackageName
        }
    }
}

#Function to download files from the GitHub repo
function Save-GitHubFiles {
    param (
        [string[]]$Files,
        [string]$DownloadDirectory
    )

    #The GitHub repo where the files are
    $Repo = "automated-winstall-scripts/automated-winstall-scripts"

    #For each file in the $Files array
    foreach ($File in $Files) {
        #Set the download URL
        $Download = "https://raw.githubusercontent.com/$Repo/main/$File"

        #Create the folder for files to download to
        if (-not (Test-Path -Path "$DownloadDirectory")) {
            New-Item -Path $DownloadDirectory -ItemType Directory | Out-Null
        }

        #Make it so that filename of download is only everything after the last /
        $File = $File -Replace ".*/", ""

        #Download file(s)
        & {
            $ProgressPreference = "SilentlyContinue"
            Invoke-WebRequest -Uri $Download -OutFile "$DownloadDirectory\$File" -UseBasicParsing
        }
    }
}

#Function to open file picker to select files
function Get-FileName($initialDirectory) {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $initialDirectory
    $OpenFileDialog.Filter = "ISO Files (*.iso)|*.iso"
    $OpenFileDialog.Multiselect = $false #Ensure only one file can be selected
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $OpenFileDialog.FileName
    } else {
        return $null #Return null if no file is selected
    }
}

# Check if running as admin, if not write error and exit
function Confirm-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "Not running as admin. Relaunch this script as an administrator."
    }
}

#endregion Functions

#region Prereqs

Confirm-Admin

$ADKDeploymentTools = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
$ADKDeploymentToolsWinPE = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment"

if ((!(Test-Path -Path "$ADKDeploymentTools")) -and (!(Test-Path -Path "$ADKDeploymentToolsWinPE"))) {
    Write-Host "You are missing both the Deployment Tools from the Windows ADK and the WinPE add-on for ADK on your system."
}

if ((Test-Path -Path "$ADKDeploymentTools") -and (!(Test-Path -Path "$ADKDeploymentToolsWinPE"))) {
    Write-Host "Deployment Tools from the Windows ADK is installed, but you are missing the WinPE add-on for ADK on your system."
}

if ((!(Test-Path -Path "$ADKDeploymentTools")) -and (Test-Path -Path "$ADKDeploymentToolsWinPE")) {
    Write-Host "The WinPE add-on for the Windows ADK is installed, but you are missing the Deployment Tools from the ADK on your system."
}

if ((!(Test-Path -Path "$ADKDeploymentTools")) -or (!(Test-Path -Path "$ADKDeploymentToolsWinPE"))) {
    ""
    Write-Host "Please refer to the instructions and documentation for this process."
    pause
    exit
}

#endregion Prereqs

#region UserInput

#Have user select their Windows 11 ISO
Write-Host "Select your Windows 11 .ISO file..." -ForegroundColor Black -BackgroundColor Yellow
Start-Sleep 2

$WindowsISO = $null

do {
    $WindowsISO = Get-FileName -initialDirectory $ENV:userprofile\Downloads

    if ($WindowsISO) {
        ""
        Write-Host "Selected Windows 11 ISO file: $WindowsISO"
    } else {
        ""
        $userChoice = Read-Host "No file selected. Do you want to try again? (Y/N)"
        if ($userChoice -ne "Y") {
            ""
            Write-Host "No Windows 11 ISO file selected. Exiting..."
            ""
            pause
            exit
        }
    }
} while (-not $WindowsISO)

""
#Have user select their Windows 11 Language and Optional Features ISO (needed for enabling WMIC)
Write-Host "Select your Windows 11 Language and Optional Features .ISO file..." -ForegroundColor Black -BackgroundColor Yellow
Start-Sleep 2

$FODISO = $null

do {
    $FODISO = Get-FileName -initialDirectory $ENV:userprofile\Downloads

    if ($FODISO) {
        ""
        Write-Host "Selected Language and Optional Features ISO file: $FODISO"
    } else {
        ""
        $userChoice = Read-Host "No file selected. Do you want to try again? (Y/N)"
        if ($userChoice -ne "Y") {
            ""
            Write-Host "Failing to select a Language and Optional Features ISO will prevent certain optional features from being enabled within the resultant Windows image (namely WMIC). Continuing regardless..."
            break
        }
    }
} while (-not $FODISO)

""
#Ask user for Win11 build number
$Build = Read-Host "Enter the build number of the Windows 11 image (i.e. 23H2, 24H2)"
$Build = $Build.ToUpper() #Force uppercase

""
#Ask user for USB version number
$CompanyNameImagingVersion = Read-Host "Enter the version number of this CompanyName Imaging USB image (i.e. 1.0.1)"

#Get date
$Date = Get-Date -format "MM-dd-yy"

#endregion UserInput

#region Getting_Started

""
#Cleanup stale wims if any are mounted
Write-Host -ForegroundColor Green "Cleaning up any stale mounted images..."
dism /cleanup-wim

""
#Create Local Directory to use as a workspace for the WIM and local workspace
Write-Host -ForegroundColor Green "Creating scratch directories..."

if (Test-Path -Path "C:\WIMPrep") {
    Remove-Item "C:\WIMPrep" -Recurse -Force
}

$LocalDir = (New-Item -Path "C:\WIMPrep" -ItemType Directory -Force).FullName
$ExtractedISO = (New-Item -Path "$LocalDir\Extracted" -ItemType Directory -Force).FullName #For extracted Windows .iso file
if ($FODISO) {
    $FODISODir = (New-Item -Path "$LocalDir\FODISO" -ItemType Directory -Force).FullName #For extracted FOD .iso file
}
$Workspace = (New-Item -Path "$LocalDir\Offline" -ItemType Directory -Force).FullName #For mounted Windows install.wim
$BootWorkspace = (New-Item -Path "$LocalDir\BootWorkspace" -ItemType Directory -Force).FullName #For mounted Windows boot.wim

#Install 7-Zip if it's not installed
if (!(Test-Path -Path "$env:ProgramFiles\7-Zip\7z.exe")) {
    ""
    Write-Host -ForegroundColor Green "Installing 7-Zip..."

    #Set TLS support for Powershell and parse the JSON request
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #Download GitHub API JSON object for 7zip repo
    $7z = Invoke-WebRequest 'https://api.GitHub.com/repos/ip7z/7zip/releases/latest' -UseBasicParsing | ConvertFrom-Json

    #Get the download URL from the JSON object
    $dlUrl = ($7z.assets | Where-Object{$_.browser_download_url -like "*-x64.msi*"}).browser_download_url

    #Get the file name
    $7zoutfile = ($7z.assets | Where-Object{$_.browser_download_url -like "*-x64.msi*"}).name
    Invoke-WebRequest -Uri $dlUrl -OutFile "$LocalDir\$7zoutfile" -UseBasicParsing

    #Install 7-Zip
    Start-Process msiexec.exe -Wait -ArgumentList "/i `"$LocalDir\$7zoutfile`" /qn /norestart"
}

""
#Use 7Zip to extract Windows ISO
Write-Host -ForegroundColor Green "Extracting Windows ISO to scratch directory: $ExtractedISO..."
& ${env:ProgramFiles}\7-Zip\7z.exe x $WindowsISO "-o$($ExtractedISO)" -y 2> $null | Out-Null

#Use 7Zip to extract FOD ISO
if ($FODISO) {
    ""
    Write-Host -ForegroundColor Green "Extracting Language and Optional Features ISO to scratch directory: $FODISODir..."
    & ${env:ProgramFiles}\7-Zip\7z.exe x $FODISO "-o$($FODISODir)" -y 2> $null | Out-Null
}

#endregion Getting_Started

#region Export_Enterprise_Index_From_Install.wim

#Determine if install.wim or install.esd
if (Test-Path -Path "$ExtractedISO\sources\install.wim") {
    $wim = "install.wim"
}
elseif (Test-Path -Path "$ExtractedISO\sources\install.esd") {
    $wim = "install.esd"
}
    else { Write-Error "No install.wim or install.esd found in extracted ISO." }

#Get index # of Enterprise in wim
$Index = (Get-WindowsImage -ImagePath "$ExtractedISO\sources\$wim" | Where-Object { $_.ImageName -eq "Windows 11 Enterprise" }).ImageIndex

if ($Index -eq $null) {
    ""
    Write-Host "Your Windows 11 ISO does not contain an Enterprise Edition index. Exiting..."
    ""
    pause
    exit
}

""
Write-Host -ForegroundColor Green "Exporting just the Enterprise edition from the Windows wim..."
dism /export-image /SourceImageFile:"$ExtractedISO\sources\$wim" /SourceIndex:$Index /DestinationImageFile:"$LocalDir\$($Build)Enterprise $Date.wim" /Compress:max /CheckIntegrity

#Get build number, this info will be placed in the registry at HKLM\Software\CompanyName later
$OSInfo = Get-WindowsImage -ImagePath "$LocalDir\$($Build)Enterprise $Date.wim" -Index 1
$WindowsBuildNumber = $OSInfo.ImageName + " " + $OSInfo.Version

#endregion Export_Enterprise_Index_From_Install.wim

#region Mount_Install.wim_and_Make_Changes

#Mount the Enterprise WIM
""
Write-Host -ForegroundColor Green "Mounting the Windows wim to $Workspace..."
Mount-WindowsImage -ImagePath "$LocalDir\$($Build)Enterprise $Date.wim" -Path "$Workspace" -Index 1

#Call Removal function
Write-Host -ForegroundColor Green "Removing preprovisioned apps/bloat..."
""
Remove-Apps -Workspace $Workspace

Write-Host -ForegroundColor Green "Applying settings..."

#Copy default apps XML
Save-GitHubFiles -Files @("files/AppAssociations.xml") -DownloadDirectory "$LocalDir"

#Import default apps XML
Dism /Image:$Workspace /Import-DefaultAppAssociations:$LocalDir\AppAssociations.xml

#Import start menu layout - create folder in Default user Appdata where we will place modified layout
New-Item -Path "$Workspace\Users\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState" -ItemType Directory -Force | Out-Null

#Download modified layout
Save-GitHubFiles -Files @("files/start2.bin") -DownloadDirectory "$Workspace\Users\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"

#C:\Windows\CompanyName folder
New-Item -Path "$Workspace\Windows\CompanyName" -ItemType Directory -Force | Out-Null

#Copy theme, .bat script, and taskbar icon layout over to apply it too
Save-GitHubFiles -Files @("files/CompanyName.deskthemepack", "files/SetTheme.bat", "files/LayoutModification.xml") -DownloadDirectory "$Workspace\Windows\CompanyName"

#Create Default user shell:startup folder if it doesn't exist
New-Item -Path "$Workspace\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" -ItemType Directory -Force | Out-Null
#Copy .vbs script to it that deletes taskbar layout xml file
Save-GitHubFiles -Files @("files/DeleteTaskbarLayout.vbs") -DownloadDirectory "$Workspace\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"

""
Write-Host -ForegroundColor Green "Applying system-wide registry settings..."
""
#Load registry of offline Windows image
reg load HKLM\WIN11OFFLINE $Workspace\Windows\System32\Config\SOFTWARE
#Allow desktop icons from network shares
reg.exe add "HKLM\WIN11OFFLINE\Policies\Microsoft\Windows\Explorer" /v "EnableShellShortcutIconRemotePath" /t REG_DWORD /d 1 /f
#Turn off Widgets
reg.exe add "HKLM\WIN11OFFLINE\Policies\Microsoft\Dsh" /v "AllowNewsAndInterests" /t REG_DWORD /d 0 /f
#Disable mouse cursor suppression
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\Windows\CurrentVersion\Policies\System" /v "EnableCursorSuppression" /t REG_DWORD /d 0 /f
#Prevent Windows update installation of drivers
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\PolicyManager\current\device\Update" /v "ExcludeWUDriversInQualityUpdate" /t REG_DWORD /d 1 /f
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\PolicyManager\default\Update" /v "ExcludeWUDriversInQualityUpdate" /t REG_DWORD /d 1 /f
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\WindowsUpdate\UX\Settings" /v "ExcludeWUDriversInQualityUpdate" /t REG_DWORD /d 1 /f
reg.exe add "HKLM\WIN11OFFLINE\Policies\Microsoft\Windows\WindowsUpdate" /v "ExcludeWUDriversInQualityUpdate" /t REG_DWORD /d 1 /f
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\PolicyManager\default\Update\ExcludeWUDriversInQualityUpdate" /v "value" /t REG_DWORD /d 1 /f
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\Windows\CurrentVersion\Device Metadata" /v "PreventDeviceMetadataFromNetwork" /t REG_DWORD /d 1 /f
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\Windows\CurrentVersion\DriverSearching" /v "SearchOrderConfig" /t REG_DWORD /d 0 /f
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\Windows\CurrentVersion\DriverSearching" /v "DontSearchWindowsUpdate" /t REG_DWORD /d 1 /f
#Use Active Setup to apply theme to new users
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\Active Setup\Installed Components\SetTheme" /v "Locale" /t REG_SZ /d "*" /f
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\Active Setup\Installed Components\SetTheme" /v "Version" /t REG_SZ /d "1" /f
reg.exe add "HKLM\WIN11OFFLINE\Microsoft\Active Setup\Installed Components\SetTheme" /v "StubPath" /t REG_SZ /d "cmd /c C:\Windows\CompanyName\SetTheme.bat" /f
#Add information to registry about CompanyName Imaging version and Windows build
reg.exe add "HKLM\WIN11OFFLINE\CompanyName\CompanyNameImaging" /v "CompanyNameImagingVersion" /t REG_SZ /d "$CompanyNameImagingVersion" /f
reg.exe add "HKLM\WIN11OFFLINE\CompanyName\CompanyNameImaging" /v "WindowsBuildAtTimeOfImaging" /t REG_SZ /d "$Build $WindowsBuildNumber" /f
#Unload registry hive
reg unload HKLM\WIN11OFFLINE

#Load registry of offline Windows image
reg load HKLM\WIN11OFFLINE $Workspace\Windows\System32\Config\SYSTEM
reg.exe add "HKLM\WIN11OFFLINE\Setup" /v "Upgrade" /t REG_DWORD /d 0 /f
#Unload registry hive
reg unload HKLM\WIN11OFFLINE

""
Write-Host -ForegroundColor Green "Applying registry settings to C:\Users\Default\NTUSER.DAT..."
""
#Load registry of default user in offline Windows image
reg load HKLM\WIN11OFFLINE $Workspace\Users\Default\NTUSER.DAT
#Align taskbar to left
reg.exe add "HKLM\WIN11OFFLINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "TaskbarAl" /t REG_DWORD /d 0 /f
#Turn off Copilot
reg.exe add "HKLM\WIN11OFFLINE\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" /v "TurnOffWindowsCopilot" /t REG_DWORD /d 1 /f
#Turn off tablet optimized taskbar
reg.exe add "HKLM\WIN11OFFLINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "ExpandableTaskbar" /t REG_DWORD /d 0 /f
#Turn off "Chat" taskbar item
reg.exe add "HKLM\WIN11OFFLINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "TaskbarMn" /t REG_DWORD /d 0 /f
#Turn off "Show recommendations for tips, shortcuts, new apps, and more" in Start Menu
reg.exe add "HKLM\WIN11OFFLINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "Start_IrisRecommendations" /t REG_DWORD /d 0 /f
#Unload registry hive
reg unload HKLM\WIN11OFFLINE

#Install .NET Framework 3.5
""
Write-Host -ForegroundColor Green "Enabling .NET Framework 3.5..."
Dism /Image:$Workspace /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:$ExtractedISO\sources\sxs

if ($FODISO) {
    ""
    Write-Host -ForegroundColor Green "Enabling WMIC Feature On Demand..."
    Dism /Image:$Workspace /Add-Capability /CapabilityName:WMIC~~~~ /LimitAccess /Source:$FODISODir\LanguagesAndOptionalFeatures
}

""
Write-Host -ForegroundColor Green "Creating shell:common startup folder in Windows image..."
#Create shell:common startup folder if it doesn't exist
$Startup = "$Workspace\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
if (-not (Test-Path -Path "$Startup")) {
    New-Item -Path $Startup -ItemType Directory | Out-Null
}

""
Write-Host -ForegroundColor Green "Cleaning up C:\ folders on wim..."
#Cleanup
$Folders = @(
    "AMD",
    "Config.msi",
    "hp",
    "Dell",
    "Intel",
    "PerfLogs",
    "SWSetup",
    "system.sav",
    "inetpub",
    "HPIALogs",
    "ProgramFilesFolder",
    "PwrMgmt"
)

foreach ($Folder in $Folders) {
    if (Test-Path -Path "$Workspace\$Folder") {
        Remove-Item $Workspace\$Folder -Recurse -Force
    }
}

""
Write-Host -ForegroundColor Green "Saving changes and unmounting Windows wim..."
#Dismount Image
Dismount-WindowsImage -Path $Workspace -Save

#Remove Workfolder
Remove-Item $Workspace -Recurse -Force

#Copy modified wim to extracted ISO directory
Write-Host -ForegroundColor Green "Replacing stock install.wim with modified one in extracted Windows ISO directory: $ExtractedISO..."
Remove-Item "$ExtractedISO\sources\$wim" -Force
Move-Item -Path "$LocalDir\$($Build)Enterprise $Date.wim" -Destination "$ExtractedISO\sources\install.wim"

#endregion Mount_Install.wim_and_Make_Changes

#region Export_Windows_Setup_Index_From_Boot.wim

""
#Export the Microsoft Windows Setup index from the boot.wim
Write-Host -ForegroundColor Green "Exporting just the Windows Setup index from the Windows ISO boot.wim..."
#Get index # of boot.wim
$BootIndex = (Get-WindowsImage -ImagePath "$ExtractedISO\sources\boot.wim" | Where-Object { $_.ImageName -like "*Microsoft Windows Setup*" }).ImageIndex
dism /export-image /SourceImageFile:"$ExtractedISO\sources\boot.wim" /SourceIndex:$BootIndex /DestinationImageFile:"$LocalDir\boot.wim" /Compress:max /CheckIntegrity

#endregion Export_Windows_Setup_Index_From_Boot.wim

#region Mount_Boot.wim_and_Make_Changes

""
#Modify boot.wim
Write-Host -ForegroundColor Green "Mounting boot.wim to $BootWorkspace..."
""
Mount-WindowsImage -ImagePath "$LocalDir\boot.wim" -Path "$BootWorkspace" -Index 1

Write-Host -ForegroundColor Green "Installing WinPE packages to boot.wim..."
#WinPE optional components path from PE add-on for the Windows ADK
$WinPEPackagesPath = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs"

#All the required components
$WinPEPackages = @(
    "WinPE-WMI.cab",
    "en-us\WinPE-WMI_en-us.cab",
    "WinPE-NetFx.cab",
    "en-us\WinPE-NetFx_en-us.cab",
    "WinPE-Scripting.cab",
    "en-us\WinPE-Scripting_en-us.cab",
    "WinPE-PowerShell.cab",
    "en-us\WinPE-PowerShell_en-us.cab",
    "WinPE-DismCmdlets.cab",
    "en-us\WinPE-DismCmdlets_en-us.cab",
    "WinPE-StorageWMI.cab",
    "en-us\WinPE-StorageWMI_en-us.cab"
)

#Install components
foreach ($WinPEPackage in $WinPEPackages) {
    dism.exe /Image:$BootWorkspace /Add-Package /PackagePath:"$WinPEPackagesPath\$WinPEPackage"
}

""
Write-Host -ForegroundColor Green "Installing fonts to boot.wim..."

#Download Segoe UI Variable font from GitHub
New-Item -Path "$LocalDir\Fonts" -ItemType Directory -Force | Out-Null
Save-GitHubFiles -Files @("files/SegUIVar.ttf") -DownloadDirectory "$LocalDir\Fonts"

#Create new com shell object
$Shell = New-Object -COMObject Shell.Application
$Fonts = "$BootWorkspace\Windows\Fonts"

#Load SOFTWARE registry hive of boot.wim
""
reg load HKLM\WIN11OFFLINE $BootWorkspace\Windows\System32\Config\SOFTWARE

foreach($File in $(Get-ChildItem -Path "$LocalDir\Fonts")) {
    if (Test-Path "$BootWorkspace\Windows\Fonts\$($File.name)") {
    }
        else {
            $Path = $File.FullName
            $Folder = Split-Path $Path
            $File = Split-Path $Path -Leaf
            $ShellFolder = $Shell.Namespace($Folder)
            $ShellFile = $ShellFolder.ParseName($File)
            $ExtAtt2 = $ShellFolder.GetDetailsOf($ShellFile, 2)

            #Set the $FontType Variable
            if ($ExtAtt2 -Like '*TrueType font file*') {
                $FontType = '(TrueType)'
            }

            $ExtAtt21 = $ShellFolder.GetDetailsOf($ShellFile, 21) + ' ' + $FontType
            New-ItemProperty -Name $ExtAtt21 -Path "HKLM:\WIN11OFFLINE\Microsoft\Windows NT\CurrentVersion\Fonts" -PropertyType string -Value $File -Force | Out-Null
            Copy-Item $Path -Destination $Fonts
        }
}

#Unload registry hive
reg unload HKLM\WIN11OFFLINE

""
Write-Host -ForegroundColor Green "Changing default display scaling in WinPE..."
""
#Load DEFAULT registry hive of boot.wim
reg load HKLM\WIN11OFFLINE $BootWorkspace\Windows\System32\Config\DEFAULT

#Set display scaling in WinPE to 150%
reg.exe add "HKLM\WIN11OFFLINE\Control Panel\Desktop" /t REG_DWORD /v LogPixels /d 144 /f
reg.exe add "HKLM\WIN11OFFLINE\Control Panel\Desktop" /v Win8DpiScaling /t REG_DWORD /d 0x00000001 /f
reg.exe add "HKLM\WIN11OFFLINE\Control Panel\Desktop" /v DpiScalingVer /t REG_DWORD /d 0x00001018 /f

#Unload registry hive
reg unload HKLM\WIN11OFFLINE

""
Write-Host -ForegroundColor Green "Making other miscellaneous changes to boot.wim..."

#Rename setup.exe in boot.wim (doing this prevents Windows Setup from launching automatically into setup.exe)
Rename-Item -Path "$BootWorkspace\setup.exe" -NewName "setup-custom.exe"

#Copy startnet.cmd to $BootWorkspace\Windows\System32
#This script is what Windows Setup will automatically launch into instead of setup.exe
if (Test-Path -Path "$BootWorkspace\Windows\System32\startnet.cmd") {
    Remove-Item "$BootWorkspace\Windows\System32\startnet.cmd" -Force
}
Save-GitHubFiles -Files @("files/startnet.cmd") -DownloadDirectory "$BootWorkspace\Windows\System32"

""
Write-Host -ForegroundColor Green "Saving changes and unmounting modified boot.wim..."
#Dismount Image
Dismount-WindowsImage -Path $BootWorkspace -Save

#Remove Workfolder
Remove-Item $BootWorkspace -Recurse -Force

#Copy modified wim to extracted ISO directory
Write-Host -ForegroundColor Green "Replacing stock boot.wim with modified one in extracted Windows ISO directory: $ExtractedISO..."
Remove-Item "$ExtractedISO\sources\boot.wim" -Force
Move-Item -Path "$LocalDir\boot.wim" -Destination "$ExtractedISO\sources\boot.wim"

#endregion Mount_Boot.wim_and_Make_Changes

#region Make_ISO_Changes

#Put txt file on root of ISO with CompanyName Imaging version number
Set-Content -Path "$ExtractedISO\CompanyNameImagingVersion.txt" -Value "$CompanyNameImagingVersion"

#Create folder for drivers in ISO
$DriversPath = New-Item -Path "$ExtractedISO\Drivers" -ItemType Directory -Force
""
Write-Host "Downloading WinPE drivers to $DriversPath..."

#Download drivers from GitHub and place on USB drive
Save-GitHubFiles -Files @("Drivers/hp.zip", "Drivers/dell.zip", "Drivers/lenovo.zip", "Drivers/docks.zip") -DownloadDirectory "$DriversPath"

#Extract dell.zip and HP.zip drivers, then remove the .zip files
foreach ($item in @("hp", "dell", "lenovo", "docks")) {
    $DriversPathFolder = New-Item -Path "$DriversPath\$item" -ItemType Directory -Force
    Expand-Archive -Path "$DriversPath\$item.zip" -DestinationPath "$DriversPathFolder" -Force
    Remove-Item "$DriversPath\$item.zip" -Force
}

#Download Test-NetworkConnectivity.ps1 script and place in $ExtractedISO\CompanyName-Imaging
$CompanyNameImagingDir = New-Item -Path "$ExtractedISO\CompanyNameImaging" -ItemType Directory -Force
Save-GitHubFiles -Files @("Test-NetworkConnectivity.ps1") -DownloadDirectory "$CompanyNameImagingDir"

#endregion Make_ISO_Changes

#region Create_ISO

""
#Create ISO
$oscdimg = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg"

Write-Host -ForegroundColor Green "Creating ISO at C:\CompanyName Imaging $Build.iso from extracted ISO contents at $ExtractedISO..."
$oscdimgCommand = "`"$oscdimg\oscdimg.exe`" -m -o -u2 -udfver102 -bootdata:2#p0,e,b$ExtractedISO\boot\etfsboot.com#pEF,e,b$ExtractedISO\efi\microsoft\boot\efisys.bin $ExtractedISO `"C:\CompanyName Imaging $Build v$CompanyNameImagingVersion.iso`""
Start-Process "cmd.exe" -ArgumentList "/c `"$oscdimgCommand`"" -Wait

""
Write-Host -ForegroundColor Green "CompanyName manual imaging Windows ISO successfully created at C:\CompanyName Imaging $Build v$CompanyNameImagingVersion.iso."

""
Write-Host -ForegroundColor Green "Cleaning up..."
#Remove WIMPrep Folder
Remove-Item $LocalDir -Recurse -Force

""
Write-Host -ForegroundColor Green "Done."
""

pause

#endregion Create_ISO