#Driver updates install

#Function to get installed programs
function Get-InstalledPrograms {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $DisplayName
    );

    if (-not $DisplayName) {
        $DisplayName = '*';
    }
    Get-ItemProperty -Path @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*';
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*';
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*';
        'HKCU:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    ) -ErrorAction 'SilentlyContinue' |
    Where-Object -Property "DisplayName" -Like "$DisplayName" |
    Select-Object -Property 'DisplayName', 'UninstallString', 'ModifyPath', 'PSChildName' |
    Sort-Object -Property 'DisplayName';
}

$DriverUpdatesDir = "$env:SystemDrive\DriverUpdates"

if (-not (Test-Path -Path $DriverUpdatesDir)) {
    New-Item -Path $DriverUpdatesDir -ItemType Directory -Force
}

#Determine manufacturer
$Manufacturer = (Get-WmiObject Win32_BIOS).Manufacturer

if ($Manufacturer -eq "HP") {

    #Download HPIA landing page
    $HPIA = Invoke-WebRequest -UseBasicParsing "https://ftp.ext.hp.com/pub/caps-softpaq/cmit/HPIA.html"

    #Get the download URL
    $dlUrl = ($HPIA.Links | Where-Object {$_.href -like "*hp-hpia-*.exe"}).href

    #Download
    Invoke-WebRequest -Uri $dlUrl -OutFile "$DriverUpdatesDir\hpia.exe" -UseBasicParsing

    #Extract HPIA
    Start-Process -FilePath "$DriverUpdatesDir\hpia.exe" -Wait -ArgumentList "/f $DriverUpdatesDir /s /e"

    #If you have a BIOS password set, you will need to generate an encrypted BIOS password .bin file with HP's Bios Config Utility.
    $PasswordFile = "\\path\to\pwd1.bin"

    if (Test-Path -Path $PasswordFile) {
        #Copy pwd1.bin (BIOS password file) to HPIA directory
        Copy-Item -Path "\\path\to\pwd1.bin" -Destination "$DriverUpdatesDir\pwd1.bin" -Force -Confirm:$false
    }

    #Install Driver Updates
    Start-Process -Wait -FilePath "$DriverUpdatesDir\HPImageAssistant.exe" -ArgumentList "/Operation:Analyze /Action:Install /Selection:All /Category:BIOS,Drivers,Firmware,Accessories /Silent /ReportFolder:$DriverUpdatesDir\DriverUpdatesLogs /SoftpaqDownloadFolder:$DriverUpdatesDir\Softpaqs /BIOSPwdFile:pwd1.bin"

}

if ($Manufacturer -eq "Dell Inc.") {

    #Check if Dell Command Update is installed
    $DCU = Get-InstalledPrograms | Where-Object {$_.DisplayName -like "*Dell Command | Update for Windows Universal*"}

    if (-not $DCU) {
        #Install Dell Command Update application
        Start-Process -Wait -FilePath "DCU_Setup.exe" -ArgumentList '/s /v"/qn/norestart"' #Wish I could just IWR this easily
    }

    #Checking for available BIOS, firmware, and driver updates with Dell Command Update
    $BiosPassword = "TestPW"
    Start-Process -Wait -FilePath "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" -ArgumentList "/configure -autoSuspendBitLocker=enable -biosPassword=`"$BiosPassword`""
    Start-Process -Wait -FilePath "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" -ArgumentList "/scan"

    #Installing all available BIOS, firmware, and driver updates
    Start-Process -Wait -FilePath "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" -ArgumentList "/applyUpdates -reboot=disable -updateType=bios,firmware,driver -outputLog=$DriverUpdatesDir\DellCommandUpdate.log"

}

if ($Manufacturer -eq "LENOVO") {

    if (-not (Get-InstalledModule -Name LSUClient -ErrorAction SilentlyContinue)) {
        #Install the LSUClient module
        Install-PackageProvider -Name NuGet -Force
        Install-Module -Name LSUClient -Force
    }

    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
    Import-Module LSUClient

    #Get only unattended updates
    $Updates = Get-LSUpdate | Where-Object { $_.Installer.Unattended }
    $Updates | Save-LSUpdate
    $Updates | Install-LSUpdate

    #Cleanup
    Remove-Item -Path "$env:TEMP\LSUPackages" -Force -Recurse
    
}