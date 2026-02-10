#region Functions

#Function to determine drive letter of USB install drive and WinPE RAM drive and set them to variables
function Find-DriveLetters {
    #Find drive letters
    $Drives = [char[]]([char]'A'..[char]'Z')

    #Find the USB drive with install media (usually D:)
    foreach ($Drive in $Drives) {
        if (Test-Path -Path "$Drive`:\sources\install.wim") {
            $ENV:USBDrive = "$Drive`:"
        }
    }

    #Find the Windows Setup RAM drive (usually X:)
    foreach ($Drive in $Drives) {
        if (Test-Path -Path "$Drive`:\setup-custom.exe") {
            $ENV:RAMDrive = "$Drive`:"
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

#Function to download the Start-Imaging script and assets from the Github repo
function Invoke-Imaging {
    ""
    #Run Find-DriveLetters function
    Find-DriveLetters

    #Test iwr over https works
    try {
        Invoke-WebRequest -Uri "https://www.google.com" -UseBasicParsing -ErrorAction Stop | Out-Null
    }
    #If not, prompt to update the system time
    catch {
        Write-Host -ForegroundColor Yellow "System time is incorrect."

        #Loop until the user enters a valid date
        do {
            $CurrentDate = Read-Host "Please provide the current date (MM/DD/YYYY)"
        
            try {
                #Parse input to verify format
                $VerifyDateFormat = [datetime]::ParseExact($CurrentDate, "MM/dd/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                $DateValid = $true
            }
            catch {
                Write-Host -ForegroundColor Red "Invalid format. Please try again."
                $DateValid = $false
            }
        } until ($DateValid)

        #Set the system date
        Set-Date -Date $CurrentDate | Out-Null

        #Attempt to correct the time
        $TimeAPI = Invoke-RestMethod -Uri "https://timeapi.io/api/Time/current/zone?timeZone=UTC" #Get time API using HTTP
        if ($TimeAPI) {
            $UTCTime = New-Object datetime ( #Convert to UTC
                $TimeAPI.year,
                $TimeAPI.month,
                $TimeAPI.day,
                $TimeAPI.hour,
                $TimeAPI.minute,
                $TimeAPI.seconds,
                $TimeAPI.milliSeconds,
                [System.DateTimeKind]::Utc
            )
            $LocalTime = $UTCTime.ToLocalTime() #Convert to local time
            Set-Date -Date $LocalTime | Out-Null #Set date and time
        }
        ""
        Write-Host "System time updated to $(Get-Date)."
        ""
    }

    #Download Start-Imaging script and assets
    Save-GithubFiles -Files @("Start-Imaging.ps1") -DownloadDirectory "$ENV:RAMDrive\CompanyNameImaging\Start-Imaging"
    Save-GithubFiles -Files @("assets/EyeOpen.png", "assets/EyeShut.png") -DownloadDirectory "$ENV:RAMDrive\CompanyNameImaging\Start-Imaging\assets"

    #Run Start-Imaging script
    & "$ENV:RAMDrive\CompanyNameImaging\Start-Imaging\Start-Imaging.ps1"
}

#endregion Functions

#Run Find-DriveLetters function
Find-DriveLetters

#Determine manufacturer (HP or Dell)
$Manufacturer = (Get-CimInstance Win32_BIOS).Manufacturer

Write-Host "======================================"
Write-Host "==========AUTOMATED WINSTALL=========="
Write-Host "======================================"

$USBVersion = Get-Content -Path "$ENV:USBDrive\CompanyNameImagingVersion.txt"
$OSInfo = Get-WindowsImage -ImagePath "$ENV:USBDrive\sources\install.wim" -Index 1
$WindowsVersion = $OSInfo.ImageName + " " + $OSInfo.Version

""
Write-Host -NoNewLine "CompanyName Imaging USB Version: "
Write-Host -ForegroundColor DarkGray "$USBVersion"
Write-Host -NoNewLine "Windows Image Version: "
Write-Host -ForegroundColor DarkGray "$WindowsVersion"

#region PEDrivers

#Give user ability to skip driver installation/injection
""
Write-Host -NoNewLine -ForegroundColor Yellow "(PRESS ANY KEY TO SKIP) "
Write-Host -NoNewLine "Installing drivers into Windows PE in"
for ($i = 5; $i -ge 1; $i--) {
    Write-Host " $i..." -NoNewline
    Start-Sleep -Milliseconds 1500
        
    #Check if a key is pressed
    if ([System.Console]::KeyAvailable) {
        $null = [System.Console]::ReadKey($true)
        ""
        Write-Host "Skipping driver installation."
        #Create marker file
        New-Item -Path "$ENV:RAMDrive\CompanyNameImaging" -Name "SkipDrivers.txt" -ItemType "File" -Force | Out-Null
        break
    }
}

#Only install drivers if SkipDrivers.txt is not present
if (!(Test-Path -Path "$ENV:RAMDrive\CompanyNameImaging\SkipDrivers.txt")) {
    #Create PEDrivers array and add dock drivers path
    $PEDriverPaths = @()
    $PEDriverPaths += "$ENV:USBDrive\drivers\docks"

    #If running on an HP, add HP drivers path
    if ($Manufacturer -eq "HP") {
        $PEDriverPaths += "$ENV:USBDrive\drivers\hp"
    }

    #If running on a Dell, add Dell drivers path
    if ($Manufacturer -eq "Dell Inc.") {
        $PEDriverPaths += "$ENV:USBDrive\drivers\dell"
    }

    #If running on a Lenovo, add Lenovo drivers path
    if ($Manufacturer -eq "LENOVO") {
        $PEDriverPaths += "$ENV:USBDrive\drivers\lenovo"
    }

    #Determine all .inf files from the selected driver folders and assign to $PEDrivers variable
    $PEDrivers = @()
    foreach ($Path in $PEDriverPaths) {
        $PEDrivers += Get-ChildItem -Path $Path -Recurse -Filter "*.inf" | Select-Object -ExpandProperty FullName
    }

    #Install all $PEDrivers using pnputil with progress bar
    if ($PEDrivers.Count -gt 0) {
        $DriverCount = $PEDrivers.Count
        $counter = 0

        foreach ($PEDriver in $PEDrivers) {
            #Update progress for each individual driver
            $percentComplete = [math]::Round(($counter / $DriverCount) * 100, 2)
            Write-Progress -Activity "Installing WinPE Drivers..." -Status "Loading driver $($counter + 1) of $DriverCount`: $PEDriver" -PercentComplete $percentComplete

            #Install driver
            pnputil /add-driver "$PEDriver" /install *> $null

            #Increase counter
            $counter++
        }

        #Complete progress bar
        Write-Progress -Activity "Installing drivers into WinPE. Please wait..." -Completed

        ""
    }
}

#endegion PEDrivers

Start-Sleep -Seconds 5

#Test network connectivity
if (Test-Connection "google.com" -Count 1 -ErrorAction SilentlyContinue) {
    #Run Invoke-Imaging function
    Invoke-Imaging
}
    else {
        #region Form

        #Windows form for configuration
        Add-Type -AssemblyName System.Windows.Forms

        $Form = New-Object -TypeName System.Windows.Forms.Form
        [System.Windows.Forms.Application]::EnableVisualStyles()
        [System.Windows.Forms.Label]$NetworkLabel = $null
        [System.Windows.Forms.Button]$OKButton = $null

        function InitializeComponent {

        $NetworkLabel = (New-Object -TypeName System.Windows.Forms.Label)
        $OKButton = (New-Object -TypeName System.Windows.Forms.Button)
        $Form.SuspendLayout()
        #
        #NetworkLabel
        #
        $NetworkLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]14))
        $NetworkLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]42))
        $NetworkLabel.Name = [System.String]'NetworkLabel'
        $NetworkLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]400,[System.Int32]100))
        $NetworkLabel.TabIndex = [System.Int32]1
        $NetworkLabel.Text = [System.String]'Please ensure a network cable is connected to the PC.'
        #
        #OKButton
        #
        $OKButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]14))
        $OKButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]181,[System.Int32]155))
        $OKButton.Name = [System.String]'OKButton'
        $OKButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]90,[System.Int32]53))
        $OKButton.TabIndex = [System.Int32]0
        $OKButton.Text = [System.String]'OK'
        $OKButton.UseVisualStyleBackColor = $true
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        #
        #Form
        #
        $Form.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]452,[System.Int32]250))
        $Form.AcceptButton = $OKButton
        $Form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
        $Form.Controls.Add($NetworkLabel)
        $Form.Controls.Add($OKButton)
        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $Form.Text = [System.String]'CompanyName Imaging'
        $Form.Add_Shown({ $Form.TopMost = $true; $Form.Activate() })
        $Form.ResumeLayout($false)
        $Form.PerformLayout()
        Add-Member -InputObject $Form -Name NetworkLabel -Value $NetworkLabel -MemberType NoteProperty
        Add-Member -InputObject $Form -Name OKButton -Value $OKButton -MemberType NoteProperty
        }
        . InitializeComponent
        $result = $Form.ShowDialog()

        #endregion Form

        #region UserHitsOK

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            #Sleep for 10 seconds
            Start-Sleep -Seconds 10

            #Test network connectivity
            if (Test-Connection "google.com" -Count 1 -ErrorAction SilentlyContinue) {
                #Run Invoke-Imaging function
                Invoke-Imaging
            }
                else {                    
                    ""
                    Write-Host -ForegroundColor Red "No network connection has been detected. Please ensure this system has a network cable attached."
                    ""
                    Write-Host "If a network cable is attached and it is not being detected, please reach out to"
                    Write-Host "[INSERT NAME HERE] for escalation, as this is likely a driver issue."
                    ""
                    pause
                    wpeutil reboot
                }
        }

        #endregion UserHitsOK
    }