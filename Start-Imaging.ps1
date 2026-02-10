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

#Function to get free drive letters
function Get-AvailableDriveLetter {
    #Store drive letters C-Z in a variable
    $allDriveLetters = [char[]]([char]'C'..[char]'Z')

    #See which of those are already taken
    $usedDriveLetters = Get-Volume | Select-Object -ExpandProperty DriveLetter

    #Inverse logic to get free drive letters
    $availableDriveLetters = $allDriveLetters | Where-Object { $_ -notin $usedDriveLetters }

    #Return first available drive letter
    return $availableDriveLetters[0] #Return the first available letter
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

#HP BIOS WMI namespace for some reason is not loaded immediately in WinPE, this gets around that
$global:HPWMIWarnedAlready = $false #Declare global var

function Test-HPBiosWMINamespace {
    $HPWMIAttempts = 0

    do {
        try {
            #Determining the value of the asset tag in the BIOS
            $CurrentAssetTag = (Get-CimInstance -Namespace root/hp/InstrumentedBIOS -Class HP_BIOSSetting -ErrorAction Stop | Where-Object Name -eq "Asset Tracking Number").Value
        }
        catch {
            #Only display the message once
            if (-not $global:HPWMIWarnedAlready) {
                Write-Host "HP WMI namespace initializing. Please wait..."
                ""
                ""
                $global:HPWMIWarnedAlready = $true
            }

            $HPWMIAttempts++
            Start-Sleep -Seconds 5
        }
    } until (($null -ne $CurrentAssetTag) -or ($CurrentAssetTag -ne "") -or ($HPWMIAttempts -ge 24)) #Maximum of 2 minutes

    return $CurrentAssetTag
}

#Function to show the main CompanyName Imaging configuration WinForm
function Show-ImagingForm {
    #Check that the assets used in the Winform are downloaded
    $RequiredAssets = @(
       "EyeOpen.png",
       "EyeShut.png",
       "QuestionButton.png",
       "CompanyNameLogo.png"
    )

    $AssetsPath = "$ENV:RAMDrive\CompanyNameImaging\Start-Imaging\assets"

    foreach ($Asset in $RequiredAssets) {
        if (!(Test-Path -Path "$AssetsPath\$Asset")) {
            Save-GithubFiles -Files @("assets/$Asset") -DownloadDirectory "$AssetsPath"
        }
    }

    #Windows form for configuration
    Add-Type -AssemblyName System.Windows.Forms

    $CompanyNameImagingForm = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Label]$CompanyNameImagingLabel = $null
    [System.Windows.Forms.Label]$DomainJoinLabel = $null
    [System.Windows.Forms.Label]$DomainJoinInstructionsLabel = $null
    [System.Windows.Forms.RadioButton]$DOM1RadioButton = $null
    [System.Windows.Forms.RadioButton]$DOM2RadioButton = $null
    [System.Windows.Forms.CheckBox]$AutoJoinCheckbox = $null
    [System.Windows.Forms.CheckBox]$SpecificOUCheckbox = $null
    [System.Windows.Forms.ComboBox]$OUDropdown = $null
    [System.Windows.Forms.Label]$SystemConfigurationLabel = $null
    [System.Windows.Forms.Label]$SystemConfigurationInstructionsLabel = $null
    [System.Windows.Forms.CheckBox]$PCHostnameCheckbox = $null
    [System.Windows.Forms.TextBox]$PCHostnameEntry = $null
    [System.Windows.Forms.CheckBox]$AssetTagCheckbox = $null
    [System.Windows.Forms.TextBox]$AssetTagEntry = $null
    [System.Windows.Forms.Label]$DomainCredentialsLabel = $null
    [System.Windows.Forms.Label]$DomainCredentialsInstructionsLabel = $null
    [System.Windows.Forms.Label]$UsernameLabel = $null
    [System.Windows.Forms.TextBox]$UsernameEntry = $null
    [System.Windows.Forms.Label]$PasswordLabel = $null
    [System.Windows.Forms.TextBox]$PasswordEntry = $null
    [System.Windows.Forms.Button]$EyeButton = $null
    [System.Windows.Forms.Button]$QuestionButton = $null
    [System.Windows.Forms.RichTextBox]$PCInfo = $null
    [System.Windows.Forms.Button]$OKButton = $null
    [System.Windows.Forms.Label]$FormBackground = $null

    #Load Images
    $CompanyNameImagingLogo = [System.Drawing.Image]::FromFile("$AssetsPath\CompanyNameLogo.png")
    $ShowPasswordImage = [System.Drawing.Image]::FromFile("$AssetsPath\EyeOpen.png")
    $HidePasswordImage = [System.Drawing.Image]::FromFile("$AssetsPath\EyeShut.png")
    $QuestionButtonImage = [System.Drawing.Image]::FromFile("$AssetsPath\QuestionButton.png")

    $Domain = "Dom1"
    $VPN = "Dom1"

    function Validate {
        if ($Dom1RadioButton.Checked -eq $true) {
            $Domain = "Dom1"
            $VPN = "Dom1"
            $SpecificOUCheckbox.Enabled = $true
        }

        if ($Dom2RadioButton.Checked -eq $true) {
            $Domain = "Dom2"
            $VPN = "Dom2"
            $SpecificOUCheckbox.Enabled = $false
            $SpecificOUCheckbox.Checked = $false
            $OUDropdown.Enabled = $false
        }

        #Update UI elements explicitly
        $UsernameLabel.Text = "Username:      $Domain \"
        $DomainCredentialsInstructionsLabel.Text = "Enter account credentials that have access to join PCs to the $Domain domain."
        $DomainJoinInstructionsLabel.Text = "Choose the domain the PC joins to. Check the option to not auto domain join if you need to log in to the local account first to connect to $VPN VPN. Check to join a specific OU in Active Directory if you have limited access in $Domain AD."

        $AutoJoinCheckbox.Text = "Don't auto domain join (I need $VPN VPN)"

        if (($UsernameEntry.Text -eq "") -or ($PasswordEntry.Text -eq "")) {
            $OKButton.Enabled = $false
        }
            else {
                $OKButton.Enabled = $true
            }

        if ($SpecificOUCheckbox.Enabled -eq $false) {
            $OUDropdown.Enabled = $false
        }

        if ($SpecificOUCheckbox.Checked -eq $true) {
            $OUDropdown.Enabled = $true
        }

        if ($SpecificOUCheckbox.Checked -eq $false) {
            $OUDropdown.Enabled = $false
        }

        if ($PCHostnameCheckbox.Checked -eq $true) {
            $PCHostnameEntry.Enabled = $true
        }

        if ($PCHostnameCheckbox.Checked -eq $false) {
            $PCHostnameEntry.Enabled = $false
        }

        if ($AssetTagCheckbox.Checked -eq $true) {
            $AssetTagEntry.Enabled = $true
        }

        if ($AssetTagCheckbox.Checked -eq $false) {
            $AssetTagEntry.Enabled = $false
        }
    }

    $CompanyNameImagingLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $DomainJoinLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $DomainJoinInstructionsLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $Dom1RadioButton = (New-Object -TypeName System.Windows.Forms.RadioButton)
    $Dom2RadioButton = (New-Object -TypeName System.Windows.Forms.RadioButton)
    $AutoJoinCheckbox = (New-Object -TypeName System.Windows.Forms.CheckBox)
    $SpecificOUCheckbox = (New-Object -TypeName System.Windows.Forms.CheckBox)
    $OUDropdown = (New-Object -TypeName System.Windows.Forms.ComboBox)
    $SystemConfigurationLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $SystemConfigurationInstructionsLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $PCHostnameCheckbox = (New-Object -TypeName System.Windows.Forms.CheckBox)
    $PCHostnameEntry = (New-Object -TypeName System.Windows.Forms.TextBox)
    $AssetTagCheckbox = (New-Object -TypeName System.Windows.Forms.CheckBox)
    $AssetTagEntry = (New-Object -TypeName System.Windows.Forms.TextBox)
    $DomainCredentialsLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $DomainCredentialsInstructionsLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $UsernameLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $UsernameEntry = (New-Object -TypeName System.Windows.Forms.TextBox)
    $PasswordLabel = (New-Object -TypeName System.Windows.Forms.Label)
    $PasswordEntry = (New-Object -TypeName System.Windows.Forms.TextBox)
    $EyeButton = (New-Object System.Windows.Forms.Button)
    $QuestionButton = (New-Object -TypeName System.Windows.Forms.Button)
    $PCInfo = (New-Object -TypeName System.Windows.Forms.RichTextBox)
    $OKButton = (New-Object -TypeName System.Windows.Forms.Button)
    $FormBackground = (New-Object -TypeName System.Windows.Forms.Label)
    $CompanyNameImagingForm.SuspendLayout()
    #
    #CompanyNameImagingLabel
    #
    $CompanyNameImagingLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]38,[System.Int32]54))
    $CompanyNameImagingLabel.Name = [System.String]'CompanyNameImagingLabel'
    $CompanyNameImagingLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]450,[System.Int32]63))
    $CompanyNameImagingLabel.TabIndex = [System.Int32]0
    $CompanyNameImagingLabel.BackgroundImage = $CompanyNameImagingLogo
    $CompanyNameImagingLabel.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
    $CompanyNameImagingLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #DomainJoinLabel
    #
    $DomainJoinLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display Semibold',[System.Single]16))
    $DomainJoinLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]31,[System.Int32]146))
    $DomainJoinLabel.Name = [System.String]'DomainJoinLabel'
    $DomainJoinLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]400,[System.Int32]60))
    $DomainJoinLabel.TabIndex = [System.Int32]1
    $DomainJoinLabel.Text = [System.String]'Domain Join'
    $DomainJoinLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #DomainJoinInstructionsLabel
    #
    $DomainJoinInstructionsLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]7))
    $DomainJoinInstructionsLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]35,[System.Int32]201))
    $DomainJoinInstructionsLabel.Name = [System.String]'DomainJoinInstructionsLabel'
    $DomainJoinInstructionsLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]620,[System.Int32]80))
    $DomainJoinInstructionsLabel.TabIndex = [System.Int32]2
    $DomainJoinInstructionsLabel.Text = "Choose the domain the PC joins to. Check the option to not auto domain join if you need to log in to the local account first to connect to $VPN VPN. Check to join a specific OU in Active Directory if you have limited access in $Domain AD."
    $DomainJoinInstructionsLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#4f4f4f")
    $DomainJoinInstructionsLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #Dom1RadioButton
    #
    $Dom1RadioButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]16))
    $Dom1RadioButton.Checked = $true
    $Dom1RadioButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]169,[System.Int32]284))
    $Dom1RadioButton.Name = [System.String]'Dom1RadioButton'
    $Dom1RadioButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]128,[System.Int32]50))
    $Dom1RadioButton.TabIndex = [System.Int32]3
    $Dom1RadioButton.Text = [System.String]'Dom1'
    $Dom1RadioButton.add_CheckedChanged({Validate})
    $Dom1RadioButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #Dom2RadioButton
    #
    $Dom2RadioButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]16))
    $Dom2RadioButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]380,[System.Int32]284))
    $Dom2RadioButton.Name = [System.String]'Dom2RadioButton'
    $Dom2RadioButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]128,[System.Int32]50))
    $Dom2RadioButton.TabIndex = [System.Int32]4
    $Dom2RadioButton.Text = [System.String]'Dom2'
    $Dom2RadioButton.add_CheckedChanged({Validate})
    $Dom2RadioButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #AutoJoinCheckbox
    #
    $AutoJoinCheckbox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $AutoJoinCheckbox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]39,[System.Int32]369))
    $AutoJoinCheckbox.Name = [System.String]'AutoJoinCheckbox'
    $AutoJoinCheckbox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]600,[System.Int32]50))
    $AutoJoinCheckbox.TabIndex = [System.Int32]5
    $AutoJoinCheckbox.Checked = $false
    $AutoJoinCheckbox.Text = "Don't auto domain join (I need $VPN VPN)"
    $AutoJoinCheckbox.UseVisualStyleBackColor = $true
    $AutoJoinCheckbox.add_CheckedChanged($AutoJoinCheckbox_CheckedChanged)
    $AutoJoinCheckbox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #SpecificOUCheckbox
    #
    $SpecificOUCheckbox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $SpecificOUCheckbox.Enabled = $true
    $SpecificOUCheckbox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]39,[System.Int32]432))
    $SpecificOUCheckbox.Name = [System.String]'SpecificOUCheckbox'
    $SpecificOUCheckbox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]600,[System.Int32]50))
    $SpecificOUCheckbox.TabIndex = [System.Int32]6
    $SpecificOUCheckbox.Checked = $false
    $SpecificOUCheckbox.Text = "Domain join to a specific OU:"
    $SpecificOUCheckbox.UseVisualStyleBackColor = $true
    $SpecificOUCheckbox.add_CheckedChanged($SpecificOUCheckbox_CheckedChanged)
    $SpecificOUCheckbox.add_CheckedChanged({Validate})
    $SpecificOUCheckbox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #OUDropdown
    #
    $OUDropdownArray = @(
        "dom1.com/ExampleOU/Computers"
    )
    $OUDropdown.Enabled = $false
    $OUDropdown.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]11))
    $OUDropdown.ForeColor = [System.Drawing.Color]::Black
    $OUDropdown.FormattingEnabled = $true
    $OUDropdown.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]39,[System.Int32]494))
    $OUDropdown.Name = [System.String]'OUDropdown'
    $OUDropdown.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]586,[System.Int32]54))
    $OUDropdown.TabIndex = [System.Int32]7
    $OUDropdown.DropDownStyle = "DropDownList"
    ForEach ($Item in $OUDropdownArray) {
        $OUDropdown.Items.Add($Item) | Out-Null
    }
    $OUDropdown.SelectedIndex = 0
    $OUDropdown.add_SelectedIndexChanged({Validate})
    $OUDropdown.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #SystemConfigurationLabel
    #
    $SystemConfigurationLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display Semibold',[System.Single]16))
    $SystemConfigurationLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]692,[System.Int32]42))
    $SystemConfigurationLabel.Name = [System.String]'SystemConfigurationLabel'
    $SystemConfigurationLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]400,[System.Int32]60))
    $SystemConfigurationLabel.TabIndex = [System.Int32]8
    $SystemConfigurationLabel.Text = [System.String]'System Configuration'
    $SystemConfigurationLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #SystemConfigurationInstructionsLabel
    #
    $SystemConfigurationInstructionsLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]7))
    $SystemConfigurationInstructionsLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]696,[System.Int32]97))
    $SystemConfigurationInstructionsLabel.Name = [System.String]'SystemConfigurationInstructionsLabel'
    $SystemConfigurationInstructionsLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]620,[System.Int32]60))
    $SystemConfigurationInstructionsLabel.TabIndex = [System.Int32]9
    $SystemConfigurationInstructionsLabel.Text = [System.String]'Enter what you would like the computer hostname and asset tag (as reported in the BIOS) to be set to. If you do not want to change one or both of these, uncheck the corresponding option(s).'
    $SystemConfigurationInstructionsLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#4f4f4f")
    $SystemConfigurationInstructionsLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #PCHostnameCheckbox
    #
    $PCHostnameCheckbox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $PCHostnameCheckbox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]700,[System.Int32]164))
    $PCHostnameCheckbox.Name = [System.String]'PCHostnameCheckbox'
    $PCHostnameCheckbox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]270,[System.Int32]50))
    $PCHostnameCheckbox.TabIndex = [System.Int32]10
    $PCHostnameCheckbox.Text = [System.String]'Set PC hostname to:'
    $PCHostnameCheckbox.Checked = $true
    $PCHostnameCheckbox.add_CheckedChanged({Validate})
    $PCHostnameCheckbox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #PCHostnameEntry
    #
    $PCHostnameEntry.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $PCHostnameEntry.Enabled = $true
    $PCHostnameEntry.ForeColor = [System.Drawing.SystemColors]::WindowText
    $PCHostnameEntry.ImeMode = [System.Windows.Forms.ImeMode]::NoControl
    $PCHostnameEntry.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]986,[System.Int32]167))
    $PCHostnameEntry.Name = [System.String]'PCHostnameEntry'
    $PCHostnameEntry.RightToLeft = [System.Windows.Forms.RightToLeft]::No
    $PCHostnameEntry.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]300,[System.Int32]26))
    $PCHostnameEntry.TabIndex = [System.Int32]11
    $PCHostnameEntry.Text = [System.String]""
    $PCHostnameEntry.ImeMode = [System.Windows.Forms.ImeMode]::NoControl
    #NetBIOS PC names can have a max length of 15 characters
    $PCHostnameEntry.MaxLength = 15
    $PCHostnameEntry.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper
    $PCHostnameEntry.add_TextChanged({Validate})
    #
    #AssetTagCheckbox
    #
    $AssetTagCheckbox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $AssetTagCheckbox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]700,[System.Int32]229))
    $AssetTagCheckbox.Name = [System.String]'AssetTagCheckbox'
    $AssetTagCheckbox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]270,[System.Int32]50))
    $AssetTagCheckbox.TabIndex = [System.Int32]12
    $AssetTagCheckbox.Text = [System.String]'Set BIOS asset tag to:'
    $AssetTagCheckbox.Checked = $false
    $AssetTagCheckbox.add_CheckedChanged({Validate})
    $AssetTagCheckbox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    if ($Manufacturer -eq "LENOVO") {
        $AssetTagCheckbox.Enabled = $false
    }
    #
    #AssetTagEntry
    #
    $AssetTagEntry.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $AssetTagEntry.Enabled = $false
    $AssetTagEntry.ForeColor = [System.Drawing.SystemColors]::WindowText
    $AssetTagEntry.ImeMode = [System.Windows.Forms.ImeMode]::NoControl
    $AssetTagEntry.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]986,[System.Int32]232))
    $AssetTagEntry.Name = [System.String]'AssetTagEntry'
    $AssetTagEntry.RightToLeft = [System.Windows.Forms.RightToLeft]::No
    $AssetTagEntry.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]300,[System.Int32]26))
    $AssetTagEntry.TabIndex = [System.Int32]13
    $AssetTagEntry.Text = [System.String]"$CurrentAssetTag"
    $AssetTagEntry.ImeMode = [System.Windows.Forms.ImeMode]::NoControl
    #Asset tags can be a max of 8 characters
    $AssetTagEntry.MaxLength = 8
    $AssetTagEntry.add_TextChanged({Validate})
    #
    #DomainCredentialsLabel
    #
    $DomainCredentialsLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display Semibold',[System.Single]16))
    $DomainCredentialsLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]692,[System.Int32]320))
    $DomainCredentialsLabel.Name = [System.String]'DomainCredentialsLabel'
    $DomainCredentialsLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]400,[System.Int32]60))
    $DomainCredentialsLabel.TabIndex = [System.Int32]14
    $DomainCredentialsLabel.Text = [System.String]'Domain Credentials'
    $DomainCredentialsLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #DomainCredentialsInstructionsLabel
    #
    $DomainCredentialsInstructionsLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]7))
    $DomainCredentialsInstructionsLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]696,[System.Int32]375))
    $DomainCredentialsInstructionsLabel.Name = [System.String]'DomainCredentialsInstructionsLabel'
    $DomainCredentialsInstructionsLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]620,[System.Int32]30))
    $DomainCredentialsInstructionsLabel.TabIndex = [System.Int32]15
    $DomainCredentialsInstructionsLabel.Text = "Enter account credentials that have access to join PCs to the $Domain domain."
    $DomainCredentialsInstructionsLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#4f4f4f")
    $DomainCredentialsInstructionsLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #UsernameLabel
    #
    $UsernameLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $UsernameLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]694,[System.Int32]432))
    $UsernameLabel.Name = [System.String]'UsernameLabel'
    $UsernameLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]250,[System.Int32]50))
    $UsernameLabel.TabIndex = [System.Int32]16
    $UsernameLabel.Text = "Username:      $Domain \"
    $UsernameLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #UsernameEntry
    #
    $UsernameEntry.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $UsernameEntry.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]950,[System.Int32]429))
    $UsernameEntry.Name = [System.String]'UsernameEntry'
    $UsernameEntry.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]336,[System.Int32]26))
    $UsernameEntry.TabIndex = [System.Int32]17
    $UsernameEntry.ImeMode = [System.Windows.Forms.ImeMode]::NoControl
    #Max sAMAccountName length
    $UsernameEntry.MaxLength = 20
    $UsernameEntry.add_TextChanged({Validate})
    #
    #PasswordLabel
    #
    $PasswordLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $PasswordLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]694,[System.Int32]497))
    $PasswordLabel.Name = [System.String]'PasswordLabel'
    $PasswordLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]145,[System.Int32]50))
    $PasswordLabel.TabIndex = [System.Int32]18
    $PasswordLabel.Text = [System.String]'Password:'
    $PasswordLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #PasswordEntry
    #
    $PasswordEntry.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]12))
    $PasswordEntry.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]845,[System.Int32]494))
    $PasswordEntry.Name = [System.String]'PasswordEntry'
    $PasswordEntry.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]441,[System.Int32]26))
    $PasswordEntry.TabIndex = [System.Int32]19
    $PasswordEntry.ImeMode = [System.Windows.Forms.ImeMode]::NoControl
    $PasswordEntry.PasswordChar = "•"
    $PasswordEntry.add_TextChanged({Validate})
    #
    #EyeButton
    #
    $EyeButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1234,[System.Int32]502))
    $EyeButton.Name = [System.String]'EyeButton'
    $EyeButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]42,[System.Int32]25))
    $EyeButton.TabIndex = [System.Int32]20
    $EyeButton.BackgroundImage = $ShowPasswordImage
    $EyeButton.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
    $EyeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $EyeButton.BackColor = $PasswordEntry.BackColor  #Matches form background
    $EyeButton.FlatAppearance.BorderSize = 0
    $EyeButton.TabStop = $false  #Prevent focus outline
    $EyeButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")

    #Toggle Password Visibility
    $EyeButton.add_Click({
        if ($PasswordEntry.UseSystemPasswordChar -or $PasswordEntry.PasswordChar -eq [char]0x2022) {
            $PasswordEntry.PasswordChar = [char]0  #Show text
            $EyeButton.BackgroundImage = $HidePasswordImage
        } else {
            $PasswordEntry.PasswordChar = [char]0x2022  #Hide text with bullet
            $EyeButton.BackgroundImage = $ShowPasswordImage
        }
    })
    #
    #QuestionButton
    #
    $QuestionButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]15))
    $QuestionButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]37,[System.Int32]618))
    $QuestionButton.Name = [System.String]'QuestionButton'
    $QuestionButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]48,[System.Int32]48))
    $QuestionButton.TabIndex = [System.Int32]21
    $QuestionButton.BackgroundImage = $QuestionButtonImage
    $QuestionButton.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
    $QuestionButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $QuestionButton.FlatAppearance.BorderSize = 0
    $QuestionButton.TabStop = $false
    $QuestionButton.UseVisualStyleBackColor = $true
    $QuestionButton.Add_Click({
        Show-HelpForm -QuestionButton $QuestionButton
    })
    #
    #PCInfo
    #
    $PCInfo.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]9))
    $PCInfo.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]122,[System.Int32]618))
    $PCInfo.Name = [System.String]'PCInfo'
    $PCInfo.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]900,[System.Int32]60))
    $PCInfo.TabIndex = [System.Int32]22
    $PCInfo.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f0f0f0")
    $PCInfo.ReadOnly = $true
    $PCInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $PCInfo.ScrollBars = [System.Windows.Forms.ScrollBars]::None
    $PCInfo.TabStop = $false
    #Append the model and line break
    $start = $PCInfo.TextLength
    $PCInfo.AppendText("Model: ")
    $PCInfo.Select($start, 6) #Make "Model:" bold
    $PCInfo.SelectionFont = (New-Object System.Drawing.Font($PCInfo.Font, [System.Drawing.FontStyle]::Bold))
    $PCInfo.Select($PCInfo.TextLength, 0)
    $PCInfo.AppendText("$Model`r`n")
    #Write out the rest of the info, make asset tag, SN, and PN titles bold
    #Asset tag
    $start = $PCInfo.TextLength
    $PCInfo.AppendText("Asset tag: ")
    $PCInfo.Select($start, 10) #Make "Asset tag:" bold
    $PCInfo.SelectionFont = (New-Object System.Drawing.Font($PCInfo.Font, [System.Drawing.FontStyle]::Bold))
    $PCInfo.Select($PCInfo.TextLength, 0)
    $PCInfo.AppendText("$CurrentAssetTag • ")
    #Serial number
    $start = $PCInfo.TextLength
    $PCInfo.AppendText("Serial number: ")
    $PCInfo.Select($start, 14) #Make "Serial number:" bold
    $PCInfo.SelectionFont = (New-Object System.Drawing.Font($PCInfo.Font, [System.Drawing.FontStyle]::Bold))
    $PCInfo.Select($PCInfo.TextLength, 0)
    $PCInfo.AppendText("$SN • ")
    #Product number
    $start = $PCInfo.TextLength
    $PCInfo.AppendText("Product number: ")
    $PCInfo.Select($start, 15) # "Product number:" is bold
    $PCInfo.SelectionFont = (New-Object System.Drawing.Font($PCInfo.Font, [System.Drawing.FontStyle]::Bold))
    $PCInfo.Select($PCInfo.TextLength, 0)
    $PCInfo.AppendText("$PN")
    # Deselect everything
    $PCInfo.Select(0,0)
    #
    #OKButton
    #
    $OKButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]11))
    $OKButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1161,[System.Int32]616))
    $OKButton.Name = [System.String]'OKButton'
    $OKButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]125,[System.Int32]52))
    $OKButton.TabIndex = [System.Int32]23
    $OKButton.Text = [System.String]'OK'
    $OKButton.Enabled = $false
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OKButton.UseVisualStyleBackColor = $true
    #
    #FormBackground
    #
    $FormBackground.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]0,[System.Int32]0))
    $FormBackground.Name = [System.String]'FormBackground'
    $FormBackground.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]3840,[System.Int32]590))
    $FormBackground.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #CompanyNameImagingForm
    #
    $CompanyNameImagingForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]1325,[System.Int32]694))
    $CompanyNameImagingForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f0f0f0")
    $CompanyNameImagingForm.Controls.Add($CompanyNameImagingLabel)
    $CompanyNameImagingForm.Controls.Add($DomainJoinLabel)
    $CompanyNameImagingForm.Controls.Add($DomainJoinInstructionsLabel)
    $CompanyNameImagingForm.Controls.Add($Dom1RadioButton)
    $CompanyNameImagingForm.Controls.Add($Dom2RadioButton)
    $CompanyNameImagingForm.Controls.Add($AutoJoinCheckbox)
    $CompanyNameImagingForm.Controls.Add($SpecificOUCheckbox)
    $CompanyNameImagingForm.Controls.Add($OUDropdown)
    $CompanyNameImagingForm.Controls.Add($SystemConfigurationLabel)
    $CompanyNameImagingForm.Controls.Add($SystemConfigurationInstructionsLabel)
    $CompanyNameImagingForm.Controls.Add($PCHostnameCheckbox)
    $CompanyNameImagingForm.Controls.Add($PCHostnameEntry)
    $CompanyNameImagingForm.Controls.Add($AssetTagCheckbox)
    $CompanyNameImagingForm.Controls.Add($AssetTagEntry)
    $CompanyNameImagingForm.Controls.Add($DomainCredentialsLabel)
    $CompanyNameImagingForm.Controls.Add($DomainCredentialsInstructionsLabel)
    $CompanyNameImagingForm.Controls.Add($UsernameLabel)
    $CompanyNameImagingForm.Controls.Add($UsernameEntry)
    $CompanyNameImagingForm.Controls.Add($PasswordLabel)
    $CompanyNameImagingForm.Controls.Add($EyeButton)
    $CompanyNameImagingForm.Controls.Add($PasswordEntry)
    $CompanyNameImagingForm.Controls.Add($QuestionButton)
    $CompanyNameImagingForm.Controls.Add($PCInfo)
    $CompanyNameImagingForm.Controls.Add($OKButton)
    $CompanyNameImagingForm.Controls.Add($FormBackground)
    $CompanyNameImagingForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $CompanyNameImagingForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $CompanyNameImagingForm.Text = [System.String]'CompanyName Imaging'
    $CompanyNameImagingForm.MaximizeBox = $false
    $CompanyNameImagingForm.MinimizeBox = $false
    $CompanyNameImagingForm.ResumeLayout($false)
    $CompanyNameImagingForm.AcceptButton = $OKButton
    $CompanyNameImagingForm.PerformLayout()

    Clear-Host
    $CompanyNameImagingFormResult = $CompanyNameImagingForm.ShowDialog()
    
    #Return variables outside of the function
    return @{
        CompanyNameImagingFormResult = $CompanyNameImagingFormResult
        
        Domain = if ($Dom1RadioButton.Checked) { "DOM1" }
            elseif ($Dom2RadioButton.Checked) { "DOM2" }
            else { $null }

        DontAutoJoin = $AutoJoinCheckbox.Checked

        SpecificOU = $SpecificOUCheckbox.Checked
        SpecificOUtoJoin = $OUDropdown.Text

        ChangePCName = $PCHostnameCheckbox.Checked
        PCName = $PCHostnameEntry.Text.Trim() #Trim any white space

        ChangeAssetTag = $AssetTagCheckbox.Checked
        AssetTag = $AssetTagEntry.Text.Trim() #Trim any white space

        DomainUsername = $UsernameEntry.Text
        DomainPassword = $PasswordEntry.Text
    }
}

#Function to show a help form when someone clicks ? button
function Show-HelpForm {
    param (
        [System.Windows.Forms.Button]$QuestionButton
    )
    #Windows form for configuration
    Add-Type -AssemblyName System.Windows.Forms

    $HelpForm = New-Object System.Windows.Forms.Form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Label]$HelpFormBackground = $null
    [System.Windows.Forms.Label]$EmailLabel = $null
    [System.Windows.Forms.TextBox]$EmailEntry = $null
    [System.Windows.Forms.Button]$SendButton = $null
    [System.Windows.Forms.Button]$CancelButton = $null

    $HelpFormBackground = (New-Object -TypeName System.Windows.Forms.Label)
    $EmailLabel = (New-Object System.Windows.Forms.Label)
    $EmailEntry = (New-Object System.Windows.Forms.TextBox)
    $SendButton = (New-Object System.Windows.Forms.Button)
    $CancelButton = (New-Object System.Windows.Forms.Button)
    $HelpForm.SuspendLayout()
    #
    #HelpFormBackground
    #
    $HelpFormBackground.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]0,[System.Int32]0))
    $HelpFormBackground.Name = [System.String]'HelpFormBackground'
    $HelpFormBackground.Name = [System.String]'FormBackground'
    $HelpFormBackground.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]3840,[System.Int32]198))
    $HelpFormBackground.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #EmailLabel
    #
    $EmailLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]11))
    $EmailLabel.Text = [System.String]'Enter your email address below if you would like the documentation for CompanyName Imaging to be sent to you.'
    $EmailLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]530,[System.Int32]70))
    $EmailLabel.TabIndex = [System.Int32]0
    $EmailLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]37))
    $EmailLabel.Name = [System.String]'EmailLabel'
    $EmailLabel.AutoSize = $false
    $EmailLabel.TextAlign = 'TopLeft'
    $EmailLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #EmailEntry
    #
    $EmailEntry.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]11))
    $EmailEntry.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]30,[System.Int32]125))
    $EmailEntry.Name = [System.String]'EmailEntry'
    $EmailEntry.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]512,[System.Int32]26))
    $EmailEntry.TabIndex = [System.Int32]1
    $EmailEntry.Text = [System.String]'Email address'
    #
    #SendButton
    #
    $SendButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]10))
    $SendButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]336,[System.Int32]218))
    $SendButton.Name = [System.String]'SendButton'
    $SendButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]42))
    $SendButton.TabIndex = [System.Int32]2
    $SendButton.Text = [System.String]'Send'
    $SendButton.UseVisualStyleBackColor = $true
    $SendButton.Add_Click({
        
        # ADD ALL THE CUSTOM LOGIC YOU WANT HERE FOR SENDING AN EMAIL

        #Show the message box
        Show-MessageForm -Message "Email has been sent."

        #Disable the QuestionButton
        if ($QuestionButton -ne $null) {
            $QuestionButton.Enabled = $false
        }

        #Close the HelpForm
        $HelpForm.Close()
    })
    #
    #CancelButton
    #
    $CancelButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]10))
    $CancelButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]444,[System.Int32]218))
    $CancelButton.Name = [System.String]'CancelButton'
    $CancelButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]42))
    $CancelButton.TabIndex = [System.Int32]3
    $CancelButton.Text = [System.String]'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CancelButton.UseVisualStyleBackColor = $true
    #
    #HelpForm
    #
    $HelpForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]572,[System.Int32]280))
    $HelpForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f0f0f0")
    $HelpForm.Controls.Add($EmailLabel)
    $HelpForm.Controls.Add($EmailEntry)
    $HelpForm.Controls.Add($SendButton)
    $HelpForm.Controls.Add($CancelButton)
    $HelpForm.Controls.Add($HelpFormBackground)
    $HelpForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $HelpForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $HelpForm.Text = "CompanyName Imaging Help"
    $HelpForm.MaximizeBox = $false
    $HelpForm.MinimizeBox = $false
    $HelpForm.ResumeLayout($false)
    $HelpForm.AcceptButton = $SendButton
    $HelpForm.CancelButton = $CancelButton
    $HelpForm.PerformLayout()

    Clear-Host
    $HelpForm.ShowDialog()
}

#Function to show custom format message box
function Show-MessageForm {
    param (
        [string]$Message,
        [string]$Title
    )

    #Windows form for configuration
    Add-Type -AssemblyName System.Windows.Forms

    $MessageForm = New-Object System.Windows.Forms.Form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Label]$MessageFormBackground = $null
    [System.Windows.Forms.Label]$MessageFormLabel = $null
    [System.Windows.Forms.Button]$OKButton = $null

    $MessageFormBackground = (New-Object -TypeName System.Windows.Forms.Label)
    $MessageFormLabel = (New-Object System.Windows.Forms.Label)
    $OKButton = (New-Object System.Windows.Forms.Button)
    $MessageForm.SuspendLayout()
    #
    #MessageFormBackground
    #
    $MessageFormBackground.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]0,[System.Int32]0))
    $MessageFormBackground.Name = [System.String]'MessageFormBackground'
    $MessageFormBackground.Name = [System.String]'FormBackground'
    $MessageFormBackground.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]3840,[System.Int32]106))
    $MessageFormBackground.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #MessageFormLabel
    #
    $MessageFormLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]11))
    $MessageFormLabel.Text = $Message
    $MessageFormLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]207,[System.Int32]32))
    $MessageFormLabel.TabIndex = [System.Int32]0
    $MessageFormLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]37))
    $MessageFormLabel.Name = [System.String]'MessageFormLabel'
    $MessageFormLabel.AutoSize = $false
    $MessageFormLabel.TextAlign = 'TopLeft'
    $MessageFormLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #OKButton
    #
    $OKButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]10))
    $OKButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]134,[System.Int32]126))
    $OKButton.Name = [System.String]'OKButton'
    $OKButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]42))
    $OKButton.TabIndex = [System.Int32]2
    $OKButton.Text = [System.String]'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OKButton.UseVisualStyleBackColor = $true
    #
    #MessageForm
    #
    $MessageForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]261,[System.Int32]186))
    $MessageForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f0f0f0")
    $MessageForm.Controls.Add($MessageFormLabel)
    $MessageForm.Controls.Add($OKButton)
    $MessageForm.Controls.Add($MessageFormBackground)
    $MessageForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $MessageForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $MessageForm.Text = $Title
    $MessageForm.MaximizeBox = $false
    $MessageForm.MinimizeBox = $false
    $MessageForm.ResumeLayout($false)
    $MessageForm.AcceptButton = $OKButton
    $MessageForm.CancelButton = $CancelButton
    $MessageForm.PerformLayout()

    Clear-Host
    $MessageForm.ShowDialog()
}

#Function that shows a form warning you about existing domain computers
function Show-ExistingADComputerForm {
    #Windows form for configuration
    Add-Type -AssemblyName System.Windows.Forms

    $ExistingADComputerForm = New-Object System.Windows.Forms.Form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Label]$ExistingADComputerFormBackground = $null
    [System.Windows.Forms.Label]$ExistingADComputerLabel = $null
    [System.Windows.Forms.Button]$RetryButton = $null
    [System.Windows.Forms.Button]$IgnoreButton = $null

    $ExistingADComputerFormBackground = (New-Object -TypeName System.Windows.Forms.Label)
    $ExistingADComputerLabel = (New-Object System.Windows.Forms.Label)
    $RetryButton = (New-Object System.Windows.Forms.Button)
    $IgnoreButton = (New-Object System.Windows.Forms.Button)
    $ExistingADComputerForm.SuspendLayout()
    #
    #ExistingADComputerFormBackground
    #
    $ExistingADComputerFormBackground.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]0,[System.Int32]0))
    $ExistingADComputerFormBackground.Name = [System.String]'ExistingADComputerFormBackground'
    $ExistingADComputerFormBackground.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]3840,[System.Int32]426))
    $ExistingADComputerFormBackground.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #ExistingADComputerLabel
    #
    $ExistingADComputerLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]11))
    $ExistingADComputerLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]37))
    $ExistingADComputerLabel.Name = [System.String]'ExistingADComputerLabel'
    $ExistingADComputerLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]530,[System.Int32]352))
    $ExistingADComputerLabel.TabIndex = [System.Int32]0
    $ExistingADComputerLabel.Text = [System.String](
        'The computer hostname you entered already exists as a computer object in AD.' + "`n`n" +
        'Press "Retry" to return to the previous form and use a different PC hostname.' + "`n`n" +
        'Press "Ignore" if you just deleted the AD computer object. The installation will continue, attempting the domain join with the computer name you entered. ' +
        'Ensure you have deleted the existing AD computer object from all domain controllers to avoid issues with replication delay.'
    )
    $ExistingADComputerLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fbfbfb")
    #
    #RetryButton
    #
    $RetryButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]10))
    $RetryButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]336,[System.Int32]446))
    $RetryButton.Name = [System.String]'RetryButton'
    $RetryButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]42))
    $RetryButton.TabIndex = [System.Int32]1
    $RetryButton.Text = [System.String]'Retry'
    $RetryButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $RetryButton.UseVisualStyleBackColor = $true
    #
    #IgnoreButton
    #
    $IgnoreButton.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Segoe UI Variable Display',[System.Single]10))
    $IgnoreButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]444,[System.Int32]446))
    $IgnoreButton.Name = [System.String]'IgnoreButton'
    $IgnoreButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]42))
    $IgnoreButton.TabIndex = [System.Int32]2
    $IgnoreButton.Text = [System.String]'Ignore'
    $IgnoreButton.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
    $IgnoreButton.UseVisualStyleBackColor = $true
    #
    #ExistingADComputerForm
    #
    $ExistingADComputerForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]572,[System.Int32]508))
    $ExistingADComputerForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f0f0f0")
    $ExistingADComputerForm.Controls.Add($ExistingADComputerLabel)
    $ExistingADComputerForm.Controls.Add($RetryButton)
    $ExistingADComputerForm.Controls.Add($IgnoreButton)
    $ExistingADComputerForm.Controls.Add($ExistingADComputerFormBackground)
    $ExistingADComputerForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $ExistingADComputerForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $ExistingADComputerForm.Text = "Existing Domain Computer Found"
    $ExistingADComputerForm.MaximizeBox = $false
    $ExistingADComputerForm.MinimizeBox = $false
    $ExistingADComputerForm.ResumeLayout($false)
    $ExistingADComputerForm.AcceptButton = $RetryButton
    $ExistingADComputerForm.PerformLayout()

    Clear-Host
    $ExistingADComputerFormResult = $ExistingADComputerForm.ShowDialog()

    return @{
        ExistingADComputerFormResult = $ExistingADComputerFormResult
    }
}

#endregion Functions

#region Prereqs

#region GetCurrentAssetTag

#Determining manufacturer - aka whether script is running on an HP or Dell
$Manufacturer = (Get-CimInstance Win32_BIOS).Manufacturer

#Also determine model, serial number, and product number
$Model = (Get-CimInstance Win32_ComputerSystem).Model
$SN = (Get-CimInstance Win32_BIOS).SerialNumber
$PN = (Get-CimInstance Win32_ComputerSystem).SystemSKUNumber

#Determining asset tag for HPs:
if ($Manufacturer -eq "HP") {
    if (Test-Path "$ENV:RAMDrive\CompanyNameImaging\SkipDrivers.txt") {
        #Skip retry logic, just try once if SkipDrivers.txt is present
        $CurrentAssetTag = (Get-CimInstance -Namespace root/hp/InstrumentedBIOS -Class HP_BIOSSetting -ErrorAction SilentlyContinue | Where-Object Name -eq "Asset Tracking Number").Value
    }
        else {
            #Test that HP BIOS WMI namespace is initialized, wait up to 2 min for it if not
            $CurrentAssetTag = Test-HPBiosWMINamespace

            if ([string]::IsNullOrWhiteSpace($CurrentAssetTag)) {
                #Even though HP WMI namespace is initialized, sometimes it will still say no
                #asset tag is set (even when it is) for a while after initialization
                $AssetTagAttempt = 0

                do {
                    Start-Sleep -Seconds 5
                    $CurrentAssetTag = Test-HPBiosWMINamespace
                    $AssetTagAttempt++
                }
                while ([string]::IsNullOrWhiteSpace($CurrentAssetTag) -and $AssetTagAttempt -lt 6)

                if ([string]::IsNullOrWhiteSpace($CurrentAssetTag)) {
                    $CurrentAssetTag = "not set"
                }
            }
        }
}
    #Determining it for Dells:
    elseif ($Manufacturer -eq "Dell Inc.") {
        $DellString = Get-CimInstance -Namespace root\dcim\sysman\biosattributes -ClassName StringAttribute
        $DellAsset = ($DellString | Where-Object AttributeName -eq "Asset" | Select-Object CurrentValue)
        $CurrentAssetTag = $DellAsset -replace ".*=" -replace "}"

        if (($CurrentAssetTag -eq "") -or ($CurrentAssetTag -eq $null)) {
	        $CurrentAssetTag = "not set"
        }
    }
    else {
        $CurrentAssetTag = "unknown"
    }

#endregion GetCurrentAssetTag

#Run Find-DriveLetters function
Find-DriveLetters

#region ImagingVersionCheck

#Connect to the automated-winstall-scripts Github project API to determine latest release
$GitHubApiUrl = "https://api.github.com/repos/automated-winstall-scripts/automated-winstall-scripts/releases/latest"

#Connect to API
$APIResponse = Invoke-WebRequest -Uri $GitHubApiUrl -UseBasicParsing

#Get Git version number
$GitImagingVersion = (ConvertFrom-Json $APIResponse.Content).tag_name
#Split the Git version into parts
$GitVersionParts = $GitImagingVersion -split '\.'
#Assign the Git major version number to a variable
$GitMajorVer = $GitVersionParts[0]

#Comparing to version on USB drive
$USBImagingVersion = (Get-Content -Path "$ENV:USBDrive\CompanyNameImagingVersion.txt")
#Split the Git version into parts
$USBVersionParts = $USBImagingVersion -split '\.'
#Assign the Git major version number to a variable
$USBMajorVer = $USBVersionParts[0]

#If the USB Imaging major version is less than the version on Github, throw warning to user telling them to update USB
if ($USBMajorVer -ne $null) {
    if ($USBMajorVer -lt $GitMajorVer) {
        Write-Host -ForegroundColor Yellow "There is a more up-to-date CompanyName Imaging Windows ISO available."
        Write-Host -ForegroundColor Yellow "Consider reflashing your USB soon."
        ""
        pause
        ""
    }
        #If it's not...
        else {
            #But the minor/patch version are different...
            if ($USBImagingVersion -ne $GitImagingVersion) {
                #Then update the USB version to be the same as Git
                $GitImagingVersion | Set-Content -Path "$ENV:USBDrive\CompanyNameImagingVersion.txt"
            }
        }
}

#endregion ImagingVersionCheck

#region UpdateTest-NetworkConnectivityScript

#Download Test.ps1 script from GitHub, replacing version that exists on USB
Save-GithubFiles -Files @("Test-NetworkConnectivity.ps1") -DownloadDirectory "$ENV:USBDrive\CompanyNameImaging"

#endregion UpdateTest-NetworkConnectivityScript

#region FindAndPartitionDisk

#Get all disks and their bus type (NVMe/USB/etc.)
$HardDrives = (Get-Disk).BusType

#If no USB drive is found, tell user to try again, then reboot
if (-not ($HardDrives -match "USB")) {
    Write-Host -ForegroundColor Red "Install USB drive has been unplugged from the system. Plug the drive in again and boot back into it."
    ""
    pause
    wpeutil reboot
    exit
}

#If no internal SSD is detected, warn user, then reboot
if (-not ($HardDrives -match "NVMe|SATA|RAID")) {
    Write-Host -ForegroundColor Red "No internal drive has been detected. Please ensure this system has an SSD installed."
    ""
    Write-Host "If an internal drive is installed and it is not being detected, please reach out to"
    Write-Host "[INSERT NAME HERE] for escalation, as this is likely a driver issue."
    ""
    pause
    wpeutil reboot
    exit
}
    #If internal SSD found...
    else {
        #Get all internal NVMe/SATA/RAID drives larger than 64GB (system requirement for Windows 11)
        $InternalDrives = @(Get-Disk | Where-Object { $_.BusType -in @('NVMe', 'SATA', 'RAID') -and $_.Size -gt 64GB } | Select-Object Number, FriendlyName, BusType, HealthStatus, OperationalStatus, Size, PartitionStyle)
        
        #If multiple internal (NVMe/SATA/RAID) drives are found
        if ($InternalDrives.Count -gt 1) {
            Write-Host -ForegroundColor Yellow "Found multiple internal drives installed in system. Specify which one to install Windows to:"
            
            #Display drives to user
            $InternalDrives | Select-Object Number, FriendlyName, BusType, HealthStatus, OperationalStatus, @{Name="Size (GB)"; Expression={[math]::Round($_.Size / 1GB, 2)}}, PartitionStyle | Sort-Object Number | Format-Table -AutoSize

            #Loop to ensure user selects an available drive number
            $DriveNumber = $null
            while ($DriveNumber -notin $InternalDrives.Number) {
                $DriveNumber = Read-Host "Refer to the `"Number`" column. Select the drive number you want to install Windows on"

                #Check if the entered drive number is valid
                if ($DriveNumber -notin $InternalDrives.Number) {
                    ""
                    Write-Host "Not an available drive number. Please try again..." -ForegroundColor Red
                    ""
                }
                ""
            }

            #Make $Disk variable what user selected
            $Disk = Get-Disk | Where-Object { $_.Number -eq $DriveNumber }
        }
            #If there is only 1 internal (NVMe/SATA/RAID) drive
            else {
                #Make $Disk variable that drive
                $Disk = $InternalDrives
            }
        
        #Create $DiskNumber variable
        $DiskNumber = $Disk.Number

        #Give final warning to user that drive is about to be wiped
        Write-Host -NoNewLine -ForegroundColor Yellow "(PRESS ANY KEY TO STOP) "
        Write-Host -NoNewline "Drive $($Disk.FriendlyName) will be wiped in"
        for ($i = 5; $i -ge 1; $i--) {
            Write-Host " $i..." -NoNewline
            Start-Sleep -Milliseconds 1500
        
            #Check if a key is pressed
            if ([System.Console]::KeyAvailable) {
                $null = [System.Console]::ReadKey($true)
                ""
                Write-Host "Drive wipe cancelled. Rebooting..."
                Start-Sleep -Seconds 5
                wpeutil reboot
                exit
            }
        }

        #Step 1: Wipe the disk
        ""
        ""
        Write-Host "Erasing drive: $($Disk.FriendlyName)..."
        #Clear disk three times
        for ($i = 1; $i -le 3; $i++) {
            #Remove all partitions
            Clear-Disk -Number $DiskNumber -RemoveData -RemoveOEM -Confirm:$false -ErrorAction SilentlyContinue
        }
        #Re-initialize as GPT
        Initialize-Disk -Number $DiskNumber -PartitionStyle GPT

        #Step 2: Create 500MB EFI System partition
        ""
        Write-Host "Creating partitions..."
        $EFI = New-Partition -DiskNumber $DiskNumber -Size 500MB -GptType "{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}"
        Format-Volume -Partition $EFI -FileSystem FAT32 -NewFileSystemLabel "System" -Force | Out-Null

        #Step 3: Create 16MB Microsoft Reserved (MSR) partition
        New-Partition -DiskNumber $DiskNumber -Size 16MB -GptType "{E3C9E316-0B5C-4DB8-817D-F92DF00215AE}" | Out-Null

        #Step 4: Create Primary Windows partition
        $Primary = New-Partition -DiskNumber $DiskNumber -UseMaximumSize
        Format-Volume -Partition $Primary -FileSystem NTFS -NewFileSystemLabel "Windows" -Force | Out-Null
        
        #Recovery partition is made automatically by Windows installer, no need to create it here

        #Step 5: Ensure no other drives are using letter C:\
        #Check if C: is in use by another drive
        $cDrive = Get-Volume -DriveLetter C -ErrorAction SilentlyContinue

        #If it is...
        if ($cDrive) {
            #Get new drive letter to move it to
            $newDriveLetter = Get-AvailableDriveLetter
        
            if ($newDriveLetter) {
                #Reassign C: to new drive letter
                Set-Partition -DriveLetter C -NewDriveLetter $newDriveLetter

                #Run Find-DriveLetters function again
                Find-DriveLetters
            }
        }

        #Step 6: Assign drive letter C: to Windows partition
        Set-Partition -DiskNumber $DiskNumber -PartitionNumber $Primary.PartitionNumber -NewDriveLetter C
    }

#endregion FindAndPartitionDisk

#region RAIDCheck

#Check for Dell computers if RAID is enabled in the BIOS instead of AHCI
if ($Manufacturer -eq "Dell Inc.") {
    #Connect to the PasswordObject WMI class
    $PasswordObject = Get-CimInstance -Namespace root\dcim\sysman\wmisecurity -ClassName PasswordObject
    
    #Check the status of the admin password
    $BIOSPWSetting = $PasswordObject | Where-Object NameId -EQ "Admin" | Select-Object -ExpandProperty IsPasswordSet

    #If a BIOS password is set...
    if ($BIOSPWSetting -eq 1) {
        $Password = "BiosPassword"
        $Encoder = New-Object System.Text.UTF8Encoding
        $Bytes = $Encoder.GetBytes($Password)

        #Connect to the SecurityInterface WMI class
        $SecurityInterface = Get-CimInstance -Namespace root\dcim\sysman\wmisecurity -Class SecurityInterface
    
        #Clear the BIOS password
        $BIOSPWClearStatus = ($SecurityInterface.SetNewPassword(1,$Bytes.Length,$Bytes,"Admin","$Password","")).Status
    }

    #If the BIOS password was already not set or if it was just successfully cleared...
    if (($BIOSPWSetting -ne 1) -or ($BIOSPWClearStatus -eq 0)) {
        #Get all currently set BIOS settings
        $DellBIOSSettings = Get-CimInstance -Namespace root\dcim\sysman\biosattributes -ClassName EnumerationAttribute

        #Determine if RAID or AHCI is enabled
        $StorageSetting = $DellBIOSSettings | Where-Object AttributeName -eq "EmbSataRaid" | Select-Object AttributeName,CurrentValue,PossibleValue

        #If the storage setting is not already set to AHCI...
        if ($StorageSetting.CurrentValue -ne "Ahci") {
            #Connect to the BIOSAttributeInterface WMI class
            $AttributeInterface = Get-CimInstance -Namespace root\dcim\sysman\biosattributes -Class BIOSAttributeInterface

            #Set storage mode to AHCI
            $StorageModeStatus = ($AttributeInterface.SetAttribute(0,0,0,"EmbSataRaid","Ahci")).Status

            if ($StorageModeStatus -eq 0) {
                ""
                Write-Host "Changed BIOS storage setting to AHCI/NVMe (disabled RAID). Rebooting now..."
                Start-Sleep -Seconds 10
                wpeutil reboot
                exit
            }
                else {
                    ""
                    Write-Host -ForegroundColor Red "Failed to change BIOS storage setting to AHCI/NVMe."
                }
        }
    }
}

#endregion RAIDCheck

#region DownloadDomainComputersCSVs

Save-GithubFiles -Files @("DomainComputers/Dom1.csv", "DomainComputers/Dom2.csv") -DownloadDirectory "$ENV:RAMDrive\CompanyNameImaging\DomainComputers"

#endregion DownloadDomainComputersCSVs

#endregion Prereqs

#region ShowForm

#Call the show form function
$FormChoices = Show-ImagingForm

#endregion ShowForm

#region UserHitsOK

if ($FormChoices.CompanyNameImagingFormResult -eq "OK") {

    #region CheckDomainComputersList

    #Only check if there is an existing AD computer with same hostname if it was selected to change hostname in the form
    if ($FormChoices.ChangePCName -eq $true) {
        if ($FormChoices.Domain -eq "Dom1") {
            $Domain = "Dom1"
        }
            elseif ($FormChoices.Domain -eq "Dom2") {
                $Domain = "Dom2"
            }
            else { $Domain = $null }

        $DomainComputers = Import-Csv -Path "$ENV:RAMDrive\CompanyNameImaging\DomainComputers\$Domain.csv"
        $DomainComputerNames = ($DomainComputers | Select-Object -ExpandProperty Name | ForEach-Object { $_.ToUpper() })

        $Repeat = $true

        while ($Repeat) {
            $PCName = $FormChoices.PCName

            if ($DomainComputerNames -contains $PCName) {
                #Show the existing AD computer form
                $ExistingADComputerFormChoices = Show-ExistingADComputerForm

                switch ($ExistingADComputerFormChoices.ExistingADComputerFormResult) {
                    "Retry" {
                        #Show the Imaging form again
                        $FormChoices = Show-ImagingForm
                    }
                    "Ignore" {
                        #Proceed anyway
                        $Repeat = $false
                    }
                }
            }
            else {
                #No match found in CSV, proceed
                $Repeat = $false
            }
        }
    }

    #endregion CheckDomainComputersList

    #region FormChoiceVariables

    #Set return variables from form as their own vars
    $CompanyNameImagingFormResult = $FormChoices.CompanyNameImagingFormResult
    $Domain = $FormChoices.Domain
    $DontAutoJoin = $FormChoices.DontAutoJoin
    $SpecificOU = $FormChoices.SpecificOU
    $SpecificOUtoJoin = $FormChoices.SpecificOUtoJoin
    $ChangePCName = $FormChoices.ChangePCName
    $PCName = $FormChoices.PCName
    $ChangeAssetTag = $FormChoices.ChangeAssetTag
    $AssetTag = $FormChoices.AssetTag
    $DomainUsername = $FormChoices.DomainUsername
    $DomainPassword = $FormChoices.DomainPassword

    $FormChoices = $null

    #endregion FormChoiceVariables

    #region CreateFoldersAndMarkerFiles

    #Create C:\CompanyNameImaging folder
    $CompanyNameImagingDir = (New-Item -Path "C:\CompanyNameImaging" -ItemType Directory -Force).FullName

    #Create C:\CompanyNameImaging\DomainJoin folder
    $DomainJoinDir = (New-Item -Path "$CompanyNameImagingDir\DomainJoin" -ItemType Directory -Force).FullName
    $DomainJoinCredsDir = (New-Item -Path "$DomainJoinDir\Creds" -ItemType Directory -Force).FullName

    #Determine domain that was selected and place marker file in C:\CompanyNameImaging\DomainJoin
    New-Item -Path "$DomainJoinDir" -Name "$Domain.txt" -ItemType "File" | Out-Null
    
    #If the box to not automatically domain join was checked, place marker file in C:\CompanyNameImaging\DomainJoin
    if ($DontAutoJoin -eq $true) {
        New-Item -Path "$DomainJoinDir" -Name "NoAutoDomainJoin.txt" -ItemType "File" | Out-Null
    }

    #If the box join to a specific OU was checked, place marker file in C:\CompanyNameImaging\DomainJoin
    if ($SpecificOU -eq $true) {
        $SpecificOUtoJoin | Set-Content -Path "$DomainJoinDir\SpecificOU.txt"
    }

    #Download the post-setup script to C:\CompanyNameImaging\DomainJoin
    Save-GithubFiles -Files @("setupcomplete.cmd") -DownloadDirectory "$DomainJoinDir"

    #Download the domain join ps1 script to C:\CompanyNameImaging\DomainJoin
    Save-GithubFiles -Files @("Join-Domain.ps1") -DownloadDirectory "$DomainJoinDir"

    #endregion CreateFoldersAndMarkerFiles

    #region HandleCredentials

    #Echo entered username to txt file in C:\CompanyNameImaging\DomainJoin
    $DomainUsername -Replace "Dom1\\" -Replace "Dom2\\" | Set-Content -Path "$DomainJoinCredsDir\Username.txt"

    #Saving credentials
    $Key = New-Object Byte[] 32
    [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
    $Key | Export-Clixml -Path "$DomainJoinCredsDir\EncryptionKey.xml"

    #Convert password entry in form to secure string
    $SecurePassword = ConvertTo-SecureString -String $DomainPassword -AsPlainText -Force

    #Convert from secure string and set content to a txt file
    $SecurePassword | ConvertFrom-SecureString -Key $Key | Set-Content -Path "$DomainJoinCredsDir\EncryptedPassword.txt"

    #Clear the password entry text
    $DomainPassword = $null

    #endregion HandleCredentials

    #region EditAutounattendFile

    #Download autounattend file for the right domain from Github
    $AutounattendFileName = "autounattend-$Domain.xml"

    #Download the autounattend file
    Save-GithubFiles -Files @("Autounattend/$AutounattendFileName") -DownloadDirectory "$ENV:RAMDrive\CompanyNameImaging\Autounattend"

    #Edit the autounattend file to use new computer hostname & set disk #
    $AutounattendFile = "$ENV:RAMDrive\CompanyNameImaging\Autounattend\$AutounattendFileName" #Load the XML file
    
    #Get content
    [xml]$AutounattendFileContent = Get-Content -Path $AutounattendFile

    #Set PC hostname
    if (($ChangePCName -eq $true) -and ($PCName -ne "")) { #Only set if user chose to and if they entered something in the entry box
        #Find ComputerName value
        $SpecializeNode = $AutounattendFileContent.unattend.settings | Where-Object { $_.pass -eq "specialize" }
        $ShellSetupNode = $SpecializeNode.component | Where-Object { $_.name -eq "Microsoft-Windows-Shell-Setup" }

        #Update the ComputerName value
        $ShellSetupNode.ComputerName = "$PCName"
    }
        else {
            #Add namespace manager
            $nsmgr = New-Object System.Xml.XmlNamespaceManager($AutounattendFileContent.NameTable)
            $nsmgr.AddNamespace("un", $AutounattendFileContent.DocumentElement.NamespaceURI)

            #Query using the namespace
            $ComputerNameNode = $AutounattendFileContent.SelectSingleNode("//un:component[@name='Microsoft-Windows-Shell-Setup']/un:ComputerName", $nsmgr)

            #Remove the entire ComputerName node
            if ($ComputerNameNode) {
                $null = $ComputerNameNode.ParentNode.RemoveChild($ComputerNameNode)
            }
        }

    #Also edit the autounattend file's DiskID to install Windows to
    $WindowsPENode = $AutounattendFileContent.unattend.settings | Where-Object { $_.pass -eq "windowsPE" }
    $WindowsSetupNode = $WindowsPENode.component | Where-Object { $_.name -eq "Microsoft-Windows-Setup" }

    #Update the DiskID value
    $WindowsSetupNode.ImageInstall.OSImage.InstallTo.DiskID = "$DiskNumber"

    #Save the changes to the autounattend XML file
    $AutounattendFileContent.Save("$AutounattendFile")
    
    #endregion EditAutounattendFile

    #region ChangeAssetTag

    if ($ChangeAssetTag -eq $true) {

        #Set asset tag for HP computers
        if ($Manufacturer -eq "HP") {
            #Test that HP BIOS WMI namespace is initialized, wait up to 2 min for it if not
            $CurrentAssetTag = Test-HPBiosWMINamespace

            #See if BIOS password is set
            $BIOSPWSetting = (Get-CimInstance -Namespace root/hp/InstrumentedBIOS -Class HP_BIOSSetting | Where-Object Name -eq "Setup Password").IsSet
    
            #Connect to the HP_BIOSSettingInterface WMI class (Get-CimInstance does not work)
            $Interface = Get-WmiObject -Namespace root/hp/InstrumentedBIOS -Class HP_BIOSSettingInterface

            #If a BIOS password is set...
            if ($BIOSPWSetting -eq 1) {
                $Password = "BiosPassword"
                #Clear the BIOS password
                $BIOSPWClearStatus = ($Interface.SetBIOSSetting("Setup Password","<utf-16/>","<utf-16/>" + "$Password")).Return
            }

            #If the BIOS password was already not set or if it was just successfully cleared...
            if (($BIOSPWSetting -ne 1) -or ($BIOSPWClearStatus -eq 0)) {
                #Set the asset tag and record status
                $AssetTagStatus = $Interface.SetBIOSSetting("Asset Tracking Number","$AssetTag","<utf-16/>" + "$Password").Return

                #Asset tag successfully set
	            if ($AssetTagStatus -eq 0) {
		            Write-Host "Asset tag successfully set to $AssetTag."
                }
		            else {
			            Write-Host -ForegroundColor Red "Failure to set asset tag."
		            }
    
                #Set the ownership tag and record status
                $OwnershipTagStatus = $Interface.SetBIOSSetting("Ownership Tag","$AssetTag","<utf-16/>" + "$Password").Return
    
                #Ownership tag successfully set
	            if ($OwnershipTagStatus -eq 0) {
		            Write-Host "Ownership tag successfully set to $AssetTag."
                }
		            else {
			            Write-Host -ForegroundColor Red "Failure to set ownership tag."
		            }
            }
        }
            #Set asset tag for Dell computers
            elseif ($Manufacturer -eq "Dell Inc.") {
                #Connect to the PasswordObject WMI class
                $PasswordObject = Get-CimInstance -Namespace root\dcim\sysman\wmisecurity -ClassName PasswordObject
    
                #Check the status of the admin password
                $BIOSPWSetting = $PasswordObject | Where-Object NameId -EQ "Admin" | Select-Object -ExpandProperty IsPasswordSet

                #If a BIOS password is set...
                if ($BIOSPWSetting -eq 1) {
                    $Password = "BiosPassword"
                    $Encoder = New-Object System.Text.UTF8Encoding
                    $Bytes = $Encoder.GetBytes($Password)

                    #Connect to the SecurityInterface WMI class
                    $SecurityInterface = Get-CimInstance -Namespace root\dcim\sysman\wmisecurity -Class SecurityInterface
    
                    #Clear the BIOS password
                    $BIOSPWClearStatus = ($SecurityInterface.SetNewPassword(1,$Bytes.Length,$Bytes,"Admin","$Password","")).Status
                }

                #If the BIOS password was already not set or if it was just successfully cleared...
                if (($BIOSPWSetting -ne 1) -or ($BIOSPWClearStatus -eq 0)) {
                    #Set the asset tag and record status
                    $AssetTagStatus = (Get-CimInstance -Namespace root\dcim\sysman\biosattributes -Class BIOSAttributeInterface).SetAttribute(1,$Bytes.Length,$Bytes,"Asset","$AssetTag").Status

                    #Asset tag successfully set
		            if ($AssetTagStatus -eq 0) {
		                Write-Host "Asset tag successfully set to $AssetTag."
                    }
		                else {
			                Write-Host -ForegroundColor Red "Failure to set asset tag."
			            }
                }
            }
    }

    #endregion ChangeAssetTag

    #region SetupDrivers

    #Don't run if SkipDrivers.txt exists
    if (!(Test-Path -Path "$ENV:RAMDrive\CompanyNameImaging\SkipDrivers.txt")) {
        #Get paths of drivers that will be installed in Windows image from USB drive and add to $SetupDrivers variable
        $SetupDrivers = @()
        $SetupDrivers += "$ENV:USBDrive\drivers\docks"
    
        #Only get path of HP/Dell/Lenovo drivers depending on PC manufacturer and add to variable
        if ($Manufacturer -eq "HP") {
            $SetupDrivers += "$ENV:USBDrive\drivers\hp"
            $DriverManufacturer = "hp"
        }

        if ($Manufacturer -eq "Dell Inc.") {
            $SetupDrivers += "$ENV:USBDrive\drivers\dell"
            $DriverManufacturer = "dell"
        }

        if ($Manufacturer -eq "LENOVO") {
            $SetupDrivers += "$ENV:USBDrive\drivers\lenovo"
            $DriverManufacturer = "lenovo"
        }

        #Create folders
        $SetupDriversDir = (New-Item -Path "$CompanyNameImagingDir\SetupDrivers" -ItemType Directory -Force).FullName
        New-Item -Path "$SetupDriversDir" -Name "$DriverManufacturer" -ItemType Directory -Force | Out-Null

        #Copy all the drivers that are needed to C:\CompanyNameImaging\SetupDrivers
        foreach ($SetupDriver in $SetupDrivers) {
            Copy-Item -Path "$SetupDriver\*" -Destination "$SetupDriversDir\$DriverManufacturer" -Recurse -Force -Confirm:$false
        }
    }

    #endregion SetupDrivers

    #region LaunchWindowsSetup

    #Launch setup.exe for Windows installation using the custom unattend file and the setup drivers
    if (Test-Path -Path "$ENV:RAMDrive\CompanyNameImaging\SkipDrivers.txt") { #Skip install drivers if SkipDrivers.txt is present
        Start-Process "$ENV:RAMDrive\setup-custom.exe" -Wait -ArgumentList "/NoReboot /Unattend:$AutounattendFile"
    }
        else { #Install drivers if SkipDrivers.txt does not exist
            Start-Process "$ENV:RAMDrive\setup-custom.exe" -Wait -ArgumentList "/NoReboot /Unattend:$AutounattendFile /InstallDrivers $SetupDriversDir"
        }

    Clear-Host

    #endregion LaunchWindowsSetup

    #region Post-Setup

    #Make directory in new Windows installation %WINDIR% for post-setup script
    $SetupScriptDir = (New-Item -Path "C:\Windows\Setup\Scripts" -ItemType Directory -Force).FullName

    #Move the post-setup script that was earlier downloaded to C:\CompanyNameImaging\DomainJoin to the setup scripts directory
    Move-Item -Path "$DomainJoinDir\setupcomplete.cmd" -Destination "$SetupScriptDir\setupcomplete.cmd"

    #Reboot
    wpeutil reboot

    #endregion Post-Setup

}
    else {
        Clear-Host
        Write-Host -ForegroundColor Yellow "Lucky you, you get to troubleshoot!"
    }

#endregion UserHitsOK