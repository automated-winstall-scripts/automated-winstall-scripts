This repo is an example of how you can ditch SCCM/MDT task sequences for imaging altogether and replace it with pure PowerShell scripts.

# Create a new imaging Windows ISO:

## Requirements:
- Stock Windows 11 ISO from Microsoft/VLSC/O365 Admin Center/Massgrave
- Optional but recommended: Language and Optional Features for Windows 11 ISO: https://learn.microsoft.com/en-us/azure/virtual-desktop/windows-11-language-packs
- An installation of Deployment Tools from the Windows ADK + the WinPE add-on for ADK: https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install
- 7-Zip installed

## Run the `New-ImagingISO.ps1` script:
Open a Terminal/PowerShell as admin and enter:

```
iwr -useb https://raw.githubusercontent.com/automated-winstall-scripts/automated-winstall-scripts/refs/heads/main/New-ImagingISO.ps1 | iex
```

This downloads and runs the `New-ImagingISO.ps1` script, which converts the normal Windows 11 ISO into your imaging ISO. If you clone or take inspiration from this repo, you can edit `New-ImagingISO.ps1` to do whatever you want (obviously), but it gives you a blueprint for what can be done.

## What `New-ImagingISO.ps1` does (at a high level):
The `New-ImagingISO.ps1` script exports only the "Windows Setup" index from `boot.wim` and mounts it. Within that mounted image, it renames the `setup.exe` within to `setup-custom.exe` (preventing it from automatically launching), and downloads the `startnet.cmd` script in the "files" folder from this repo to `C:\PathToMount\Windows\System32` (this is the script that automatically launches upon booting into Windows Setup).
3. `startnet.cmd` calls another script, `Test-NetworkConnectivity.ps1`, which kicks off the entire PowerShell imaging sequence.

# Basic workflow:

1. `New-ImagingISO.ps1` creates your imaging ISO, which you can then flash to a USB drive with Rufus or a similar tool or serve over PXE boot.
2. Upon booting into your new imaging ISO, `startnet.cmd` is called, which calls `Test-NetworkConnectivity.ps1`.
3. `Test-NetworkConnectivity.ps1` calls `Start-Imaging.ps1`, which handles the actual Windows install.
4. After Windows installation, `Start-Imaging.ps1` copies `setupcomplete.cmd` from this repo and places it into `C:\Windows\Setup\Scripts` (this is the path for custom Windows setup scripts, which run after booting into the Windows install but pre-OOBE).
5. `setupcomplete.cmd` calls `Join-Domain.ps1`.
6. If you need to run even more scripts, you can copy them to `shell:common startup` (`C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup`), where they will automatically launch after logging in with a user account.

# Private repo option:

If you want your GitHub repo that contains all your scripts to be private, the `Save-GitHubFiles` function can be changed to specify a personal access token when downloading files:

```
function Save-GitHubFiles {
    param (
        [string[]]$Files,
        [string]$DownloadDirectory
    )

    #Put your personal access token here. You can create one here: https://github.com/settings/tokens
    $Credentials = "ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

    #The GitHub repo where the files are (replace with your repo)
    $Repo = "automated-winstall-scripts/automated-winstall-scripts"

    #Create headers for GitHub download
    $Headers = @{
        Authorization = "token $Credentials"
        Accept        = "application/json"
    }

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
            Invoke-WebRequest -Uri $Download -Headers $Headers -OutFile "$DownloadDirectory\$File" -UseBasicParsing
        }
    }
}
```
