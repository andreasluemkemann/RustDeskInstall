$DownloadControl = @{
    Owner   = 'rustdesk'
    Project = 'rustdesk'
    Tag     = 'nightly'
}

$StandardFilter = 'x86_64.exe'
$ScriptPath = $($PSCommandPath)
$CurDir = $(Get-Location)

$global:RustdeskConfig = @'
rendezvous_server = '192.168.71.194:21116'
nat_type = 1
serial = 0

[options]
api-server = "192.168.71.194"
custom-rendezvous-server = "192.168.71.194"
relay-server = "192.168.71.194"
direct-server = 'Y'
enable-audio = 'N'
key = "MMkE1J1ALaiQGwQ7u64Rl0QHlAzEo+cEX3UKGu0Koss="
allow-remove-wallpaper = 'Y'
stop-service = 'N'
'@

$global:RustdeskDefault = @'
[options]
disable_audio = 'Y'
show_remote_cursor = 'Y'
collapse_toolbar = 'Y'
view_style = 'adaptive'
image_quality = 'low'
enable_file_transfer = 'Y'
'@

function New-MessageBox {
    <#
.SYNOPSIS
    New-Popup will display a message box. If a timeout is requested it uses Wscript.Shell PopUp method. If a default button is requested it uses the ::Show method from 'Windows.Forms.MessageBox'
.DESCRIPTION
    The New-Popup command uses the Wscript.Shell PopUp method to display a graphical message
    box. You can customize its appearance of icons and buttons. By default the user
    must click a button to dismiss but you can set a timeout value in seconds to
    automatically dismiss the popup.

    The command will write the return value of the clicked button to the pipeline:
    Timeout = -1
    OK      =  1
    Cancel  =  2
    Abort   =  3
    Retry   =  4
    Ignore  =  5
    Yes     =  6
    No      =  7

    If no button is clicked, the return value is -1.
.PARAMETER Message
    The message you want displayed
.PARAMETER Title
    The text to appear in title bar of dialog box
.PARAMETER Time
    The time to display the message. Defaults to 0 (zero) which will keep dialog open until a button is clicked
.PARAMETER  Buttons
    Valid values for -Buttons include:
    "OK"
    "OKCancel"
    "AbortRetryIgnore"
    "YesNo"
    "YesNoCancel"
    "RetryCancel"
.PARAMETER  Icon
    Valid values for -Icon include:
    "Stop"
    "Question"
    "Exclamation"
    "Information"
    "None"
.PARAMETER  ShowOnTop
    Switch which will force the popup window to appear on top of all other windows.
.PARAMETER  AsString
    Will return a human readable representation of which button was pressed as opposed to an integer value.
.EXAMPLE
    new-popup -message "The update script has completed" -title "Finished" -time 5

    This will display a popup message using the default OK button and default
    Information icon. The popup will automatically dismiss after 5 seconds.
.EXAMPLE
    $answer = new-popup -Message "Please pick" -Title "form" -buttons "OKCancel" -icon "information"

    If the user clicks "OK" the $answer variable will be equal to 1. If the user clicks "Cancel" the
    $answer variable will be equal to 2.
.EXAMPLE
    $answer = new-popup -Message "Please pick" -Title "form" -buttons "OKCancel" -icon "information" -AsString

    If the user clicks "OK" the $answer variable will be equal to 'OK'. If the user clicks "Cancel" the
    $answer variable will be 'Cancel'
.OUTPUTS
    An integer with the following value depending upon the button pushed.

    Timeout = -1    # Value when timer finishes countdown.
    OK      =  1
    Cancel  =  2
    Abort   =  3
    Retry   =  4
    Ignore  =  5
    Yes     =  6
    No      =  7
.LINK
    Wscript.Shell
.NOTES
    Fixed issue with -AsString and a timeout not reporting correctly.
#>

    #region Parameters
    [CmdletBinding(DefaultParameterSetName = 'Timeout')]
    [OutputType('int')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    Param (
        [Parameter(Mandatory, HelpMessage = 'Enter a message for the message box', ParameterSetName = 'DefaultButton')]
        [Parameter(Mandatory, HelpMessage = 'Enter a message for the message box', ParameterSetName = 'Timeout')]
        [ValidateNotNullorEmpty()]
        [string] $Message,

        [Parameter(Mandatory, HelpMessage = 'Enter a title for the message box', ParameterSetName = 'DefaultButton')]
        [Parameter(Mandatory, HelpMessage = 'Enter a title for the message box', ParameterSetName = 'Timeout')]
        [ValidateNotNullorEmpty()]
        [string] $Title,

        [Parameter(ParameterSetName = 'Timeout')]
        [ValidateScript({ $_ -ge 0 })]
        [int] $Time = 0,

        [Parameter(ParameterSetName = 'DefaultButton')]
        [Parameter(ParameterSetName = 'Timeout')]
        [ValidateNotNullorEmpty()]
        [ValidateSet('OK', 'OKCancel', 'AbortRetryIgnore', 'YesNo', 'YesNoCancel', 'RetryCancel')]
        [string] $Buttons = 'OK',

        [Parameter(ParameterSetName = 'DefaultButton')]
        [Parameter(ParameterSetName = 'Timeout')]
        [ValidateNotNullorEmpty()]
        [ValidateSet('None', 'Stop', 'Hand', 'Error', 'Question', 'Exclamation', 'Warning', 'Asterisk', 'Information')]
        [string] $Icon = 'None',

        [Parameter(ParameterSetName = 'Timeout')]
        [switch] $ShowOnTop,

        [Parameter(ParameterSetName = 'DefaultButton')]
        [ValidateSet('Button1', 'Button2', 'Button2')]
        [string] $DefaultButton = 'Button1',

        [Parameter(ParameterSetName = 'DefaultButton')]
        [Parameter(ParameterSetName = 'Timeout')]
        [switch] $AsString

    )
    #endregion Parameters

    begin {
        Write-Verbose -Message "Starting [$($MyInvocation.Mycommand)]"
        Write-Verbose -Message "ParameterSetName [$($PsCmdlet.ParameterSetName)]"

        # set $ShowOnTopValue based on switch
        if ($ShowOnTop) {
            $ShowOnTopValue = 4096
        }
        else {
            $ShowOnTopValue = 0
        }

        #lookup key to convert buttons to their integer equivalents
        $ButtonsKey = ([ordered] @{
                'OK'               = 0
                'OKCancel'         = 1
                'AbortRetryIgnore' = 2
                'YesNo'            = 4
                'YesNoCancel'      = 3
                'RetryCancel'      = 5
            })

        #lookup key to convert icon to their integer equivalents
        $IconKey = ([ordered] @{
                'None'        = 0
                'Stop'        = 16
                'Hand'        = 16
                'Error'       = 16
                'Question'    = 32
                'Exclamation' = 48
                'Warning'     = 48
                'Asterisk'    = 64
                'Information' = 64
            })

        #lookup key to convert return value to human readable button press
        $ReturnKey = ([ordered] @{
                -1 = 'Timeout'
                1  = 'OK'
                2  = 'Cancel'
                3  = 'Abort'
                4  = 'Retry'
                5  = 'Ignore'
                6  = 'Yes'
                7  = 'No'
            })
    }

    process {
        switch ($PsCmdlet.ParameterSetName) {
            'Timeout' {
                try {
                    $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
                    #Button and icon type values are added together to create an integer value
                    $return = $wshell.Popup($Message, $Time, $Title, $ButtonsKey[$Buttons] + $Iconkey[$Icon] + $ShowOnTopValue)
                    if ($return -eq -1) {
                        Write-Verbose -Message "User timedout [$($returnkey[$return])] after [$time] seconds"
                    }
                    else {
                        Write-Verbose -Message "User pressed [$($returnkey[$return])]"
                    }
                    if ($AsString) {
                        $ReturnKey.$return
                    }
                    else {
                        $return
                    }
                }
                catch {
                    #You should never really run into an exception in normal usage
                    Write-Error -Message 'Failed to create Wscript.Shell COM object'
                    Write-Error -Message ($_.exception.message)
                }
            }

            'DefaultButton' {
                try {
                    $MessageBox = [Windows.Forms.MessageBox]
                    $Return = ($MessageBox::Show($Message, $Title, $ButtonsKey[$Buttons], $Iconkey[$Icon], $DefaultButton)).Value__
                    Write-Verbose -Message "User pressed [$($returnkey[$return])]"
                    if ($AsString) {
                        $ReturnKey.$return
                    }
                    else {
                        $return
                    }
                }
                catch {
                    #You should never really run into an exception in normal usage
                    Write-Error -Message 'Failed to create MessageBox'
                    Write-Error -Message ($_.exception.message)
                }
            }
        }
    }

    end {
        Write-Verbose -Message "Ending [$($MyInvocation.Mycommand)]"
    }

} # EndFunction New-MessageBox


Function test-RunAsAdministrator() {
    $ScriptPath = $($PSCommandPath)
    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        # Relaunch as an elevated process:
        Start-Process pwsh.exe "-File", ('"{0}"' -f $ScriptPath) -Verb RunAs
        exit
    }
}

function RustdeskWaitService {
    $global:ServiceName = 'Rustdesk'
    $global:arrService = $(Get-Service -Name $($global:ServiceName) -ErrorAction SilentlyContinue)
    while ($($global:arrService).Status -ne 'Running') {
        $RustdeskInstalled = Test-Path "C:\Program Files\RustDesk"
        if ($RustdeskInstalled) {
            Set-Location $env:ProgramFiles\RustDesk
            Start-Process .\rustdesk.exe --install-service -Verb RunAs
            Start-Sleep -Seconds 6
        }
        Start-Sleep -Seconds 6
        Start-Service $global:ServiceName -ErrorAction SilentlyContinue
        return
    }
}

function DownloadLegacy {
    param (
        [string]$url,
        [string]$targetFile
    )
    Write-Verbose "Legacy Downloading: $url to $targetFile"
    $progressPreference = 'silentlyContinue'
    Invoke-WebRequest -Uri $url -OutFile $targetFile
    $progressPreference = 'Continue'
    return $targetFile
}

function get-DownloadSize {
    param (
        [string]$URL
    )
    $DownloadSizeByte = [int]::Parse(((Invoke-WebRequest $URL -Method Head).Headers.'Content-Length'))
    $DownloadSizeMB = [math]::Round($DownloadSizeByte / 1MB, 2)
    Write-Verbose "URL: $URL Size: $DownloadSizeMB MB"
    return $DownloadSizeMB
}

function DownloadFn($url, $targetFile) {
    Write-Host "DownloadFn Time: $(Get-Date)" -ForegroundColor Yellow
    Write-Host "Downloading: $url" -ForegroundColor yellow
    Write-Host "To: $targetFile" -ForegroundColor cyan
    $uri = New-Object "System.Uri" "$url"
    $request = [System.Net.HttpWebRequest]::Create($uri)
    $request.set_Timeout(15000) #15 second timeout
    $response = $request.GetResponse()
    $totalLength = [System.Math]::Floor($response.get_ContentLength() / 1024)
    $responseStream = $response.GetResponseStream()
    $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create
    $buffer = New-Object byte[] 10KB
    $count = $responseStream.Read($buffer, 0, $buffer.length)
    $downloadedBytes = $count
    while ($count -gt 0) {
        $targetStream.Write($buffer, 0, $count)
        $count = $responseStream.Read($buffer, 0, $buffer.length)
        $downloadedBytes = $downloadedBytes + $count
        Write-Progress -Activity "Downloading file '$($url.split('/') | Select-Object -Last 1)'" -Status "Downloaded ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes / 1024)) / $totalLength) * 100)
    }
    Write-Progress -Activity "Finished downloading file '$($url.split('/') | Select-Object -Last 1)'"
    $targetStream.Flush()
    $targetStream.Close()
    $targetStream.Dispose()
    $responseStream.Dispose()
}

function get-GithubRelease {
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Github Repo Owner')]
        [string]$Owner,
        [Parameter(Mandatory = $true, HelpMessage = 'Github Repo Project Name to')]
        [string]$Project,
        [Parameter(HelpMessage = 'Github Repo Tag to download, defaults to latest')]
        [string]$Tag = "latest",
        [Parameter(HelpMessage = 'Path to download to, defaults to current directory')]
        [string]$Destination = $PWD,
        [Parameter(HelpMessage = 'No filter')]
        [switch]$NoFilter,
        [Parameter(HelpMessage = 'no GUI')]
        [switch]$NoGui
    )
    Set-Location ~
    Write-Verbose "get-GithubRelease Time: $(Get-Date)"
    $Releases = @()
    $DownloadList = @()

    if ($Tag -eq "latest") {
        $URL = "https://api.github.com/repos/$Owner/$Project/releases/$Tag"
    }
    else {
        $URL = "https://api.github.com/repos/$Owner/$Project/releases/tags/$Tag"
    }
    $Releases = (Invoke-RestMethod -Uri $URL).assets.browser_download_url | Where-Object { $_ -like "*$($StandardFilter)" }
    $i = 0

    $Releases | ForEach-Object {
        $i++
        $DownloadList += @(
            [PSCustomObject]@{
                Id          = $i
                File        = ($Filepart = $($_.split('/') | Select-Object -Last 1))
                URL         = $_
                Destination = "$($PWD)\$($Filepart)"
                Size        = $(get-DownloadSize -URL $($_))
            }
        )
    }


    $DownloadList | Out-GridView -Title "Select file to download" -OutputMode Single -OutVariable DownloadSelection
    if ($null -eq $DownloadSelection) {
        Write-Host "No file selected, exiting" -ForegroundColor Red
        exit
    }
    $DownloadList | Where-Object -Property File -EQ $($DownloadSelection.File) | Select-Object -Property File, URL, Destination | ForEach-Object {
        Write-Verbose "Downloading $($_.File) from $($_.URL) to $($_.Destination)"
        DownloadLegacy -url $($_.URL) -targetFile $($_.Destination)
        $global:RustdeskUpdateExe = $($_.Destination)
    }
}

function RustdeskMenu {
    param (
        [string]$RustdeskPath = "C:\Program Files\RustDesk"
    )

    $RustdeskMenu = @{
        'AiO'        = 'Install or Upgrade RustDesk and Configure with your Rendezvous server'
        'Upgrade'    = 'upgrade or install RustDesk, leave configuration as is'
        'Configure'  = 'only configure installed RustDesk to use your own Rustdesk server'
        'Chocolatey' = 'Install Chocolatey'
    }
    $RustdeskMenu | Out-GridView -Title "Select RustDesk action" -OutputMode Single -OutVariable RustdeskAction
    Write-Verbose "RustdeskAction: $($RustdeskAction.Key)"
    switch ($($RustdeskAction.Key)) {
        "AiO" {
            Write-Verbose "Installing RustDesk and configuring with your Rendezvous server"
            get-GithubRelease @DownloadControl -Destination $targetdir
            Get-Service -ErrorAction SilentlyContinue | Where-Object -Property Name -Like "rustdesk" | Stop-Service -ErrorAction SilentlyContinue
            Get-Process -ErrorAction SilentlyContinue | Where-Object -Property Name -Like "rustdesk" | Stop-Process -ErrorAction SilentlyContinue
            $global:RustdeskConfig | Out-File -FilePath "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\RustDesk\config\RustDesk2.toml" -ErrorAction SilentlyContinue -Force
            $global:RustdeskConfig | Out-File -FilePath "$env:USERPROFILE\AppData\Roaming\RustDesk\config\RustDesk2.toml" -ErrorAction SilentlyContinue -Force
            $global:RustdeskDefault | Out-File -FilePath "$env:USERPROFILE\AppData\Roaming\RustDesk\config\RustDesk_default.toml" -ErrorAction SilentlyContinue -Force
            Start-Process -FilePath $global:RustdeskUpdateExe -ArgumentList "--silent-install" -Verb RunAs
            RustdeskWaitService
            Set-Location "$env:ProgramFiles\RustDesk"
            .\rustdesk.exe --get-id | Write-Output -OutVariable RustdeskID
            $rustdeskResult = "Successfully Installed Rustdesk, your ID is $RustdeskID"
            Write-Host $rustdeskResult -ForegroundColor Green
            # import Windows Forms assembly

        }
        "Upgrade" {
            Write-Verbose "Upgrading RustDesk"
            get-GithubRelease @DownloadControl -Destination $targetdir
            Get-Service -ErrorAction SilentlyContinue | Where-Object -Property Name -Like "rustdesk" | Stop-Service -ErrorAction SilentlyContinue
            Get-Process -ErrorAction SilentlyContinue | Where-Object -Property Name -Like "rustdesk" | Stop-Process -ErrorAction SilentlyContinue
            Start-Process -FilePath $global:RustdeskUpdateExe -ArgumentList "--silent-install" -Verb RunAs
            Start-Process -FilePath "C:\Program Files\RustDesk\rustdesk.exe" -ArgumentList "--install-service" -Verb RunAs -WorkingDirectory $RustdeskPath -ErrorAction SilentlyContinue
        }
        "Configure" {
            Write-Verbose "Configuring RustDesk"
            Get-Service -ErrorAction SilentlyContinue | Where-Object -Property Name -Like "rustdesk" | Stop-Service -ErrorAction SilentlyContinue
            Get-Process -ErrorAction SilentlyContinue | Where-Object -Property Name -Like "rustdesk" | Stop-Process -Force -ErrorAction SilentlyContinue
            $RustdeskConfig | Out-File -FilePath "$env:USERPROFILE\AppData\Roaming\RustDesk\config\RustDesk2.toml" -ErrorAction SilentlyContinue
            $RustdeskDefault | Out-File -FilePath "$env:USERPROFILE\AppData\Roaming\RustDesk\config\RustDesk_default.toml" -ErrorAction SilentlyContinue -Force
            $RustdeskConfig | Out-File -FilePath "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\RustDesk\config\RustDesk2.toml" -ErrorAction SilentlyContinue
            if (Test-Path "C:\Program Files\RustDesk\rustdesk.exe") {
                Start-Process -FilePath "C:\Program Files\RustDesk\rustdesk.exe" -ArgumentList "--install-service" -Verb RunAs -WorkingDirectory $RustdeskPath -ErrorAction SilentlyContinue
            }
            else {
                $installMsgBox = New-MessageBox -Message "Rustdesk Configured but not yet installed. Install Rustdesk?" -Title "Rustdesk Install" -buttons "YesNo" -icon "information" -AsString
                if ($installMsgBox -eq "Yes") {
                    get-GithubRelease @DownloadControl -Destination $targetdir
                    Start-Process -FilePath $global:RustdeskUpdateExe -ArgumentList "--silent-install" -Verb RunAs
                    Start-Process -FilePath "C:\Program Files\RustDesk\rustdesk.exe" -ArgumentList "--install-service" -Verb RunAs -WorkingDirectory $RustdeskPath -ErrorAction SilentlyContinue
                }
            }
        }
        "Chocolatey" {
            Write-Verbose "Install Chocolatey"
            Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        }
    }
}

#Set-Location ~
#Check Script is running with Elevated Privileges
RustdeskMenu
Set-Location $CurDir
