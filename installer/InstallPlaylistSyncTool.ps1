Add-Type -AssemblyName System.Windows.Forms

function Make-Shortcut ([String]$LinkPath, [String]$TargetPath, [String]$IconPath){
    $iconIndex = 0
    $shell = New-Object -comObject WScript.Shell
    $shortcut = $shell.CreateShortcut($LinkPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.IconLocation = "$IconPath,$iconIndex"
    $shortcut.Save()
}

function Install-Actions {
    $rootPath = "$($env:LocalAppData)\PlaylistSyncTool\"
    $playlistsPath = Join-Path -Path $rootPath -ChildPath "playlists\"
    $zipName = "PlaylistSync.zip"
    $uri = "https://nas.benpruyne.com/d/s/15Ut9tNz3TQQTWbcUk28VuF3nCz3XOT0/webapi/entry.cgi/SYNO.SynologyDrive.Files/PlaylistSync.zip?api=SYNO.SynologyDrive.Files&method=download&version=2&files=%5B%22id%3A913729252514379413%22%5D&force_download=true&json_error=true&download_type=%22download%22&c2_offload=%22allow%22&_dc=1761101370221&sharing_token=%2272T_a3eSuS3Hy4xRrpsNTL378ezRfDfNgOuh9YENZELzDcY8svChB32AH0ZdIN1QBUYgApsjMGymIiEyRug4e5SWNHXE_9xZUyVz6Z190mmJb2UQergXwNx5gMik8rtbGkGE8jtkHnqB4qIpHB6XhaqQlF6rrWc.Bxn4a8H0PvvxaFcssKIlyQdJcSjHq4KGbZArSynxux1mZS.j2Fee_ajKLAgTOT1LuCqtLgFdEmnAs9ul83pQ.sKg%22&SynoToken=2RIaHrbjwwRb."
    $uriDest = Join-Path -Path $rootPath -ChildPath $zipName
    $iconName = "playlistsync.ico"
    $iconPath = Join-Path -Path $rootPath -ChildPath $iconName
    $exeName = "PlaylistSync.exe"
    $exePath = Join-Path -Path $rootPath -ChildPath $exeName

    #Create Directories in the local AppData folder
    New-Item -Path $rootPath -ItemType Directory
    New-Item -Path $playlistsPath -ItemType Directory

    #Download the script, the exe and the config XMl, place them in the newly created folder in local AppData
    Invoke-WebRequest -Uri $uri -OutFile $uriDest -UseBasicParsing
    Expand-Archive $uriDest -DestinationPath $rootPath
    Remove-Item -Path $uriDest -Force
    
    #Create shortcut on Public desktop to the exe file
    Make-Shortcut -LinkPath "C:\Users\Public\Desktop\PlaylistSync.lnk"-TargetPath $exePath -IconPath $iconPath
}
function Show-OK-Dialog ([String]$Title, [String]$BodyText) {
    [void][System.Windows.Forms.MessageBox]::Show(
        $BodyText,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}
function Show-YesNo-Dialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [string]$Title = 'Confirm',
        [ValidateSet('Question','Information','Warning','Error','None')]
        [string]$Icon = 'Question',
        [switch]$DefaultNo,          # Enter defaults to No instead of Yes
        [switch]$TopMost,            # Keep dialog above other windows
        [switch]$ReturnDialogResult  # Return DialogResult instead of $true/$false
    )

    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw "GUI dialogs require STA. Start PowerShell with -STA (e.g., 'pwsh -STA')."
    }

    $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
    $icon    = [System.Windows.Forms.MessageBoxIcon]::$Icon
    $defBtn  = if ($DefaultNo) {
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    } else {
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1
    }

    # Optional owner form so we can force TopMost reliably
    $owner = $null
    if ($TopMost) {
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.ShowInTaskbar = $false
        $owner.StartPosition = 'CenterScreen'
        $owner.Size = [Drawing.Size]::new(0,0)
        $owner.Show(); $owner.Hide()
    }

    $res = if ($owner) {
        [System.Windows.Forms.MessageBox]::Show($owner, $Message, $Title, $buttons, $icon, $defBtn)
        $owner.Dispose()
    } else {
        [System.Windows.Forms.MessageBox]::Show($Message, $Title, $buttons, $icon, $defBtn)
    }

    if ($ReturnDialogResult) { return $res }
    return ($res -eq [System.Windows.Forms.DialogResult]::Yes)
}

if (Show-YesNo-Dialog -Message "Do you want to install 'Playlist Sync Tool'?" -Title "Install Confirmation") {
    Install-Actions
    Show-OK-Dialog -Title "Installing Playlist Sync" -BodyText "Playlist Sync has been successfully installed, use the Icon on your Desktop to run the program."
}else {
    break
}
