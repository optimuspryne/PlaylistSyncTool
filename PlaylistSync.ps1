Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

class ConfigXML {
    [String]$ConfigXMLName
    [String]$ConfigXMLPath
    [xml]   $ConfigXMLContent
    [String]$FirstRunFlagElePath
    [String]$MusicSrcElePath
    [String]$MusicDestElePath
    [String]$PlaylistsSrcElePath

    ConfigXML([Context]$context) {
        $this.ConfigXMLName = "playlist_sync_conf.xml"
        $this.ConfigXMLPath = Join-Path $context.RootPath -ChildPath $this.ConfigXMLName
        $this.ConfigXMLContent = Get-Content -Path $this.ConfigXMLPath -Raw
        $this.FirstRunFlagElePath = "/config/first_run/flag"
        $this.MusicSrcElePath = "/config/m_src_path/path"
        $this.MusicDestElePath = "/config/m_dest_path/path"
        $this.PlaylistsSrcElePath = "/config/playlists_src_path/path"
    }

    [String[]]GetXMLElements ([string]$elementPath) {
            $nodes = $this.ConfigXMLContent.SelectNodes($elementPath)

        if ($nodes -eq $null -or $nodes.Count -eq 0) {
            Write-Host "No nodes found for XPath '$elementPath'"
            return @()
        }

        return $nodes | ForEach-Object { $_.InnerText.ToString() }
    }

    ModifyXMLElement ([String]$element, [String]$elementPath){
        $this.ConfigXMLContent.SelectSingleNode($elementPath).InnerText = $element
        $this.ConfigXMLContent.Save($this.ConfigXMLPath)
    }
}

class Context {
    [String]$RootPath
    [String]$PackageURI
    [String]$PackageDestination
    [String]$ScriptName
    [String]$ScriptPath
    [String]$BatName
    [String]$BatPath
    [String]$BatIconName
    [String]$BatIconPath
    [String]$PlaylistsPath
    $Playlists

    Context () {
        $this.RootPath = "$($env:LocalAppData)\PlaylistSyncTool\"       
        $this.PackageURI = "https://tjb-public.s3.us-west-004.backblazeb2.com/PlaylistSync.zip"
        $this.PackageDestination = "C:\Windows\Temp\PlaylistSync.zip"
        $this.ScriptName = "Copy-M3UPlaylistFiles.ps1"
        $this.ScriptPath = Join-Path -Path $this.RootPath -ChildPath $this.ScriptName
        $this.BatName = "PlaylistSync.bat"
        $this.BatPath = Join-Path -Path $this.RootPath -ChildPath $this.BatName
        $this.BatIconName = "playlistsync.ico"
        $this.BatIconPath = Join-Path -Path $this.RootPath -ChildPath $this.BatIconName
        $this.PlaylistsPath = Join-Path -Path $this.RootPath -ChildPath "playlists\"   
    }

    UpdatePlaylists() {
        $this.Playlists = Get-ChildItem -Path $this.PlaylistsPath
    }

    AddPlaylist($playlist) {
        $this.Playlists += $playlist
    }
}

function Make-Shortcut ([String]$LinkPath, [String]$TargetPath, [String]$IconPath){
    $iconIndex = 0
    $shell = New-Object -comObject WScript.Shell
    $shortcut = $shell.CreateShortcut($LinkPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.IconLocation = "$IconPath,$iconIndex"
    $shortcut.Save()
}

function Folder-Select () {
    try {
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Please choose a folder"
        $folderBrowser.ShowNewFolderButton = $false

        if ($folderBrowser.ShowDialog() -eq "OK") {
            $selectedPath = $folderBrowser.SelectedPath
            return $selectedPath
        } else {
            Write-Host "No folder selected."
            break
        }
    }catch {
        Write-Host "Error: Unable to use selected folder" -ForegroundColor Red
        break
    }
}

function Show-OK-Dialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [string]$BodyText
    )

    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw "GUI dialogs require STA. Start PowerShell with -STA (e.g., 'pwsh -STA')."
    }

    # Create a tiny invisible owner that is always on top
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost       = $true
    $owner.ShowInTaskbar = $false
    $owner.StartPosition = 'CenterScreen'
    $owner.Size          = [System.Drawing.Size]::new(0,0)
    $owner.Show(); $owner.Hide()

    try {
        [void][System.Windows.Forms.MessageBox]::Show(
            $owner,                    # owner keeps it on top
            $BodyText,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    finally {
        $owner.Dispose()
    }
}

function Show-ActionDialog {
  param(
    [ConfigXML]$XML,
    [string]$Title = "Playlist Sync Tool",
    [string]$Message = "What would you like to do?"
  )

  if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw "GUI dialogs require STA. Start PowerShell with -STA."
  }

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles()

  $form = New-Object System.Windows.Forms.Form
  $form.Text = $Title
  $form.StartPosition = 'CenterScreen'
  $form.FormBorderStyle = 'FixedDialog'
  $form.MinimizeBox = $false
  $form.MaximizeBox = $false
  $form.ClientSize = [System.Drawing.Size]::new(560,200)   # slightly taller to fit path text

  $label = New-Object System.Windows.Forms.Label
  # Show current paths under the message
  $label.Text = ($Message + "`r`n`r`n" +
                 "Playlists Source:        $($XML.GetXMLElements($XML.PlaylistsSrcElePath))`r`n" +
                 "Music Source:     $($XML.GetXMLElements($XML.MusicSrcElePath))`r`n" +
                 "Music Destination: $($XML.GetXMLElements($XML.MusicDestElePath))")
  $label.AutoSize = $true
  $label.Location = [System.Drawing.Point]::new(15,15)
  $form.Controls.Add($label)

  $btnResync = New-Object System.Windows.Forms.Button
  $btnResync.Text = "Re-Sync Playlists"
  $btnResync.Size = [System.Drawing.Size]::new(150,32)
  $btnResync.Location = [System.Drawing.Point]::new(20,110)
  $btnResync.Add_Click({ $form.Tag = 'Resync'; $form.Close() })
  $form.Controls.Add($btnResync)

  $btnPaths = New-Object System.Windows.Forms.Button
  $btnPaths.Text = "Change Paths"
  $btnPaths.Size = [System.Drawing.Size]::new(150,32)
  $btnPaths.Location = [System.Drawing.Point]::new(205,110)
  $btnPaths.Add_Click({ $form.Tag = 'ChangePaths'; $form.Close() })
  $form.Controls.Add($btnPaths)

  $btnExit = New-Object System.Windows.Forms.Button
  $btnExit.Text = "Exit"
  $btnExit.Size = [System.Drawing.Size]::new(150,32)
  $btnExit.Location = [System.Drawing.Point]::new(390,110)
  $btnExit.Add_Click({ $form.Tag = 'Exit'; $form.Close() })
  $form.Controls.Add($btnExit)

  $form.AcceptButton = $btnResync   # Enter
  $form.CancelButton = $btnExit     # Esc

  [void]$form.ShowDialog()
  return $form.Tag   # 'Resync' | 'ChangePaths' | 'Exit' | $null
}

function Populate-Config-XML ([ConfigXML]$XML) {
    Show-OK-Dialog -Title "Please Confirm" -BodyText "I need to know where your Playlist files located are located..."
    $playlistSrcPath = Folder-Select
    $XML.ModifyXMLElement($playlistSrcPath, $XML.PlaylistsSrcElePath)

    Show-OK-Dialog -Title "Please Confirm" -BodyText "I need to know where your Music files are located..."
    $musicSrcPath = Folder-Select
    $XML.ModifyXMLElement($musicSrcPath, $XML.MusicSrcElePath)

    Show-OK-Dialog -Title "Please Confirm" -BodyText "I need to know what drive letter your MP3 Player is using.  Please select ONLY the drive, not the 'music' folder"
    $musicDestPath = Folder-Select
    $XML.ModifyXMLElement($musicDestPath, $XML.MusicDestElePath)
}

function Run-Playlist-Sync ([Context]$Context, [ConfigXML]$XML) {
    [string]$playlistSrcPath = $XML.GetXMLElements($XML.PlaylistsSrcElePath)

    [string]$musicSrcPath = $XML.GetXMLElements($XML.MusicSrcElePath)

    [string]$musicDestPath = $XML.GetXMLElements($XML.MusicDestElePath)

    #Check For New Playlists.  If new playlists exist, add them to the Playlists variable in $currentContext
    $newPlaylists = Get-ChildItem -Path $playlistSrcPath

    foreach ($new in $newPlaylists) {
      $exists = $Context.Playlists -contains $new

      if ($exists -eq $false) {
        $copyPath = Join-Path -Path $Context.PlaylistsPath -ChildPath $new.Name
        Copy-Item -Path $new.FullName -Destination $copyPath
      }   
    }

    $Context.UpdatePlaylists()

    if ($Context.Playlists -eq $null -or $Context.Playlists -eq '') {
        Show-OK-Dialog -Title "ERROR" -BodyText "Unable to locate any playlist files..."
    }

    foreach ($playlist in $Context.Playlists) {
        Write-Host ""
        Write-Host "Playlist being Synced: $($playlist.FullName)" -ForegroundColor Blue

        $params = @{
            PlaylistPath   = $playlist.FullName
            SourceRoot     = $musicSrcPath
            DestinationRoot= $musicDestPath
            RewritePlaylist= $true
        }
        & $Context.ScriptPath @params
    }
}

$currentContext = [Context]::new()
$configXML = [ConfigXML]::new($currentContext)
$firstRunFlag = $configXML.GetXMLElements($configXML.FirstRunFlagElePath)

if ($firstRunFlag -eq "Yes") {
    Show-OK-Dialog -Title "Information" -BodyText "Detected this is the first time you've run this program.  You will need to provide paths for your playlists, music files and the drive letter your MP3 player is using."
    Populate-Config-XML -XML $configXML
    $configXML.ModifyXMLElement("No", $configXML.FirstRunFlagElePath)
    $configXML = [ConfigXML]::new($currentContext)
    $currentContext.UpdatePlaylists()
}else {
    $currentContext.UpdatePlaylists()
    $configXML = [ConfigXML]::new($currentContext)
}

do {
    $choice = Show-ActionDialog -XML $configXML

    switch ($choice) {
        "ReSync" {
           Run-Playlist-Sync -Context $currentContext -XML $configXML  
        }"ChangePaths" {
           Populate-Config-XML -XML $configXML
        }
    }
}until ($choice -eq "Exit")