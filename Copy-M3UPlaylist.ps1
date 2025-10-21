[CmdletBinding(SupportsShouldProcess)]
param(
  [string]$PlaylistPath,
  [string]$SourceRoot,
  [string]$DestinationRoot,
  [switch]$RewritePlaylist,
  [switch]$UsePickers
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------- GUI helpers ----------------
function Assert-STA {
  if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw "GUI pickers require STA. Re-run PowerShell with -STA (e.g., 'pwsh -STA')."
  }
}
function Ensure-WinForms { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
function Show-OpenFileDialog {
  param([string]$Title="Select M3U playlist",[string]$Filter="M3U playlist (*.m3u;*.m3u8)|*.m3u;*.m3u8|All files (*.*)|*.*",[string]$InitialDirectory=$null)
  Ensure-WinForms
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $Title; $dlg.Filter = $Filter; $dlg.Multiselect = $false
  if ($InitialDirectory) { $dlg.InitialDirectory = $InitialDirectory }
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
  return $null
}
function Show-FolderBrowserDialog {
  param([string]$Description="Select a folder")
  Ensure-WinForms
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description = $Description; $dlg.ShowNewFolderButton = $true
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
  return $null
}
function Ask-YesNo {
  param([string]$Text,[string]$Caption="Question")
  Ensure-WinForms
  $res = [System.Windows.Forms.MessageBox]::Show($Text,$Caption,[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
  return ($res -eq [System.Windows.Forms.DialogResult]::Yes)
}

# --------- Collect inputs via pickers if requested or missing ----------
if ($UsePickers -or -not $PlaylistPath -or -not $SourceRoot -or -not $DestinationRoot) {
  Assert-STA
  if (-not $PlaylistPath) {
    $PlaylistPath = Show-OpenFileDialog -Title "Select the .m3u/.m3u8 playlist"
    if (-not $PlaylistPath) { throw "No playlist selected." }
  }
  if (-not $SourceRoot) {
    $SourceRoot = Show-FolderBrowserDialog -Description "Select the SOURCE ROOT (folder that contains 'music', etc.)"
    if (-not $SourceRoot) { throw "No source root selected." }
  }
  if (-not $DestinationRoot) {
    $DestinationRoot = Show-FolderBrowserDialog -Description "Select the DESTINATION ROOT (where to copy files)"
    if (-not $DestinationRoot) { throw "No destination root selected." }
  }
  if (-not $PSBoundParameters.ContainsKey('RewritePlaylist')) {
    $RewritePlaylist = Ask-YesNo -Text "Rewrite playlist paths to be relative to the destination folder?" -Caption "Rewrite Playlist?"
  }
}

# ---------------- Core helpers ----------------
function Normalize-RelPath {
  param([string]$PathText)
  if ([string]::IsNullOrWhiteSpace($PathText)) { return $null }
  if ($PathText.TrimStart().StartsWith('#'))   { return $null }
  $trimmed = $PathText.Trim()
  if ($trimmed -match '^[\\/]+') { return $trimmed.TrimStart('\','/') }              # leading / or \
  if (-not [System.IO.Path]::IsPathRooted($trimmed)) { return $trimmed.TrimStart('\','/') }  # relative
  return $trimmed                                                                       # rooted path
}
function Resolve-SourceFile {
  param([string]$ItemPath,[string]$SourceRoot)
  if ([System.IO.Path]::IsPathRooted($ItemPath) -and $ItemPath -notmatch '^[\\/]+') {
    if (Test-Path -LiteralPath $ItemPath) { return (Resolve-Path -LiteralPath $ItemPath).Path }
    return $null
  }
  $rel = $ItemPath.TrimStart('\','/')
  $candidate = Join-Path -Path $SourceRoot -ChildPath $rel
  if (Test-Path -LiteralPath $candidate) { return (Resolve-Path -LiteralPath $candidate).Path }
  return $null
}

# --------------- Prep destination ---------------
if (-not (Test-Path -LiteralPath $DestinationRoot)) {
  New-Item -ItemType Directory -Path $DestinationRoot | Out-Null
}

# --------------- Read playlist source ---------------
try {
  $rawLines = Get-Content -LiteralPath $PlaylistPath -Encoding UTF8
} catch {
  Write-Error "Failed to read playlist: $($_.Exception.Message)"; exit 1
}

# Build mapping: Source -> Dest, and a Windows-style Relative ('Rel') with backslashes
$resolvedMap  = New-Object System.Collections.Generic.List[Hashtable]
$missingItems = New-Object System.Collections.Generic.List[string]

foreach ($line in $rawLines) {
  $norm = Normalize-RelPath -PathText $line
  if ($null -eq $norm) { continue }

  $src = Resolve-SourceFile -ItemPath $norm -SourceRoot $SourceRoot
  if ($null -eq $src) { $missingItems.Add($line); continue }

  if ([System.IO.Path]::IsPathRooted($norm) -and $norm -notmatch '^[\\/]+') {
    $relFromSource = $null
    try {
      $srcFull = (Resolve-Path -LiteralPath $src).Path
      $srcRootFull = (Resolve-Path -LiteralPath $SourceRoot).Path
      if ($srcFull.StartsWith($srcRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relFromSource = $srcFull.Substring($srcRootFull.Length).TrimStart('\','/')
      }
    } catch {}
    $rel = if ($relFromSource) { $relFromSource } else { Split-Path -Path $src -NoQualifier | TrimStart '\','/' }
    $dest = Join-Path $DestinationRoot $rel
  } else {
    $rel = $norm
    $dest = Join-Path $DestinationRoot $rel
  }

  $resolvedMap.Add(@{
    Source = $src
    Dest   = $dest
    Rel    = ($rel -replace '/', '\')   # force Windows-style separators for playlist output
  })
}

# --------------- Copy files (skip existing) ----------------
$copied  = 0
$skipped = 0

foreach ($item in $resolvedMap) {
  $destDir = Split-Path -Path $item.Dest -Parent
  if (-not (Test-Path -LiteralPath $destDir)) {
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
  }

  # Skip if the destination file already exists
  if (Test-Path -LiteralPath $item.Dest) {
    $skipped++
    Write-Verbose "Skip (exists): $($item.Dest)"
    continue
  }

  if ($PSCmdlet.ShouldProcess($item.Source, "Copy to $($item.Dest)")) {
    Copy-Item -LiteralPath $item.Source -Destination $item.Dest
    $copied++
  }
}

# --------------- Write / copy playlist with CRLF + backslashes ---------------
$playlistName = Split-Path -Path $PlaylistPath -Leaf
$destPlaylist = Join-Path $DestinationRoot $playlistName
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)  # UTF-8 without BOM

if ($RewritePlaylist) {
  $outputLines = New-Object System.Collections.Generic.List[string]
  foreach ($line in $rawLines) {
    $norm = Normalize-RelPath -PathText $line
    if ($null -eq $norm) {
      # keep comment/blank lines; they’ll be joined with CRLF below
      $outputLines.Add($line)
      continue
    }
    $match = $resolvedMap | Where-Object { $_.Rel -ieq ($norm -replace '/', '\') } | Select-Object -First 1
    if ($null -ne $match) {
      # Windows backslashes in rewritten entries
      $outputLines.Add($match.Rel)
    } else {
      # if missing, leave original text (still normalized to CRLF by join)
      $outputLines.Add($line)
    }
  }
  $crlfText = ($outputLines -join "`r`n")
  [System.IO.File]::WriteAllText($destPlaylist, $crlfText, $utf8NoBom)
}
else {
  # Copy original but normalize line endings to CRLF for maximum compatibility
  $origText = Get-Content -LiteralPath $PlaylistPath -Raw -Encoding UTF8
  $crlfText = $origText -replace "`r?`n", "`r`n"
  [System.IO.File]::WriteAllText($destPlaylist, $crlfText, $utf8NoBom)
}

# --------------- Summary ----------------
Write-Host ""
Write-Host "Copied $copied file(s), skipped $skipped existing file(s) to '$DestinationRoot'."
if ($missingItems.Count -gt 0) {
  Write-Warning "Some playlist entries could not be found under SourceRoot:"
  $missingItems | ForEach-Object { Write-Warning "  $_" }
}
Write-Host ("Playlist saved to: " + $destPlaylist)
