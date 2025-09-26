# Retro Theme Assistant
# Guided script for applying classic Windows themes to Windows 11

# --- GUI Setup ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Main Window ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Retro Theme Assistant"
$Form.Size = New-Object System.Drawing.Size(500, 450) # Increased height for better layout
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = 'FixedDialog'
$Form.MaximizeBox = $false
$Form.MinimizeBox = $true

# --- GUI Controls ---
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Welcome! Initializing..."
$StatusLabel.Location = New-Object System.Drawing.Point(10, 10)
$StatusLabel.Size = New-Object System.Drawing.Size(460, 50)
$Form.Controls.Add($StatusLabel)

$ThemeListBox = New-Object System.Windows.Forms.ListBox
$ThemeListBox.Location = New-Object System.Drawing.Point(10, 60)
$ThemeListBox.Size = New-Object System.Drawing.Size(340, 280)
$Form.Controls.Add($ThemeListBox)

$ApplyButton = New-Object System.Windows.Forms.Button
$ApplyButton.Text = "Apply Theme"
$ApplyButton.Location = New-Object System.Drawing.Point(360, 60)
$ApplyButton.Size = New-Object System.Drawing.Size(110, 40)
$Form.Controls.Add($ApplyButton)

$RestoreButton = New-Object System.Windows.Forms.Button
$RestoreButton.Text = "Restore Win11"
$RestoreButton.Location = New-Object System.Drawing.Point(360, 110)
$RestoreButton.Size = New-Object System.Drawing.Size(110, 40)
$Form.Controls.Add($RestoreButton)

$RefreshButton = New-Object System.Windows.Forms.Button
$RefreshButton.Text = "Refresh List"
$RefreshButton.Location = New-Object System.Drawing.Point(10, 350)
$RefreshButton.Size = New-Object System.Drawing.Size(110, 30)
$Form.Controls.Add($RefreshButton)

$GuidanceBox = New-Object System.Windows.Forms.GroupBox
$GuidanceBox.Text = "Action Requise"
$GuidanceBox.Location = $ThemeListBox.Location
$GuidanceBox.Size = $ThemeListBox.Size
$GuidanceBox.Visible = $false
$Form.Controls.Add($GuidanceBox)

$GuidanceLabel = New-Object System.Windows.Forms.Label
$GuidanceLabel.Text = ""
$GuidanceLabel.Location = New-Object System.Drawing.Point(10, 20)
$GuidanceLabel.Size = New-Object System.Drawing.Size(320, 120)
$GuidanceBox.Controls.Add($GuidanceLabel)

$DownloadLink = New-Object System.Windows.Forms.LinkLabel
$DownloadLink.Text = "Download Link"
$DownloadLink.Location = New-Object System.Drawing.Point(10, 140)
$DownloadLink.Size = New-Object System.Drawing.Size(320, 40)
$GuidanceBox.Controls.Add($DownloadLink)

# --- Configuration ---
$ToolsDirectory = Join-Path $PSScriptRoot "Tools"
$ThemesDirectory = Join-Path $PSScriptRoot "Themes"
$RequiredThemes = @{
    "WindowsXP" = @{
        DisplayName = "Windows XP (niivu)"
        SearchPattern = "*XP*niivu*.zip"
        DownloadUrl = "https://www.deviantart.com/niivu/art/Windows-11-XP-994351926"
        RetroBarTheme = "Windows XP"
        Found = $false
        FilePath = ""
    }
}

# --- Functions ---

function Write-Log {
    param ([string]$Message)
    $StatusLabel.Text = $Message
    $Form.Refresh()
}

function Get-LatestReleaseInfo {
    param ([string]$RepoPath, [string]$AssetFilter)
    $apiUrl = "https://api.github.com/repos/$RepoPath/releases/latest"
    Write-Log "Fetching latest release from $RepoPath..."
    try {
        $releaseInfo = Invoke-RestMethod -Uri $apiUrl -Method Get
        $asset = $releaseInfo.assets | Where-Object { $_.name -like $AssetFilter } | Select-Object -First 1
        if ($asset) {
            return @{ Url = $asset.browser_download_url; FileName = $asset.name }
        } else {
            Write-Log "ERROR: No asset matching '$AssetFilter' found for $RepoPath."
            return $null
        }
    } catch {
        Write-Log "ERROR: Could not fetch release info from GitHub API for $RepoPath."
        return $null
    }
}

function Download-Tool {
    param ([string]$Url, [string]$OutFile, [string]$ToolName)
    Write-Log "Downloading $ToolName..."
    try {
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
        return $true
    } catch {
        Write-Log "ERROR: Failed to download $ToolName. Please check your internet connection."
        return $false
    }
}

function Install-Tools {
    Write-Log "Step 1/5: Installing required tools..."
    if (-not (Test-Path $ToolsDirectory)) { New-Item -Path $ToolsDirectory -ItemType Directory | Out-Null }

    # --- SecureUxTheme ---
    $secureUxThemePath = "C:\Program Files\SecureUxTheme\SecureUxTheme.exe"
    if (-not (Test-Path $secureUxThemePath)) {
        $releaseInfo = Get-LatestReleaseInfo -RepoPath "namazso/SecureUxTheme" -AssetFilter "*.msi"
        if (-not $releaseInfo) { return $false }
        $secureInstaller = Join-Path $ToolsDirectory $releaseInfo.FileName
        if (-not (Test-Path $secureInstaller)) {
            if (-not (Download-Tool -Url $releaseInfo.Url -OutFile $secureInstaller -ToolName "SecureUxTheme")) { return $false }
        }
        Write-Log "Installing SecureUxTheme..."
        Start-Process msiexec.exe -ArgumentList "/i `"$secureInstaller`" /qn" -Wait
        if (-not (Test-Path $secureUxThemePath)) { Write-Log "ERROR: SecureUxTheme installation failed."; return $false }
    }

    # --- RetroBar ---
    $retroBarPath = "$env:ProgramFiles\RetroBar\RetroBar.exe"
    if (-not (Test-Path $retroBarPath)) {
        $releaseInfo = Get-LatestReleaseInfo -RepoPath "dremin/RetroBar" -AssetFilter "*Installer.exe"
        if (-not $releaseInfo) { return $false }
        $retroInstaller = Join-Path $ToolsDirectory $releaseInfo.FileName
        if (-not (Test-Path $retroInstaller)) {
            if (-not (Download-Tool -Url $releaseInfo.Url -OutFile $retroInstaller -ToolName "RetroBar")) { return $false }
        }
        Write-Log "Installing RetroBar..."
        Start-Process -FilePath $retroInstaller -ArgumentList "/VERYSILENT /NORESTART" -Wait
        if (-not (Test-Path $retroBarPath)) { Write-Log "ERROR: RetroBar installation failed."; return $false }
    }

    # --- Open-Shell ---
    $openShellPath = "$env:ProgramFiles\Open-Shell\StartMenu.exe"
    if (-not (Test-Path $openShellPath)) {
        $releaseInfo = Get-LatestReleaseInfo -RepoPath "Open-Shell/Open-Shell-Menu" -AssetFilter "*Setup.exe"
        if (-not $releaseInfo) { return $false }
        $openInstaller = Join-Path $ToolsDirectory $releaseInfo.FileName
        if (-not (Test-Path $openInstaller)) {
            if (-not (Download-Tool -Url $releaseInfo.Url -OutFile $openInstaller -ToolName "Open-Shell")) { return $false }
        }
        Write-Log "Installing Open-Shell..."
        Start-Process -FilePath $openInstaller -ArgumentList "/VERYSILENT /NORESTART" -Wait
        if (-not (Test-Path $openShellPath)) { Write-Log "ERROR: Open-Shell installation failed."; return $false }
    }

    Write-Log "Tool installation complete."
    return $true
}

function Extract-ThemeArchive {
    param([string]$ThemeKey)
    $themeInfo = $RequiredThemes[$ThemeKey]
    $archivePath = $themeInfo.FilePath
    $destinationPath = Join-Path $ThemesDirectory "Extracted\$ThemeKey"

    if (Test-Path $destinationPath) { Remove-Item -Path $destinationPath -Recurse -Force }
    New-Item -Path $destinationPath -ItemType Directory | Out-Null
    Write-Log "Step 2/5: Extracting theme..."

    if ($archivePath -like "*.zip") {
        try {
            Expand-Archive -Path $archivePath -DestinationPath $destinationPath -Force
            $RequiredThemes[$ThemeKey].ExtractedPath = $destinationPath
            return $true
        } catch { Write-Log "ERROR: Failed to extract ZIP archive. It may be corrupt or in use."; return $false }
    } else { Write-Log "ERROR: Unsupported archive type. Only .zip files are supported."; return $false }
}

function Guide-SystemPatch {
    $patcherPath = "C:\Program Files\SecureUxTheme\SecureUxTheme.exe"
    $flagFile = Join-Path $ToolsDirectory "patch.flag"
    if (Test-Path $flagFile) { Write-Log "System patch previously confirmed. Skipping."; return $true }
    if (-not (Test-Path $patcherPath)) { Write-Log "ERROR: SecureUxTheme not found."; return $false }

    Write-Log "Step 3/5: Guiding user through system patch..."
    $message = "Manual Action Required (One Time Only):`n`nThe SecureUxTheme tool will now open. Please follow these steps:`n1. In the SecureUxTheme window, click the 'Install' button.`n2. Once it's done, close the tool and then click 'OK' on this message box."

    Start-Process -FilePath $patcherPath
    [System.Windows.Forms.MessageBox]::Show($message, "System Patch Required", "OK", "Information")

    $confirmResult = [System.Windows.Forms.MessageBox]::Show("Please confirm: Did you successfully click 'Install' in SecureUxTheme?", "Confirm Patch", "YesNo", "Question")
    if ($confirmResult -eq 'Yes') {
        New-Item -Path $flagFile -ItemType File -Force | Out-Null
        Write-Log "Patching process confirmed."
        return $true
    } else {
        Write-Log "Patching process was not confirmed. Halting."
        return $false
    }
}

function Apply-VisualStyle {
    param([string]$ThemeKey)
    Write-Log "Step 4/5: Applying visual style..."
    $themeInfo = $RequiredThemes[$ThemeKey]
    $extractedPath = $themeInfo.ExtractedPath
    $msstylesFile = Get-ChildItem -Path $extractedPath -Filter "*.msstyles" -Recurse | Select-Object -First 1
    if (-not $msstylesFile) { Write-Log "ERROR: No .msstyles file found in theme archive."; return $false }

    $themeSourceFolder = $msstylesFile.Directory.FullName
    $themeName = $msstylesFile.Directory.Name
    $themeDestinationFolder = Join-Path "C:\Windows\Resources\Themes" $themeName
    $destinationMsstylesPath = Join-Path $themeDestinationFolder $msstylesFile.Name

    try { Copy-Item -Path "$themeSourceFolder\*" -Destination $themeDestinationFolder -Recurse -Force } catch { Write-Log "ERROR: Failed to copy theme files. Check permissions."; return $false }
    try {
        Start-Process "rundll32.exe" -ArgumentList "themecpl.dll,OpenThemeAction `"$destinationMsstylesPath`"" -Wait
        Start-Sleep -Seconds 2
        return $true
    } catch { Write-Log "ERROR: Failed to apply visual style."; return $false }
}

function Configure-RetroTools {
    param([string]$ThemeKey)
    Write-Log "Step 5/5: Configuring retro tools..."
    $themeInfo = $RequiredThemes[$ThemeKey]
    $extractedPath = $themeInfo.ExtractedPath
    Get-Process -Name "RetroBar" -ErrorAction SilentlyContinue | Stop-Process -Force

    $openShellSettingsFile = Get-ChildItem -Path $extractedPath -Filter "*.xml" -Recurse | Select-Object -First 1
    $openShellExecutable = "$env:ProgramFiles\Open-Shell\StartMenu.exe"
    if ($openShellSettingsFile -and (Test-Path $openShellExecutable)) {
        Start-Process -FilePath $openShellExecutable -ArgumentList "-load `"$($openShellSettingsFile.FullName)`"" -Wait
    }

    $retroBarSettingsPath = "$env:APPDATA\RetroBar\Settings.yaml"
    if (Test-Path $retroBarSettingsPath) {
        (Get-Content $retroBarSettingsPath) -replace 'Theme: .*', "Theme: $($themeInfo.RetroBarTheme)" | Set-Content $retroBarSettingsPath
    }

    $retroBarPath = "$env:ProgramFiles\RetroBar\RetroBar.exe"
    if (Test-Path $retroBarPath) { Start-Process -FilePath $retroBarPath } else { Write-Log "WARNING: RetroBar.exe not found."; return $false }
    return $true
}

function Restore-DefaultTheme {
    Write-Log "Restoring default Windows 11 theme..."
    Get-Process -Name "RetroBar" -ErrorAction SilentlyContinue | Stop-Process -Force
    try { Set-ItemProperty -Path "HKCU:\Software\OpenShell\StartMenu" -Name "Enable" -Value 0 } catch {}
    $aeroThemePath = "C:\Windows\Resources\Themes\aero.theme"
    if (Test-Path $aeroThemePath) { Start-Process "rundll32.exe" -ArgumentList "themecpl.dll,OpenThemeAction `"$aeroThemePath`"" -Wait }

    $confirmResult = [System.Windows.Forms.MessageBox]::Show("To finalize the restoration, Windows Explorer needs to be restarted. This will close all your folder windows.`n`nIs it okay to proceed?", "Confirm Restart", "YesNo", "Warning")
    if ($confirmResult -eq 'Yes') {
        Write-Log "Restarting Windows Explorer..."
        Get-Process -Name "explorer" | Stop-Process -Force
        Write-Log "Windows 11 default theme restored."
    } else {
        Write-Log "Explorer restart cancelled. A manual restart is required to see all changes."
    }
}

function Apply-SelectedTheme {
    $selectedItem = $ThemeListBox.SelectedItem
    if (-not $selectedItem) { [System.Windows.Forms.MessageBox]::Show("Please select a theme first.", "No Theme Selected", "OK", "Warning"); return }
    $themeKey = ($RequiredThemes.GetEnumerator() | Where-Object { $_.Value.DisplayName -eq $selectedItem }).Key
    if (-not $themeKey) { Write-Log "ERROR: Could not find theme data for '$selectedItem'."; return }

    $Form.Enabled = $false
    Write-Log "Starting theme application process for $selectedItem..."
    if (-not (Install-Tools)) { Write-Log "Halting due to tool installation failure."; $Form.Enabled = $true; return }
    if (-not (Extract-ThemeArchive -ThemeKey $themeKey)) { Write-Log "Halting due to extraction failure."; $Form.Enabled = $true; return }
    if (-not (Guide-SystemPatch)) { Write-Log "Halting due to patch failure."; $Form.Enabled = $true; return }
    if (-not (Apply-VisualStyle -ThemeKey $themeKey)) { Write-Log "Halting due to visual style failure."; $Form.Enabled = $true; return }
    if (-not (Configure-RetroTools -ThemeKey $themeKey)) { Write-Log "Process finished with warnings." }

    $Form.Enabled = $true
    Write-Log "Theme '$selectedItem' applied successfully!"
}

function Refresh-ThemeList {
    $StatusLabel.Text = "Checking for themes in 'Themes' folder..."
    $ThemeListBox.Items.Clear()
    $foundAnyTheme = $false
    if (-not (Test-Path $ThemesDirectory)) { New-Item -Path $ThemesDirectory -ItemType Directory | Out-Null }

    foreach ($themeKey in $RequiredThemes.Keys) {
        $theme = $RequiredThemes[$themeKey]
        $theme.Found = $false
        $foundFile = Get-ChildItem -Path $ThemesDirectory -Filter $theme.SearchPattern -Recurse | Select-Object -First 1
        if ($foundFile) {
            $ThemeListBox.Items.Add($theme.DisplayName)
            $theme.Found = $true
            $theme.FilePath = $foundFile.FullName
            $foundAnyTheme = $true
        }
    }

    if (-not $foundAnyTheme) {
        $missingTheme = $RequiredThemes.GetEnumerator() | Where-Object { -not $_.Value.Found } | Select-Object -First 1
        if ($missingTheme) {
            $StatusLabel.Text = "Action requise pour continuer."
            $GuidanceLabel.Text = "Le thème $($missingTheme.Value.DisplayName) est manquant. Veuillez le télécharger (fichier .zip), le placer dans le dossier 'Themes' (à côté de ce programme), puis cliquer sur 'Refresh List'."
            $DownloadLink.Links.Clear()
            $DownloadLink.Text = "Ouvrir la page de téléchargement pour $($missingTheme.Value.DisplayName)"
            $DownloadLink.Links.Add(0, $DownloadLink.Text.Length, $missingTheme.Value.DownloadUrl) | Out-Null
            $GuidanceBox.Visible = $true
        }
        $ThemeListBox.Visible = $false
    } else {
        $StatusLabel.Text = "Thèmes chargés. Prêt à appliquer."
        $ThemeListBox.Visible = $true
        $GuidanceBox.Visible = $false
    }
}

# --- Event Handlers ---
$DownloadLink.Add_LinkClicked({ param($sender, $e) if ($e.Link.LinkData) { Start-Process $e.Link.LinkData } })
$Form.Add_Shown({ Refresh-ThemeList })
$RefreshButton.Add_Click({ Refresh-ThemeList })
$ApplyButton.Add_Click({ Apply-SelectedTheme })
$RestoreButton.Add_Click({ Restore-DefaultTheme })

# --- Show the Form ---
[System.Windows.Forms.Application]::EnableVisualStyles()
$Form.ShowDialog()