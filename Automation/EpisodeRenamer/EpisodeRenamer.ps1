
<#
.SYNOPSIS
    Normalizes TV show episode filenames using a graphical interface.

.DESCRIPTION
    Episode Renamer is a PowerShell Windows Forms application designed
    to scan a selected folder and normalize TV show episode filenames
    using the standard SxxExx naming convention.

    The script automatically detects episode numbers from common naming
    patterns, including:
        - S01E01
        - S1E1
        - 1x01
        - 01x01

    Files are renamed using the following format:
        SxxExx.extension

    The application provides:
        - A graphical user interface with a dark theme.
        - Folder selection through a dialog or drag and drop.
        - Optional recursive folder scanning.
        - Configurable file extension filtering.
        - Preview mode enabled by default.
        - Automatic conflict handling.
        - Rename operation with session-based undo support.
        - Operation logging with export capability.
        - Progress reporting and stop request support.

    No files are modified while preview mode is enabled.

.EXAMPLE
    .\EpisodeRenamer.ps1

    Launches the graphical application.

.EXAMPLE
    Drag and drop a folder onto the application.

    Selects the dropped folder as the source directory.

.EXAMPLE
    Drag and drop a video file onto the application.

    Selects the parent folder of the dropped file as the source directory.

.NOTES
    Name        : EpisodeRenamer.ps1
    Author      : KikeCorp
    Website     : https://kikecorp.eu/
    Repository  : https://github.com/KikeCorp/powershell-toolbox/tree/main/Automation/EpisodeRenamer
    Version     : 1.0.0
    Requires    : Windows PowerShell 5.1 or later
    Platform    : Windows
    UI          : Windows Forms

    Supported default extensions:
        MKV, MP4, AVI, MOV

    Leaving the extension filter empty allows all file extensions
    to be processed.

    Undo functionality is available only during the current
    application session.

.LINK
    https://kikecorp.eu/

.LINK
    https://github.com/KikeCorp/powershell-toolbox/tree/main/Automation/EpisodeRenamer
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ─────────────────────────────────────────────
#  COULEURS & STYLE
# ─────────────────────────────────────────────
$C_BG        = [System.Drawing.Color]::FromArgb(18, 18, 28)
$C_BG2       = [System.Drawing.Color]::FromArgb(24, 24, 38)
$C_SURFACE   = [System.Drawing.Color]::FromArgb(28, 28, 42)
$C_SURFACE2  = [System.Drawing.Color]::FromArgb(38, 38, 56)
$C_BORDER    = [System.Drawing.Color]::FromArgb(55, 55, 80)
$C_ACCENT    = [System.Drawing.Color]::FromArgb(99, 102, 241)   # indigo
$C_ACCENT2   = [System.Drawing.Color]::FromArgb(139, 92, 246)   # violet (pour le dégradé)
$C_SUCCESS   = [System.Drawing.Color]::FromArgb(52, 211, 153)   # vert
$C_WARN      = [System.Drawing.Color]::FromArgb(251, 191, 36)   # jaune
$C_ERROR     = [System.Drawing.Color]::FromArgb(248, 113, 113)  # rouge
$C_TEXT      = [System.Drawing.Color]::FromArgb(226, 232, 240)
$C_SUBTEXT   = [System.Drawing.Color]::FromArgb(148, 163, 184)
$C_BTN_HOVER = [System.Drawing.Color]::FromArgb(79, 82, 220)

$script:HasSemibold = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name -contains "Segoe UI Semibold"
$BOLD_FAMILY = if ($script:HasSemibold) { "Segoe UI Semibold" } else { "Segoe UI" }
$BOLD_STYLE  = if ($script:HasSemibold) { [System.Drawing.FontStyle]::Regular } else { [System.Drawing.FontStyle]::Bold }

$FONT_MAIN   = New-Object System.Drawing.Font("Segoe UI", 9)
$FONT_BOLD   = New-Object System.Drawing.Font($BOLD_FAMILY, 9, $BOLD_STYLE)
$FONT_TITLE  = New-Object System.Drawing.Font("Segoe UI Light", 15)
$FONT_MONO   = New-Object System.Drawing.Font("Consolas", 8.5)
$FONT_SMALL  = New-Object System.Drawing.Font("Segoe UI", 8)
$FONT_ICON   = New-Object System.Drawing.Font("Segoe UI Emoji", 9)

# ─────────────────────────────────────────────
#  FONCTIONS HELPERS UI
# ─────────────────────────────────────────────
function Set-RoundedRegion {
    param($Control, [int]$Radius = 8)
    if ($Control.Width -le 0 -or $Control.Height -le 0) { return }
    $d = $Radius * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $rect = New-Object System.Drawing.Rectangle(0, 0, $Control.Width, $Control.Height)
    $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
    $path.AddArc($rect.Right - $d, $rect.Y, $d, $d, 270, 90)
    $path.AddArc($rect.Right - $d, $rect.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($rect.X, $rect.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    $Control.Region = New-Object System.Drawing.Region($path)
}

function New-StyledButton {
    param($Text, $X, $Y, $W = 130, $H = 32, $Primary = $false)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text       = $Text
    $btn.Location   = New-Object System.Drawing.Point($X, $Y)
    $btn.Size       = New-Object System.Drawing.Size($W, $H)
    $btn.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $btn.Font       = $FONT_BOLD
    $btn.Cursor     = [System.Windows.Forms.Cursors]::Hand
    $btn.FlatAppearance.BorderSize = 1
    $btn.TextAlign  = [System.Drawing.ContentAlignment]::MiddleCenter

    if ($Primary) {
        $btn.BackColor = $C_ACCENT
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.FlatAppearance.BorderColor = $C_ACCENT
        $btn.Tag = "primary"
        $btn.Add_MouseEnter({ if ($this.Enabled) { $this.BackColor = $C_BTN_HOVER } })
        $btn.Add_MouseLeave({ if ($this.Enabled) { $this.BackColor = $this.Tag2 } })
    } else {
        $btn.BackColor = $C_SURFACE
        $btn.ForeColor = $C_TEXT
        $btn.FlatAppearance.BorderColor = $C_BORDER
        $btn.Add_MouseEnter({ if ($this.Enabled) { $this.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 65) } })
        $btn.Add_MouseLeave({ if ($this.Enabled) { $this.BackColor = $C_SURFACE } })
    }
    # Stocke la couleur "normale" courante pour que le survol puisse y revenir
    # même si la couleur de fond change dynamiquement (ex: bouton Lancer/Stop).
    $btn | Add-Member -MemberType NoteProperty -Name Tag2 -Value $btn.BackColor -Force

    $btn.Add_Resize({ Set-RoundedRegion $this 6 })
    Set-RoundedRegion $btn 6
    return $btn
}

function New-Label {
    param($Text, $X, $Y, $Bold = $false, $Color = $null)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Text
    $lbl.Location  = New-Object System.Drawing.Point($X, $Y)
    $lbl.AutoSize  = $true
    $lbl.Font      = if ($Bold) { $FONT_BOLD } else { $FONT_MAIN }
    $lbl.ForeColor = if ($Color) { $Color } else { $C_TEXT }
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    return $lbl
}

function New-Separator {
    param($X, $Y, $W)
    $sep = New-Object System.Windows.Forms.Panel
    $sep.Location  = New-Object System.Drawing.Point($X, $Y)
    $sep.Size      = New-Object System.Drawing.Size($W, 1)
    $sep.BackColor = $C_BORDER
    return $sep
}

function Set-ButtonColor {
    param($Button, $BackColor, $ForeColor = [System.Drawing.Color]::White, $BorderColor = $null)
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.FlatAppearance.BorderColor = if ($BorderColor) { $BorderColor } else { $BackColor }
    $Button.Tag2 = $BackColor
}

# ─────────────────────────────────────────────
#  LOGIQUE MÉTIER
# ─────────────────────────────────────────────
function Get-CleanEpisodeName {
    param([string]$FileName, [string]$ExtFilter)
    $file = [System.IO.FileInfo]$FileName
    if ($ExtFilter -ne "" -and ($ExtFilter -split "," | ForEach-Object { $_.Trim().ToLower() }) -notcontains $file.Extension.ToLower().TrimStart('.')) {
        return $null
    }

    # Format standard : S01E01, S1E1, S01.E01, S01-E01, S01_E01 ...
    if ($file.Name -match '(?i)S(\d{1,2})[.\-_ ]?E(\d{1,2})') {
        $s = $Matches[1].PadLeft(2, '0')
        $e = $Matches[2].PadLeft(2, '0')
        return "S${s}E${e}$($file.Extension.ToLower())"
    }
    # Format alternatif : 1x01, 01x01
    if ($file.Name -match '(?i)(?<![0-9])(\d{1,2})x(\d{2})(?![0-9])') {
        $s = $Matches[1].PadLeft(2, '0')
        $e = $Matches[2].PadLeft(2, '0')
        return "S${s}E${e}$($file.Extension.ToLower())"
    }
    return $null
}

function Get-UniqueFilePath {
    param([string]$Folder, [string]$Name, [string]$ExcludePath = $null)
    $base  = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $ext   = [System.IO.Path]::GetExtension($Name)
    $path  = Join-Path $Folder $Name
    $count = 2
    while ((Test-Path $path) -and ($path -ne $ExcludePath)) {
        $path = Join-Path $Folder "${base}_${count}${ext}"
        $count++
    }
    return $path
}

function Invoke-SafeRename {
    <#
        Renomme en gérant le cas particulier d'un changement de casse uniquement
        (ex: "s01e01.mkv" -> "S01E01.mkv"), que NTFS peut refuser en une seule étape
        car il considère les deux noms comme identiques.
    #>
    param([string]$FullPath, [string]$NewName)
    $oldName = [System.IO.Path]::GetFileName($FullPath)
    if ($oldName -ceq $NewName) { return }  # rien à faire, déjà strictement identique

    if ($oldName -eq $NewName) {
        # Différence de casse uniquement : renommage en deux temps via un nom temporaire
        $tempName = "__tmp_$([Guid]::NewGuid().ToString('N'))"
        Rename-Item -Path $FullPath -NewName $tempName -ErrorAction Stop
        $tempPath = Join-Path ([System.IO.Path]::GetDirectoryName($FullPath)) $tempName
        Rename-Item -Path $tempPath -NewName $NewName -ErrorAction Stop
    } else {
        Rename-Item -Path $FullPath -NewName $NewName -ErrorAction Stop
    }
}

# ─────────────────────────────────────────────
#  FORMULAIRE PRINCIPAL
# ─────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Episode Renamer"
$form.Size            = New-Object System.Drawing.Size(720, 640)
$form.MinimumSize     = New-Object System.Drawing.Size(720, 640)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $C_BG
$form.ForeColor       = $C_TEXT
$form.Font            = $FONT_MAIN
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox     = $false

# ── En-tête (avec léger dégradé) ──
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Location  = New-Object System.Drawing.Point(0, 0)
$pnlHeader.Size      = New-Object System.Drawing.Size(720, 70)
$pnlHeader.BackColor = $C_SURFACE
$pnlHeader.Add_Paint({
    param($s, $e)
    $rect  = $s.ClientRectangle
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $C_SURFACE2, $C_SURFACE, 0)
    $e.Graphics.FillRectangle($brush, $rect)
    $brush.Dispose()
    $accentBrush = New-Object System.Drawing.SolidBrush($C_ACCENT)
    $e.Graphics.FillRectangle($accentBrush, 0, 0, 4, $s.Height)
    $accentBrush.Dispose()
})
$form.Controls.Add($pnlHeader)

$lblTitle = New-Label "▶  Episode Renamer" 24 12 $false $C_TEXT
$lblTitle.Font = $FONT_TITLE
$pnlHeader.Controls.Add($lblTitle)

$lblSub = New-Label "Renommage automatique au format SxxExx  ·  glissez-déposez un dossier" 27 42 $false $C_SUBTEXT
$lblSub.Font = $FONT_SMALL
$pnlHeader.Controls.Add($lblSub)

$sepHeader = New-Separator 0 69 720
$form.Controls.Add($sepHeader)

# ── Zone dossier ──
$pnlFolder = New-Object System.Windows.Forms.Panel
$pnlFolder.Location  = New-Object System.Drawing.Point(20, 85)
$pnlFolder.Size      = New-Object System.Drawing.Size(675, 60)
$pnlFolder.BackColor = [System.Drawing.Color]::Transparent
$pnlFolder.AllowDrop = $true
$form.Controls.Add($pnlFolder)

$lblFolder = New-Label "DOSSIER CIBLE" 0 0 $true $C_SUBTEXT
$lblFolder.Font = $FONT_SMALL
$pnlFolder.Controls.Add($lblFolder)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location    = New-Object System.Drawing.Point(0, 20)
$txtPath.Size        = New-Object System.Drawing.Size(530, 26)
$txtPath.BackColor   = $C_SURFACE
$txtPath.ForeColor   = $C_TEXT
$txtPath.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtPath.Font        = $FONT_MAIN
$txtPath.AllowDrop   = $true
$pnlFolder.Controls.Add($txtPath)

$btnBrowse = New-StyledButton "📂  Parcourir" 540 18 130 28
$btnBrowse.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "Choisissez le dossier contenant vos épisodes"
    if ($fb.ShowDialog() -eq "OK") { $txtPath.Text = $fb.SelectedPath }
})
$pnlFolder.Controls.Add($btnBrowse)

$lblFileCount = New-Label "" 0 50 $false $C_ACCENT
$lblFileCount.Font = $FONT_SMALL
$pnlFolder.Controls.Add($lblFileCount)

# ── Glisser-déposer d'un dossier (panneau + champ texte) ──
$dragHandlerEnter = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
}
$dragHandlerDrop = {
    param($s, $e)
    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($paths -and $paths.Count -gt 0) {
        $p = $paths[0]
        if (Test-Path $p -PathType Container) {
            $txtPath.Text = $p
        } else {
            $txtPath.Text = [System.IO.Path]::GetDirectoryName($p)
        }
    }
}
$pnlFolder.Add_DragEnter($dragHandlerEnter)
$pnlFolder.Add_DragDrop($dragHandlerDrop)
$txtPath.Add_DragEnter($dragHandlerEnter)
$txtPath.Add_DragDrop($dragHandlerDrop)

# ── Options ──
$sepOpt = New-Separator 20 155 675
$form.Controls.Add($sepOpt)

$pnlOptions = New-Object System.Windows.Forms.Panel
$pnlOptions.Location  = New-Object System.Drawing.Point(20, 165)
$pnlOptions.Size      = New-Object System.Drawing.Size(675, 80)
$pnlOptions.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($pnlOptions)

$lblOptions = New-Label "OPTIONS" 0 0 $true $C_SUBTEXT
$lblOptions.Font = $FONT_SMALL
$pnlOptions.Controls.Add($lblOptions)

# Check récursif
$chkRecurse = New-Object System.Windows.Forms.CheckBox
$chkRecurse.Text      = "Inclure les sous-dossiers"
$chkRecurse.Location  = New-Object System.Drawing.Point(0, 22)
$chkRecurse.AutoSize  = $true
$chkRecurse.ForeColor = $C_TEXT
$chkRecurse.BackColor = [System.Drawing.Color]::Transparent
$chkRecurse.Font      = $FONT_MAIN
$pnlOptions.Controls.Add($chkRecurse)

# Check dry-run
$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text      = "Prévisualisation (sans renommer)"
$chkDryRun.Location  = New-Object System.Drawing.Point(200, 22)
$chkDryRun.AutoSize  = $true
$chkDryRun.ForeColor = $C_WARN
$chkDryRun.BackColor = [System.Drawing.Color]::Transparent
$chkDryRun.Font      = $FONT_MAIN
$chkDryRun.Checked   = $true
$pnlOptions.Controls.Add($chkDryRun)

# Filtre extensions
$lblExt = New-Label "Extensions (ex: mkv,mp4,avi) :" 0 52 $false $C_SUBTEXT
$pnlOptions.Controls.Add($lblExt)

$txtExt = New-Object System.Windows.Forms.TextBox
$txtExt.Location    = New-Object System.Drawing.Point(200, 48)
$txtExt.Size        = New-Object System.Drawing.Size(200, 22)
$txtExt.BackColor   = $C_SURFACE
$txtExt.ForeColor   = $C_TEXT
$txtExt.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtExt.Font        = $FONT_MAIN
$txtExt.Text        = "mkv,mp4,avi,mov"
$pnlOptions.Controls.Add($txtExt)

$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($chkDryRun, "Simule les renommages sans toucher aux fichiers")
$tip.SetToolTip($chkRecurse, "Parcourt aussi les sous-dossiers du dossier cible")
$tip.SetToolTip($txtExt, "Laissez vide pour traiter toutes les extensions")

# ── Boutons d'action ──
$sepAction = New-Separator 20 253 675
$form.Controls.Add($sepAction)

$btnRename  = New-StyledButton "▶  Lancer" 20 265 150 34 $true
$btnUndo    = New-StyledButton "↩  Annuler" 178 265 120 34
$btnExport  = New-StyledButton "💾  Exporter log" 306 265 140 34
$btnClear   = New-StyledButton "✕  Vider log" 454 265 110 34

$btnUndo.Enabled   = $false
$btnExport.Enabled = $false

$form.Controls.Add($btnRename)
$form.Controls.Add($btnUndo)
$form.Controls.Add($btnExport)
$form.Controls.Add($btnClear)

# ── Log RichTextBox ──
$sepLog = New-Separator 20 308 675
$form.Controls.Add($sepLog)

$lblLog = New-Label "JOURNAL DES OPÉRATIONS" 20 316 $true $C_SUBTEXT
$lblLog.Font = $FONT_SMALL
$form.Controls.Add($lblLog)

$rtb = New-Object System.Windows.Forms.RichTextBox
$rtb.Location    = New-Object System.Drawing.Point(20, 336)
$rtb.Size        = New-Object System.Drawing.Size(675, 195)
$rtb.BackColor   = $C_SURFACE
$rtb.ForeColor   = $C_TEXT
$rtb.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$rtb.ReadOnly    = $true
$rtb.Font        = $FONT_MONO
$rtb.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$form.Controls.Add($rtb)

# ── Barre de statut ──
$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Location  = New-Object System.Drawing.Point(0, 540)
$pnlStatus.Size      = New-Object System.Drawing.Size(720, 42)
$pnlStatus.BackColor = $C_SURFACE
$form.Controls.Add($pnlStatus)

$pbr = New-Object System.Windows.Forms.ProgressBar
$pbr.Location  = New-Object System.Drawing.Point(20, 10)
$pbr.Size      = New-Object System.Drawing.Size(350, 8)
$pbr.Style     = [System.Windows.Forms.ProgressBarStyle]::Continuous
$pbr.BackColor = $C_BG
$pbr.ForeColor = $C_ACCENT
$pnlStatus.Controls.Add($pbr)

$lblStats = New-Label "Prêt." 385 12 $false $C_SUBTEXT
$lblStats.Font = $FONT_SMALL
$pnlStatus.Controls.Add($lblStats)

# ─────────────────────────────────────────────
#  ÉTAT GLOBAL
# ─────────────────────────────────────────────
$script:undoStack        = @()   # stocke les paires [newPath, oldName]
$script:logLines         = @()   # pour l'export
$script:isRunning        = $false
$script:cancelRequested  = $false

# ─────────────────────────────────────────────
#  HELPERS LOG
# ─────────────────────────────────────────────
function Write-Log {
    param([string]$Text, [System.Drawing.Color]$Color)
    $rtb.SelectionStart  = $rtb.TextLength
    $rtb.SelectionLength = 0
    $rtb.SelectionColor  = $Color
    $rtb.AppendText("$Text`n")
    $rtb.ScrollToCaret()
    $script:logLines += $Text
}

function Write-LogInfo    { param($t) Write-Log $t $C_SUBTEXT }
function Write-LogSuccess { param($t) Write-Log $t $C_SUCCESS }
function Write-LogWarn    { param($t) Write-Log $t $C_WARN }
function Write-LogError   { param($t) Write-Log $t $C_ERROR }

# ─────────────────────────────────────────────
#  APERÇU DU NOMBRE DE FICHIERS CORRESPONDANTS
# ─────────────────────────────────────────────
$previewTimer = New-Object System.Windows.Forms.Timer
$previewTimer.Interval = 400

function Update-FilePreview {
    $folderPath = $txtPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folderPath) -or -not (Test-Path -Path $folderPath -PathType Container)) {
        $lblFileCount.Text = ""
        return
    }
    try {
        $extFilter = $txtExt.Text.Trim()
        $getParams = @{ Path = $folderPath; File = $true; ErrorAction = "SilentlyContinue" }
        if ($chkRecurse.Checked) { $getParams['Recurse'] = $true }
        $matchCount = 0
        $totalCount = 0
        foreach ($f in Get-ChildItem @getParams) {
            $totalCount++
            if (Get-CleanEpisodeName $f.FullName $extFilter) { $matchCount++ }
        }
        $lblFileCount.Text = "🔎  $matchCount épisode(s) détecté(s) sur $totalCount fichier(s)"
    } catch {
        $lblFileCount.Text = "⚠  Impossible de lire ce dossier"
    }
}

$previewTimer.Add_Tick({
    $previewTimer.Stop()
    Update-FilePreview
})
$txtPath.Add_TextChanged({ $previewTimer.Stop(); $previewTimer.Start() })
$txtExt.Add_TextChanged({ $previewTimer.Stop(); $previewTimer.Start() })
$chkRecurse.Add_CheckedChanged({ Update-FilePreview })

# ─────────────────────────────────────────────
#  ÉTAT DE L'INTERFACE (pendant le traitement)
# ─────────────────────────────────────────────
function Set-UIRunning {
    param([bool]$Running)
    $script:isRunning = $Running
    $btnBrowse.Enabled   = -not $Running
    $txtPath.Enabled     = -not $Running
    $txtExt.Enabled      = -not $Running
    $chkRecurse.Enabled  = -not $Running
    $chkDryRun.Enabled   = -not $Running
    $btnUndo.Enabled     = (-not $Running) -and ($script:undoStack.Count -gt 0)
    $btnExport.Enabled   = (-not $Running) -and ($script:logLines.Count -gt 0)
    $btnClear.Enabled    = -not $Running

    if ($Running) {
        Set-ButtonColor $btnRename $C_ERROR ([System.Drawing.Color]::White) $C_ERROR
        $btnRename.Text = "■  Arrêter"
    } else {
        Set-ButtonColor $btnRename $C_ACCENT ([System.Drawing.Color]::White) $C_ACCENT
        $btnRename.Text = "▶  Lancer"
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# ─────────────────────────────────────────────
#  ACTION : RENOMMER / PRÉVISUALISER / ARRÊTER
# ─────────────────────────────────────────────
$btnRename.Add_Click({
    if ($script:isRunning) {
        $script:cancelRequested = $true
        return
    }

    $folderPath = $txtPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folderPath) -or -not (Test-Path -Path $folderPath -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show("Le dossier spécifié n'existe pas.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $isDryRun  = $chkDryRun.Checked
    $isRecurse = $chkRecurse.Checked
    $extFilter = $txtExt.Text.Trim()

    $rtb.Clear()
    $script:logLines        = @()
    $script:undoStack       = @()
    $script:cancelRequested = $false
    Set-UIRunning $true
    $form.Cursor = [System.Windows.Forms.Cursors]::AppStarting

    $mode = if ($isDryRun) { "PRÉVISUALISATION" } else { "RENOMMAGE" }
    Write-LogInfo "══════════════════════════════════════════"
    Write-LogInfo "  $mode  —  $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
    Write-LogInfo "  Dossier : $folderPath"
    Write-LogInfo "══════════════════════════════════════════"

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        $getParams = @{ Path = $folderPath; File = $true; ErrorAction = "Stop" }
        if ($isRecurse) { $getParams['Recurse'] = $true }
        $files = @(Get-ChildItem @getParams)
    } catch {
        Write-LogError "  ✖  Impossible de lister les fichiers : $_"
        $files = @()
    }

    $total    = $files.Count
    $cRenamed = 0; $cIgnored = 0; $cConflict = 0; $cError = 0
    $wasCancelled = $false

    $pbr.Maximum = [Math]::Max($total, 1)
    $pbr.Value   = 0

    foreach ($file in $files) {
        if ($script:cancelRequested) {
            $wasCancelled = $true
            Write-LogWarn "  ⏹  Opération arrêtée par l'utilisateur."
            break
        }

        $pbr.Value = [Math]::Min($pbr.Value + 1, $pbr.Maximum)
        $lblStats.Text = "Traitement $($pbr.Value) / $total ..."
        [System.Windows.Forms.Application]::DoEvents()

        $newName = Get-CleanEpisodeName $file.FullName $extFilter

        if ($null -eq $newName) {
            Write-LogInfo "  ⊘  IGNORÉ    $($file.Name)"
            $cIgnored++
            continue
        }

        if ($newName -ceq $file.Name) {
            Write-LogInfo "  ✓  DÉJÀ OK   $($file.Name)"
            $cIgnored++
            continue
        }

        $destPath = Get-UniqueFilePath $file.DirectoryName $newName $file.FullName
        $destName = [System.IO.Path]::GetFileName($destPath)

        if ($destName -ne $newName) {
            Write-LogWarn "  ⚠  CONFLIT   $($file.Name)  →  $destName  (renommage auto)"
            $cConflict++
        }

        if ($isDryRun) {
            Write-LogSuccess "  →  APERÇU    $($file.Name)  →  $destName"
            $cRenamed++
        } else {
            try {
                Invoke-SafeRename -FullPath $file.FullName -NewName $destName
                Write-LogSuccess "  ✔  RENOMMÉ   $($file.Name)  →  $destName"
                $script:undoStack += ,@($destPath, $file.Name)
                $cRenamed++
            } catch {
                Write-LogError "  ✖  ERREUR    $($file.Name) : $_"
                $cError++
            }
        }
    }

    $sw.Stop()
    Write-LogInfo "──────────────────────────────────────────"
    $summary = "  Résultat : $cRenamed renommés  |  $cIgnored ignorés  |  $cConflict conflits  |  $cError erreurs  |  $([Math]::Round($sw.Elapsed.TotalSeconds,1))s"
    Write-LogInfo $summary
    Write-LogInfo "══════════════════════════════════════════"

    $lblStats.Text = "$cRenamed renommés · $cIgnored ignorés · $cConflict conflits · $cError erreurs"

    Set-UIRunning $false
    if (-not $isDryRun -and $script:undoStack.Count -gt 0) {
        $btnUndo.Enabled = $true
    }
    $btnExport.Enabled = $script:logLines.Count -gt 0
})

# ─────────────────────────────────────────────
#  ACTION : ANNULER
# ─────────────────────────────────────────────
$btnUndo.Add_Click({
    if ($script:undoStack.Count -eq 0) { return }
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Êtes-vous sûr de vouloir annuler les $($script:undoStack.Count) renommages ?",
        "Confirmer l'annulation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($result -ne "Yes") { return }

    Write-LogInfo ""
    Write-LogInfo "══ ANNULATION ══"
    $ok = 0; $err = 0

    # On annule dans l'ordre inverse pour limiter les collisions de noms
    for ($i = $script:undoStack.Count - 1; $i -ge 0; $i--) {
        $pair = $script:undoStack[$i]
        $currentPath  = $pair[0]
        $originalName = $pair[1]
        $folder       = [System.IO.Path]::GetDirectoryName($currentPath)
        try {
            if (Test-Path $currentPath) {
                Invoke-SafeRename -FullPath $currentPath -NewName $originalName
                Write-LogSuccess "  ↩  RESTAURÉ  $([System.IO.Path]::GetFileName($currentPath))  →  $originalName"
                $ok++
            } else {
                Write-LogWarn "  ⚠  INTROUVABLE  $([System.IO.Path]::GetFileName($currentPath)) (déjà déplacé/renommé)"
            }
        } catch {
            Write-LogError "  ✖  ÉCHEC     $([System.IO.Path]::GetFileName($currentPath)) : $_"
            $err++
        }
    }

    Write-LogInfo "  $ok restaurés · $err erreurs"
    $script:undoStack = @()
    $btnUndo.Enabled   = $false
})

# ─────────────────────────────────────────────
#  ACTION : EXPORTER LOG
# ─────────────────────────────────────────────
$btnExport.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "Fichier texte (*.txt)|*.txt"
    $sfd.FileName = "rename-log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($sfd.ShowDialog() -eq "OK") {
        try {
            $script:logLines | Out-File -FilePath $sfd.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Log exporté avec succès.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Échec de l'export : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# ─────────────────────────────────────────────
#  ACTION : VIDER LOG
# ─────────────────────────────────────────────
$btnClear.Add_Click({
    $rtb.Clear()
    $script:logLines  = @()
    $pbr.Value         = 0
    $lblStats.Text     = "Prêt."
    $btnExport.Enabled = $false
})

# ─────────────────────────────────────────────
#  LANCEMENT
# ─────────────────────────────────────────────
[void]$form.ShowDialog()