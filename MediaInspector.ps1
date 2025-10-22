Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
chcp 65001 > $null
$ErrorActionPreference = "Stop"

# PowerShell ウィンドウを非表示
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) | Out-Null  # 0 = SW_HIDE (完全に非表示)

# コマンドライン引数を取得
$droppedFiles = $args

# 設定ファイルのパス
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configDir = Join-Path $scriptDir "Profile"
$configFile = Join-Path $configDir "MediaInspector.ini"
$historyFile = Join-Path $configDir "history.txt"

# 設定フォルダがなければ作成
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
}

# --- 設定を読み込む関数 ---
function Load-Config {
    $config = @{
        Theme = "Dark"
        FontName = "Consolas"
        FontSize = 10
        WindowOpacity = 1.0
        YtDlpPath = ""
        MediaInfoPath = ""
        IncludeSubfolders = $false
        IncludeAudioFiles = $false
        MaxHistoryCount = 20
        RememberWindowPosition = $false
        WindowX = ""
        WindowY = ""
        NewShortcut = "Ctrl+N"
        OpenFileShortcut = "Ctrl+O"
        AnalyzeShortcut = "Ctrl+R"
        ShowWindowShortcut = "Ctrl+W"
        CloseAllWindowsShortcut = "Ctrl+Q"
        OptionsShortcut = "Ctrl+P"
        SearchShortcut = "Ctrl+F"
        FindNextShortcut = "F3"
        FindPreviousShortcut = "Shift+F3"
        ClearHighlightShortcut = "Alt+F3"
        ShowYtDlpTitle = $true
        ShowYtDlpUploader = $true
        ShowYtDlpUploadDate = $true
        ShowYtDlpDuration = $true
        ShowYtDlpChapters = $true
        ShowYtDlpSubtitles = $true
        ShowYtDlpFormats = $true
        ShowDuration = $true
        ShowBitrate = $true
        ShowArtist = $true
        ShowChapters = $true
        ShowVideoCodec = $true
        ShowResolution = $true
        ShowFPS = $true
        ShowColorSpace = $true
        ShowChromaSubsampling = $true
        ShowBitDepth = $true
        ShowScanType = $true
        ShowColorRange = $true
        ShowHDR = $true
        ShowVideoBitrate = $true
        ShowVideoLanguage = $true
        ShowVideoStreamSize = $true
        ShowVideoWritingLibrary = $true
        ShowVideoEncodingSettings = $true
        ShowAudioCodec = $true
        ShowSampleRate = $true
        ShowAudioBitrate = $true
        ShowAudioLanguage = $true
        ShowAudioStreamSize = $true
        ShowAudioWritingLibrary = $true
        ShowCoverImage = $true
        ShowTextStream = $true
        ShowComment = $true
        ShowReplayGain = $true
        SearchMatchCase = $false
        SearchWrapAround = $false
    }
    
    if (Test-Path $configFile) {
        Get-Content $configFile | ForEach-Object {
            if ($_ -match '^(.+?)=(.+)$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $config[$key] = $value
            }
        }
    }
    
    return $config
}

# --- 設定を保存する関数 ---
function Save-Config {
    $content = @"
Theme=$($script:currentTheme)
FontName=$($script:currentFontName)
FontSize=$($script:currentFontSize)
WindowOpacity=$($script:windowOpacity)
YtDlpPath=$($script:ytDlpPath)
MediaInfoPath=$($script:mediaInfoPath)
IncludeSubfolders=$($script:includeSubfolders)
IncludeAudioFiles=$($script:includeAudioFiles)
MaxHistoryCount=$($script:maxHistoryCount)
RememberWindowPosition=$($script:rememberWindowPosition)
WindowX=$($script:windowX)
WindowY=$($script:windowY)
NewShortcut=$($script:newShortcut)
OpenFileShortcut=$($script:openFileShortcut)
AnalyzeShortcut=$($script:analyzeShortcut)
ShowWindowShortcut=$($script:showWindowShortcut)
CloseAllWindowsShortcut=$($script:closeAllWindowsShortcut)
OptionsShortcut=$($script:optionsShortcut)
SearchShortcut=$($script:searchShortcut)
FindNextShortcut=$($script:findNextShortcut)
FindPreviousShortcut=$($script:findPreviousShortcut)
ClearHighlightShortcut=$($script:clearHighlightShortcut)
ShowYtDlpTitle=$($script:showYtDlpTitle)
ShowYtDlpUploader=$($script:showYtDlpUploader)
ShowYtDlpUploadDate=$($script:showYtDlpUploadDate)
ShowYtDlpDuration=$($script:showYtDlpDuration)
ShowYtDlpChapters=$($script:showYtDlpChapters)
ShowYtDlpSubtitles=$($script:showYtDlpSubtitles)
ShowYtDlpFormats=$($script:showYtDlpFormats)
ShowDuration=$($script:showDuration)
ShowBitrate=$($script:showBitrate)
ShowArtist=$($script:showArtist)
ShowChapters=$($script:showChapters)
ShowVideoCodec=$($script:showVideoCodec)
ShowResolution=$($script:showResolution)
ShowFPS=$($script:showFPS)
ShowColorSpace=$($script:showColorSpace)
ShowChromaSubsampling=$($script:showChromaSubsampling)
ShowBitDepth=$($script:showBitDepth)
ShowScanType=$($script:showScanType)
ShowColorRange=$($script:showColorRange)
ShowHDR=$($script:showHDR)
ShowVideoBitrate=$($script:showVideoBitrate)
ShowVideoLanguage=$($script:showVideoLanguage)
ShowVideoStreamSize=$($script:showVideoStreamSize)
ShowVideoWritingLibrary=$($script:showVideoWritingLibrary)
ShowVideoEncodingSettings=$($script:showVideoEncodingSettings)
ShowAudioCodec=$($script:showAudioCodec)
ShowSampleRate=$($script:showSampleRate)
ShowAudioBitrate=$($script:showAudioBitrate)
ShowAudioLanguage=$($script:showAudioLanguage)
ShowAudioStreamSize=$($script:showAudioStreamSize)
ShowAudioWritingLibrary=$($script:showAudioWritingLibrary)
ShowCoverImage=$($script:showCoverImage)
ShowTextStream=$($script:showTextStream)
ShowComment=$($script:showComment)
ShowReplayGain=$($script:showReplayGain)
SearchMatchCase=$($script:searchMatchCase)
SearchWrapAround=$($script:searchWrapAround)
"@
    Set-Content -Path $configFile -Value $content -Encoding UTF8
}

# --- 履歴を読み込む関数 ---
function Load-History {
    if (Test-Path $historyFile) {
        $history = Get-Content -Path $historyFile -Encoding UTF8 | Where-Object { $_.Trim() }
        return $history
    }
    return @()
}

# --- 履歴を保存する関数 ---
function Save-History($items) {
    # 最新N件まで保存
    $uniqueItems = $items | Select-Object -Unique | Select-Object -First $script:maxHistoryCount
    Set-Content -Path $historyFile -Value $uniqueItems -Encoding UTF8
}

# --- 設定を読み込み ---
$config = Load-Config
$script:currentTheme = $config.Theme
$script:currentFontName = $config.FontName
$script:currentFontSize = [int]$config.FontSize
$script:windowOpacity = [double]$config.WindowOpacity
$script:ytDlpPath = $config.YtDlpPath
$script:mediaInfoPath = $config.MediaInfoPath
$script:includeSubfolders = [bool]::Parse($config.IncludeSubfolders)
$script:includeAudioFiles = [bool]::Parse($config.IncludeAudioFiles)
$script:maxHistoryCount = [int]$config.MaxHistoryCount
$script:rememberWindowPosition = [bool]::Parse($config.RememberWindowPosition)
$script:windowX = $config.WindowX
$script:windowY = $config.WindowY
$script:newShortcut = $config.NewShortcut
$script:openFileShortcut = $config.OpenFileShortcut
$script:analyzeShortcut = $config.AnalyzeShortcut
$script:showWindowShortcut = $config.ShowWindowShortcut
$script:closeAllWindowsShortcut = $config.CloseAllWindowsShortcut
$script:optionsShortcut = $config.OptionsShortcut
$script:searchShortcut = $config.SearchShortcut
$script:findNextShortcut = $config.FindNextShortcut
$script:findPreviousShortcut = $config.FindPreviousShortcut
$script:clearHighlightShortcut = $config.ClearHighlightShortcut
$script:showYtDlpTitle = [bool]::Parse($config.ShowYtDlpTitle)
$script:showYtDlpUploader = [bool]::Parse($config.ShowYtDlpUploader)
$script:showYtDlpUploadDate = [bool]::Parse($config.ShowYtDlpUploadDate)
$script:showYtDlpDuration = [bool]::Parse($config.ShowYtDlpDuration)
$script:showYtDlpChapters = [bool]::Parse($config.ShowYtDlpChapters)
$script:showYtDlpSubtitles = [bool]::Parse($config.ShowYtDlpSubtitles)
$script:showYtDlpFormats = [bool]::Parse($config.ShowYtDlpFormats)
$script:showDuration = [bool]::Parse($config.ShowDuration)
$script:showBitrate = [bool]::Parse($config.ShowBitrate)
$script:showArtist = [bool]::Parse($config.ShowArtist)
$script:showChapters = [bool]::Parse($config.ShowChapters)
$script:showVideoCodec = [bool]::Parse($config.ShowVideoCodec)
$script:showResolution = [bool]::Parse($config.ShowResolution)
$script:showFPS = [bool]::Parse($config.ShowFPS)
$script:showColorSpace = [bool]::Parse($config.ShowColorSpace)
$script:showChromaSubsampling = [bool]::Parse($config.ShowChromaSubsampling)
$script:showBitDepth = [bool]::Parse($config.ShowBitDepth)
$script:showScanType = [bool]::Parse($config.ShowScanType)
$script:showColorRange = [bool]::Parse($config.ShowColorRange)
$script:showHDR = [bool]::Parse($config.ShowHDR)
$script:showVideoBitrate = [bool]::Parse($config.ShowVideoBitrate)
$script:showVideoLanguage = [bool]::Parse($config.ShowVideoLanguage)
$script:showVideoStreamSize = [bool]::Parse($config.ShowVideoStreamSize)
$script:showVideoWritingLibrary = [bool]::Parse($config.ShowVideoWritingLibrary)
$script:showVideoEncodingSettings = [bool]::Parse($config.ShowVideoEncodingSettings)
$script:showAudioCodec = [bool]::Parse($config.ShowAudioCodec)
$script:showSampleRate = [bool]::Parse($config.ShowSampleRate)
$script:showAudioBitrate = [bool]::Parse($config.ShowAudioBitrate)
$script:showAudioLanguage = [bool]::Parse($config.ShowAudioLanguage)
$script:showAudioStreamSize = [bool]::Parse($config.ShowAudioStreamSize)
$script:showAudioWritingLibrary = [bool]::Parse($config.ShowAudioWritingLibrary)
$script:showCoverImage = [bool]::Parse($config.ShowCoverImage)
$script:showTextStream = [bool]::Parse($config.ShowTextStream)
$script:showComment = [bool]::Parse($config.ShowComment)
$script:showReplayGain = [bool]::Parse($config.ShowReplayGain)
$script:searchMatchCase = [bool]::Parse($config.SearchMatchCase)
$script:searchWrapAround = [bool]::Parse($config.SearchWrapAround)

# --- ツールパスチェック ---
foreach ($tool in @($script:ytDlpPath, $script:mediaInfoPath)) {
    if ($tool -and -not (Test-Path $tool)) {
        [System.Windows.Forms.MessageBox]::Show("$tool が見つかりません。設定でパスを確認してください。")
        # ここでは終了せずに続行(設定で修正可能なため)
    }
}

function Format-Time($seconds) {
    if (-not $seconds) { return "不明" }
    $t = [timespan]::FromSeconds([double]$seconds)
    $parts = @()
    if ($t.Hours -gt 0) { $parts += "$($t.Hours)時間" }
    if ($t.Minutes -gt 0) { $parts += "$($t.Minutes)分" }
    $parts += "$($t.Seconds)秒"
    return ($parts -join "")
}

function Format-FileSize($bytes) {
    if (-not $bytes) { return "不明" }
    $units = @("B", "KiB", "MiB", "GiB", "TiB")
    $size = [double]$bytes
    $unitIndex = 0
    while ($size -ge 1024 -and $unitIndex -lt $units.Length - 1) {
        $size /= 1024
        $unitIndex++
    }
    return "{0:N2} {1}" -f $size, $units[$unitIndex]
}

function Format-Bitrate($bitrate) {
    if (-not $bitrate) { return "不明" }
    if ($bitrate -match '([\d\s,]+)\s*kb/s') {
        $value = $matches[1] -replace '\s+', '' -replace ',', ''
        $numericValue = [double]$value
        return "{0:N0} kb/s" -f $numericValue
    }
    return $bitrate
}

function Convert-DurationToJapanese($duration) {
    if (-not $duration) { return "不明" }
    
    if ($duration -match '(\d+)\s*h\s*(\d+)\s*min') {
        $hours = $matches[1]
        $minutes = $matches[2]
        return "${hours}時間${minutes}分"
    }
    elseif ($duration -match '(\d+)\s*min\s*(\d+)\s*s') {
        $minutes = $matches[1]
        $seconds = $matches[2]
        return "${minutes}分${seconds}秒"
    }
    elseif ($duration -match '(\d+)\s*h') {
        $hours = $matches[1]
        return "${hours}時間"
    }
    
    return $duration
}

function Get-HDRInfo($colorPrimaries, $transferCharacteristics, $matrixCoefficients) {
    $hdrType = ""
    $colorSpace = ""
    
    if ($transferCharacteristics -match "PQ|SMPTE ST 2084") {
        $hdrType = "HDR10"
    } elseif ($transferCharacteristics -match "HLG") {
        $hdrType = "HLG"
    } else {
        $hdrType = "SDR"
    }
    
    if ($colorPrimaries -match "BT.2020") {
        $colorSpace = "BT.2020"
    } elseif ($colorPrimaries -match "BT.709") {
        $colorSpace = "BT.709"
    } elseif ($matrixCoefficients -match "BT.2020") {
        $colorSpace = "BT.2020"
    } elseif ($matrixCoefficients -match "BT.709") {
        $colorSpace = "BT.709"
    } else {
        $colorSpace = "不明"
    }
    
    return "[$hdrType] $colorSpace"
}

# テーマを適用する関数
function Apply-Theme {
    if ($script:currentTheme -eq "Dark") {
        $script:bgColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
        $script:fgColor = [System.Drawing.Color]::WhiteSmoke
        $script:inputBgColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
        $script:outputBgColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
        $script:menuBgColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
    } else {
        $script:bgColor = [System.Drawing.Color]::White
        $script:fgColor = [System.Drawing.Color]::Black
        $script:inputBgColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
        $script:outputBgColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
        $script:menuBgColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    }
    
    $form.BackColor = $script:bgColor
    $form.ForeColor = $script:fgColor
    $textBox.BackColor = $script:inputBgColor
    $textBox.ForeColor = $script:fgColor
    $outputBox.BackColor = $script:outputBgColor
    $outputBox.ForeColor = $script:fgColor
    $menuStrip.BackColor = $script:menuBgColor
    $menuStrip.ForeColor = $script:fgColor
    $label.ForeColor = $script:fgColor
    $progressLabel.ForeColor = $script:fgColor
    
    if ($script:currentTheme -eq "Dark") {
        $menuHoverColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
    } else {
        $menuHoverColor = [System.Drawing.Color]::FromArgb(200, 220, 240)
    }
    $colorTable = New-Object CustomColorTable($menuHoverColor, $menuHoverColor, $menuHoverColor)
    $menuStrip.Renderer = New-Object System.Windows.Forms.ToolStripProfessionalRenderer($colorTable)
}

# 初期テーマを適用
$script:bgColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
$script:fgColor = [System.Drawing.Color]::WhiteSmoke
$script:inputBgColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$script:outputBgColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$script:menuBgColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$script:accent = [System.Drawing.Color]::FromArgb(70, 130, 180)

if ($script:currentTheme -eq "Light") {
    $script:bgColor = [System.Drawing.Color]::White
    $script:fgColor = [System.Drawing.Color]::Black
    $script:inputBgColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $script:outputBgColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $script:menuBgColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
}

# バージョン情報
$script:version = "1.0"

# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "MediaInspector v$script:version"
$form.Size = New-Object System.Drawing.Size(850, 600)

if ($script:rememberWindowPosition -and $script:windowX -and $script:windowY) {
    $form.StartPosition = "Manual"
    $form.Location = New-Object System.Drawing.Point([int]$script:windowX, [int]$script:windowY)
} else {
    $form.StartPosition = "CenterScreen"
}

$form.BackColor = $script:bgColor
$form.ForeColor = $script:fgColor
$form.Font = New-Object System.Drawing.Font("Meiryo UI", 9)
$form.Opacity = $script:windowOpacity
$form.KeyPreview = $true

# カスタム ColorTable クラスを定義
Add-Type -TypeDefinition @'
using System;
using System.Drawing;
using System.Windows.Forms;

public class CustomColorTable : ProfessionalColorTable
{
    private Color menuItemSelectedColor;
    private Color menuItemPressedColor;
    private Color menuItemBorderColor;
    
    public CustomColorTable(Color selectedColor, Color pressedColor, Color borderColor)
    {
        this.menuItemSelectedColor = selectedColor;
        this.menuItemPressedColor = pressedColor;
        this.menuItemBorderColor = borderColor;
    }
    
    public override Color MenuItemSelected
    {
        get { return menuItemSelectedColor; }
    }
    
    public override Color MenuItemSelectedGradientBegin
    {
        get { return menuItemSelectedColor; }
    }
    
    public override Color MenuItemSelectedGradientEnd
    {
        get { return menuItemSelectedColor; }
    }
    
    public override Color MenuItemPressedGradientBegin
    {
        get { return menuItemPressedColor; }
    }
    
    public override Color MenuItemPressedGradientMiddle
    {
        get { return menuItemPressedColor; }
    }
    
    public override Color MenuItemPressedGradientEnd
    {
        get { return menuItemPressedColor; }
    }
    
    public override Color MenuItemBorder
    {
        get { return menuItemBorderColor; }
    }
}
'@ -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.Drawing.Primitives

# メニューストリップを作成
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = $script:menuBgColor
$menuStrip.ForeColor = $script:fgColor

if ($script:currentTheme -eq "Dark") {
    $menuHoverColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
} else {
    $menuHoverColor = [System.Drawing.Color]::FromArgb(200, 220, 240)
}
$colorTable = New-Object CustomColorTable($menuHoverColor, $menuHoverColor, $menuHoverColor)
$menuStrip.Renderer = New-Object System.Windows.Forms.ToolStripProfessionalRenderer($colorTable)

# 「ファイル」メニュー
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "ファイル(&F)"

# 新規
$newItem = New-Object System.Windows.Forms.ToolStripMenuItem
$newItem.Text = "新規(&N)"
$newItem.Add_Click({
    if ($textBox.Text.Trim() -or $outputBox.Text.Trim()) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "入力内容と出力結果をクリアしますか？",
            "確認",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $textBox.Clear()
            $outputBox.Clear()
            $script:analysisResults = @()
            $showWindowButton.Enabled = $false
            $progress.Value = 0
            $progressLabel.Text = "0%"
            $textBox.Focus()
        }
    } else {
        $textBox.Clear()
        $outputBox.Clear()
        $script:analysisResults = @()
        $showWindowButton.Enabled = $false
        $progress.Value = 0
        $progressLabel.Text = "0%"
        $textBox.Focus()
    }
})
$fileMenu.DropDownItems.Add($newItem)

# セパレーター
$fileMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

# ファイルを追加
$addFileItem = New-Object System.Windows.Forms.ToolStripMenuItem
$addFileItem.Text = "ファイルを追加(&F)..."
$addFileItem.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "動画ファイル (*.mp4;*.mkv;*.avi;*.mov;*.wmv;*.flv;*.webm;*.m4v;*.ts;*.m2ts)|*.mp4;*.mkv;*.avi;*.mov;*.wmv;*.flv;*.webm;*.m4v;*.ts;*.m2ts|すべてのファイル (*.*)|*.*"
    $openFileDialog.Multiselect = $true
    $openFileDialog.Title = "ファイルを選択"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $existingText = $textBox.Text.Trim()
        $quotedFiles = $openFileDialog.FileNames | ForEach-Object { "`"$_`"" }
        $newFiles = $quotedFiles -join "`r`n"
        
        if ($existingText) {
            $textBox.Text = $existingText + "`r`n" + $newFiles
        } else {
            $textBox.Text = $newFiles
        }
        
        $textBox.SelectionStart = $textBox.Text.Length
        $textBox.ScrollToCaret()
    }
})
$fileMenu.DropDownItems.Add($addFileItem)

# フォルダを追加
$addFolderItem = New-Object System.Windows.Forms.ToolStripMenuItem
$addFolderItem.Text = "フォルダを追加(&D)..."
$addFolderItem.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "フォルダを選択"
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $existingText = $textBox.Text.Trim()
        $selectedFolder = "`"$($folderBrowser.SelectedPath)`""
        
        if ($existingText) {
            $textBox.Text = $existingText + "`r`n" + $selectedFolder
        } else {
            $textBox.Text = $selectedFolder
        }
        
        $textBox.SelectionStart = $textBox.Text.Length
        $textBox.ScrollToCaret()
    }
})
$fileMenu.DropDownItems.Add($addFolderItem)

# セパレーター
$fileMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

# 終了
$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "終了(&X)"
$exitItem.Add_Click({
    $form.Close()
})
$fileMenu.DropDownItems.Add($exitItem)

$menuStrip.Items.Add($fileMenu)

# 「検索」メニュー
$searchMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$searchMenu.Text = "検索(&S)"

# 検索
$searchItem = New-Object System.Windows.Forms.ToolStripMenuItem
$searchItem.Text = "検索(&F)...`tCtrl+F"
$searchItem.Add_Click({
    Show-SearchDialog
})
$searchMenu.DropDownItems.Add($searchItem)

# 次を検索
$findNextItem = New-Object System.Windows.Forms.ToolStripMenuItem
$findNextItem.Text = "次を検索(&N)`tF3"
$findNextItem.Add_Click({
    Find-Next
})
$searchMenu.DropDownItems.Add($findNextItem)

# 前を検索
$findPreviousItem = New-Object System.Windows.Forms.ToolStripMenuItem
$findPreviousItem.Text = "前を検索(&P)`tShift+F3"
$findPreviousItem.Add_Click({
    Find-Previous
})
$searchMenu.DropDownItems.Add($findPreviousItem)

# ハイライト解除
$clearHighlightItem = New-Object System.Windows.Forms.ToolStripMenuItem
$clearHighlightItem.Text = "ハイライト解除(&C)`tAlt+F3"
$clearHighlightItem.Add_Click({
    Clear-SearchHighlight
})
$searchMenu.DropDownItems.Add($clearHighlightItem)

$menuStrip.Items.Add($searchMenu)

# 「ツール」メニュー
$toolMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$toolMenu.Text = "ツール(&T)"

# 動画ファイルを整理
$organizeItem = New-Object System.Windows.Forms.ToolStripMenuItem
$organizeItem.Text = "動画ファイルを整理(&M)..."
$organizeItem.ToolTipText = "指定したフォルダにある動画ファイルのタグ情報からフォルダを作成し、整理します"
$organizeItem.Add_Click({
    Show-FileOrganizer
})
$toolMenu.DropDownItems.Add($organizeItem)

# 解析結果をリスト表示
$filterItem = New-Object System.Windows.Forms.ToolStripMenuItem
$filterItem.Text = "解析結果をリスト表示(&F)..."
$filterItem.Add_Click({
    Show-AllResultsList
})
$toolMenu.DropDownItems.Add($filterItem)

# セパレーター
$toolMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

# オプション
$optionsItem = New-Object System.Windows.Forms.ToolStripMenuItem
$optionsItem.Text = "オプション(&O)..."
$optionsItem.Add_Click({
    # オプションダイアログを作成
    $optionsForm = New-Object System.Windows.Forms.Form
    $optionsForm.Text = "オプション"
    $optionsForm.Size = New-Object System.Drawing.Size(500, 720)
    $optionsForm.StartPosition = "CenterParent"
    $optionsForm.FormBorderStyle = "FixedDialog"
    $optionsForm.MaximizeBox = $false
    $optionsForm.MinimizeBox = $false
    $optionsForm.BackColor = $script:bgColor
    $optionsForm.ForeColor = $script:fgColor
    $optionsForm.AutoScroll = $false
    
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(464, 610)
    $tabControl.Anchor = "Top,Bottom,Left,Right"
    $optionsForm.Controls.Add($tabControl)
    
    # 全般タブ
    $generalTab = New-Object System.Windows.Forms.TabPage
    $generalTab.Text = "全般"
    $generalTab.BackColor = $script:bgColor
    $generalTab.AutoScroll = $true
    $tabControl.TabPages.Add($generalTab)
    
    $analysisTab = New-Object System.Windows.Forms.TabPage
    $analysisTab.Text = "解析"
    $analysisTab.BackColor = $script:bgColor
    $analysisTab.AutoScroll = $true
    $tabControl.TabPages.Add($analysisTab)
    
    $keyboardTab = New-Object System.Windows.Forms.TabPage
    $keyboardTab.Text = "キーボード"
    $keyboardTab.BackColor = $script:bgColor
    $keyboardTab.AutoScroll = $true
    $tabControl.TabPages.Add($keyboardTab)
    
    # テーマ設定
    $themeLabel = New-Object System.Windows.Forms.Label
    $themeLabel.Text = "テーマ:"
    $themeLabel.Location = New-Object System.Drawing.Point(20, 20)
    $themeLabel.Size = New-Object System.Drawing.Size(100, 20)
    $themeLabel.ForeColor = $script:fgColor
    $generalTab.Controls.Add($themeLabel)
    
    $themeCombo = New-Object System.Windows.Forms.ComboBox
    $themeCombo.Location = New-Object System.Drawing.Point(130, 18)
    $themeCombo.Size = New-Object System.Drawing.Size(280, 25)
    $themeCombo.DropDownStyle = "DropDownList"
    $themeCombo.Items.AddRange(@("ダークテーマ", "ライトテーマ"))
    $themeCombo.SelectedIndex = if ($script:currentTheme -eq "Dark") { 0 } else { 1 }
    $generalTab.Controls.Add($themeCombo)
    
    # フォント名設定
    $fontNameLabel = New-Object System.Windows.Forms.Label
    $fontNameLabel.Text = "フォント名:"
    $fontNameLabel.Location = New-Object System.Drawing.Point(20, 60)
    $fontNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    $fontNameLabel.ForeColor = $script:fgColor
    $generalTab.Controls.Add($fontNameLabel)

    # インストール済みフォントを取得
    $installedFonts = New-Object System.Drawing.Text.InstalledFontCollection
    $fontFamilies = $installedFonts.Families | Where-Object { $_.IsStyleAvailable([System.Drawing.FontStyle]::Regular) } | Sort-Object Name

    $fontNameCombo = New-Object System.Windows.Forms.ComboBox
    $fontNameCombo.Location = New-Object System.Drawing.Point(130, 58)
    $fontNameCombo.Size = New-Object System.Drawing.Size(280, 25)
    $fontNameCombo.DropDownStyle = "DropDown"
    $fontNameCombo.AutoCompleteSource = "ListItems"
    $fontNameCombo.AutoCompleteMode = "SuggestAppend"

    # フォントを追加（既存のフォントも含める）
    $fontList = New-Object System.Collections.ArrayList
    foreach ($fontFamily in $fontFamilies) {
        [void]$fontList.Add($fontFamily.Name)
    }

    # 既存のフォントがリストにない場合は追加
    $defaultFonts = @("Consolas", "Courier New", "MS Gothic", "Meiryo", "Yu Gothic")
    foreach ($font in $defaultFonts) {
        if ($fontList -notcontains $font) {
            [void]$fontList.Add($font)
        }
    }

    # アルファベット順にソート
    $fontList.Sort()

    # コンボボックスにフォントを追加
    foreach ($font in $fontList) {
        [void]$fontNameCombo.Items.Add($font)
    }

    # 現在のフォントを選択
    if ($fontNameCombo.Items.Contains($script:currentFontName)) {
        $fontNameCombo.SelectedItem = $script:currentFontName
    } else {
        # 現在のフォントが見つからない場合はConsolasを選択
        $fontNameCombo.SelectedItem = "Consolas"
    }

    $generalTab.Controls.Add($fontNameCombo)
    
    # フォントサイズ設定
    $fontSizeLabel = New-Object System.Windows.Forms.Label
    $fontSizeLabel.Text = "フォントサイズ:"
    $fontSizeLabel.Location = New-Object System.Drawing.Point(20, 100)
    $fontSizeLabel.Size = New-Object System.Drawing.Size(100, 20)
    $fontSizeLabel.ForeColor = $script:fgColor
    $generalTab.Controls.Add($fontSizeLabel)
    
    $fontSizeCombo = New-Object System.Windows.Forms.ComboBox
    $fontSizeCombo.Location = New-Object System.Drawing.Point(130, 98)
    $fontSizeCombo.Size = New-Object System.Drawing.Size(280, 25)
    $fontSizeCombo.DropDownStyle = "DropDownList"
    $fontSizeCombo.Items.AddRange(@("8", "9", "10", "11", "12", "14", "16"))
    $fontSizeCombo.SelectedItem = $script:currentFontSize.ToString()
    $generalTab.Controls.Add($fontSizeCombo)
    
    # yt-dlpパス設定
    $ytDlpLabel = New-Object System.Windows.Forms.Label
    $ytDlpLabel.Text = "yt-dlp パス:"
    $ytDlpLabel.Location = New-Object System.Drawing.Point(20, 140)
    $ytDlpLabel.Size = New-Object System.Drawing.Size(100, 20)
    $ytDlpLabel.ForeColor = $script:fgColor
    $generalTab.Controls.Add($ytDlpLabel)
    
    $ytDlpTextBox = New-Object System.Windows.Forms.TextBox
    $ytDlpTextBox.Location = New-Object System.Drawing.Point(130, 138)
    $ytDlpTextBox.Size = New-Object System.Drawing.Size(245, 25)
    $ytDlpTextBox.Text = $script:ytDlpPath
    $ytDlpTextBox.BackColor = $script:inputBgColor
    $ytDlpTextBox.ForeColor = $script:fgColor
    $generalTab.Controls.Add($ytDlpTextBox)
    
    $ytDlpBrowseButton = New-Object System.Windows.Forms.Button
    $ytDlpBrowseButton.Text = "参照"
    $ytDlpBrowseButton.Location = New-Object System.Drawing.Point(380, 138)
    $ytDlpBrowseButton.Size = New-Object System.Drawing.Size(50, 25)
    $ytDlpBrowseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*"
        $openFileDialog.Title = "yt-dlp のパスを選択"
        if ($script:ytDlpPath -and (Test-Path $script:ytDlpPath)) {
            $openFileDialog.InitialDirectory = Split-Path -Parent $script:ytDlpPath
        }
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $ytDlpTextBox.Text = $openFileDialog.FileName
        }
    })
    $generalTab.Controls.Add($ytDlpBrowseButton)
    
    # MediaInfoパス設定
    $mediaInfoLabel = New-Object System.Windows.Forms.Label
    $mediaInfoLabel.Text = "MediaInfo パス:"
    $mediaInfoLabel.Location = New-Object System.Drawing.Point(20, 180)
    $mediaInfoLabel.Size = New-Object System.Drawing.Size(100, 20)
    $mediaInfoLabel.ForeColor = $script:fgColor
    $generalTab.Controls.Add($mediaInfoLabel)
    
    $mediaInfoTextBox = New-Object System.Windows.Forms.TextBox
    $mediaInfoTextBox.Location = New-Object System.Drawing.Point(130, 178)
    $mediaInfoTextBox.Size = New-Object System.Drawing.Size(245, 25)
    $mediaInfoTextBox.Text = $script:mediaInfoPath
    $mediaInfoTextBox.BackColor = $script:inputBgColor
    $mediaInfoTextBox.ForeColor = $script:fgColor
    $generalTab.Controls.Add($mediaInfoTextBox)
    
    $mediaInfoBrowseButton = New-Object System.Windows.Forms.Button
    $mediaInfoBrowseButton.Text = "参照"
    $mediaInfoBrowseButton.Location = New-Object System.Drawing.Point(380, 178)
    $mediaInfoBrowseButton.Size = New-Object System.Drawing.Size(50, 25)
    $mediaInfoBrowseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*"
        $openFileDialog.Title = "MediaInfo のパスを選択"
        if ($script:mediaInfoPath -and (Test-Path $script:mediaInfoPath)) {
            $openFileDialog.InitialDirectory = Split-Path -Parent $script:mediaInfoPath
        }
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $mediaInfoTextBox.Text = $openFileDialog.FileName
        }
    })
    $generalTab.Controls.Add($mediaInfoBrowseButton)
    
    # 透明度設定
    $opacityLabel = New-Object System.Windows.Forms.Label
    $opacityLabel.Text = "ウィンドウの透明度:"
    $opacityLabel.Location = New-Object System.Drawing.Point(20, 220)
    $opacityLabel.Size = New-Object System.Drawing.Size(150, 20)
    $opacityLabel.ForeColor = $script:fgColor
    $generalTab.Controls.Add($opacityLabel)
    
    $opacityTrackBar = New-Object System.Windows.Forms.TrackBar
    $opacityTrackBar.Location = New-Object System.Drawing.Point(20, 245)
    $opacityTrackBar.Size = New-Object System.Drawing.Size(300, 45)
    $opacityTrackBar.Minimum = 50
    $opacityTrackBar.Maximum = 100
    $opacityTrackBar.TickFrequency = 10
    $opacityTrackBar.Value = [int]($script:windowOpacity * 100)
    $generalTab.Controls.Add($opacityTrackBar)
    
    $opacityValueLabel = New-Object System.Windows.Forms.Label
    $opacityValueLabel.Location = New-Object System.Drawing.Point(330, 250)
    $opacityValueLabel.Size = New-Object System.Drawing.Size(80, 20)
    $opacityValueLabel.Text = "$([int]($script:windowOpacity * 100))%"
    $opacityValueLabel.ForeColor = $script:fgColor
    $generalTab.Controls.Add($opacityValueLabel)
    
    $opacityTrackBar.Add_ValueChanged({
        $opacityValueLabel.Text = "$($opacityTrackBar.Value)%"
        $form.Opacity = $opacityTrackBar.Value / 100.0
    })
    
    # ウィンドウの位置とサイズを記憶する
    $rememberWindowPositionCheckBox = New-Object System.Windows.Forms.CheckBox
    $rememberWindowPositionCheckBox.Text = "ウィンドウの位置を記憶する"
    $rememberWindowPositionCheckBox.Location = New-Object System.Drawing.Point(20, 300)
    $rememberWindowPositionCheckBox.Size = New-Object System.Drawing.Size(400, 25)
    $rememberWindowPositionCheckBox.Checked = $script:rememberWindowPosition
    $rememberWindowPositionCheckBox.ForeColor = $script:fgColor
    $generalTab.Controls.Add($rememberWindowPositionCheckBox)
    
    # サブフォルダを含むオプション
    $includeSubfoldersCheckBox = New-Object System.Windows.Forms.CheckBox
    $includeSubfoldersCheckBox.Text = "フォルダ解析時にサブフォルダを含める"
    $includeSubfoldersCheckBox.Location = New-Object System.Drawing.Point(20, 20)
    $includeSubfoldersCheckBox.Size = New-Object System.Drawing.Size(400, 25)
    $includeSubfoldersCheckBox.Checked = $script:includeSubfolders
    $includeSubfoldersCheckBox.ForeColor = $script:fgColor
    $analysisTab.Controls.Add($includeSubfoldersCheckBox)
    
    $includeAudioFilesCheckBox = New-Object System.Windows.Forms.CheckBox
    $includeAudioFilesCheckBox.Text = "フォルダ解析時に音声ファイルを含める"
    $includeAudioFilesCheckBox.Location = New-Object System.Drawing.Point(20, 50)
    $includeAudioFilesCheckBox.Size = New-Object System.Drawing.Size(400, 25)
    $includeAudioFilesCheckBox.Checked = $script:includeAudioFiles
    $includeAudioFilesCheckBox.ForeColor = $script:fgColor
    $analysisTab.Controls.Add($includeAudioFilesCheckBox)
    
    $maxHistoryLabel = New-Object System.Windows.Forms.Label
    $maxHistoryLabel.Text = "履歴の最大保存数:"
    $maxHistoryLabel.Location = New-Object System.Drawing.Point(20, 85)
    $maxHistoryLabel.Size = New-Object System.Drawing.Size(150, 20)
    $maxHistoryLabel.ForeColor = $script:fgColor
    $analysisTab.Controls.Add($maxHistoryLabel)
    
    $maxHistoryNumeric = New-Object System.Windows.Forms.NumericUpDown
    $maxHistoryNumeric.Location = New-Object System.Drawing.Point(180, 83)
    $maxHistoryNumeric.Size = New-Object System.Drawing.Size(80, 25)
    $maxHistoryNumeric.Minimum = 1
    $maxHistoryNumeric.Maximum = 100
    $maxHistoryNumeric.Value = $script:maxHistoryCount
    $maxHistoryNumeric.BackColor = $script:inputBgColor
    $maxHistoryNumeric.ForeColor = $script:fgColor
    $analysisTab.Controls.Add($maxHistoryNumeric)
    
    $ytDlpItemsGroupBox = New-Object System.Windows.Forms.GroupBox
    $ytDlpItemsGroupBox.Text = "yt-dlp 解析表示"
    $ytDlpItemsGroupBox.Location = New-Object System.Drawing.Point(20, 120)
    $ytDlpItemsGroupBox.Size = New-Object System.Drawing.Size(410, 150)
    $ytDlpItemsGroupBox.ForeColor = $script:fgColor
    $analysisTab.Controls.Add($ytDlpItemsGroupBox)
    
    $ytDlpItemsPanel = New-Object System.Windows.Forms.Panel
    $ytDlpItemsPanel.Location = New-Object System.Drawing.Point(10, 25)
    $ytDlpItemsPanel.Size = New-Object System.Drawing.Size(390, 115)
    $ytDlpItemsPanel.AutoScroll = $false
    $ytDlpItemsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $ytDlpItemsGroupBox.Controls.Add($ytDlpItemsPanel)
    
    $ytDlpCheckBoxes = @{}
    
    $ytDlpItems = @(
        @{Key="ShowYtDlpTitle"; Label="タイトル"; X=5; Y=5},
        @{Key="ShowYtDlpUploader"; Label="アップローダー"; X=230; Y=5},
        @{Key="ShowYtDlpUploadDate"; Label="投稿日時"; X=5; Y=30},
        @{Key="ShowYtDlpDuration"; Label="再生時間"; X=230; Y=30},
        @{Key="ShowYtDlpChapters"; Label="チャプター"; X=5; Y=55},
        @{Key="ShowYtDlpSubtitles"; Label="字幕情報"; X=230; Y=55},
        @{Key="ShowYtDlpFormats"; Label="フォーマット一覧"; X=5; Y=80}
    )
    
    foreach ($item in $ytDlpItems) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = $item.Label
        $checkBox.Location = New-Object System.Drawing.Point($item.X, $item.Y)
        $checkBox.Size = New-Object System.Drawing.Size(210, 25)
        $varName = $item.Key.Substring(0,1).ToLower() + $item.Key.Substring(1)
        $checkBox.Checked = (Get-Variable -Name $varName -Scope Script).Value
        $checkBox.ForeColor = $script:fgColor
        $ytDlpItemsPanel.Controls.Add($checkBox)
        $ytDlpCheckBoxes[$item.Key] = $checkBox
    }
    
    $analysisItemsGroupBox = New-Object System.Windows.Forms.GroupBox
    $analysisItemsGroupBox.Text = "MediaInfo 解析表示"
    $analysisItemsGroupBox.Location = New-Object System.Drawing.Point(20, 280)
    $analysisItemsGroupBox.Size = New-Object System.Drawing.Size(410, 470)
    $analysisItemsGroupBox.ForeColor = $script:fgColor
    $analysisTab.Controls.Add($analysisItemsGroupBox)
    
    $analysisItemsPanel = New-Object System.Windows.Forms.Panel
    $analysisItemsPanel.Location = New-Object System.Drawing.Point(10, 25)
    $analysisItemsPanel.Size = New-Object System.Drawing.Size(390, 435)
    $analysisItemsPanel.AutoScroll = $false
    $analysisItemsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $analysisItemsGroupBox.Controls.Add($analysisItemsPanel)
    
    $analysisCheckBoxes = @{}
    
    $analysisItems = @(
        @{Key="ShowDuration"; Label="再生時間"; X=5; Y=5},
        @{Key="ShowBitrate"; Label="ビットレート (全体)"; X=230; Y=5},
        @{Key="ShowArtist"; Label="作成者"; X=5; Y=30},
        @{Key="ShowComment"; Label="コメント"; X=230; Y=30},
        @{Key="ShowChapters"; Label="チャプター"; X=5; Y=55},
        @{Key="Separator1"; Label=""; Y=80},
        @{Key="ShowVideoCodec"; Label="映像: コーデック"; X=5; Y=95},
        @{Key="ShowVideoBitrate"; Label="映像: ビットレート"; X=230; Y=95},
        @{Key="ShowResolution"; Label="映像: 解像度"; X=5; Y=120},
        @{Key="ShowFPS"; Label="映像: フレームレート"; X=230; Y=120},
        @{Key="ShowColorSpace"; Label="映像: 色空間"; X=5; Y=145},
        @{Key="ShowChromaSubsampling"; Label="映像: クロマサブサンプリング"; X=230; Y=145},
        @{Key="ShowBitDepth"; Label="映像: ビット深度"; X=5; Y=170},
        @{Key="ShowScanType"; Label="映像: スキャンタイプ"; X=230; Y=170},
        @{Key="ShowColorRange"; Label="映像: 色範囲"; X=5; Y=195},
        @{Key="ShowHDR"; Label="映像: HDR/SDR"; X=230; Y=195},
        @{Key="ShowVideoLanguage"; Label="映像: 言語"; X=5; Y=220},
        @{Key="ShowVideoWritingLibrary"; Label="映像: ライブラリ"; X=230; Y=220},
        @{Key="ShowVideoStreamSize"; Label="映像: ストリームサイズ"; X=5; Y=245},
        @{Key="ShowVideoEncodingSettings"; Label="映像: エンコードの設定"; X=230; Y=245},
        @{Key="Separator2"; Label=""; Y=270},
        @{Key="ShowAudioCodec"; Label="音声: コーデック"; X=5; Y=285},
        @{Key="ShowAudioBitrate"; Label="音声: ビットレート"; X=230; Y=285},
        @{Key="ShowSampleRate"; Label="音声: サンプリングレート"; X=5; Y=310},
        @{Key="ShowReplayGain"; Label="音声: リプレイゲイン"; X=230; Y=310},
        @{Key="ShowAudioLanguage"; Label="音声: 言語"; X=5; Y=335},
        @{Key="ShowAudioWritingLibrary"; Label="音声: ライブラリ"; X=230; Y=335},
        @{Key="ShowAudioStreamSize"; Label="音声: ストリームサイズ"; X=5; Y=360},
        @{Key="Separator3"; Label=""; Y=385},
        @{Key="ShowTextStream"; Label="テキストストリーム (字幕)"; X=5; Y=400},
        @{Key="ShowCoverImage"; Label="カバー画像"; X=230; Y=400}
    )
    
    foreach ($item in $analysisItems) {
        if ($item.Key -match '^Separator\d+$') {
            $separator = New-Object System.Windows.Forms.Label
            $separator.Location = New-Object System.Drawing.Point(5, $item.Y)
            $separator.Size = New-Object System.Drawing.Size(410, 2)
            $separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
            $analysisItemsPanel.Controls.Add($separator)
        } else {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = $item.Label
            $checkBox.Location = New-Object System.Drawing.Point($item.X, $item.Y)
            $checkBox.Size = New-Object System.Drawing.Size(210, 25)
            $varName = $item.Key.Substring(0,1).ToLower() + $item.Key.Substring(1)
            $checkBox.Checked = (Get-Variable -Name $varName -Scope Script).Value
            $checkBox.ForeColor = $script:fgColor
            $analysisItemsPanel.Controls.Add($checkBox)
            $analysisCheckBoxes[$item.Key] = $checkBox
        }
    }
    
    $newShortcutLabel = New-Object System.Windows.Forms.Label
    $newShortcutLabel.Text = "新規作成:"
    $newShortcutLabel.Location = New-Object System.Drawing.Point(20, 20)
    $newShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $newShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($newShortcutLabel)
    
    $newShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $newShortcutTextBox.Location = New-Object System.Drawing.Point(180, 18)
    $newShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $newShortcutTextBox.Text = $script:newShortcut
    $newShortcutTextBox.BackColor = $script:inputBgColor
    $newShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($newShortcutTextBox)
    
    $openFileShortcutLabel = New-Object System.Windows.Forms.Label
    $openFileShortcutLabel.Text = "ファイルを追加:"
    $openFileShortcutLabel.Location = New-Object System.Drawing.Point(20, 60)
    $openFileShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $openFileShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($openFileShortcutLabel)
    
    $openFileShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $openFileShortcutTextBox.Location = New-Object System.Drawing.Point(180, 58)
    $openFileShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $openFileShortcutTextBox.Text = $script:openFileShortcut
    $openFileShortcutTextBox.BackColor = $script:inputBgColor
    $openFileShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($openFileShortcutTextBox)
    
    $analyzeShortcutLabel = New-Object System.Windows.Forms.Label
    $analyzeShortcutLabel.Text = "解析開始:"
    $analyzeShortcutLabel.Location = New-Object System.Drawing.Point(20, 100)
    $analyzeShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $analyzeShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($analyzeShortcutLabel)
    
    $analyzeShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $analyzeShortcutTextBox.Location = New-Object System.Drawing.Point(180, 98)
    $analyzeShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $analyzeShortcutTextBox.Text = $script:analyzeShortcut
    $analyzeShortcutTextBox.BackColor = $script:inputBgColor
    $analyzeShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($analyzeShortcutTextBox)
    
    $showWindowShortcutLabel = New-Object System.Windows.Forms.Label
    $showWindowShortcutLabel.Text = "結果を別ウィンドウ表示:"
    $showWindowShortcutLabel.Location = New-Object System.Drawing.Point(20, 140)
    $showWindowShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $showWindowShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($showWindowShortcutLabel)
    
    $showWindowShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $showWindowShortcutTextBox.Location = New-Object System.Drawing.Point(180, 138)
    $showWindowShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $showWindowShortcutTextBox.Text = $script:showWindowShortcut
    $showWindowShortcutTextBox.BackColor = $script:inputBgColor
    $showWindowShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($showWindowShortcutTextBox)
    
    $closeAllWindowsShortcutLabel = New-Object System.Windows.Forms.Label
    $closeAllWindowsShortcutLabel.Text = "全ウィンドウを閉じる:"
    $closeAllWindowsShortcutLabel.Location = New-Object System.Drawing.Point(20, 180)
    $closeAllWindowsShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $closeAllWindowsShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($closeAllWindowsShortcutLabel)
    
    $closeAllWindowsShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $closeAllWindowsShortcutTextBox.Location = New-Object System.Drawing.Point(180, 178)
    $closeAllWindowsShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $closeAllWindowsShortcutTextBox.Text = $script:closeAllWindowsShortcut
    $closeAllWindowsShortcutTextBox.BackColor = $script:inputBgColor
    $closeAllWindowsShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($closeAllWindowsShortcutTextBox)
    
    $optionsShortcutLabel = New-Object System.Windows.Forms.Label
    $optionsShortcutLabel.Text = "オプション:"
    $optionsShortcutLabel.Location = New-Object System.Drawing.Point(20, 220)
    $optionsShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $optionsShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($optionsShortcutLabel)
    
    $optionsShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $optionsShortcutTextBox.Location = New-Object System.Drawing.Point(180, 218)
    $optionsShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $optionsShortcutTextBox.Text = $script:optionsShortcut
    $optionsShortcutTextBox.BackColor = $script:inputBgColor
    $optionsShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($optionsShortcutTextBox)
    
    $searchShortcutLabel = New-Object System.Windows.Forms.Label
    $searchShortcutLabel.Text = "検索:"
    $searchShortcutLabel.Location = New-Object System.Drawing.Point(20, 260)
    $searchShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $searchShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($searchShortcutLabel)
    
    $searchShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $searchShortcutTextBox.Location = New-Object System.Drawing.Point(180, 258)
    $searchShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $searchShortcutTextBox.Text = $script:searchShortcut
    $searchShortcutTextBox.BackColor = $script:inputBgColor
    $searchShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($searchShortcutTextBox)
    
    $findNextShortcutLabel = New-Object System.Windows.Forms.Label
    $findNextShortcutLabel.Text = "次を検索:"
    $findNextShortcutLabel.Location = New-Object System.Drawing.Point(20, 300)
    $findNextShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $findNextShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($findNextShortcutLabel)
    
    $findNextShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $findNextShortcutTextBox.Location = New-Object System.Drawing.Point(180, 298)
    $findNextShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $findNextShortcutTextBox.Text = $script:findNextShortcut
    $findNextShortcutTextBox.BackColor = $script:inputBgColor
    $findNextShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($findNextShortcutTextBox)
    
    $findPreviousShortcutLabel = New-Object System.Windows.Forms.Label
    $findPreviousShortcutLabel.Text = "前を検索:"
    $findPreviousShortcutLabel.Location = New-Object System.Drawing.Point(20, 340)
    $findPreviousShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $findPreviousShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($findPreviousShortcutLabel)
    
    $findPreviousShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $findPreviousShortcutTextBox.Location = New-Object System.Drawing.Point(180, 338)
    $findPreviousShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $findPreviousShortcutTextBox.Text = $script:findPreviousShortcut
    $findPreviousShortcutTextBox.BackColor = $script:inputBgColor
    $findPreviousShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($findPreviousShortcutTextBox)
    
    $clearHighlightShortcutLabel = New-Object System.Windows.Forms.Label
    $clearHighlightShortcutLabel.Text = "ハイライト解除:"
    $clearHighlightShortcutLabel.Location = New-Object System.Drawing.Point(20, 380)
    $clearHighlightShortcutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $clearHighlightShortcutLabel.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($clearHighlightShortcutLabel)
    
    $clearHighlightShortcutTextBox = New-Object System.Windows.Forms.TextBox
    $clearHighlightShortcutTextBox.Location = New-Object System.Drawing.Point(180, 378)
    $clearHighlightShortcutTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $clearHighlightShortcutTextBox.Text = $script:clearHighlightShortcut
    $clearHighlightShortcutTextBox.BackColor = $script:inputBgColor
    $clearHighlightShortcutTextBox.ForeColor = $script:fgColor
    $keyboardTab.Controls.Add($clearHighlightShortcutTextBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(250, 635)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Anchor = "Bottom,Right"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Add_Click({
        # テーマ適用
        $newTheme = if ($themeCombo.SelectedIndex -eq 0) { "Dark" } else { "Light" }
        if ($newTheme -ne $script:currentTheme) {
            $script:currentTheme = $newTheme
            Apply-Theme
        }
        
        # フォント適用
        $newFontName = $fontNameCombo.SelectedItem
        $newFontSize = [int]$fontSizeCombo.SelectedItem
        
        if ($newFontName -ne $script:currentFontName -or $newFontSize -ne $script:currentFontSize) {
            $script:currentFontName = $newFontName
            $script:currentFontSize = $newFontSize
            $outputBox.Font = New-Object System.Drawing.Font($script:currentFontName, $script:currentFontSize)
        }
        
        # ツールパス適用
        $newYtDlpPath = $ytDlpTextBox.Text.Trim()
        $newMediaInfoPath = $mediaInfoTextBox.Text.Trim()
        
        if ($newYtDlpPath) { $script:ytDlpPath = $newYtDlpPath }
        if ($newMediaInfoPath) { $script:mediaInfoPath = $newMediaInfoPath }
        
        # 透明度適用
        $script:windowOpacity = $opacityTrackBar.Value / 100.0
        $form.Opacity = $script:windowOpacity
        
        # サブフォルダオプション適用
        $script:includeSubfolders = $includeSubfoldersCheckBox.Checked
        
        # フォルダの音声ファイル適用
        $script:includeAudioFiles = $includeAudioFilesCheckBox.Checked
        
        # 履歴の最大保存数適用
        $script:maxHistoryCount = [int]$maxHistoryNumeric.Value
        
        # ウィンドウの位置とサイズ記憶適用
        $script:rememberWindowPosition = $rememberWindowPositionCheckBox.Checked
        
        $script:newShortcut = $newShortcutTextBox.Text.Trim()
        $script:openFileShortcut = $openFileShortcutTextBox.Text.Trim()
        $script:analyzeShortcut = $analyzeShortcutTextBox.Text.Trim()
        $script:showWindowShortcut = $showWindowShortcutTextBox.Text.Trim()
        $script:closeAllWindowsShortcut = $closeAllWindowsShortcutTextBox.Text.Trim()
        $script:optionsShortcut = $optionsShortcutTextBox.Text.Trim()
        $script:searchShortcut = $searchShortcutTextBox.Text.Trim()
        $script:findNextShortcut = $findNextShortcutTextBox.Text.Trim()
        $script:findPreviousShortcut = $findPreviousShortcutTextBox.Text.Trim()
        $script:clearHighlightShortcut = $clearHighlightShortcutTextBox.Text.Trim()
        
        foreach ($key in $ytDlpCheckBoxes.Keys) {
            $varName = $key.Substring(0,1).ToLower() + $key.Substring(1)
            Set-Variable -Name $varName -Value $ytDlpCheckBoxes[$key].Checked -Scope Script
        }
        
        foreach ($key in $analysisCheckBoxes.Keys) {
            $varName = $key.Substring(0,1).ToLower() + $key.Substring(1)
            Set-Variable -Name $varName -Value $analysisCheckBoxes[$key].Checked -Scope Script
        }
        
        Save-Config
        
        # ツールの存在確認
        foreach ($tool in @($script:ytDlpPath, $script:mediaInfoPath)) {
            if ($tool -and -not (Test-Path $tool)) {
                [System.Windows.Forms.MessageBox]::Show("$tool が見つかりません。パスを確認してください。")
            }
        }
    })
    $optionsForm.Controls.Add($okButton)
    
    # キャンセルボタン
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "キャンセル"
    $cancelButton.Location = New-Object System.Drawing.Point(340, 635)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Anchor = "Bottom,Right"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Add_Click({
        # 透明度を元に戻す
        $form.Opacity = $script:windowOpacity
    })
    $optionsForm.Controls.Add($cancelButton)
    
    $optionsForm.AcceptButton = $okButton
    $optionsForm.CancelButton = $cancelButton
    
    [void]$optionsForm.ShowDialog($form)
})

$toolMenu.DropDownItems.Add($optionsItem)
$menuStrip.Items.Add($toolMenu)

# 「ヘルプ」メニュー
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "ヘルプ(&H)"

# yt-dlp について
$ytDlpInfoItem = New-Object System.Windows.Forms.ToolStripMenuItem
$ytDlpInfoItem.Text = "yt-dlp について(&Y)..."
$ytDlpInfoItem.Add_Click({
    Show-YtDlpInfo
})
$helpMenu.DropDownItems.Add($ytDlpInfoItem)

# MediaInfo CLI について
$mediaInfoInfoItem = New-Object System.Windows.Forms.ToolStripMenuItem
$mediaInfoInfoItem.Text = "MediaInfo CLI について(&M)..."
$mediaInfoInfoItem.Add_Click({
    Show-MediaInfoInfo
})
$helpMenu.DropDownItems.Add($mediaInfoInfoItem)

# セパレーター
$helpMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

# MediaInspector について
$aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutItem.Text = "MediaInspector について(&A)..."
$aboutItem.Add_Click({
    Show-AboutMediaInspector
})
$helpMenu.DropDownItems.Add($aboutItem)

$menuStrip.Items.Add($helpMenu)

$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

$label = New-Object System.Windows.Forms.Label
$label.Text = "URL または ローカルファイル (複数入力する場合、改行またはスペースで区切る / ドラッグ && ドロップ 対応)"
$label.Location = New-Object System.Drawing.Point(10, 35)
$label.AutoSize = $true
$label.Anchor = "Top,Left"
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 60)
$textBox.Size = New-Object System.Drawing.Size(810, 80)
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$textBox.BackColor = $script:inputBgColor
$textBox.ForeColor = $script:fgColor
$textBox.AllowDrop = $true
$textBox.Anchor = "Top,Left,Right"
$form.Controls.Add($textBox)

# コマンドライン引数からドロップされたファイルを挿入
if ($droppedFiles -and $droppedFiles.Count -gt 0) {
    $quotedFiles = $droppedFiles | ForEach-Object { "`"$_`"" }
    $textBox.Text = $quotedFiles -join "`r`n"
}

$textBox.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [Windows.Forms.DragDropEffects]::Copy
    } else {
        $e.Effect = [Windows.Forms.DragDropEffects]::None
    }
})

$textBox.Add_DragDrop({
    param($sender, $e)
    $files = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    if ($files) {
        $existingText = $textBox.Text.Trim()
        $quotedFiles = $files | ForEach-Object { "`"$_`"" }
        $newFiles = $quotedFiles -join "`r`n"
        
        if ($existingText) {
            $textBox.Text = $existingText + "`r`n" + $newFiles
        } else {
            $textBox.Text = $newFiles
        }
        
        $textBox.SelectionStart = $textBox.Text.Length
        $textBox.ScrollToCaret()
    }
})

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "消去"
$clearButton.Location = New-Object System.Drawing.Point(690, 30)
$clearButton.Size = New-Object System.Drawing.Size(60, 25)
$clearButton.BackColor = [System.Drawing.Color]::FromArgb(200, 60, 60)
$clearButton.ForeColor = $script:fgColor
$clearButton.Anchor = "Top,Right"
$clearButton.Add_Click({
    $textBox.Clear()
    $textBox.Focus()
})
$form.Controls.Add($clearButton)

# 履歴ボタンを追加
$historyButton = New-Object System.Windows.Forms.Button
$historyButton.Text = "履歴"
$historyButton.Location = New-Object System.Drawing.Point(760, 30)
$historyButton.Size = New-Object System.Drawing.Size(60, 25)
$historyButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$historyButton.ForeColor = $script:fgColor
$historyButton.Anchor = "Top,Right"
$historyButton.Add_Click({
    Show-HistoryDialog
})
$form.Controls.Add($historyButton)

$button = New-Object System.Windows.Forms.Button
$button.Text = "解析開始"
$button.Location = New-Object System.Drawing.Point(10, 150)
$button.Size = New-Object System.Drawing.Size(120, 30)
$button.BackColor = $script:accent
$button.ForeColor = $script:fgColor
$button.Anchor = "Top,Left"
$form.Controls.Add($button)

$showWindowButton = New-Object System.Windows.Forms.Button
$showWindowButton.Text = "結果を別ウィンドウ表示"
$showWindowButton.Location = New-Object System.Drawing.Point(140, 150)
$showWindowButton.Size = New-Object System.Drawing.Size(180, 30)
$showWindowButton.BackColor = [System.Drawing.Color]::FromArgb(90, 150, 90)
$showWindowButton.ForeColor = $script:fgColor
$showWindowButton.Enabled = $false
$showWindowButton.Anchor = "Top,Left"
$form.Controls.Add($showWindowButton)

$closeAllWindowsButton = New-Object System.Windows.Forms.Button
$closeAllWindowsButton.Text = "全ウィンドウを閉じる"
$closeAllWindowsButton.Location = New-Object System.Drawing.Point(330, 150)
$closeAllWindowsButton.Size = New-Object System.Drawing.Size(150, 30)
$closeAllWindowsButton.BackColor = [System.Drawing.Color]::FromArgb(180, 60, 60)
$closeAllWindowsButton.ForeColor = $script:fgColor
$closeAllWindowsButton.Enabled = $false
$closeAllWindowsButton.Anchor = "Top,Left"
$form.Controls.Add($closeAllWindowsButton)

$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = "結果をコピー"
$copyButton.Location = New-Object System.Drawing.Point(490, 150)
$copyButton.Size = New-Object System.Drawing.Size(120, 30)
$copyButton.BackColor = [System.Drawing.Color]::FromArgb(100, 120, 140)
$copyButton.ForeColor = $script:fgColor
$copyButton.Anchor = "Top,Left"
$copyButton.Add_Click({ Copy-OutputToClipboard })
$form.Controls.Add($copyButton)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(620, 150)
$progress.Size = New-Object System.Drawing.Size(120, 30)
$progress.Style = 'Continuous'
$progress.Anchor = "Top,Left,Right"
$form.Controls.Add($progress)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(750, 150)
$progressLabel.Size = New-Object System.Drawing.Size(70, 30)
$progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$progressLabel.Font = New-Object System.Drawing.Font("Meiryo UI", 10, [System.Drawing.FontStyle]::Bold)
$progressLabel.ForeColor = $script:fgColor
$progressLabel.Text = "0%"
$progressLabel.Anchor = "Top,Right"
$form.Controls.Add($progressLabel)

$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font($script:currentFontName, $script:currentFontSize)
$outputBox.Location = New-Object System.Drawing.Point(10, 190)
$outputBox.Size = New-Object System.Drawing.Size(810, 360)
$outputBox.BackColor = $script:outputBgColor
$outputBox.ForeColor = $script:fgColor
$outputBox.Anchor = "Top,Bottom,Left,Right"
$outputBox.HideSelection = $false
$form.Controls.Add($outputBox)

# 解析結果を保存するグローバル変数
$script:analysisResults = @()
$script:resultForms = @()

function Write-OutputBox($msg) {
    $outputBox.AppendText($msg + "`r`n")
    $outputBox.ScrollToCaret()
}

function Set-Progress($v) {
    if ($v -gt 100) { $v = 100 }
    $progress.Value = $v
    $progressLabel.Text = "$v%"
    $form.Refresh()
}

# 別ウィンドウで結果を表示する関数
function Show-ResultWindows {
    if ($script:analysisResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("表示する解析結果がありません。")
        return
    }
    
    # 既存のウィンドウを閉じる
    Close-AllResultWindows
    
    # 固定サイズを設定
    $windowWidth = 300
    $windowHeight = 400
    $spacing = 10
    
    # 作業領域（タスクバーを除いた領域）のサイズを取得
    $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $screenWidth = $workingArea.Width
    $screenHeight = $workingArea.Height
    
    # 横方向に配置できるウィンドウ数を計算（最小マージンを20pxとする）
    $minMargin = 20
    $columnsPerRow = [math]::Floor(($screenWidth - $minMargin * 2 + $spacing) / ($windowWidth + $spacing))
    if ($columnsPerRow -lt 1) { $columnsPerRow = 1 }
    
    # 縦方向に配置できるウィンドウ数を計算
    $rowsPerScreen = [math]::Floor(($screenHeight - $minMargin * 2 + $spacing) / ($windowHeight + $spacing))
    if ($rowsPerScreen -lt 1) { $rowsPerScreen = 1 }
    
    # 実際に必要な幅と高さを計算
    $totalWidth = $columnsPerRow * $windowWidth + ($columnsPerRow - 1) * $spacing
    $totalHeight = $rowsPerScreen * $windowHeight + ($rowsPerScreen - 1) * $spacing
    
    # 中央配置のための開始位置を計算
    $startX = [math]::Max($minMargin, [math]::Floor(($screenWidth - $totalWidth) / 2))
    $startY = [math]::Max($minMargin, [math]::Floor(($screenHeight - $totalHeight) / 2))
    
    # 画面内に収まる最大ウィンドウ数
    $maxWindowsPerScreen = $columnsPerRow * $rowsPerScreen
    
    # 表示するウィンドウ数を決定
    $totalCount = $script:analysisResults.Count
    $displayCount = [math]::Min($totalCount, $maxWindowsPerScreen)
    
    # 表示数が制限される場合は確認メッセージ
    if ($totalCount -gt $maxWindowsPerScreen) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "解析結果が ${totalCount} 件ありますが、画面サイズの制限により最大 ${maxWindowsPerScreen} 件まで表示できます。`n`n最初の ${maxWindowsPerScreen} 件を表示しますか？",
            "確認",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }
    
    # 表示位置の初期化
    $xOffset = $startX
    $yOffset = $startY
    $currentColumn = 0
    
    # ウィンドウを表示
    for ($i = 0; $i -lt $displayCount; $i++) {
        $result = $script:analysisResults[$i]
        
        $resultForm = New-Object System.Windows.Forms.Form
        $resultForm.Text = "解析結果: $($result.Title)"
        $resultForm.Size = New-Object System.Drawing.Size($windowWidth, $windowHeight)
        $resultForm.StartPosition = "Manual"
        $resultForm.Location = New-Object System.Drawing.Point($xOffset, $yOffset)
        $resultForm.BackColor = $script:bgColor
        $resultForm.ForeColor = $script:fgColor
        $resultForm.Font = New-Object System.Drawing.Font("Meiryo UI", 9)
        
        $resultTextBox = New-Object System.Windows.Forms.TextBox
        $resultTextBox.Multiline = $true
        $resultTextBox.ScrollBars = "Vertical"
        $resultTextBox.ReadOnly = $true
        $resultTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
        $resultTextBox.Dock = "Fill"
        $resultTextBox.BackColor = $script:outputBgColor
        $resultTextBox.ForeColor = $script:fgColor
        $resultTextBox.Text = $result.Content
        $resultForm.Controls.Add($resultTextBox)
        
        # FormClosedイベントでリストから削除
        $resultForm.Add_FormClosed({
            param($sender, $e)
            $script:resultForms = $script:resultForms | Where-Object { $_ -ne $sender }
            if ($script:resultForms.Count -eq 0) {
                $closeAllWindowsButton.Enabled = $false
            }
        })
        
        $resultForm.Show()
        $script:resultForms += $resultForm
        
        # 次のウィンドウの位置を計算
        $currentColumn++
        
        if ($currentColumn -ge $columnsPerRow) {
            # 次の行へ
            $currentColumn = 0
            $xOffset = $startX
            $yOffset += $windowHeight + $spacing
        } else {
            # 右へ移動
            $xOffset += $windowWidth + $spacing
        }
    }
    
    # ウィンドウが開かれたらボタンを有効化
    if ($script:resultForms.Count -gt 0) {
        $closeAllWindowsButton.Enabled = $true
    }
}

# 全ての結果ウィンドウを閉じる関数
function Close-AllResultWindows {
    foreach ($form in $script:resultForms) {
        if ($form -and -not $form.IsDisposed) {
            $form.Close()
        }
    }
    $script:resultForms = @()
    $closeAllWindowsButton.Enabled = $false
}

# --- 検索関連のグローバル変数 ---
$script:searchText = ""
$script:searchMatchCase = $false
$script:searchWrapAround = $false
$script:lastSearchIndex = -1

# --- 検索ダイアログを表示する関数 ---
function Show-SearchDialog {
    $searchForm = New-Object System.Windows.Forms.Form
    $searchForm.Text = "検索"
    $searchForm.Size = New-Object System.Drawing.Size(405, 180)
    $searchForm.StartPosition = "CenterParent"
    $searchForm.FormBorderStyle = "FixedDialog"
    $searchForm.MaximizeBox = $false
    $searchForm.MinimizeBox = $false
    $searchForm.BackColor = $script:bgColor
    $searchForm.ForeColor = $script:fgColor
    
    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "検索する文字列:(&F)"
    $searchLabel.Location = New-Object System.Drawing.Point(10, 15)
    $searchLabel.Size = New-Object System.Drawing.Size(110, 20)
    $searchLabel.ForeColor = $script:fgColor
    $searchForm.Controls.Add($searchLabel)
    
    $hitCountLabel = New-Object System.Windows.Forms.Label
    $hitCountLabel.Text = ""
    $hitCountLabel.Location = New-Object System.Drawing.Point(120, 15)
    $hitCountLabel.Size = New-Object System.Drawing.Size(170, 20)
    $hitCountLabel.ForeColor = $script:fgColor
    $hitCountLabel.AutoSize = $false
    $searchForm.Controls.Add($hitCountLabel)
    
    $searchTextBox = New-Object System.Windows.Forms.TextBox
    $searchTextBox.Location = New-Object System.Drawing.Point(10, 40)
    $searchTextBox.Size = New-Object System.Drawing.Size(280, 25)
    $searchTextBox.Text = $script:searchText
    $searchTextBox.BackColor = $script:inputBgColor
    $searchTextBox.ForeColor = $script:fgColor
    $searchForm.Controls.Add($searchTextBox)
    
    function Update-HitCount {
        $searchString = $searchTextBox.Text
        
        if ([string]::IsNullOrEmpty($searchString)) {
            $hitCountLabel.Text = ""
            return
        }
        
        $text = $outputBox.Text
        if ([string]::IsNullOrEmpty($text)) {
            $hitCountLabel.Text = "0個、ヒットしました。"
            return
        }
        
        $comparisonType = if ($matchCaseCheckBox.Checked) {
            [System.StringComparison]::Ordinal
        } else {
            [System.StringComparison]::OrdinalIgnoreCase
        }
        
        $count = 0
        $startIndex = 0
        
        while ($startIndex -lt $text.Length) {
            $foundIndex = $text.IndexOf($searchString, $startIndex, $comparisonType)
            if ($foundIndex -ge 0) {
                $count++
                $startIndex = $foundIndex + 1
            } else {
                break
            }
        }
        
        if ($count -gt 0) {
            $hitCountLabel.Text = "$($count)個、ヒットしました。"
        } else {
            $hitCountLabel.Text = "0個、ヒットしました。"
        }
    }
    
    $searchTextBox.Add_TextChanged({
        Update-HitCount
    })
    
    $matchCaseCheckBox = New-Object System.Windows.Forms.CheckBox
    $matchCaseCheckBox.Text = "大文字と小文字を区別(&C)"
    $matchCaseCheckBox.Location = New-Object System.Drawing.Point(10, 75)
    $matchCaseCheckBox.Size = New-Object System.Drawing.Size(200, 25)
    $matchCaseCheckBox.Checked = $script:searchMatchCase
    $matchCaseCheckBox.ForeColor = $script:fgColor
    $matchCaseCheckBox.Add_CheckedChanged({
        Update-HitCount
    })
    $searchForm.Controls.Add($matchCaseCheckBox)
    
    $wrapAroundCheckBox = New-Object System.Windows.Forms.CheckBox
    $wrapAroundCheckBox.Text = "折り返す(&D)"
    $wrapAroundCheckBox.Location = New-Object System.Drawing.Point(10, 100)
    $wrapAroundCheckBox.Size = New-Object System.Drawing.Size(200, 25)
    $wrapAroundCheckBox.Checked = $script:searchWrapAround
    $wrapAroundCheckBox.ForeColor = $script:fgColor
    $searchForm.Controls.Add($wrapAroundCheckBox)
    
    $searchForm.Add_FormClosed({
        # ダイアログを閉じた時点でチェック状態を script 変数に反映して一括保存
        $script:searchMatchCase = $matchCaseCheckBox.Checked
        $script:searchWrapAround = $wrapAroundCheckBox.Checked
        Save-Config
    })
    
    $findPreviousButton = New-Object System.Windows.Forms.Button
    $findPreviousButton.Text = "前を検索(P)"
    $findPreviousButton.Location = New-Object System.Drawing.Point(300, 15)
    $findPreviousButton.Size = New-Object System.Drawing.Size(80, 25)
    $findPreviousButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $findPreviousButton.ForeColor = $script:fgColor
    $findPreviousButton.Add_Click({
        $script:searchText = $searchTextBox.Text
        $script:searchMatchCase = $matchCaseCheckBox.Checked
        $script:searchWrapAround = $wrapAroundCheckBox.Checked
        $outputBox.Focus()
        Find-Previous
        $searchTextBox.Focus()
    })
    $searchForm.Controls.Add($findPreviousButton)
    
    $findNextButton = New-Object System.Windows.Forms.Button
    $findNextButton.Text = "次を検索(N)"
    $findNextButton.Location = New-Object System.Drawing.Point(300, 45)
    $findNextButton.Size = New-Object System.Drawing.Size(80, 25)
    $findNextButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $findNextButton.ForeColor = $script:fgColor
    $findNextButton.Add_Click({
        $script:searchText = $searchTextBox.Text
        $script:searchMatchCase = $matchCaseCheckBox.Checked
        $script:searchWrapAround = $wrapAroundCheckBox.Checked
        $outputBox.Focus()
        Find-Next
        $searchTextBox.Focus()
    })
    $searchForm.Controls.Add($findNextButton)
    
    $clearHighlightButton = New-Object System.Windows.Forms.Button
    $clearHighlightButton.Text = "解除"
    $clearHighlightButton.Location = New-Object System.Drawing.Point(300, 75)
    $clearHighlightButton.Size = New-Object System.Drawing.Size(80, 25)
    $clearHighlightButton.BackColor = [System.Drawing.Color]::FromArgb(150, 100, 50)
    $clearHighlightButton.ForeColor = $script:fgColor
    $clearHighlightButton.Add_Click({
        Clear-SearchHighlight
    })
    $searchForm.Controls.Add($clearHighlightButton)
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "閉じる"
    $closeButton.Location = New-Object System.Drawing.Point(300, 105)
    $closeButton.Size = New-Object System.Drawing.Size(80, 25)
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $closeButton.ForeColor = $script:fgColor
    $closeButton.Add_Click({
        $searchForm.Close()
    })
    $searchForm.Controls.Add($closeButton)
    
    $searchForm.AcceptButton = $findPreviousButton
    $searchForm.CancelButton = $closeButton
    
    Update-HitCount
    
    [void]$searchForm.ShowDialog($form)
}

# --- 次を検索する関数 ---
function Find-Next {
    if ([string]::IsNullOrEmpty($script:searchText)) {
        [System.Windows.Forms.MessageBox]::Show("検索する文字列を入力してください。", "情報")
        return
    }

    if ($script:prevSearchText -ne $script:searchText) {
        Clear-SearchHighlight
        $script:prevSearchText = $script:searchText
    }
    
    $text = $outputBox.Text
    $startIndex = if ($script:lastSearchIndex -ge 0) { $script:lastSearchIndex + $script:searchText.Length } else { 0 }
    
    $comparisonType = if ($script:searchMatchCase) {
        [System.StringComparison]::Ordinal
    } else {
        [System.StringComparison]::OrdinalIgnoreCase
    }
    
    $foundIndex = $text.IndexOf($script:searchText, $startIndex, $comparisonType)
    
    if ($foundIndex -eq -1 -and $script:searchWrapAround -and $startIndex -gt 0) {
        $foundIndex = $text.IndexOf($script:searchText, 0, $comparisonType)
    }
    
    if ($foundIndex -ge 0) {
        Highlight-SearchResult $foundIndex
        $script:lastSearchIndex = $foundIndex
    } else {
        [System.Windows.Forms.MessageBox]::Show("検索文字列が見つかりませんでした。", "情報")
        $script:lastSearchIndex = -1
    }
}

# --- 前を検索する関数 ---
function Find-Previous {
    if ([string]::IsNullOrEmpty($script:searchText)) {
        [System.Windows.Forms.MessageBox]::Show("検索する文字列を入力してください。", "情報")
        return
    }
    
    if ($script:prevSearchText -ne $script:searchText) {
        Clear-SearchHighlight
        $script:prevSearchText = $script:searchText
    }
    
    $text = $outputBox.Text
    $startIndex = if ($script:lastSearchIndex -gt 0) { $script:lastSearchIndex - 1 } else { $text.Length - 1 }
    
    $comparisonType = if ($script:searchMatchCase) {
        [System.StringComparison]::Ordinal
    } else {
        [System.StringComparison]::OrdinalIgnoreCase
    }
    
    $foundIndex = $text.LastIndexOf($script:searchText, $startIndex, $comparisonType)
    
    if ($foundIndex -eq -1 -and $script:searchWrapAround -and $startIndex -lt $text.Length - 1) {
        $foundIndex = $text.LastIndexOf($script:searchText, $text.Length - 1, $comparisonType)
    }
    
    if ($foundIndex -ge 0) {
        Highlight-SearchResult $foundIndex
        $script:lastSearchIndex = $foundIndex
    } else {
        [System.Windows.Forms.MessageBox]::Show("検索文字列が見つかりませんでした。", "情報")
        $script:lastSearchIndex = -1
    }
}

# --- 検索結果をハイライト表示する関数 ---
function Highlight-SearchResult($foundIndex) {
    $text = $outputBox.Text
    $searchLength = $script:searchText.Length

    $nativeDef = @'
using System;
using System.Runtime.InteropServices;
namespace Win32 {
  public static class NativeMethods {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
  }
}
'@
    Add-Type -TypeDefinition $nativeDef -ErrorAction SilentlyContinue

    $outputBox.SelectionStart = $foundIndex
    $outputBox.SelectionLength = $searchLength
    $outputBox.SelectionBackColor = if ($script:currentTheme -eq "Dark") {
        [System.Drawing.Color]::FromArgb(255, 165, 0)
    } else {
        [System.Drawing.Color]::FromArgb(255, 255, 0)
    }

    $lineIndex = $outputBox.GetLineFromCharIndex($foundIndex)

    $EM_GETFIRSTVISIBLELINE = 0x00CE
    $EM_LINESCROLL = 0x00B6

    $firstVisible = [Win32.NativeMethods]::SendMessage($outputBox.Handle, $EM_GETFIRSTVISIBLELINE, 0, 0)

    $fontHeight = [int][Math]::Ceiling($outputBox.Font.GetHeight())
    if ($fontHeight -le 0) { $fontHeight = 16 }

    $visibleLines = [int][Math]::Floor($outputBox.Height / $fontHeight)
    if ($visibleLines -lt 1) { $visibleLines = 1 }

    $desiredFirst = $lineIndex - [math]::Floor($visibleLines / 2)
    if ($desiredFirst -lt 0) { $desiredFirst = 0 }

    $delta = $desiredFirst - $firstVisible

    if ($delta -ne 0) {
        [Win32.NativeMethods]::SendMessage($outputBox.Handle, $EM_LINESCROLL, 0, $delta)
    } else {
        $outputBox.ScrollToCaret()
    }
}

# --- ハイライトを解除する関数 ---
function Clear-SearchHighlight {
    $outputBox.SelectionStart = 0
    $outputBox.SelectionLength = $outputBox.Text.Length
    $outputBox.SelectionBackColor = $outputBox.BackColor
    $outputBox.SelectionStart = 0
    $outputBox.SelectionLength = 0
    
    $script:lastSearchIndex = -1
}

# --- クリップボードにコピーする関数 ---
function Copy-OutputToClipboard {
    $text = $outputBox.Text
    
    if ([string]::IsNullOrWhiteSpace($text)) {
        [System.Windows.Forms.MessageBox]::Show("コピーする内容がありません。")
        return
    }
    
    # 不要なメッセージを削除
    $removePatterns = @(
        "解析開始\.\.\. 少々お待ちください。",
        "ローカルファイルとして解析します。",
        "フォルダ内のファイルを解析します。",
        "見つかったファイル数:.*",
        "解析完了。",
        "=== 全ファイル解析完了 ==="
    )
    
    $lines = $text -split "`r?`n"
    $filteredLines = $lines | Where-Object {
        $line = $_.Trim()
        $shouldKeep = $true
        foreach ($pattern in $removePatterns) {
            if ($line -match "^$pattern`$") {
                $shouldKeep = $false
                break
            }
        }
        $shouldKeep
    }
    
    $cleanedText = $filteredLines -join "`r`n"
    
    # 「--- 利用可能なコーデック一覧 ---」の後に「⚠ フォーマット情報を解析できませんでした。」しかない場合は両方削除
    $cleanedText = $cleanedText -replace "--- 利用可能なコーデック一覧 ---\s*`r?`n\s*⚠ フォーマット情報を解析できませんでした。\s*`r?`n", ""
    
    # 連続する空行を1つにまとめる
    $cleanedText = $cleanedText -replace "(`r?`n){3,}", "`r`n`r`n"
    
    try {
        [System.Windows.Forms.Clipboard]::SetText($cleanedText)
        
        # 自動的に閉じるダイアログを表示
        $notifyForm = New-Object System.Windows.Forms.Form
        $notifyForm.Text = "コピー完了"
        $notifyForm.Size = New-Object System.Drawing.Size(300, 140)
        $notifyForm.StartPosition = "CenterParent"
        $notifyForm.FormBorderStyle = "FixedDialog"
        $notifyForm.MaximizeBox = $false
        $notifyForm.MinimizeBox = $false
        $notifyForm.BackColor = $script:bgColor
        $notifyForm.ForeColor = $script:fgColor
        
        $notifyLabel = New-Object System.Windows.Forms.Label
        $notifyLabel.Text = "出力内容をクリップボードに`r`nコピーしました。"
        $notifyLabel.Location = New-Object System.Drawing.Point(20, 15)
        $notifyLabel.Size = New-Object System.Drawing.Size(260, 50)
        $notifyLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $notifyLabel.ForeColor = $script:fgColor
        $notifyForm.Controls.Add($notifyLabel)
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(110, 70)
        $okButton.Size = New-Object System.Drawing.Size(80, 25)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $notifyForm.Controls.Add($okButton)
        $notifyForm.AcceptButton = $okButton
        
        # 3秒後に自動的に閉じるタイマー
        $autoCloseTimer = New-Object System.Windows.Forms.Timer
        $autoCloseTimer.Interval = 3000
        $autoCloseTimer.Add_Tick({
            param($sender, $e)
            if ($notifyForm -and -not $notifyForm.IsDisposed) {
                $notifyForm.Close()
            }
            $sender.Stop()
            $sender.Dispose()
        })
        
        # フォームが閉じられたときにタイマーを停止
        $notifyForm.Add_FormClosed({
            if ($autoCloseTimer) {
                $autoCloseTimer.Stop()
                $autoCloseTimer.Dispose()
            }
        })
        
        $autoCloseTimer.Start()
        
        [void]$notifyForm.ShowDialog($form)
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("クリップボードへのコピーに失敗しました。`n$($_.Exception.Message)")
    }
}

# --- 履歴ダイアログを表示する関数 ---
function Show-HistoryDialog {
    $history = Load-History
    
    if ($history.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("履歴がありません。")
        return
    }
    
    $historyForm = New-Object System.Windows.Forms.Form
    $historyForm.Text = "入力履歴"
    $historyForm.Size = New-Object System.Drawing.Size(700, 500)
    $historyForm.StartPosition = "CenterParent"
    $historyForm.BackColor = $script:bgColor
    $historyForm.ForeColor = $script:fgColor
    
    $historyLabel = New-Object System.Windows.Forms.Label
    $historyLabel.Text = "履歴から選択してください（ダブルクリックで追加）："
    $historyLabel.Location = New-Object System.Drawing.Point(10, 10)
    $historyLabel.Size = New-Object System.Drawing.Size(680, 20)
    $historyLabel.ForeColor = $script:fgColor
    $historyForm.Controls.Add($historyLabel)
    
    $historyListBox = New-Object System.Windows.Forms.ListBox
    $historyListBox.Location = New-Object System.Drawing.Point(10, 35)
    $historyListBox.Size = New-Object System.Drawing.Size(660, 350)
    $historyListBox.BackColor = $script:inputBgColor
    $historyListBox.ForeColor = $script:fgColor
    $historyListBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $historyListBox.SelectionMode = "MultiExtended"
    
    foreach ($item in $history) {
        [void]$historyListBox.Items.Add($item)
    }
    
    $historyForm.Controls.Add($historyListBox)
    
    # ダブルクリックで追加
    $historyListBox.Add_DoubleClick({
        if ($historyListBox.SelectedItems.Count -gt 0) {
            $selectedItems = $historyListBox.SelectedItems
            $existingText = $textBox.Text.Trim()
            # ダブルクォートで囲む処理を追加
            $quotedItems = $selectedItems | ForEach-Object { "`"$_`"" }
            $newItems = $quotedItems -join "`r`n"
            
            if ($existingText) {
                $textBox.Text = $existingText + "`r`n" + $newItems
            } else {
                $textBox.Text = $newItems
            }
            
            $historyForm.Close()
        }
    })
    
    # 追加ボタン
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "追加"
    $addButton.Location = New-Object System.Drawing.Point(210, 395)
    $addButton.Size = New-Object System.Drawing.Size(80, 30)
    $addButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $addButton.ForeColor = $script:fgColor
    $addButton.Add_Click({
        if ($historyListBox.SelectedItems.Count -gt 0) {
            $selectedItems = $historyListBox.SelectedItems
            $existingText = $textBox.Text.Trim()
            # ダブルクォートで囲む処理を追加
            $quotedItems = $selectedItems | ForEach-Object { "`"$_`"" }
            $newItems = $quotedItems -join "`r`n"
            
            if ($existingText) {
                $textBox.Text = $existingText + "`r`n" + $newItems
            } else {
                $textBox.Text = $newItems
            }
        }
    })
    $historyForm.Controls.Add($addButton)
    
    # 選択項目を削除ボタン
    $deleteSelectedButton = New-Object System.Windows.Forms.Button
    $deleteSelectedButton.Text = "選択項目を削除"
    $deleteSelectedButton.Location = New-Object System.Drawing.Point(300, 395)
    $deleteSelectedButton.Size = New-Object System.Drawing.Size(120, 30)
    $deleteSelectedButton.BackColor = [System.Drawing.Color]::FromArgb(200, 100, 60)
    $deleteSelectedButton.ForeColor = $script:fgColor
    $deleteSelectedButton.Add_Click({
        if ($historyListBox.SelectedItems.Count -gt 0) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "選択した $($historyListBox.SelectedItems.Count) 件の履歴を削除しますか？", 
                "確認", 
                [System.Windows.Forms.MessageBoxButtons]::YesNo
            )
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                # 選択項目を取得（削除中にインデックスが変わるため配列化）
                $itemsToDelete = @($historyListBox.SelectedItems)
                
                # リストボックスから削除
                foreach ($item in $itemsToDelete) {
                    $historyListBox.Items.Remove($item)
                }
                
                # ファイルに保存
                $remainingItems = @()
                foreach ($item in $historyListBox.Items) {
                    $remainingItems += $item
                }
                Save-History $remainingItems
                
                [System.Windows.Forms.MessageBox]::Show("選択した履歴を削除しました。")
                
                # 履歴が空になったらダイアログを閉じる
                if ($historyListBox.Items.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("履歴が空になりました。")
                    $historyForm.Close()
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("削除する項目を選択してください。")
        }
    })
    $historyForm.Controls.Add($deleteSelectedButton)
    
    # 全ての履歴を削除ボタン
    $clearHistoryButton = New-Object System.Windows.Forms.Button
    $clearHistoryButton.Text = "全ての履歴を削除"
    $clearHistoryButton.Location = New-Object System.Drawing.Point(430, 395)
    $clearHistoryButton.Size = New-Object System.Drawing.Size(140, 30)
    $clearHistoryButton.BackColor = [System.Drawing.Color]::FromArgb(200, 60, 60)
    $clearHistoryButton.ForeColor = $script:fgColor
    $clearHistoryButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show("全ての履歴を削除しますか？", "確認", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            if (Test-Path $historyFile) {
                Remove-Item -Path $historyFile -Force
            }
            $historyListBox.Items.Clear()
            [System.Windows.Forms.MessageBox]::Show("全ての履歴を削除しました。")
            $historyForm.Close()
        }
    })
    $historyForm.Controls.Add($clearHistoryButton)
    
    # 閉じるボタン
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "閉じる"
    $closeButton.Location = New-Object System.Drawing.Point(580, 395)
    $closeButton.Size = New-Object System.Drawing.Size(80, 30)
    $closeButton.Add_Click({
        $historyForm.Close()
    })
    $historyForm.Controls.Add($closeButton)
    
    [void]$historyForm.ShowDialog($form)
}

$showWindowButton.Add_Click({ Show-ResultWindows })
$closeAllWindowsButton.Add_Click({ Close-AllResultWindows })

# --- ツール情報表示機能 ---
function Show-YtDlpInfo {
    $version = "不明"
    $url = "https://github.com/yt-dlp/yt-dlp"
    
    if (Test-Path $script:ytDlpPath) {
        try {
            $versionOutput = & $script:ytDlpPath --version 2>$null
            if ($versionOutput) {
                $version = $versionOutput.Trim()
            }
        } catch {
            $version = "取得失敗"
        }
    } else {
        $version = "ツールが見つかりません"
    }
    
    $infoForm = New-Object System.Windows.Forms.Form
    $infoForm.Text = "yt-dlp について"
    $infoForm.Size = New-Object System.Drawing.Size(450, 220)
    $infoForm.StartPosition = "CenterParent"
    $infoForm.FormBorderStyle = "FixedDialog"
    $infoForm.MaximizeBox = $false
    $infoForm.MinimizeBox = $false
    $infoForm.BackColor = $script:bgColor
    $infoForm.ForeColor = $script:fgColor
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "yt-dlp"
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(400, 25)
    $titleLabel.Font = New-Object System.Drawing.Font("Meiryo UI", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = $script:fgColor
    $infoForm.Controls.Add($titleLabel)
    
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "バージョン: $version"
    $versionLabel.Location = New-Object System.Drawing.Point(20, 55)
    $versionLabel.Size = New-Object System.Drawing.Size(400, 20)
    $versionLabel.ForeColor = $script:fgColor
    $infoForm.Controls.Add($versionLabel)
    
    $urlLabel = New-Object System.Windows.Forms.Label
    $urlLabel.Text = "URL:"
    $urlLabel.Location = New-Object System.Drawing.Point(20, 85)
    $urlLabel.Size = New-Object System.Drawing.Size(400, 20)
    $urlLabel.ForeColor = $script:fgColor
    $infoForm.Controls.Add($urlLabel)
    
    $urlLink = New-Object System.Windows.Forms.LinkLabel
    $urlLink.Text = $url
    $urlLink.Location = New-Object System.Drawing.Point(20, 105)
    $urlLink.Size = New-Object System.Drawing.Size(400, 20)
    $urlLink.LinkColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $urlLink.Add_LinkClicked({
        Start-Process $url
    })
    $infoForm.Controls.Add($urlLink)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(175, 140)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $infoForm.Controls.Add($okButton)
    $infoForm.AcceptButton = $okButton
    
    [void]$infoForm.ShowDialog($form)
}

function Show-MediaInfoInfo {
    $version = "不明"
    $url = "https://mediaarea.net/en/MediaInfo/Download/Windows"
    
    if (Test-Path $script:mediaInfoPath) {
        try {
            # ファイルのバージョン情報から取得
            $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($script:mediaInfoPath)
            if ($fileInfo.FileVersion) {
                $version = $fileInfo.FileVersion
            } elseif ($fileInfo.ProductVersion) {
                $version = $fileInfo.ProductVersion
            }
        } catch {
            $version = "取得失敗"
        }
    } else {
        $version = "ツールが見つかりません"
    }
    
    $infoForm = New-Object System.Windows.Forms.Form
    $infoForm.Text = "MediaInfo CLI について"
    $infoForm.Size = New-Object System.Drawing.Size(450, 220)
    $infoForm.StartPosition = "CenterParent"
    $infoForm.FormBorderStyle = "FixedDialog"
    $infoForm.MaximizeBox = $false
    $infoForm.MinimizeBox = $false
    $infoForm.BackColor = $script:bgColor
    $infoForm.ForeColor = $script:fgColor
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "MediaInfo CLI"
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(400, 25)
    $titleLabel.Font = New-Object System.Drawing.Font("Meiryo UI", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = $script:fgColor
    $infoForm.Controls.Add($titleLabel)
    
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "バージョン: $version"
    $versionLabel.Location = New-Object System.Drawing.Point(20, 55)
    $versionLabel.Size = New-Object System.Drawing.Size(400, 20)
    $versionLabel.ForeColor = $script:fgColor
    $infoForm.Controls.Add($versionLabel)
    
    $urlLabel = New-Object System.Windows.Forms.Label
    $urlLabel.Text = "URL:"
    $urlLabel.Location = New-Object System.Drawing.Point(20, 85)
    $urlLabel.Size = New-Object System.Drawing.Size(400, 20)
    $urlLabel.ForeColor = $script:fgColor
    $infoForm.Controls.Add($urlLabel)
    
    $urlLink = New-Object System.Windows.Forms.LinkLabel
    $urlLink.Text = $url
    $urlLink.Location = New-Object System.Drawing.Point(20, 105)
    $urlLink.Size = New-Object System.Drawing.Size(400, 20)
    $urlLink.LinkColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $urlLink.Add_LinkClicked({
        Start-Process $url
    })
    $infoForm.Controls.Add($urlLink)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(175, 140)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $infoForm.Controls.Add($okButton)
    $infoForm.AcceptButton = $okButton
    
    [void]$infoForm.ShowDialog($form)
}

function Show-AboutMediaInspector {
    $url = "https://github.com/shion255/MediaInspector"
    
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "MediaInspector について"
    $aboutForm.Size = New-Object System.Drawing.Size(450, 220)
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false
    $aboutForm.MinimizeBox = $false
    $aboutForm.BackColor = $script:bgColor
    $aboutForm.ForeColor = $script:fgColor
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "MediaInspector"
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(400, 25)
    $titleLabel.Font = New-Object System.Drawing.Font("Meiryo UI", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = $script:fgColor
    $aboutForm.Controls.Add($titleLabel)
    
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "バージョン: $script:version"
    $versionLabel.Location = New-Object System.Drawing.Point(20, 55)
    $versionLabel.Size = New-Object System.Drawing.Size(400, 20)
    $versionLabel.ForeColor = $script:fgColor
    $aboutForm.Controls.Add($versionLabel)
    
    $urlLabel = New-Object System.Windows.Forms.Label
    $urlLabel.Text = "URL:"
    $urlLabel.Location = New-Object System.Drawing.Point(20, 85)
    $urlLabel.Size = New-Object System.Drawing.Size(400, 20)
    $urlLabel.ForeColor = $script:fgColor
    $aboutForm.Controls.Add($urlLabel)
    
    $urlLink = New-Object System.Windows.Forms.LinkLabel
    $urlLink.Text = $url
    $urlLink.Location = New-Object System.Drawing.Point(20, 105)
    $urlLink.Size = New-Object System.Drawing.Size(400, 20)
    $urlLink.LinkColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $urlLink.Add_LinkClicked({
        Start-Process $url
    })
    $aboutForm.Controls.Add($urlLink)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(175, 140)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $aboutForm.Controls.Add($okButton)
    $aboutForm.AcceptButton = $okButton
    
    [void]$aboutForm.ShowDialog($form)
}

# --- 解析結果をリスト表示 ---
function Show-AllResultsList {
    if ($script:analysisResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("解析結果がありません。先に解析を実行してください。", "情報")
        return
    }

    if ($script:analysisListForm -and -not $script:analysisListForm.IsDisposed) {
        $listView = $script:analysisListListView
        $listView.Items.Clear()
        $rowIndex = 0
        foreach ($result in $script:analysisResults) {
            $item = New-Object System.Windows.Forms.ListViewItem($result.Title)
            $item.Tag = $result
            if ($rowIndex % 2 -eq 0) {
                if ($script:currentTheme -eq "Dark") {
                    $item.BackColor = [System.Drawing.Color]::FromArgb(55, 60, 65)
                } else {
                    $item.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
                }
            }
            [void]$listView.Items.Add($item)
            $rowIndex++
        }
        $script:analysisListForm.Text = "解析結果一覧 - $($listView.Items.Count)件"
        $script:analysisListForm.BringToFront()
        $script:analysisListForm.Focus()
        return
    }
    
    $resultForm = New-Object System.Windows.Forms.Form
    $resultForm.Text = "解析結果一覧 - $($script:analysisResults.Count)件"
    $resultForm.Size = New-Object System.Drawing.Size(800, 650)
    $resultForm.StartPosition = "CenterScreen"
    $resultForm.BackColor = $script:bgColor
    $resultForm.ForeColor = $script:fgColor

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 10)
    $listView.Size = New-Object System.Drawing.Size(760, 550)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.BackColor = $script:inputBgColor
    $listView.ForeColor = $script:fgColor
    $listView.Anchor = "Top,Bottom,Left,Right"

    [void]$listView.Columns.Add("ファイル名", 700)

    $rowIndex = 0
    foreach ($result in $script:analysisResults) {
        $item = New-Object System.Windows.Forms.ListViewItem($result.Title)
        $item.Tag = $result
        if ($rowIndex % 2 -eq 0) {
            if ($script:currentTheme -eq "Dark") {
                $item.BackColor = [System.Drawing.Color]::FromArgb(55, 60, 65)
            } else {
                $item.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
            }
        }
        [void]$listView.Items.Add($item)
        $rowIndex++
    }

    $listView.Add_DoubleClick({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedResult = $listView.SelectedItems[0].Tag
            Show-ResultDetail $selectedResult
        }
    })

    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $contextMenu.BackColor = $script:menuBgColor
    $contextMenu.ForeColor = $script:fgColor

    $openFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $openFileMenuItem.Text = "ファイルを開く(&O)"
    $openFileMenuItem.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedResult = $listView.SelectedItems[0].Tag
            if ($selectedResult.ContainsKey('FullPath') -and $selectedResult.FullPath) {
                $filePath = $selectedResult.FullPath
            } else {
                $filePath = $selectedResult.Title
            }
            
            if ($filePath -and (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                try {
                    Start-Process $filePath
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("ファイルを開けませんでした。`n$($_.Exception.Message)", "エラー")
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("ファイルが見つかりません:`n$filePath", "エラー")
            }
        }
    })
    $contextMenu.Items.Add($openFileMenuItem)

    $openFolderMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $openFolderMenuItem.Text = "フォルダを開く(&F)"
    $openFolderMenuItem.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedResult = $listView.SelectedItems[0].Tag
            if ($selectedResult.ContainsKey('FullPath') -and $selectedResult.FullPath) {
                $filePath = $selectedResult.FullPath
            } else {
                $filePath = $selectedResult.Title
            }
            
            if ($filePath -and (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                try {
                    Start-Process "explorer.exe" -ArgumentList "/select,`"$filePath`""
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("フォルダを開けませんでした。`n$($_.Exception.Message)", "エラー")
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("ファイルが見つかりません:`n$filePath", "エラー")
            }
        }
    })
    $contextMenu.Items.Add($openFolderMenuItem)

    $contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $copyFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyFileMenuItem.Text = "ファイルをコピー(&C)..."
    $copyFileMenuItem.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("ファイルを選択してください。", "情報")
            return
        }
        
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "コピー先のフォルダを選択してください"
        $folderBrowser.ShowNewFolderButton = $true
        
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $destFolder = $folderBrowser.SelectedPath
            $successCount = 0
            $errorCount = 0
            
            foreach ($item in $listView.SelectedItems) {
                $selectedResult = $item.Tag
                if ($selectedResult.ContainsKey('FullPath') -and $selectedResult.FullPath) {
                    $filePath = $selectedResult.FullPath
                } else {
                    $filePath = $selectedResult.Title
                }
                
                if ($filePath -and (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                    $fileName = [System.IO.Path]::GetFileName($filePath)
                    $destPath = Join-Path $destFolder $fileName
                    
                    try {
                        Copy-Item -LiteralPath $filePath -Destination $destPath -ErrorAction Stop
                        $successCount++
                    } catch {
                        $errorCount++
                    }
                }
            }
            
            $message = "処理完了`n`nコピー: $successCount 件"
            if ($errorCount -gt 0) {
                $message += "`nエラー: $errorCount 件"
            }
            [System.Windows.Forms.MessageBox]::Show($message, "処理結果")
        }
    })
    $contextMenu.Items.Add($copyFileMenuItem)

    $moveFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $moveFileMenuItem.Text = "ファイルを移動(&M)..."
    $moveFileMenuItem.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("ファイルを選択してください。", "情報")
            return
        }
        
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "移動先のフォルダを選択してください"
        $folderBrowser.ShowNewFolderButton = $true
        
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $destFolder = $folderBrowser.SelectedPath
            $successCount = 0
            $errorCount = 0
            
            $itemsToRemove = @($listView.SelectedItems)
            foreach ($item in $itemsToRemove) {
                $selectedResult = $item.Tag
                if ($selectedResult.ContainsKey('FullPath') -and $selectedResult.FullPath) {
                       $filePath = $selectedResult.FullPath
                } else {
                       $filePath = $selectedResult.Title
                }
                
                if ($filePath -and (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                    $fileName = [System.IO.Path]::GetFileName($filePath)
                    $destPath = Join-Path $destFolder $fileName
                    
                    try {
                        Move-Item -LiteralPath $filePath -Destination $destPath -ErrorAction Stop
                        $listView.Items.Remove($item)
                        $successCount++
                    } catch {
                        $errorCount++
                    }
                }
            }
            
            $message = "処理完了`n`n移動: $successCount 件"
            if ($errorCount -gt 0) {
                $message += "`nエラー: $errorCount 件"
            }
            [System.Windows.Forms.MessageBox]::Show($message, "処理結果")
            
            if ($listView.Items.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("すべてのファイルが移動されました。", "情報")
                $resultForm.Close()
            } else {
                $resultForm.Text = "解析結果一覧 - $($listView.Items.Count)件"
            }
        }
    })
    $contextMenu.Items.Add($moveFileMenuItem)

    $contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $deleteFileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $deleteFileMenuItem.Text = "ファイルを削除(ごみ箱)(&D)"
    $deleteFileMenuItem.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("ファイルを選択してください。", "情報")
            return
        }
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            "$($listView.SelectedItems.Count)件のファイルをごみ箱に移動しますか？",
            "確認",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Add-Type -AssemblyName Microsoft.VisualBasic
            $successCount = 0
            $errorCount = 0
            
            $itemsToRemove = @($listView.SelectedItems)
            foreach ($item in $itemsToRemove) {
                $selectedResult = $item.Tag
                if ($selectedResult.ContainsKey('FullPath') -and $selectedResult.FullPath) {
                    $filePath = $selectedResult.FullPath
                } else {
                    $filePath = $selectedResult.Title
                }
                
                if ($filePath -and (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                    try {
                        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($filePath, 'OnlyErrorDialogs', 'SendToRecycleBin')
                        $listView.Items.Remove($item)
                        $successCount++
                    } catch {
                        $errorCount++
                    }
                }
            }
            
            $message = "処理完了`n`nごみ箱に移動: $successCount 件"
            if ($errorCount -gt 0) {
                $message += "`nエラー: $errorCount 件"
            }
            [System.Windows.Forms.MessageBox]::Show($message, "処理結果")
            
            if ($listView.Items.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("すべてのファイルが削除されました。", "情報")
                $resultForm.Close()
            } else {
                $resultForm.Text = "解析結果一覧 - $($listView.Items.Count)件"
            }
        }
    })
    $contextMenu.Items.Add($deleteFileMenuItem)

    $contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $copyUrlMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyUrlMenuItem.Text = "コメントタグのURLをコピー(&U)"
    $copyUrlMenuItem.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("ファイルを選択してください。", "情報")
            return
        }
        
        $urls = @()
        
        foreach ($item in $listView.SelectedItems) {
            $selectedResult = $item.Tag
            if ($selectedResult.ContainsKey('FullPath') -and $selectedResult.FullPath) {
                $filePath = $selectedResult.FullPath
            } else {
                $filePath = $selectedResult.Title
            }
            
            if ($filePath -and (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                $mediaInfoOutput = Invoke-MediaInfo "$filePath"
                
                if ($mediaInfoOutput) {
                    foreach ($line in $mediaInfoOutput) {
                        if ($line -match '^\s*Comment\s*:\s*(.+)$') {
                            $comment = $matches[1].Trim()
                            if ($comment -match 'https?://[^\s]+') {
                                $urls += $matches[0]
                            }
                            break
                        }
                    }
                }
            }
        }
        
        if ($urls.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("コメントタグにURLが見つかりませんでした。", "情報")
            return
        }
        
        $urlText = $urls -join "`r`n"
        
        try {
            [System.Windows.Forms.Clipboard]::SetText($urlText)
            
            $notifyForm = New-Object System.Windows.Forms.Form
            $notifyForm.Text = "コピー完了"
            $notifyForm.Size = New-Object System.Drawing.Size(300, 140)
            $notifyForm.StartPosition = "CenterParent"
            $notifyForm.FormBorderStyle = "FixedDialog"
            $notifyForm.MaximizeBox = $false
            $notifyForm.MinimizeBox = $false
            $notifyForm.BackColor = $script:bgColor
            $notifyForm.ForeColor = $script:fgColor
            
            $notifyLabel = New-Object System.Windows.Forms.Label
            $notifyLabel.Text = "$($urls.Count)件のURLを`r`nクリップボードにコピーしました。"
            $notifyLabel.Location = New-Object System.Drawing.Point(20, 15)
            $notifyLabel.Size = New-Object System.Drawing.Size(260, 50)
            $notifyLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $notifyLabel.ForeColor = $script:fgColor
            $notifyForm.Controls.Add($notifyLabel)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "OK"
            $okButton.Location = New-Object System.Drawing.Point(110, 70)
            $okButton.Size = New-Object System.Drawing.Size(80, 25)
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $notifyForm.Controls.Add($okButton)
            $notifyForm.AcceptButton = $okButton
            
            $autoCloseTimer = New-Object System.Windows.Forms.Timer
            $autoCloseTimer.Interval = 3000
            $autoCloseTimer.Add_Tick({
                param($sender, $e)
                if ($notifyForm -and -not $notifyForm.IsDisposed) {
                    $notifyForm.Close()
                }
                $sender.Stop()
                $sender.Dispose()
            })
            
            $notifyForm.Add_FormClosed({
                if ($autoCloseTimer) {
                    $autoCloseTimer.Stop()
                    $autoCloseTimer.Dispose()
                }
            })
            
            $autoCloseTimer.Start()
            
            [void]$notifyForm.ShowDialog($resultForm)
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("クリップボードへのコピーに失敗しました。`n$($_.Exception.Message)", "エラー")
        }
    })
    $contextMenu.Items.Add($copyUrlMenuItem)

    $copyPathMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyPathMenuItem.Text = "フルパスをコピー(&P)"
    $copyPathMenuItem.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("ファイルを選択してください。", "情報")
            return
        }
        
        $paths = @()
        
        foreach ($item in $listView.SelectedItems) {
            $selectedResult = $item.Tag
            if ($selectedResult.ContainsKey('FullPath') -and $selectedResult.FullPath) {
                $paths += "`"$($selectedResult.FullPath)`""
            } else {
                $paths += "`"$($selectedResult.Title)`""
            }
        }
        
        $pathText = $paths -join "`r`n"
        
        try {
            [System.Windows.Forms.Clipboard]::SetText($pathText)
            
            $notifyForm = New-Object System.Windows.Forms.Form
            $notifyForm.Text = "コピー完了"
            $notifyForm.Size = New-Object System.Drawing.Size(300, 140)
            $notifyForm.StartPosition = "CenterParent"
            $notifyForm.FormBorderStyle = "FixedDialog"
            $notifyForm.MaximizeBox = $false
            $notifyForm.MinimizeBox = $false
            $notifyForm.BackColor = $script:bgColor
            $notifyForm.ForeColor = $script:fgColor
            
            $notifyLabel = New-Object System.Windows.Forms.Label
            $notifyLabel.Text = "$($paths.Count)件のパスを`r`nクリップボードにコピーしました。"
            $notifyLabel.Location = New-Object System.Drawing.Point(20, 15)
            $notifyLabel.Size = New-Object System.Drawing.Size(260, 50)
            $notifyLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $notifyLabel.ForeColor = $script:fgColor
            $notifyForm.Controls.Add($notifyLabel)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "OK"
            $okButton.Location = New-Object System.Drawing.Point(110, 70)
            $okButton.Size = New-Object System.Drawing.Size(80, 25)
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $notifyForm.Controls.Add($okButton)
            $notifyForm.AcceptButton = $okButton
            
            $autoCloseTimer = New-Object System.Windows.Forms.Timer
            $autoCloseTimer.Interval = 3000
            $autoCloseTimer.Add_Tick({
                param($sender, $e)
                if ($notifyForm -and -not $notifyForm.IsDisposed) {
                    $notifyForm.Close()
                }
                $sender.Stop()
                $sender.Dispose()
            })
            
            $notifyForm.Add_FormClosed({
                if ($autoCloseTimer) {
                    $autoCloseTimer.Stop()
                    $autoCloseTimer.Dispose()
                }
            })
            
            $autoCloseTimer.Start()
            
            [void]$notifyForm.ShowDialog($resultForm)
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("クリップボードへのコピーに失敗しました。`n$($_.Exception.Message)", "エラー")
        }
    })
    $contextMenu.Items.Add($copyPathMenuItem)

    $listView.ContextMenuStrip = $contextMenu

    $resultForm.Controls.Add($listView)

    $showWindowButton = New-Object System.Windows.Forms.Button
    $showWindowButton.Text = "結果を別ウィンドウ表示"
    $showWindowButton.Location = New-Object System.Drawing.Point(50, 570)
    $showWindowButton.Size = New-Object System.Drawing.Size(180, 30)
    $showWindowButton.BackColor = [System.Drawing.Color]::FromArgb(90, 150, 90)
    $showWindowButton.ForeColor = $script:fgColor
    $showWindowButton.Anchor = "Bottom"
    $showWindowButton.Add_Click({
        $previousResults = $script:analysisResults
        $script:analysisResults = @()
        
        foreach ($item in $listView.Items) {
            $script:analysisResults += $item.Tag
        }
        
        Show-ResultWindows
        
        $script:analysisResults = $previousResults
    })
    $resultForm.Controls.Add($showWindowButton)

    $closeAllButton = New-Object System.Windows.Forms.Button
    $closeAllButton.Text = "全ウィンドウを閉じる"
    $closeAllButton.Location = New-Object System.Drawing.Point(240, 570)
    $closeAllButton.Size = New-Object System.Drawing.Size(150, 30)
    $closeAllButton.BackColor = [System.Drawing.Color]::FromArgb(180, 60, 60)
    $closeAllButton.ForeColor = $script:fgColor
    $closeAllButton.Anchor = "Bottom"
    $closeAllButton.Add_Click({
        Close-AllResultWindows
    })
    $resultForm.Controls.Add($closeAllButton)

    $filterButton = New-Object System.Windows.Forms.Button
    $filterButton.Text = "絞り込み"
    $filterButton.Location = New-Object System.Drawing.Point(400, 570)
    $filterButton.Size = New-Object System.Drawing.Size(100, 30)
    $filterButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $filterButton.ForeColor = $script:fgColor
    $filterButton.Anchor = "Bottom"
    $filterButton.Add_Click({
        Show-FilterDialog
    })
    $resultForm.Controls.Add($filterButton)

    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "結果をコピー"
    $copyButton.Location = New-Object System.Drawing.Point(510, 570)
    $copyButton.Size = New-Object System.Drawing.Size(120, 30)
    $copyButton.BackColor = [System.Drawing.Color]::FromArgb(100, 120, 140)
    $copyButton.ForeColor = $script:fgColor
    $copyButton.Anchor = "Bottom"
    $copyButton.Add_Click({
        if ($listView.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("コピーする内容がありません。", "情報")
            return
        }
        
        $copyText = ""
        foreach ($item in $listView.Items) {
            $result = $item.Tag
            $copyText += $result.Content + "`r`n"
        }
        
        $removePatterns = @(
            "解析開始\.\.\. 少々お待ちください。",
            "ローカルファイルとして解析します。",
            "フォルダ内のファイルを解析します。",
            "見つかったファイル数:.*",
            "解析完了。",
            "=== 全ファイル解析完了 ==="
        )
        
        $lines = $copyText -split "`r?`n"
        $filteredLines = $lines | Where-Object {
            $line = $_.Trim()
            $shouldKeep = $true
            foreach ($pattern in $removePatterns) {
                if ($line -match "^$pattern`$") {
                    $shouldKeep = $false
                    break
                }
            }
            $shouldKeep
        }
        
        $cleanedText = $filteredLines -join "`r`n"
        $cleanedText = $cleanedText -replace "--- 利用可能なコーデック一覧 ---\s*`r?`n\s*⚠ フォーマット情報を解析できませんでした。\s*`r?`n", ""
        $cleanedText = $cleanedText -replace "(`r?`n){3,}", "`r`n`r`n"
        
        try {
            [System.Windows.Forms.Clipboard]::SetText($cleanedText)
            
            $notifyForm = New-Object System.Windows.Forms.Form
            $notifyForm.Text = "コピー完了"
            $notifyForm.Size = New-Object System.Drawing.Size(300, 140)
            $notifyForm.StartPosition = "CenterParent"
            $notifyForm.FormBorderStyle = "FixedDialog"
            $notifyForm.MaximizeBox = $false
            $notifyForm.MinimizeBox = $false
            $notifyForm.BackColor = $script:bgColor
            $notifyForm.ForeColor = $script:fgColor
            
            $notifyLabel = New-Object System.Windows.Forms.Label
            $notifyLabel.Text = "出力内容をクリップボードに`r`nコピーしました。"
            $notifyLabel.Location = New-Object System.Drawing.Point(20, 15)
            $notifyLabel.Size = New-Object System.Drawing.Size(260, 50)
            $notifyLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $notifyLabel.ForeColor = $script:fgColor
            $notifyForm.Controls.Add($notifyLabel)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "OK"
            $okButton.Location = New-Object System.Drawing.Point(110, 70)
            $okButton.Size = New-Object System.Drawing.Size(80, 25)
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $notifyForm.Controls.Add($okButton)
            $notifyForm.AcceptButton = $okButton
            
            $autoCloseTimer = New-Object System.Windows.Forms.Timer
            $autoCloseTimer.Interval = 3000
            $autoCloseTimer.Add_Tick({
                param($sender, $e)
                if ($notifyForm -and -not $notifyForm.IsDisposed) {
                    $notifyForm.Close()
                }
                $sender.Stop()
                $sender.Dispose()
            })
            
            $notifyForm.Add_FormClosed({
                if ($autoCloseTimer) {
                    $autoCloseTimer.Stop()
                    $autoCloseTimer.Dispose()
                }
            })
            
            $autoCloseTimer.Start()
            
            [void]$notifyForm.ShowDialog($resultForm)
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("クリップボードへのコピーに失敗しました。`n$($_.Exception.Message)", "エラー")
        }
    })
    $resultForm.Controls.Add($copyButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "閉じる"
    $closeButton.Location = New-Object System.Drawing.Point(640, 570)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $closeButton.ForeColor = $script:fgColor
    $closeButton.Anchor = "Bottom"
    $closeButton.Add_Click({
        $resultForm.Close()
    })
    $resultForm.Controls.Add($closeButton)

    $resultForm.Add_FormClosed({
        $script:analysisListForm = $null
        $script:analysisListListView = $null
    })

    $script:analysisListForm = $resultForm
    $script:analysisListListView = $listView

    [void]$resultForm.ShowDialog($form)
}

# --- 解析結果絞り込み機能 ---
function Show-FilterDialog {
    if ($script:analysisResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("解析結果がありません。先に解析を実行してください。", "情報")
        return
    }
    
    $videoCodecs = @{}
    $audioCodecs = @{}
    $hdrTypes = @{}
    $hasChapterCount = 0
    $hasSubtitleCount = 0
    $hasCoverImageCount = 0
    $noChapterCount = 0
    $noSubtitleCount = 0
    $noCoverImageCount = 0
    
    foreach ($result in $script:analysisResults) {
        $content = $result.Content
        
        if ($content -match '映像\d+:\s*([^\s]+)') {
            $codec = $matches[1]
            if (-not $videoCodecs.ContainsKey($codec)) {
                $videoCodecs[$codec] = 0
            }
            $videoCodecs[$codec]++
        }
        
        if ($content -match '音声\d+:\s*([^\s]+)') {
            $codec = $matches[1]
            if (-not $audioCodecs.ContainsKey($codec)) {
                $audioCodecs[$codec] = 0
            }
            $audioCodecs[$codec]++
        }
        
        if ($content -match '\[(HDR10|HLG|SDR)\]') {
            $hdrType = $matches[1]
            if (-not $hdrTypes.ContainsKey($hdrType)) {
                $hdrTypes[$hdrType] = 0
            }
            $hdrTypes[$hdrType]++
        }
        
        if ($content -match '✅ チャプターあり') {
            $hasChapterCount++
        } elseif ($content -match '❌ チャプターなし') {
            $noChapterCount++
        }
        
        if ($content -match 'テキスト\d+:') {
            $hasSubtitleCount++
        } else {
            $noSubtitleCount++
        }
        
        if ($content -match 'カバー画像\d+:') {
            $hasCoverImageCount++
        } else {
            $noCoverImageCount++
        }
    }
    
    if ($videoCodecs.Count -eq 0 -and $audioCodecs.Count -eq 0 -and $hdrTypes.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("コーデック情報が見つかりませんでした。", "情報")
        return
    }
    
    $filterForm = New-Object System.Windows.Forms.Form
    $filterForm.Text = "解析結果から絞り込み"
    $filterForm.Size = New-Object System.Drawing.Size(850, 500)
    $filterForm.StartPosition = "CenterParent"
    $filterForm.BackColor = $script:bgColor
    $filterForm.ForeColor = $script:fgColor
    
    $videoGroupBox = New-Object System.Windows.Forms.GroupBox
    $videoGroupBox.Text = "映像コーデック"
    $videoGroupBox.Location = New-Object System.Drawing.Point(20, 20)
    $videoGroupBox.Size = New-Object System.Drawing.Size(240, 350)
    $videoGroupBox.ForeColor = $script:fgColor
    $filterForm.Controls.Add($videoGroupBox)
    
    $videoPanel = New-Object System.Windows.Forms.Panel
    $videoPanel.Location = New-Object System.Drawing.Point(10, 25)
    $videoPanel.Size = New-Object System.Drawing.Size(205, 310)
    $videoPanel.AutoScroll = $true
    $videoGroupBox.Controls.Add($videoPanel)
    
    $videoCheckBoxes = @{}
    $yPos = 5
    foreach ($codec in ($videoCodecs.Keys | Sort-Object)) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = "$codec ($($videoCodecs[$codec]))"
        $checkBox.Location = New-Object System.Drawing.Point(5, $yPos)
        $checkBox.Size = New-Object System.Drawing.Size(195, 25)
        $checkBox.ForeColor = $script:fgColor
        $videoPanel.Controls.Add($checkBox)
        $videoCheckBoxes[$codec] = $checkBox
        $yPos += 30
    }
    
    $audioGroupBox = New-Object System.Windows.Forms.GroupBox
    $audioGroupBox.Text = "音声コーデック"
    $audioGroupBox.Location = New-Object System.Drawing.Point(280, 20)
    $audioGroupBox.Size = New-Object System.Drawing.Size(240, 350)
    $audioGroupBox.ForeColor = $script:fgColor
    $filterForm.Controls.Add($audioGroupBox)
    
    $audioPanel = New-Object System.Windows.Forms.Panel
    $audioPanel.Location = New-Object System.Drawing.Point(10, 25)
    $audioPanel.Size = New-Object System.Drawing.Size(205, 310)
    $audioPanel.AutoScroll = $true
    $audioGroupBox.Controls.Add($audioPanel)
    
    $audioCheckBoxes = @{}
    $yPos = 5
    foreach ($codec in ($audioCodecs.Keys | Sort-Object)) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = "$codec ($($audioCodecs[$codec]))"
        $checkBox.Location = New-Object System.Drawing.Point(5, $yPos)
        $checkBox.Size = New-Object System.Drawing.Size(195, 25)
        $checkBox.ForeColor = $script:fgColor
        $audioPanel.Controls.Add($checkBox)
        $audioCheckBoxes[$codec] = $checkBox
        $yPos += 30
    }
    
    $otherGroupBox = New-Object System.Windows.Forms.GroupBox
    $otherGroupBox.Text = "その他"
    $otherGroupBox.Location = New-Object System.Drawing.Point(540, 20)
    $otherGroupBox.Size = New-Object System.Drawing.Size(280, 350)
    $otherGroupBox.ForeColor = $script:fgColor
    $filterForm.Controls.Add($otherGroupBox)
    
    $otherPanel = New-Object System.Windows.Forms.Panel
    $otherPanel.Location = New-Object System.Drawing.Point(10, 25)
    $otherPanel.Size = New-Object System.Drawing.Size(260, 310)
    $otherPanel.AutoScroll = $true
    $otherGroupBox.Controls.Add($otherPanel)
    
    $otherCheckBoxes = @{}
    $yPos = 5
    
    $hdrOrder = @("HDR10", "HLG", "SDR")
    foreach ($hdrType in $hdrOrder) {
        if ($hdrTypes.ContainsKey($hdrType)) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = "$hdrType ($($hdrTypes[$hdrType]))"
            $checkBox.Location = New-Object System.Drawing.Point(5, $yPos)
            $checkBox.Size = New-Object System.Drawing.Size(250, 25)
            $checkBox.ForeColor = $script:fgColor
            $otherPanel.Controls.Add($checkBox)
            $otherCheckBoxes[$hdrType] = $checkBox
            $yPos += 30
        }
    }
    
    if ($hasChapterCount -gt 0 -or $noChapterCount -gt 0) {
        $yPos += 10
        
        if ($hasChapterCount -gt 0) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = "チャプターあり ($hasChapterCount)"
            $checkBox.Location = New-Object System.Drawing.Point(5, $yPos)
            $checkBox.Size = New-Object System.Drawing.Size(120, 25)
            $checkBox.ForeColor = $script:fgColor
            $otherPanel.Controls.Add($checkBox)
            $otherCheckBoxes["HasChapter"] = $checkBox
        }
        
        if ($noChapterCount -gt 0) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = "チャプターなし ($noChapterCount)"
            $checkBox.Location = New-Object System.Drawing.Point(130, $yPos)
            $checkBox.Size = New-Object System.Drawing.Size(120, 25)
            $checkBox.ForeColor = $script:fgColor
            $otherPanel.Controls.Add($checkBox)
            $otherCheckBoxes["NoChapter"] = $checkBox
        }
        
        $yPos += 30
    }
    
    if ($hasSubtitleCount -gt 0 -or $noSubtitleCount -gt 0) {
        if ($hasSubtitleCount -gt 0) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = "字幕あり ($hasSubtitleCount)"
            $checkBox.Location = New-Object System.Drawing.Point(5, $yPos)
            $checkBox.Size = New-Object System.Drawing.Size(120, 25)
            $checkBox.ForeColor = $script:fgColor
            $otherPanel.Controls.Add($checkBox)
            $otherCheckBoxes["HasSubtitle"] = $checkBox
        }
        
        if ($noSubtitleCount -gt 0) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = "字幕なし ($noSubtitleCount)"
            $checkBox.Location = New-Object System.Drawing.Point(130, $yPos)
            $checkBox.Size = New-Object System.Drawing.Size(120, 25)
            $checkBox.ForeColor = $script:fgColor
            $otherPanel.Controls.Add($checkBox)
            $otherCheckBoxes["NoSubtitle"] = $checkBox
        }
        
        $yPos += 30
    }
    
    if ($hasCoverImageCount -gt 0 -or $noCoverImageCount -gt 0) {
        if ($hasCoverImageCount -gt 0) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = "カバー画像あり ($hasCoverImageCount)"
            $checkBox.Location = New-Object System.Drawing.Point(5, $yPos)
            $checkBox.Size = New-Object System.Drawing.Size(120, 25)
            $checkBox.ForeColor = $script:fgColor
            $otherPanel.Controls.Add($checkBox)
            $otherCheckBoxes["HasCoverImage"] = $checkBox
        }
        
        if ($noCoverImageCount -gt 0) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = "カバー画像なし ($noCoverImageCount)"
            $checkBox.Location = New-Object System.Drawing.Point(130, $yPos)
            $checkBox.Size = New-Object System.Drawing.Size(120, 25)
            $checkBox.ForeColor = $script:fgColor
            $otherPanel.Controls.Add($checkBox)
            $otherCheckBoxes["NoCoverImage"] = $checkBox
        }
        
        $yPos += 30
    }
    
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "検索"
    $searchButton.Location = New-Object System.Drawing.Point(300, 390)
    $searchButton.Size = New-Object System.Drawing.Size(100, 35)
    $searchButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $searchButton.ForeColor = $script:fgColor
    $searchButton.Add_Click({
        $selectedVideoCodecs = @()
        foreach ($codec in $videoCheckBoxes.Keys) {
            if ($videoCheckBoxes[$codec].Checked) {
                $selectedVideoCodecs += $codec
            }
        }
        
        $selectedAudioCodecs = @()
        foreach ($codec in $audioCheckBoxes.Keys) {
            if ($audioCheckBoxes[$codec].Checked) {
                $selectedAudioCodecs += $codec
            }
        }
        
        $selectedOtherFilters = @{}
        foreach ($key in $otherCheckBoxes.Keys) {
            if ($otherCheckBoxes[$key].Checked) {
                $selectedOtherFilters[$key] = $true
            }
        }
        
        if ($selectedVideoCodecs.Count -eq 0 -and $selectedAudioCodecs.Count -eq 0 -and $selectedOtherFilters.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("少なくとも1つの条件を選択してください。", "情報")
            return
        }

        $filteredResults = @()
        foreach ($result in $script:analysisResults) {
            $content = $result.Content
            $videoMatch = $false
            $audioMatch = $false
            $otherMatch = $false
            
            if ($selectedVideoCodecs.Count -gt 0) {
                foreach ($codec in $selectedVideoCodecs) {
                    if ($content -match "映像\d+:\s*$([regex]::Escape($codec))") {
                        $videoMatch = $true
                        break
                    }
                }
            } else {
                $videoMatch = $true
            }
            
            if ($selectedAudioCodecs.Count -gt 0) {
                foreach ($codec in $selectedAudioCodecs) {
                    if ($content -match "音声\d+:\s*$([regex]::Escape($codec))") {
                        $audioMatch = $true
                        break
                    }
                }
            } else {
                $audioMatch = $true
            }
            
            if ($selectedOtherFilters.Count -gt 0) {
                $allOtherMatch = $true
                foreach ($key in $selectedOtherFilters.Keys) {
                    $match = $false
                    switch ($key) {
                        "HDR10" { if ($content -match "\[HDR10\]") { $match = $true } }
                        "HLG" { if ($content -match "\[HLG\]") { $match = $true } }
                        "SDR" { if ($content -match "\[SDR\]") { $match = $true } }
                        "HasChapter" { if ($content -match "✅ チャプターあり") { $match = $true } }
                        "NoChapter" { if ($content -match "❌ チャプターなし") { $match = $true } }
                        "HasSubtitle" { if ($content -match "テキスト\d+:") { $match = $true } }
                        "NoSubtitle" { if ($content -notmatch "テキスト\d+:" -and $content -match "(再生時間|ビットレート)") { $match = $true } }
                        "HasCoverImage" { if ($content -match "カバー画像\d+:") { $match = $true } }
                        "NoCoverImage" { if ($content -notmatch "カバー画像\d+:" -and $content -match "(再生時間|ビットレート)") { $match = $true } }
                    }
                    if (-not $match) {
                        $allOtherMatch = $false
                        break
                    }
                }
                $otherMatch = $allOtherMatch
            } else {
                $otherMatch = $true
            }
            
            if ($videoMatch -and $audioMatch -and $otherMatch) {
                $filteredResults += $result
            }
        }

        if ($filteredResults.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("条件に一致する結果が見つかりませんでした。", "情報")
            return
        }

        if ($script:analysisListForm -and -not $script:analysisListForm.IsDisposed) {
            $lv = $script:analysisListListView
            $lv.Items.Clear()
            $rowIndex = 0
            foreach ($r in $filteredResults) {
                $item = New-Object System.Windows.Forms.ListViewItem($r.Title)
                $item.Tag = $r
                if ($rowIndex % 2 -eq 0) {
                    if ($script:currentTheme -eq "Dark") {
                        $item.BackColor = [System.Drawing.Color]::FromArgb(55, 60, 65)
                    } else {
                        $item.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
                    }
                }
                [void]$lv.Items.Add($item)
                $rowIndex++
            }
            $script:analysisListForm.Text = "解析結果一覧 - $($lv.Items.Count)件 (絞込)"
        } else {
            $previousResults = $script:analysisResults
            $script:analysisResults = $filteredResults
            Show-AllResultsList
            $script:analysisResults = $previousResults
        }

        $filterForm.Close()
    })
    $filterForm.Controls.Add($searchButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "キャンセル"
    $cancelButton.Location = New-Object System.Drawing.Point(450, 390)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $cancelButton.ForeColor = $script:fgColor
    $cancelButton.Add_Click({
        $filterForm.Close()
    })
    $filterForm.Controls.Add($cancelButton)
    
    [void]$filterForm.ShowDialog($form)
}

function Show-ResultDetail($result) {
    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = "解析結果: $($result.Title)"
    $detailForm.Size = New-Object System.Drawing.Size(700, 600)
    $detailForm.StartPosition = "CenterScreen"
    $detailForm.BackColor = $script:bgColor
    $detailForm.ForeColor = $script:fgColor
    
    $detailTextBox = New-Object System.Windows.Forms.TextBox
    $detailTextBox.Multiline = $true
    $detailTextBox.ScrollBars = "Vertical"
    $detailTextBox.ReadOnly = $true
    $detailTextBox.Font = New-Object System.Drawing.Font($script:currentFontName, $script:currentFontSize)
    $detailTextBox.Dock = "Fill"
    $detailTextBox.BackColor = $script:outputBgColor
    $detailTextBox.ForeColor = $script:fgColor
    $detailTextBox.Text = $result.Content
    $detailForm.Controls.Add($detailTextBox)
    
    [void]$detailForm.ShowDialog()
}

# --- 動画ファイル整理機能 ---
function Get-MediaInfoArtist($filePath) {
    try {
        $tempFile = [System.IO.Path]::GetTempFileName()
        $null = & $script:mediaInfoPath --Output=Text --LogFile="$tempFile" "$filePath" 2>$null
        
        if (Test-Path $tempFile) {
            $output = Get-Content -Path $tempFile -Encoding UTF8
            Remove-Item -Path $tempFile -Force
            
            # Artist/Performerを検索
            foreach ($line in $output) {
                if ($line -match '^\s*(Performer|Artist)\s*:\s*(.+)$') {
                    return $matches[2].Trim()
                }
            }
        }
        return $null
    } catch {
        return $null
    }
}

function Show-FileOrganizer {
    # フォルダ選択ダイアログ
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "動画ファイルが保存されているフォルダを選択してください"
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($folderBrowser.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }
    
    $targetPath = $folderBrowser.SelectedPath
    
    # 動画ファイルを検索
    Write-Host "動画ファイルを検索中: $targetPath"
    $videoExtensions = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.webm', '.m4v', '.ts', '.m2ts')
    $files = Get-ChildItem -LiteralPath $targetPath -File | Where-Object {
        $videoExtensions -contains $_.Extension.ToLower()
    }
    
    if ($files.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("動画ファイルが見つかりませんでした。", "情報")
        return
    }
    
    # 作成者情報を取得
    $fileInfoList = @()
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "作成者情報を取得中..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 120)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    $progressForm.BackColor = $script:bgColor
    
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(360, 20)
    $progressLabel.ForeColor = $script:fgColor
    $progressForm.Controls.Add($progressLabel)
    
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 50)
    $progressBar.Size = New-Object System.Drawing.Size(360, 25)
    $progressForm.Controls.Add($progressBar)
    
    $progressForm.Show()
    $progressForm.Refresh()
    
    $count = 0
    foreach ($file in $files) {
        $count++
        $progressLabel.Text = "処理中: $count / $($files.Count)"
        $progressBar.Value = [int](($count / $files.Count) * 100)
        $progressForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        $artist = Get-MediaInfoArtist $file.FullName
        
        if ($artist) {
            # ファイル名に使用できない文字を置換
            $safeArtist = $artist -replace '[<>:"/\\|?*]', '-'
            
            $fileInfoList += [PSCustomObject]@{
                FileName = $file.Name
                FullPath = $file.FullName
                Artist = $artist
                SafeArtist = $safeArtist
            }
        }
    }
    
    $progressForm.Close()
    
    if ($fileInfoList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("作成者情報を持つ動画ファイルが見つかりませんでした。", "情報")
        return
    }
    
    # 整理ダイアログを表示
    $organizerForm = New-Object System.Windows.Forms.Form
    $organizerForm.Text = "動画ファイル整理 - $($fileInfoList.Count)件"
    $organizerForm.Size = New-Object System.Drawing.Size(900, 600)
    $organizerForm.StartPosition = "CenterScreen"
    $organizerForm.BackColor = $script:bgColor
    $organizerForm.ForeColor = $script:fgColor
    
    # ListView
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 10)
    $listView.Size = New-Object System.Drawing.Size(860, 480)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.Sorting = [System.Windows.Forms.SortOrder]::None
    $listView.BackColor = $script:inputBgColor
    $listView.ForeColor = $script:fgColor
    $listView.Anchor = "Top,Bottom,Left,Right"
    
    # 列を追加
    [void]$listView.Columns.Add("ファイル名", 500)
    [void]$listView.Columns.Add("作成者", 340)
    
    # ソート用の変数
    $script:sortColumn = 1  # 作成者列を初期ソート対象に
    $script:sortOrder = $true  # true = 昇順, false = 降順
    
    # データを作成者順にソートして追加
    $sortedList = $fileInfoList | Sort-Object Artist
    $rowIndex = 0
    foreach ($info in $sortedList) {
        $item = New-Object System.Windows.Forms.ListViewItem($info.FileName)
        $item.SubItems.Add($info.Artist) | Out-Null
        $item.Tag = $info
        
        # 偶数行に背景色を設定
        if ($rowIndex % 2 -eq 0) {
            if ($script:currentTheme -eq "Dark") {
                $item.BackColor = [System.Drawing.Color]::FromArgb(55, 60, 65)
            } else {
                $item.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
            }
        }
        
        [void]$listView.Items.Add($item)
        $rowIndex++
    }
    
    # 列ヘッダークリックでソート
    $listView.Add_ColumnClick({
        param($sender, $e)
        
        $columnIndex = $e.Column
        
        # 同じ列をクリックした場合は昇順/降順を切り替え
        if ($script:sortColumn -eq $columnIndex) {
            $script:sortOrder = -not $script:sortOrder
        } else {
            $script:sortColumn = $columnIndex
            $script:sortOrder = $true
        }
        
        # ソート実行
        $items = @($listView.Items | ForEach-Object { $_ })
        $listView.Items.Clear()
        
        $sortedItems = if ($columnIndex -eq 0) {
            # ファイル名でソート
            if ($script:sortOrder) {
                $items | Sort-Object { $_.Tag.FileName }
            } else {
                $items | Sort-Object { $_.Tag.FileName } -Descending
            }
        } else {
            # 作成者でソート
            if ($script:sortOrder) {
                $items | Sort-Object { $_.Tag.Artist }
            } else {
                $items | Sort-Object { $_.Tag.Artist } -Descending
            }
        }
        
        $rowIndex = 0
        foreach ($item in $sortedItems) {
            # 偶数行に背景色を再設定
            if ($rowIndex % 2 -eq 0) {
                if ($script:currentTheme -eq "Dark") {
                    $item.BackColor = [System.Drawing.Color]::FromArgb(55, 60, 65)
                } else {
                    $item.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
                }
            } else {
                $item.BackColor = $script:inputBgColor
            }
            
            [void]$listView.Items.Add($item)
            $rowIndex++
        }
    })
    
    $organizerForm.Controls.Add($listView)
    
    # ボタンパネル
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Location = New-Object System.Drawing.Point(10, 500)
    $buttonPanel.Size = New-Object System.Drawing.Size(860, 50)
    $buttonPanel.Anchor = "Bottom,Left,Right"
    $organizerForm.Controls.Add($buttonPanel)
    
    # 移動先パス変更ボタン
    $changePathButton = New-Object System.Windows.Forms.Button
    $changePathButton.Text = "移動先パスを変更"
    $changePathButton.Location = New-Object System.Drawing.Point(0, 10)
    $changePathButton.Size = New-Object System.Drawing.Size(150, 35)
    $changePathButton.BackColor = [System.Drawing.Color]::FromArgb(100, 120, 140)
    $changePathButton.ForeColor = $script:fgColor
    $changePathButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "フォルダの作成先を選択してください"
        $folderBrowser.SelectedPath = $targetPath
        $folderBrowser.ShowNewFolderButton = $true
        
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $targetPath = $folderBrowser.SelectedPath
            [System.Windows.Forms.MessageBox]::Show("移動先パスを変更しました：`n$targetPath", "確認")
        }
    })
    $buttonPanel.Controls.Add($changePathButton)
    
    # 選択したファイルを移動ボタン
    $moveSelectedButton = New-Object System.Windows.Forms.Button
    $moveSelectedButton.Text = "選択したファイルを移動"
    $moveSelectedButton.Location = New-Object System.Drawing.Point(160, 10)
    $moveSelectedButton.Size = New-Object System.Drawing.Size(200, 35)
    $moveSelectedButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $moveSelectedButton.ForeColor = $script:fgColor
    $moveSelectedButton.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("移動するファイルを選択してください。", "情報")
            return
        }
        
        $selectedFiles = @()
        foreach ($item in $listView.SelectedItems) {
            $selectedFiles += $item.Tag
        }
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            "$($selectedFiles.Count)件のファイルを作成者別フォルダに移動しますか？", 
            "確認", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Move-FilesToChannelFolders $selectedFiles $targetPath $listView $organizerForm
        }
    })
    $buttonPanel.Controls.Add($moveSelectedButton)
    
    # 全てのファイルを移動ボタン
    $moveAllButton = New-Object System.Windows.Forms.Button
    $moveAllButton.Text = "全てのファイルを移動"
    $moveAllButton.Location = New-Object System.Drawing.Point(370, 10)
    $moveAllButton.Size = New-Object System.Drawing.Size(200, 35)
    $moveAllButton.BackColor = [System.Drawing.Color]::FromArgb(90, 150, 90)
    $moveAllButton.ForeColor = $script:fgColor
    $moveAllButton.Add_Click({
        $allFiles = @()
        foreach ($item in $listView.Items) {
            $allFiles += $item.Tag
        }
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            "$($allFiles.Count)件のファイルを作成者別フォルダに移動しますか？", 
            "確認", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Move-FilesToChannelFolders $allFiles $targetPath $listView $organizerForm
        }
    })
    $buttonPanel.Controls.Add($moveAllButton)
    
    # 閉じるボタン
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "閉じる"
    $closeButton.Location = New-Object System.Drawing.Point(580, 10)
    $closeButton.Size = New-Object System.Drawing.Size(100, 35)
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $closeButton.ForeColor = $script:fgColor
    $closeButton.Anchor = "Bottom,Right"
    $closeButton.Add_Click({
        $organizerForm.Close()
    })
    $buttonPanel.Controls.Add($closeButton)
    
    [void]$organizerForm.ShowDialog($form)
}

function Move-FilesToChannelFolders($files, $basePath, $listView, $parentForm) {
    $movedCount = 0
    $skippedCount = 0
    $errorCount = 0
    
    # 作成者別にグループ化
    $groupedFiles = $files | Group-Object SafeArtist
    
    foreach ($group in $groupedFiles) {
        $channelFolder = Join-Path $basePath $group.Name
        
        # フォルダを作成
        if (-not (Test-Path -LiteralPath $channelFolder)) {
            try {
                New-Item -ItemType Directory -Path $channelFolder -Force | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "フォルダ作成エラー: $channelFolder`n$($_.Exception.Message)", 
                    "エラー"
                )
                $errorCount += $group.Group.Count
                continue
            }
        }
        
        # ファイルを移動
        foreach ($fileInfo in $group.Group) {
            try {
                $destination = Join-Path $channelFolder ([System.IO.Path]::GetFileName($fileInfo.FullPath))
                
                # 同名ファイルが既に存在するかチェック
                if (Test-Path -LiteralPath $destination) {
                    $skippedCount++
                    continue
                }
                
                Move-Item -LiteralPath $fileInfo.FullPath -Destination $destination -ErrorAction Stop
                $movedCount++
                
                # ListViewから削除
                $itemToRemove = $listView.Items | Where-Object { $_.Tag.FullPath -eq $fileInfo.FullPath }
                if ($itemToRemove) {
                    $listView.Items.Remove($itemToRemove)
                }
                
            } catch {
                $errorCount++
            }
        }
    }
    
    # 結果を表示
    $message = "処理完了`n`n"
    $message += "移動: $movedCount 件`n"
    if ($skippedCount -gt 0) {
        $message += "スキップ (同名ファイル): $skippedCount 件`n"
    }
    if ($errorCount -gt 0) {
        $message += "エラー: $errorCount 件`n"
    }
    
    [System.Windows.Forms.MessageBox]::Show($message, "処理結果")
    
    # リストが空になったらダイアログを閉じる
    if ($listView.Items.Count -eq 0) {
        $parentForm.Close()
    } else {
        # タイトルを更新
        $parentForm.Text = "動画ファイル整理 - $($listView.Items.Count)件"
    }
}

# --- MediaInfo 呼び出し ---
function Invoke-MediaInfo($filePath) {
    try {
        # 一時ファイルに出力してUTF-8で読み込む
        $tempFile = [System.IO.Path]::GetTempFileName()
        $null = & $script:mediaInfoPath --Output=Text --LogFile="$tempFile" "$filePath" 2>$null
        
        if (Test-Path $tempFile) {
            $output = Get-Content -Path $tempFile -Encoding UTF8
            Remove-Item -Path $tempFile -Force
            return $output
        }
        return $null
    } catch {
        Write-OutputBox("⚠ MediaInfo エラー: $($_.Exception.Message)")
        return $null
    }
}

# MediaInfo出力をパースする関数
function Parse-MediaInfo($mediaInfoOutput) {
    $currentSection = ""
    $duration = ""
    $overallBitrate = ""
    $fileSize = ""
    $artist = ""
    $comment = ""
    $replayGain = ""
    $replayGainPeak = ""
    $videoStreams = @()
    $audioStreams = @()
    $textStreams = @()
    $imageStreams = @()
    $hasChapters = $false
    $chapterCount = 0
    
    $videoInfo = @{}
    $audioInfo = @{}
    $textInfo = @{}
    $imageInfo = @{}
    
    foreach ($line in $mediaInfoOutput) {
        $line = $line.Trim()
        if (-not $line) { continue }
        
        # セクションヘッダー検出
        if ($line -match '^(General|Video|Audio|Text|Menu|Image)') {
            if ($currentSection -eq "Video" -and $videoInfo.Count -gt 0) {
                $videoStreams += $videoInfo.Clone()
                $videoInfo.Clear()
            }
            if ($currentSection -eq "Audio" -and $audioInfo.Count -gt 0) {
                $audioStreams += $audioInfo.Clone()
                $audioInfo.Clear()
            }
            if ($currentSection -eq "Text" -and $textInfo.Count -gt 0) {
                $textStreams += $textInfo.Clone()
                $textInfo.Clear()
            }
            if ($currentSection -eq "Image" -and $imageInfo.Count -gt 0) {
                $imageStreams += $imageInfo.Clone()
                $imageInfo.Clear()
            }
            
            $currentSection = $matches[1]
            
            if ($currentSection -eq "Menu") {
                $hasChapters = $true
            }
            continue
        }
        
        if ($currentSection -eq "Menu") {
            if ($line -match '^\d+\s*:\s*\d') {
                $chapterCount++
            }
        }
        
        if ($line -match '^(.+?)\s*:\s*(.+)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            
            switch ($currentSection) {
                "General" {
                    if ($key -eq "Duration") { $duration = $value }
                    if ($key -eq "Overall bit rate") { $overallBitrate = $value }
                    if ($key -eq "File size") { $fileSize = $value }
                    if ($key -match "^(Performer|Artist)$") { $artist = $value }
                    if ($key -eq "Comment") { $comment = $value }
                }
                "Video" {
                    if ($key -eq "Format") { $videoInfo["format"] = $value }
                    if ($key -eq "Width") { $videoInfo["width"] = $value -replace '\D', '' }
                    if ($key -eq "Height") { $videoInfo["height"] = $value -replace '\D', '' }
                    if ($key -eq "Frame rate") { $videoInfo["fps"] = $value }
                    if ($key -eq "Frame rate mode") { $videoInfo["fps_mode"] = $value }
                    if ($key -eq "Bit rate") { $videoInfo["bitrate"] = $value }
                    if ($key -eq "Maximum bit rate") { $videoInfo["max_bitrate"] = $value }
                    if ($key -eq "Bit rate mode") { $videoInfo["bitrate_mode"] = $value }
                    if ($key -eq "Stream size") { $videoInfo["stream_size"] = $value }
                    if ($key -eq "Color space") { $videoInfo["color_space"] = $value }
                    if ($key -eq "Chroma subsampling") { $videoInfo["chroma_subsampling"] = $value }
                    if ($key -eq "Bit depth") { $videoInfo["bit_depth"] = $value }
                    if ($key -eq "Scan type") { $videoInfo["scan_type"] = $value }
                    if ($key -eq "Color range") { $videoInfo["color_range"] = $value }
                    if ($key -eq "Color primaries") { $videoInfo["color_primaries"] = $value }
                    if ($key -eq "Transfer characteristics") { $videoInfo["transfer_characteristics"] = $value }
                    if ($key -eq "Matrix coefficients") { $videoInfo["matrix_coefficients"] = $value }
                    if ($key -eq "Language") { $videoInfo["language"] = $value }
                    if ($key -eq "Writing library") { $videoInfo["writing_library"] = $value }
                    if ($key -eq "Encoding settings") { $videoInfo["encoding_settings"] = $value }
                }
                "Audio" {
                    if ($key -eq "Format") { $audioInfo["format"] = $value }
                    if ($key -match "Channel") { $audioInfo["channels"] = $value }
                    if ($key -eq "Sampling rate") { $audioInfo["samplerate"] = $value }
                    if ($key -eq "Bit rate") { $audioInfo["bitrate"] = $value }
                    if ($key -eq "Stream size") { $audioInfo["stream_size"] = $value }
                    if ($key -eq "Replay gain") { 
                        if (-not $replayGain) { $replayGain = $value }
                    }
                    if ($key -eq "Replay gain peak") { 
                        if (-not $replayGainPeak) { $replayGainPeak = $value }
                    }
                    if ($key -eq "Language") { $audioInfo["language"] = $value }
                    if ($key -eq "Writing library") { $audioInfo["writing_library"] = $value }
                    if ($key -eq "Encoding settings") { $audioInfo["encoding_settings"] = $value }
                }
                "Text" {
                    if ($key -eq "Language") { $textInfo["language"] = $value }
                    if ($key -eq "Default") { $textInfo["default"] = $value }
                    if ($key -eq "Forced") { $textInfo["forced"] = $value }
                    if ($key -eq "Stream size") { $textInfo["stream_size"] = $value }
                }
                "Image" {
                    if ($key -eq "Format") { $imageInfo["format"] = $value }
                    if ($key -eq "Width") { $imageInfo["width"] = $value -replace '\D', '' }
                    if ($key -eq "Height") { $imageInfo["height"] = $value -replace '\D', '' }
                    if ($key -eq "Stream size") { $imageInfo["stream_size"] = $value }
                }
            }
        }
    }
    
    # 最後のストリームを追加
    if ($videoInfo.Count -gt 0) { $videoStreams += $videoInfo }
    if ($audioInfo.Count -gt 0) { $audioStreams += $audioInfo }
    if ($textInfo.Count -gt 0) { $textStreams += $textInfo }
    if ($imageInfo.Count -gt 0) { $imageStreams += $imageInfo }
    
    return @{
        Duration = $duration
        OverallBitrate = $overallBitrate
        FileSize = $fileSize
        Artist = $artist
        Comment = $comment
        ReplayGain = $replayGain
        ReplayGainPeak = $replayGainPeak
        VideoStreams = $videoStreams
        AudioStreams = $audioStreams
        TextStreams = $textStreams
        ImageStreams = $imageStreams
        HasChapters = $hasChapters
        ChapterCount = $chapterCount
    }
}

# MediaInfo情報を表示する関数（設定に基づく）
function Display-MediaInfo($parsedInfo, [ref]$resultContentRef) {
    # 再生時間の日本語表記に変換
    $duration = Convert-DurationToJapanese $parsedInfo.Duration
    
    # ビットレートのフォーマット
    $overallBitrate = Format-Bitrate $parsedInfo.OverallBitrate
    
    # 基本情報を表示
    if ($script:showDuration) {
        Write-OutputBox("再生時間: $duration")
        $resultContentRef.Value += "再生時間: $duration`r`n"
    }
    
    if ($script:showBitrate) {
        Write-OutputBox("ビットレート: $overallBitrate")
        $resultContentRef.Value += "ビットレート: $overallBitrate`r`n"
    }
    
    # 作成者情報を表示
    if ($script:showArtist -and $parsedInfo.Artist) {
        Write-OutputBox("作成者: $($parsedInfo.Artist)")
        $resultContentRef.Value += "作成者: $($parsedInfo.Artist)`r`n"
    }
    
    # コメント情報を表示
    if ($script:showComment -and $parsedInfo.Comment) {
        Write-OutputBox("コメント: $($parsedInfo.Comment)")
        $resultContentRef.Value += "コメント: $($parsedInfo.Comment)`r`n"
    }
    
    # チャプター情報を表示
    if ($script:showChapters) {
        if ($parsedInfo.HasChapters) {
            if ($parsedInfo.ChapterCount -gt 0) {
                Write-OutputBox("✅ チャプターあり ($($parsedInfo.ChapterCount)個)")
                $resultContentRef.Value += "✅ チャプターあり ($($parsedInfo.ChapterCount)個)`r`n"
            } else {
                Write-OutputBox("✅ チャプターあり")
                $resultContentRef.Value += "✅ チャプターあり`r`n"
            }
        }
    }
    
    # 映像ストリーム情報（番号付き）
    $videoIndex = 1
    foreach ($v in $parsedInfo.VideoStreams) {
        $videoLine = ""
        $parts = @()
        
        # コーデック
        if ($script:showVideoCodec) {
            $format = if ($v["format"]) { $v["format"] } else { "不明" }
            $parts += $format
        }
        
        # ビットレート
        if ($script:showVideoBitrate) {
            $bitrateMode = if ($v["bitrate_mode"]) { "[$($v['bitrate_mode'])]" } else { "" }
            $bitrate = if ($v["bitrate"]) { (Format-Bitrate $v["bitrate"]) } else { "不明" }
            if ($v["max_bitrate"]) {
                $maxBitrate = Format-Bitrate $v["max_bitrate"]
                $parts += "$bitrateMode $bitrate (Max: $maxBitrate)"
            } else {
                $parts += "$bitrateMode $bitrate"
            }
        }
        
        # 解像度
        if ($script:showResolution) {
            $res = if ($v["width"] -and $v["height"]) { "$($v['width'])x$($v['height'])" } else { "不明" }
            $parts += $res
        }
        
        # FPS
        if ($script:showFPS) {
            $fpsMode = if ($v["fps_mode"]) { "[$($v['fps_mode'])]" } else { "" }
            $fps = if ($v["fps"]) { 
                ($v["fps"] -replace '\s*FPS', '') + " fps"
            } else { 
                "不明" 
            }
            $parts += "$fpsMode $fps"
        }
        
        # 色空間
        if ($script:showColorSpace -and $v["color_space"]) {
            $parts += $v["color_space"]
        }
        
        # クロマサブサンプリング
        if ($script:showChromaSubsampling -and $v["chroma_subsampling"]) {
            $parts += $v["chroma_subsampling"]
        }
        
        # ビット深度
        if ($script:showBitDepth -and $v["bit_depth"]) {
            $bitDepthValue = $v["bit_depth"] -replace '\s*bits', ''
            $parts += "${bitDepthValue}bit"
        }
        
        # スキャンタイプ
        if ($script:showScanType -and $v["scan_type"]) {
            $scanType = $v["scan_type"]
            $scanTypeJp = switch ($scanType) {
                "Progressive" { "プログレッシブ" }
                "Interlaced" { "インターレース" }
                default { $scanType }
            }
            $parts += $scanTypeJp
        }
        
        # 色範囲
        if ($script:showColorRange -and $v["color_range"]) {
            $parts += $v["color_range"]
        }
        
        # HDR/SDR情報
        if ($script:showHDR) {
            $hdrInfo = Get-HDRInfo $v["color_primaries"] $v["transfer_characteristics"] $v["matrix_coefficients"]
            $parts += $hdrInfo
        }
        
        # 言語
        if ($script:showVideoLanguage -and $v["language"]) {
            $parts += $v["language"]
        }
        
        # ライブラリ
        if ($script:showVideoWritingLibrary -and $v["writing_library"]) {
            $parts += $v["writing_library"]
        }
        
        # ストリームサイズ
        if ($script:showVideoStreamSize -and $v["stream_size"]) {
            $parts += $v["stream_size"]
        }
        
        if ($parts.Count -gt 0) {
            $videoLine = "映像${videoIndex}: " + ($parts -join " | ")
            Write-OutputBox($videoLine)
            $resultContentRef.Value += $videoLine + "`r`n"
        }
        
        if ($script:showVideoEncodingSettings -and $v["encoding_settings"]) {
            Write-OutputBox("")
            $resultContentRef.Value += "`r`n"
            Write-OutputBox($v["encoding_settings"])
            $resultContentRef.Value += $v["encoding_settings"] + "`r`n"
            Write-OutputBox("")
            $resultContentRef.Value += "`r`n"
        }
        
        $videoIndex++
    }
    
    # 音声ストリーム情報（番号付き）
    $audioIndex = 1
    foreach ($a in $parsedInfo.AudioStreams) {
        $audioLine = ""
        $parts = @()
        
        # コーデック
        if ($script:showAudioCodec) {
            $format = if ($a["format"]) { $a["format"] } else { "不明" }
            $parts += $format
        }
        
        # ビットレート
        if ($script:showAudioBitrate) {
            $bitrate = if ($a["bitrate"]) { (Format-Bitrate $a["bitrate"]) } else { "不明" }
            $parts += $bitrate
        }
        
        # サンプリングレート
        if ($script:showSampleRate) {
            $samplerate = if ($a["samplerate"]) { $a["samplerate"] } else { "不明" }
            $parts += $samplerate
        }
        
        # リプレイゲイン
        if ($script:showReplayGain) {
            if ($parsedInfo.ReplayGain) {
                $parts += "RG: $($parsedInfo.ReplayGain)"
            }
            if ($parsedInfo.ReplayGainPeak) {
                $parts += "Peak: $($parsedInfo.ReplayGainPeak)"
            }
        }
        
        # 言語
        if ($script:showAudioLanguage -and $a["language"]) {
            $parts += $a["language"]
        }
        
        # ライブラリ
        if ($script:showAudioWritingLibrary -and $a["writing_library"]) {
            $parts += $a["writing_library"]
        }
        
        # ストリームサイズ
        if ($script:showAudioStreamSize -and $a["stream_size"]) {
            $parts += $a["stream_size"]
        }
        
        if ($parts.Count -gt 0) {
            $audioLine = "音声${audioIndex}: " + ($parts -join " | ")
            Write-OutputBox($audioLine)
            $resultContentRef.Value += $audioLine + "`r`n"
        }
        
        $audioIndex++
    }
    
    # テキストストリーム情報（番号付き）
    if ($script:showTextStream) {
        $textIndex = 1
        foreach ($txt in $parsedInfo.TextStreams) {
            $language = if ($txt["language"]) { $txt["language"] } else { "不明" }
            $default = if ($txt["default"] -eq "Yes") { "はい" } else { "いいえ" }
            $forced = if ($txt["forced"] -eq "Yes") { "はい" } else { "いいえ" }
            
            $textLine = "テキスト${textIndex}: $language | Default - $default | Forced - $forced"
            Write-OutputBox($textLine)
            $resultContentRef.Value += $textLine + "`r`n"
            $textIndex++
        }
    }
    
    # 画像ストリーム情報（番号付き）
    if ($script:showCoverImage) {
        $imageIndex = 1
        foreach ($img in $parsedInfo.ImageStreams) {
            $format = if ($img["format"]) { $img["format"] } else { "不明" }
            $res = if ($img["width"] -and $img["height"]) { "$($img['width'])x$($img['height'])" } else { "不明" }
            $streamSize = if ($img["stream_size"]) { $img["stream_size"] } else { "" }
            
            $imageLine = "カバー画像${imageIndex}: $format $res"
            if ($streamSize) { $imageLine += " | $streamSize" }
            Write-OutputBox($imageLine)
            $resultContentRef.Value += $imageLine + "`r`n"
            $imageIndex++
        }
    }
}

# --- 解析処理 ---
function Analyze-Video {
    $inputsRaw = $textBox.Text.Trim()
    if (-not $inputsRaw) {
        [System.Windows.Forms.MessageBox]::Show("URLまたはファイルパスを入力してください。")
        return
    }

    # 解析結果をクリア
    $script:analysisResults = @()
    $showWindowButton.Enabled = $false

    $inputs = @()
    $lines = $inputsRaw -split "`r?`n"
    foreach ($line in $lines) {
        $line = $line.Trim()
        if (-not $line) { continue }
        
        # ダブルクォートで囲まれた部分とそうでない部分を抽出
        $currentPos = 0
        while ($currentPos -lt $line.Length) {
            # 先頭の空白をスキップ
            while ($currentPos -lt $line.Length -and $line[$currentPos] -eq ' ') {
                $currentPos++
            }
            if ($currentPos -ge $line.Length) { break }
            
            # ダブルクォートで始まる場合
            if ($line[$currentPos] -eq '"') {
                $endQuote = $line.IndexOf('"', $currentPos + 1)
                if ($endQuote -gt $currentPos) {
                    $path = $line.Substring($currentPos + 1, $endQuote - $currentPos - 1)
                    if ($path) { $inputs += $path }
                    $currentPos = $endQuote + 1
                } else {
                    # 閉じクォートがない場合は残り全体を取得
                    $path = $line.Substring($currentPos + 1).Trim()
                    if ($path) { $inputs += $path }
                    break
                }
            }
            # ダブルクォートなしの場合
            else {
                $nextSpace = $line.IndexOf(' ', $currentPos)
                $nextQuote = $line.IndexOf('"', $currentPos)
                
                # 次の区切り位置を決定
                $endPos = $line.Length
                if ($nextSpace -gt $currentPos -and $nextQuote -gt $currentPos) {
                    $endPos = [math]::Min($nextSpace, $nextQuote)
                } elseif ($nextSpace -gt $currentPos) {
                    $endPos = $nextSpace
                } elseif ($nextQuote -gt $currentPos) {
                    $endPos = $nextQuote
                }
                
                $path = $line.Substring($currentPos, $endPos - $currentPos).Trim()
                if ($path) { $inputs += $path }
                $currentPos = $endPos
            }
        }
    }
    
    if ($inputs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("有効な入力がありません。")
        return
    }
    
    $outputBox.Clear()

    $total = $inputs.Count
    $count = 0

    foreach ($inputRaw in $inputs) {
        $input = $inputRaw.Trim()
        $count++
        
        # 各ファイルの結果を保存する変数
        $resultContent = ""
        $resultTitle = $input
        
        Write-OutputBox("--------------------------------------------------")
        $resultContent += "--------------------------------------------------`r`n"
        Write-OutputBox("入力: $input")
        $resultContent += "入力: $input`r`n"
        Write-OutputBox("解析開始... 少々お待ちください。")
        $resultContent += "解析開始... 少々お待ちください。`r`n"
        Set-Progress([math]::Round($count / $total * 100 / 2))

        $isUrl = $input -match '^https?://'
        $target = $null

        try {
            if ($isUrl) {
                # yt-dlpの出力を適切に処理（余計な出力を抑制）
                $infoJson = & $script:ytDlpPath --dump-json --no-warnings "$input" 2>$null | ConvertFrom-Json
                if ($infoJson) {
                    $resultTitle = $infoJson.title
                    
                    if ($script:showYtDlpTitle) {
                        Write-OutputBox("タイトル: $($infoJson.title)")
                        $resultContent += "タイトル: $($infoJson.title)`r`n"
                    }
                    
                    if ($script:showYtDlpUploader) {
                        Write-OutputBox("アップローダー: $($infoJson.uploader)")
                        $resultContent += "アップローダー: $($infoJson.uploader)`r`n"
                    }
                    
                    if ($script:showYtDlpUploadDate -and $infoJson.upload_date) {
                        $dateStr = $infoJson.upload_date
                        if ($dateStr -match '^(\d{4})(\d{2})(\d{2})$') {
                            $formattedDate = "$($matches[1])年$($matches[2])月$($matches[3])日"
                            Write-OutputBox("投稿日時: $formattedDate")
                            $resultContent += "投稿日時: $formattedDate`r`n"
                        } else {
                            Write-OutputBox("投稿日時: $dateStr")
                            $resultContent += "投稿日時: $dateStr`r`n"
                        }
                    }
                    
                    if ($script:showYtDlpDuration) {
                        $durationText = Format-Time $infoJson.duration
                        Write-OutputBox("再生時間: $durationText")
                        $resultContent += "再生時間: $durationText`r`n"
                    }
                    
                    if ($script:showYtDlpChapters) {
                        if ($infoJson.chapters -and $infoJson.chapters.Count -gt 0) {
                            Write-OutputBox("✅ チャプターあり ($($infoJson.chapters.Count)個)")
                            $resultContent += "✅ チャプターあり ($($infoJson.chapters.Count)個)`r`n"
                        } else {
                            Write-OutputBox("❌ チャプターなし")
                            $resultContent += "❌ チャプターなし`r`n"
                        }
                    }
                    
                    if ($script:showYtDlpSubtitles) {
                        $subs = $infoJson.subtitles.PSObject.Properties.Name
                        if ($subs -match "ja|jpn|Japanese") { 
                            Write-OutputBox("✅ 日本語字幕あり")
                            $resultContent += "✅ 日本語字幕あり`r`n"
                        } else { 
                            Write-OutputBox("❌ 日本語字幕なし")
                            $resultContent += "❌ 日本語字幕なし`r`n"
                        }
                    }
                    
                    if ($script:showYtDlpFormats) {
                    # フォーマット一覧を取得（エラー出力も含める）
                    $formatOutput = & $script:ytDlpPath -F "$input" 2>&1
                    
                    if ($formatOutput) {
                        $videoFormats = @()
                        $audioFormats = @()
                        $parseStarted = $false
                        
                        foreach ($line in $formatOutput) {
                            $lineStr = $line.ToString()
                            
                            # ヘッダー行を検出（複数パターンに対応）
                            if ($lineStr -match 'ID\s+(EXT|EXTENSION).*RESOLUTION' -or 
                                $lineStr -match '^-+\s+-+\s+-+') {
                                $parseStarted = $true
                                continue
                            }
                            
                            # パース開始後、フォーマット行を処理
                            if ($parseStarted -and $lineStr.Trim() -and -not ($lineStr -match '^-+$')) {
                                # 行の先頭が数字またはフォーマットIDで始まる場合のみ処理
                                if ($lineStr -match '^\s*(\d+|sb\d+|dash\d+)\s+') {
                                    # より柔軟な正規表現でパース
                                    $parts = $lineStr -split '\s{2,}|\s+\|'
                                    
                                    if ($parts.Count -ge 3) {
                                        $formatId = $parts[0].Trim()
                                        
                                        # storyboardをスキップ
                                        if ($formatId -match '^sb\d+') {
                                            continue
                                        }
                                        
                                        $ext = $parts[1].Trim()
                                        $resolution = if ($parts[2]) { $parts[2].Trim() } else { "" }
                                        
                                        # 残りの部分から情報を抽出
                                        $remainingText = $lineStr
                                        
                                        $vcodec = "none"
                                        $acodec = "none"
                                        $tbr = ""
                                        $filesize = ""
                                        
                                        # コーデック情報を抽出
                                        if ($remainingText -match '\|\s*([a-z0-9\.]+)\s*\|\s*([a-z0-9\.]+)') {
                                            $vcodec = $matches[1]
                                            $acodec = $matches[2]
                                        }
                                        
                                        # ビットレートとファイルサイズを抽出
                                        if ($remainingText -match '(\d+\.?\d*k)') { $tbr = $matches[1] }
                                        if ($remainingText -match '~?\s*(\d+\.?\d*[KMG]iB)') { $filesize = $matches[1] }
                                        
                                        $formatInfo = @{
                                            ID = $formatId
                                            Ext = $ext
                                            Resolution = $resolution
                                            VCodec = $vcodec
                                            ACodec = $acodec
                                            TBR = $tbr
                                            FileSize = $filesize
                                        }
                                        
                                        # 映像または音声として分類
                                        if ($resolution -match '\d+x\d+' -or $resolution -match '\d+p') {
                                            $videoFormats += $formatInfo
                                        } elseif ($resolution -match 'audio' -or $acodec -ne "none") {
                                            $audioFormats += $formatInfo
                                        }
                                    }
                                }
                            }
                        }
                        
                        # フォーマット情報が取得できた場合のみヘッダーと内容を表示
                        if ($videoFormats.Count -gt 0 -or $audioFormats.Count -gt 0) {
                            Write-OutputBox("")
                            $resultContent += "`r`n"
                            Write-OutputBox("--- 利用可能なコーデック一覧 ---")
                            $resultContent += "--- 利用可能なコーデック一覧 ---`r`n"
                        
                            # 映像フォーマットを表示
                            if ($videoFormats.Count -gt 0) {
                                Write-OutputBox("")
                                $resultContent += "`r`n"
                                Write-OutputBox("【映像フォーマット】")
                                $resultContent += "【映像フォーマット】`r`n"
                                $videoFormats | ForEach-Object {
                                    $line = "  ID: $($_.ID) | 拡張子: $($_.Ext)"
                                    if ($_.Resolution) { $line += " | 解像度: $($_.Resolution)" }
                                    if ($_.VCodec -and $_.VCodec -ne "none") { $line += " | Vコーデック: $($_.VCodec)" }
                                    if ($_.ACodec -and $_.ACodec -ne "none") { $line += " | Aコーデック: $($_.ACodec)" }
                                    if ($_.TBR) { $line += " | $($_.TBR)bps" }
                                    if ($_.FileSize) { $line += " | $($_.FileSize)" }
                                    Write-OutputBox($line)
                                    $resultContent += $line + "`r`n"
                                }
                            }
                            
                            # 音声フォーマットを表示
                            if ($audioFormats.Count -gt 0) {
                                Write-OutputBox("")
                                $resultContent += "`r`n"
                                Write-OutputBox("【音声フォーマット】")
                                $resultContent += "【音声フォーマット】`r`n"
                                $audioFormats | ForEach-Object {
                                    $line = "  ID: $($_.ID) | 拡張子: $($_.Ext)"
                                    if ($_.ACodec -and $_.ACodec -ne "none") { $line += " | コーデック: $($_.ACodec)" }
                                    if ($_.TBR) { $line += " | $($_.TBR)bps" }
                                    if ($_.FileSize) { $line += " | $($_.FileSize)" }
                                    Write-OutputBox($line)
                                    $resultContent += $line + "`r`n"
                                }
                            }
                        }
                    }
                    }

                    # yt-dlpの出力を適切に処理
                    $target = & $script:ytDlpPath -f best -g "$input" 2>$null | Select-Object -First 1
                    if (-not $target) { Write-OutputBox("") }
                } else { 
                    Write-OutputBox("yt-dlp で情報取得に失敗")
                    $resultContent += "yt-dlp で情報取得に失敗`r`n"
                }
            } else {
                # ローカルファイルの場合
                if (-not (Test-Path -LiteralPath $input)) {
                    Write-OutputBox("ファイルが存在しません: $input")
                    $resultContent += "ファイルが存在しません: $input`r`n"
                    continue
                }
                
                # ディレクトリの場合は中のファイルを取得
                if (Test-Path -LiteralPath $input -PathType Container) {
                    Write-OutputBox("フォルダ内のファイルを解析します。")
                    $resultContent += "フォルダ内のファイルを解析します。`r`n"
                    
                    if ($script:includeSubfolders) {
                        Write-OutputBox("(サブフォルダを含む)")
                        $resultContent += "(サブフォルダを含む)`r`n"
                    }
                    
                    # ファイル拡張子リストを作成
                    $videoExtensions = @('*.mp4', '*.mkv', '*.avi', '*.mov', '*.wmv', '*.flv', '*.webm', '*.m4v', '*.ts', '*.m2ts')
                    $audioExtensions = @('*.mp3', '*.flac', '*.wav', '*.m4a', '*.aac', '*.ogg', '*.opus', '*.wma', '*.ape', '*.tak', '*.tta', '*.dsd', '*.dsf', '*.dff')
                    
                    $extensions = $videoExtensions
                    if ($script:includeAudioFiles) {
                        $extensions = $videoExtensions + $audioExtensions
                    }
                    
                    if ($script:includeSubfolders) {
                        $files = Get-ChildItem -LiteralPath $input -File -Recurse | Where-Object {
                            $ext = $_.Extension.ToLower()
                            $extensions | ForEach-Object { $_ -replace '\*', '' } | Where-Object { $ext -eq $_ }
                        } | Sort-Object FullName
                    } else {
                        $files = Get-ChildItem -LiteralPath $input -File | Where-Object {
                            $ext = $_.Extension.ToLower()
                            $extensions | ForEach-Object { $_ -replace '\*', '' } | Where-Object { $ext -eq $_ }
                        } | Sort-Object Name
                    }
                    
                    if ($files.Count -eq 0) {
                        $fileTypeMsg = if ($script:includeAudioFiles) { "動画または音声ファイル" } else { "動画ファイル" }
                        Write-OutputBox("⚠ フォルダ内に${fileTypeMsg}が見つかりませんでした。")
                        $resultContent += "⚠ フォルダ内に${fileTypeMsg}が見つかりませんでした。`r`n"
                        continue
                    }
                    
                    Write-OutputBox("見つかったファイル数: $($files.Count)")
                    $resultContent += "見つかったファイル数: $($files.Count)`r`n"
                    Write-OutputBox("")
                    $resultContent += "`r`n"
                    
                    # 各ファイルを解析
                    $fileIndex = 0
                    foreach ($file in $files) {
                        $fileIndex++
                        $target = $file.FullName
                        $resultTitle = $file.Name
                        $resultContent = ""
                        
                        # サブフォルダを含む場合は相対パスを表示
                        if ($script:includeSubfolders) {
                            $relativePath = $file.FullName.Substring($input.Length).TrimStart('\')
                            Write-OutputBox("[$fileIndex/$($files.Count)] $relativePath")
                            $resultContent += "[$fileIndex/$($files.Count)] $relativePath`r`n"
                        } else {
                            Write-OutputBox("[$fileIndex/$($files.Count)] $($file.Name)")
                            $resultContent += "[$fileIndex/$($files.Count)] $($file.Name)`r`n"
                        }
                        
                        Set-Progress(50 + [math]::Round(($count - 1 + $fileIndex / $files.Count) / $total * 50))
                        $mediaInfoOutput = Invoke-MediaInfo "$target"
                        
                        if ($mediaInfoOutput) {
                            $parsedInfo = Parse-MediaInfo $mediaInfoOutput
                            Display-MediaInfo $parsedInfo ([ref]$resultContent)
                        } else {
                            Write-OutputBox("  ⚠ MediaInfo で情報を取得できませんでした。")
                            $resultContent += "  ⚠ MediaInfo で情報を取得できませんでした。`r`n"
                        }
                        
                        Write-OutputBox("")
                        $resultContent += "`r`n"
                        
                        # フォルダ解析結果を保存
                        $script:analysisResults += @{
                            Title = $resultTitle
                            Content = $resultContent
                            FullPath = $target
                        }
                    }
                    
                    continue
                }
                
                Write-OutputBox("ローカルファイルとして解析します。")
                $resultContent += "ローカルファイルとして解析します。`r`n"
                $target = $input
                $resultTitle = [System.IO.Path]::GetFileName($input)
            }

            if ($target) {
                Set-Progress(50 + [math]::Round($count / $total * 50))
                
                # URLの場合はMediaInfoをスキップ
                if (-not $isUrl) {
                    $mediaInfoOutput = Invoke-MediaInfo "$target"
                } else {
                    $mediaInfoOutput = $null
                }

                if ($mediaInfoOutput) {
                    Write-OutputBox("--- 詳細情報 ---")
                    $resultContent += "--- 詳細情報 ---`r`n"
                    
                    $parsedInfo = Parse-MediaInfo $mediaInfoOutput
                    Display-MediaInfo $parsedInfo ([ref]$resultContent)
                    
                } elseif (-not $isUrl) {
                    # ローカルファイルでMediaInfoが失敗した場合のみエラー表示
                    Write-OutputBox("⚠ MediaInfo で情報を取得できませんでした。")
                    $resultContent += "⚠ MediaInfo で情報を取得できませんでした。`r`n"
                }
            }

        } catch { 
            Write-OutputBox("エラー: $_")
            $resultContent += "エラー: $_`r`n"
        }

        Write-OutputBox("")
        $resultContent += "`r`n"
        Write-OutputBox("解析完了。`r`n")
        $resultContent += "解析完了。`r`n"
        
        # 結果を配列に追加
        $script:analysisResults += @{
            Title = $resultTitle
            Content = $resultContent
            FullPath = $target
        }
    }

    Set-Progress(100)
    Write-OutputBox("=== 全ファイル解析完了 ===")
    
    # 成功した入力のみ履歴に保存
    $successfulInputs = @()
    $processedFolders = @{}
    
    foreach ($result in $script:analysisResults) {
        if ($result.Content -notmatch "ファイルが存在しません|yt-dlp で情報取得に失敗|MediaInfo で情報を取得できませんでした|エラー:") {
            if ($result.ContainsKey('FullPath') -and $result.FullPath) {
                # フォルダ解析の場合はフォルダパスを取得
                $filePath = $result.FullPath
                $parentFolder = Split-Path -Parent $filePath
                
                # 元の入力がフォルダかどうかを確認
                $isFromFolder = $false
                foreach ($inputRaw in $inputs) {
                    $cleanInput = $inputRaw.Trim()
                    if ((Test-Path -LiteralPath $cleanInput -PathType Container) -and 
                        $filePath.StartsWith($cleanInput, [StringComparison]::OrdinalIgnoreCase)) {
                        # フォルダ解析の場合
                        if (-not $processedFolders.ContainsKey($cleanInput)) {
                            $successfulInputs += $cleanInput
                            $processedFolders[$cleanInput] = $true
                        }
                        $isFromFolder = $true
                        break
                    }
                }
                
                # フォルダ解析ではない場合は個別ファイルとして保存
                if (-not $isFromFolder) {
                    $successfulInputs += $filePath
                }
            } else {
                $successfulInputs += $result.Title
            }
        }
    }
    
    if ($successfulInputs.Count -gt 0) {
        $currentHistory = Load-History
        $newHistory = @($successfulInputs) + $currentHistory
        Save-History $newHistory
    }
    
    # 解析完了後にボタンを有効化
    $showWindowButton.Enabled = $true
}

$button.Add_Click({ Analyze-Video })

# キーボードショートカット
function Parse-Shortcut($shortcutString) {
    $parts = $shortcutString -split '\+'
    $modifiers = @{
        Control = $false
        Alt = $false
        Shift = $false
    }
    $key = ""
    
    foreach ($part in $parts) {
        $part = $part.Trim()
        switch ($part) {
            "Ctrl" { $modifiers.Control = $true }
            "Control" { $modifiers.Control = $true }
            "Alt" { $modifiers.Alt = $true }
            "Shift" { $modifiers.Shift = $true }
            default { $key = $part }
        }
    }
    
    return @{
        Control = $modifiers.Control
        Alt = $modifiers.Alt
        Shift = $modifiers.Shift
        Key = $key
    }
}

$form.Add_KeyDown({
    param($sender, $e)
    
    $searchShortcutParsed = Parse-Shortcut $script:searchShortcut
    if ($e.Control -eq $searchShortcutParsed.Control -and
        $e.Alt -eq $searchShortcutParsed.Alt -and
        $e.Shift -eq $searchShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $searchShortcutParsed.Key) {
        $e.Handled = $true
        $searchItem.PerformClick()
        return
    }
    
    $findNextShortcutParsed = Parse-Shortcut $script:findNextShortcut
    if ($e.Control -eq $findNextShortcutParsed.Control -and
        $e.Alt -eq $findNextShortcutParsed.Alt -and
        $e.Shift -eq $findNextShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $findNextShortcutParsed.Key) {
        $e.Handled = $true
        $findNextItem.PerformClick()
        return
    }
    
    $findPreviousShortcutParsed = Parse-Shortcut $script:findPreviousShortcut
    if ($e.Control -eq $findPreviousShortcutParsed.Control -and
        $e.Alt -eq $findPreviousShortcutParsed.Alt -and
        $e.Shift -eq $findPreviousShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $findPreviousShortcutParsed.Key) {
        $e.Handled = $true
        $findPreviousItem.PerformClick()
        return
    }
    
    $clearHighlightShortcutParsed = Parse-Shortcut $script:clearHighlightShortcut
    if ($e.Control -eq $clearHighlightShortcutParsed.Control -and
        $e.Alt -eq $clearHighlightShortcutParsed.Alt -and
        $e.Shift -eq $clearHighlightShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $clearHighlightShortcutParsed.Key) {
        $e.Handled = $true
        $clearHighlightItem.PerformClick()
        return
    }
    
    $newShortcutParsed = Parse-Shortcut $script:newShortcut
    if ($e.Control -eq $newShortcutParsed.Control -and
        $e.Alt -eq $newShortcutParsed.Alt -and
        $e.Shift -eq $newShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $newShortcutParsed.Key) {
        $e.Handled = $true
        $newItem.PerformClick()
        return
    }
    
    $openFileShortcutParsed = Parse-Shortcut $script:openFileShortcut
    if ($e.Control -eq $openFileShortcutParsed.Control -and
        $e.Alt -eq $openFileShortcutParsed.Alt -and
        $e.Shift -eq $openFileShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $openFileShortcutParsed.Key) {
        $e.Handled = $true
        $addFileItem.PerformClick()
        return
    }
    
    $analyzeShortcutParsed = Parse-Shortcut $script:analyzeShortcut
    if ($e.Control -eq $analyzeShortcutParsed.Control -and
        $e.Alt -eq $analyzeShortcutParsed.Alt -and
        $e.Shift -eq $analyzeShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $analyzeShortcutParsed.Key) {
        $e.Handled = $true
        $button.PerformClick()
        return
    }
    
    $showWindowShortcutParsed = Parse-Shortcut $script:showWindowShortcut
    if ($e.Control -eq $showWindowShortcutParsed.Control -and
        $e.Alt -eq $showWindowShortcutParsed.Alt -and
        $e.Shift -eq $showWindowShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $showWindowShortcutParsed.Key) {
        $e.Handled = $true
        if ($showWindowButton.Enabled) {
            $showWindowButton.PerformClick()
        }
        return
    }
    
    $closeAllWindowsShortcutParsed = Parse-Shortcut $script:closeAllWindowsShortcut
    if ($e.Control -eq $closeAllWindowsShortcutParsed.Control -and
        $e.Alt -eq $closeAllWindowsShortcutParsed.Alt -and
        $e.Shift -eq $closeAllWindowsShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $closeAllWindowsShortcutParsed.Key) {
        $e.Handled = $true
        if ($closeAllWindowsButton.Enabled) {
            $closeAllWindowsButton.PerformClick()
        }
        return
    }
    
    $optionsShortcutParsed = Parse-Shortcut $script:optionsShortcut
    if ($e.Control -eq $optionsShortcutParsed.Control -and
        $e.Alt -eq $optionsShortcutParsed.Alt -and
        $e.Shift -eq $optionsShortcutParsed.Shift -and
        $e.KeyCode.ToString() -eq $optionsShortcutParsed.Key) {
        $e.Handled = $true
        $optionsItem.PerformClick()
        return
    }
})

$form.Add_FormClosing({
    if ($script:rememberWindowPosition) {
        $script:windowX = $form.Location.X
        $script:windowY = $form.Location.Y
        Save-Config
    }
})

[void]$form.ShowDialog()