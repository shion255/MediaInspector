Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
chcp 65001 > $null
$ErrorActionPreference = "Stop"

# PowerShell ウィンドウを最小化
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 6) | Out-Null  # 6 = SW_MINIMIZE

# コマンドライン引数を取得
$droppedFiles = $args

# 設定ファイルのパス
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configDir = Join-Path $scriptDir "ini"
$configFile = Join-Path $configDir "MediaInspector.ini"
$historyFile = Join-Path $configDir "MediaInspector_history.txt"

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
        YtDlpPath = "C:\encode\tools\yt-dlp.exe"
        MediaInfoPath = "C:\DTV\tools\MediaInfo_CLI\MediaInfo.exe"
        IncludeSubfolders = $false
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
    # 最新20件まで保存
    $maxHistory = 20
    $uniqueItems = $items | Select-Object -Unique | Select-Object -First $maxHistory
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

# --- ツールパスチェック ---
foreach ($tool in @($script:ytDlpPath, $script:mediaInfoPath)) {
    if (-not (Test-Path $tool)) {
        [System.Windows.Forms.MessageBox]::Show("$tool が見つかりません。設定でパスを確認してください。")
        # ここでは終了せずに続行（設定で修正可能なため）
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

# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "MediaInspector"
$form.Size = New-Object System.Drawing.Size(850, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = $script:bgColor
$form.ForeColor = $script:fgColor
$form.Font = New-Object System.Drawing.Font("Meiryo UI", 9)
$form.Opacity = $script:windowOpacity

# メニューストリップを作成
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = $script:menuBgColor
$menuStrip.ForeColor = $script:fgColor

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
            $textBox.Focus()
        }
    } else {
        $textBox.Clear()
        $outputBox.Clear()
        $script:analysisResults = @()
        $showWindowButton.Enabled = $false
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

# 解析結果から絞り込み
$filterItem = New-Object System.Windows.Forms.ToolStripMenuItem
$filterItem.Text = "解析結果から絞り込み(&F)..."
$filterItem.Add_Click({
    Show-FilterDialog
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
    $optionsForm.Size = New-Object System.Drawing.Size(450, 460)
    $optionsForm.StartPosition = "CenterParent"
    $optionsForm.FormBorderStyle = "FixedDialog"
    $optionsForm.MaximizeBox = $false
    $optionsForm.MinimizeBox = $false
    $optionsForm.BackColor = $script:bgColor
    $optionsForm.ForeColor = $script:fgColor
    
    # テーマ設定
    $themeLabel = New-Object System.Windows.Forms.Label
    $themeLabel.Text = "テーマ："
    $themeLabel.Location = New-Object System.Drawing.Point(20, 20)
    $themeLabel.Size = New-Object System.Drawing.Size(100, 20)
    $themeLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($themeLabel)
    
    $themeCombo = New-Object System.Windows.Forms.ComboBox
    $themeCombo.Location = New-Object System.Drawing.Point(130, 18)
    $themeCombo.Size = New-Object System.Drawing.Size(280, 25)
    $themeCombo.DropDownStyle = "DropDownList"
    $themeCombo.Items.AddRange(@("ダークテーマ", "ライトテーマ"))
    $themeCombo.SelectedIndex = if ($script:currentTheme -eq "Dark") { 0 } else { 1 }
    $optionsForm.Controls.Add($themeCombo)
    
    # フォント名設定
    $fontNameLabel = New-Object System.Windows.Forms.Label
    $fontNameLabel.Text = "フォント名："
    $fontNameLabel.Location = New-Object System.Drawing.Point(20, 60)
    $fontNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    $fontNameLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($fontNameLabel)

    # インストール済みフォントを取得
    $installedFonts = New-Object System.Drawing.Text.InstalledFontCollection
    $fontFamilies = $installedFonts.Families | Where-Object { $_.IsStyleAvailable([System.Drawing.FontStyle]::Regular) } | Sort-Object Name

    $fontNameCombo = New-Object System.Windows.Forms.ComboBox
    $fontNameCombo.Location = New-Object System.Drawing.Point(130, 58)
    $fontNameCombo.Size = New-Object System.Drawing.Size(280, 25)
    $fontNameCombo.DropDownStyle = "DropDown"  # 編集可能に変更
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

    $optionsForm.Controls.Add($fontNameCombo)
    
    # フォントサイズ設定
    $fontSizeLabel = New-Object System.Windows.Forms.Label
    $fontSizeLabel.Text = "フォントサイズ："
    $fontSizeLabel.Location = New-Object System.Drawing.Point(20, 100)
    $fontSizeLabel.Size = New-Object System.Drawing.Size(100, 20)
    $fontSizeLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($fontSizeLabel)
    
    $fontSizeCombo = New-Object System.Windows.Forms.ComboBox
    $fontSizeCombo.Location = New-Object System.Drawing.Point(130, 98)
    $fontSizeCombo.Size = New-Object System.Drawing.Size(280, 25)
    $fontSizeCombo.DropDownStyle = "DropDownList"
    $fontSizeCombo.Items.AddRange(@("8", "9", "10", "11", "12", "14", "16"))
    $fontSizeCombo.SelectedItem = $script:currentFontSize.ToString()
    $optionsForm.Controls.Add($fontSizeCombo)
    
    # yt-dlpパス設定
    $ytDlpLabel = New-Object System.Windows.Forms.Label
    $ytDlpLabel.Text = "yt-dlp パス："
    $ytDlpLabel.Location = New-Object System.Drawing.Point(20, 140)
    $ytDlpLabel.Size = New-Object System.Drawing.Size(100, 20)
    $ytDlpLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($ytDlpLabel)
    
    $ytDlpTextBox = New-Object System.Windows.Forms.TextBox
    $ytDlpTextBox.Location = New-Object System.Drawing.Point(130, 138)
    $ytDlpTextBox.Size = New-Object System.Drawing.Size(245, 25)
    $ytDlpTextBox.Text = $script:ytDlpPath
    $ytDlpTextBox.BackColor = $script:inputBgColor
    $ytDlpTextBox.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($ytDlpTextBox)
    
    $ytDlpBrowseButton = New-Object System.Windows.Forms.Button
    $ytDlpBrowseButton.Text = "参照"
    $ytDlpBrowseButton.Location = New-Object System.Drawing.Point(380, 138)
    $ytDlpBrowseButton.Size = New-Object System.Drawing.Size(50, 25)
    $ytDlpBrowseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*"
        $openFileDialog.Title = "yt-dlp のパスを選択"
        $openFileDialog.InitialDirectory = Split-Path -Parent $script:ytDlpPath
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $ytDlpTextBox.Text = $openFileDialog.FileName
        }
    })
    $optionsForm.Controls.Add($ytDlpBrowseButton)
    
    # MediaInfoパス設定
    $mediaInfoLabel = New-Object System.Windows.Forms.Label
    $mediaInfoLabel.Text = "MediaInfo パス："
    $mediaInfoLabel.Location = New-Object System.Drawing.Point(20, 180)
    $mediaInfoLabel.Size = New-Object System.Drawing.Size(100, 20)
    $mediaInfoLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($mediaInfoLabel)
    
    $mediaInfoTextBox = New-Object System.Windows.Forms.TextBox
    $mediaInfoTextBox.Location = New-Object System.Drawing.Point(130, 178)
    $mediaInfoTextBox.Size = New-Object System.Drawing.Size(245, 25)
    $mediaInfoTextBox.Text = $script:mediaInfoPath
    $mediaInfoTextBox.BackColor = $script:inputBgColor
    $mediaInfoTextBox.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($mediaInfoTextBox)
    
    $mediaInfoBrowseButton = New-Object System.Windows.Forms.Button
    $mediaInfoBrowseButton.Text = "参照"
    $mediaInfoBrowseButton.Location = New-Object System.Drawing.Point(380, 178)
    $mediaInfoBrowseButton.Size = New-Object System.Drawing.Size(50, 25)
    $mediaInfoBrowseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*"
        $openFileDialog.Title = "MediaInfo のパスを選択"
        $openFileDialog.InitialDirectory = Split-Path -Parent $script:mediaInfoPath
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $mediaInfoTextBox.Text = $openFileDialog.FileName
        }
    })
    $optionsForm.Controls.Add($mediaInfoBrowseButton)
    
    # 透明度設定
    $opacityLabel = New-Object System.Windows.Forms.Label
    $opacityLabel.Text = "ウィンドウの透明度："
    $opacityLabel.Location = New-Object System.Drawing.Point(20, 220)
    $opacityLabel.Size = New-Object System.Drawing.Size(150, 20)
    $opacityLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($opacityLabel)
    
    $opacityTrackBar = New-Object System.Windows.Forms.TrackBar
    $opacityTrackBar.Location = New-Object System.Drawing.Point(20, 245)
    $opacityTrackBar.Size = New-Object System.Drawing.Size(300, 45)
    $opacityTrackBar.Minimum = 50
    $opacityTrackBar.Maximum = 100
    $opacityTrackBar.TickFrequency = 10
    $opacityTrackBar.Value = [int]($script:windowOpacity * 100)
    $optionsForm.Controls.Add($opacityTrackBar)
    
    $opacityValueLabel = New-Object System.Windows.Forms.Label
    $opacityValueLabel.Location = New-Object System.Drawing.Point(330, 250)
    $opacityValueLabel.Size = New-Object System.Drawing.Size(80, 20)
    $opacityValueLabel.Text = "$([int]($script:windowOpacity * 100))%"
    $opacityValueLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($opacityValueLabel)
    
    $opacityTrackBar.Add_ValueChanged({
        $opacityValueLabel.Text = "$($opacityTrackBar.Value)%"
        $form.Opacity = $opacityTrackBar.Value / 100.0
    })
    
    # サブフォルダを含めるオプション
    $includeSubfoldersCheckBox = New-Object System.Windows.Forms.CheckBox
    $includeSubfoldersCheckBox.Text = "フォルダ解析時にサブフォルダを含める"
    $includeSubfoldersCheckBox.Location = New-Object System.Drawing.Point(20, 290)
    $includeSubfoldersCheckBox.Size = New-Object System.Drawing.Size(400, 25)
    $includeSubfoldersCheckBox.Checked = $script:includeSubfolders
    $includeSubfoldersCheckBox.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($includeSubfoldersCheckBox)
    
    # OKボタン
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(240, 370)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
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
        $script:ytDlpPath = $ytDlpTextBox.Text.Trim()
        $script:mediaInfoPath = $mediaInfoTextBox.Text.Trim()
        
        # 透明度適用
        $script:windowOpacity = $opacityTrackBar.Value / 100.0
        $form.Opacity = $script:windowOpacity
        
        # サブフォルダオプション適用
        $script:includeSubfolders = $includeSubfoldersCheckBox.Checked
        
        # 設定を保存
        Save-Config
        
        # ツールの存在確認
        foreach ($tool in @($script:ytDlpPath, $script:mediaInfoPath)) {
            if (-not (Test-Path $tool)) {
                [System.Windows.Forms.MessageBox]::Show("$tool が見つかりません。パスを確認してください。")
            }
        }
    })
    $optionsForm.Controls.Add($okButton)
    
    # キャンセルボタン
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "キャンセル"
    $cancelButton.Location = New-Object System.Drawing.Point(330, 370)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
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

$menuStrip.Items.Add($helpMenu)

$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

$label = New-Object System.Windows.Forms.Label
$label.Text = "URL または ローカルファイル（複数可、改行で区切る / ドラッグ&ドロップ可）："
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
        $newFiles = $files -join "`r`n"
        
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

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font($script:currentFontName, $script:currentFontSize)
$outputBox.Location = New-Object System.Drawing.Point(10, 190)
$outputBox.Size = New-Object System.Drawing.Size(810, 360)
$outputBox.BackColor = $script:outputBgColor
$outputBox.ForeColor = $script:fgColor
$outputBox.Anchor = "Top,Bottom,Left,Right"
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
    
    # 画面サイズを取得
    $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
    $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height
    
    $totalCount = $script:analysisResults.Count
    
    # 2段組みで表示する場合の計算
    if ($totalCount -le 5) {
        # 5個以下：1段のみ
        $row1Count = $totalCount
        $row2Count = 0
        $windowHeight = [int]($screenHeight - 50)
    } else {
        # 6個以上：2段組み
        $row1Count = [Math]::Min($totalCount, 5)
        $row2Count = $totalCount - $row1Count
        $windowHeight = [int](($screenHeight - 60) / 2)
    }
    
    $currentIndex = 0
    
    # 1段目を表示
    if ($row1Count -gt 0) {
        $windowWidth = [int]([Math]::Floor($screenWidth / $row1Count) - 10)
        $xOffset = 0
        $yOffset = 20
        
        for ($i = 0; $i -lt $row1Count; $i++) {
            $result = $script:analysisResults[$currentIndex]
            
            $resultForm = New-Object System.Windows.Forms.Form
            $resultForm.Text = "解析結果: $($result.Title)"
            $resultForm.Size = New-Object System.Drawing.Size([int]$windowWidth, [int]$windowHeight)
            $resultForm.StartPosition = "Manual"
            $resultForm.Location = New-Object System.Drawing.Point([int]$xOffset, [int]$yOffset)
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
            
            $xOffset = [int]($xOffset + $windowWidth + 5)
            $currentIndex++
        }
    }
    
    # 2段目を表示
    if ($row2Count -gt 0) {
        $windowWidth = [int]([Math]::Floor($screenWidth / $row2Count) - 10)
        $xOffset = 0
        $yOffset = [int](20 + $windowHeight + 10)
        
        for ($i = 0; $i -lt $row2Count; $i++) {
            $result = $script:analysisResults[$currentIndex]
            
            $resultForm = New-Object System.Windows.Forms.Form
            $resultForm.Text = "解析結果: $($result.Title)"
            $resultForm.Size = New-Object System.Drawing.Size([int]$windowWidth, [int]$windowHeight)
            $resultForm.StartPosition = "Manual"
            $resultForm.Location = New-Object System.Drawing.Point([int]$xOffset, [int]$yOffset)
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
            
            $xOffset = [int]($xOffset + $windowWidth + 5)
            $currentIndex++
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

# --- 解析結果絞り込み機能 ---
function Show-FilterDialog {
    if ($script:analysisResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("解析結果がありません。先に解析を実行してください。", "情報")
        return
    }
    
    # 利用可能なコーデックを収集
    $videoCodecs = @{}
    $audioCodecs = @{}
    
    foreach ($result in $script:analysisResults) {
        $content = $result.Content
        
        # 映像コーデックを抽出
        if ($content -match '映像\d+:\s*([^\s]+)') {
            $codec = $matches[1]
            if (-not $videoCodecs.ContainsKey($codec)) {
                $videoCodecs[$codec] = 0
            }
            $videoCodecs[$codec]++
        }
        
        # 音声コーデックを抽出
        if ($content -match '音声\d+:\s*([^\s]+)') {
            $codec = $matches[1]
            if (-not $audioCodecs.ContainsKey($codec)) {
                $audioCodecs[$codec] = 0
            }
            $audioCodecs[$codec]++
        }
    }
    
    if ($videoCodecs.Count -eq 0 -and $audioCodecs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("コーデック情報が見つかりませんでした。", "情報")
        return
    }
    
    # 絞り込みダイアログ
    $filterForm = New-Object System.Windows.Forms.Form
    $filterForm.Text = "解析結果から絞り込み"
    $filterForm.Size = New-Object System.Drawing.Size(600, 500)
    $filterForm.StartPosition = "CenterParent"
    $filterForm.BackColor = $script:bgColor
    $filterForm.ForeColor = $script:fgColor
    
    # 映像コーデックグループ
    $videoGroupBox = New-Object System.Windows.Forms.GroupBox
    $videoGroupBox.Text = "映像コーデック"
    $videoGroupBox.Location = New-Object System.Drawing.Point(20, 20)
    $videoGroupBox.Size = New-Object System.Drawing.Size(260, 350)
    $videoGroupBox.ForeColor = $script:fgColor
    $filterForm.Controls.Add($videoGroupBox)
    
    $videoPanel = New-Object System.Windows.Forms.Panel
    $videoPanel.Location = New-Object System.Drawing.Point(10, 25)
    $videoPanel.Size = New-Object System.Drawing.Size(240, 315)
    $videoPanel.AutoScroll = $true
    $videoGroupBox.Controls.Add($videoPanel)
    
    $videoCheckBoxes = @{}
    $yPos = 5
    foreach ($codec in ($videoCodecs.Keys | Sort-Object)) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = "$codec ($($videoCodecs[$codec]))"
        $checkBox.Location = New-Object System.Drawing.Point(5, $yPos)
        $checkBox.Size = New-Object System.Drawing.Size(220, 25)
        $checkBox.ForeColor = $script:fgColor
        $videoPanel.Controls.Add($checkBox)
        $videoCheckBoxes[$codec] = $checkBox
        $yPos += 30
    }
    
    # 音声コーデックグループ
    $audioGroupBox = New-Object System.Windows.Forms.GroupBox
    $audioGroupBox.Text = "音声コーデック"
    $audioGroupBox.Location = New-Object System.Drawing.Point(300, 20)
    $audioGroupBox.Size = New-Object System.Drawing.Size(260, 350)
    $audioGroupBox.ForeColor = $script:fgColor
    $filterForm.Controls.Add($audioGroupBox)
    
    $audioPanel = New-Object System.Windows.Forms.Panel
    $audioPanel.Location = New-Object System.Drawing.Point(10, 25)
    $audioPanel.Size = New-Object System.Drawing.Size(240, 315)
    $audioPanel.AutoScroll = $true
    $audioGroupBox.Controls.Add($audioPanel)
    
    $audioCheckBoxes = @{}
    $yPos = 5
    foreach ($codec in ($audioCodecs.Keys | Sort-Object)) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = "$codec ($($audioCodecs[$codec]))"
        $checkBox.Location = New-Object System.Drawing.Point(5, $yPos)
        $checkBox.Size = New-Object System.Drawing.Size(220, 25)
        $checkBox.ForeColor = $script:fgColor
        $audioPanel.Controls.Add($checkBox)
        $audioCheckBoxes[$codec] = $checkBox
        $yPos += 30
    }
    
    # 検索ボタン
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "検索"
    $searchButton.Location = New-Object System.Drawing.Point(200, 390)
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
        
        if ($selectedVideoCodecs.Count -eq 0 -and $selectedAudioCodecs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("少なくとも1つのコーデックを選択してください。", "情報")
            return
        }
        
        Show-FilteredResults $selectedVideoCodecs $selectedAudioCodecs
        $filterForm.Close()
    })
    $filterForm.Controls.Add($searchButton)
    
    # キャンセルボタン
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "キャンセル"
    $cancelButton.Location = New-Object System.Drawing.Point(310, 390)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $cancelButton.ForeColor = $script:fgColor
    $cancelButton.Add_Click({
        $filterForm.Close()
    })
    $filterForm.Controls.Add($cancelButton)
    
    [void]$filterForm.ShowDialog($form)
}

function Show-FilteredResults($videoCodecs, $audioCodecs) {
    # 絞り込み実行
    $filteredResults = @()
    
    foreach ($result in $script:analysisResults) {
        $content = $result.Content
        $match = $false
        
        # 映像コーデックチェック
        if ($videoCodecs.Count -gt 0) {
            foreach ($codec in $videoCodecs) {
                if ($content -match "映像\d+:\s*$([regex]::Escape($codec))") {
                    $match = $true
                    break
                }
            }
        }
        
        # 音声コーデックチェック
        if (-not $match -and $audioCodecs.Count -gt 0) {
            foreach ($codec in $audioCodecs) {
                if ($content -match "音声\d+:\s*$([regex]::Escape($codec))") {
                    $match = $true
                    break
                }
            }
        }
        
        # OR条件: 映像または音声のいずれかに一致
        if ($videoCodecs.Count -gt 0 -and $audioCodecs.Count -gt 0) {
            $match = $false
            $videoMatch = $false
            $audioMatch = $false
            
            foreach ($codec in $videoCodecs) {
                if ($content -match "映像\d+:\s*$([regex]::Escape($codec))") {
                    $videoMatch = $true
                    break
                }
            }
            
            foreach ($codec in $audioCodecs) {
                if ($content -match "音声\d+:\s*$([regex]::Escape($codec))") {
                    $audioMatch = $true
                    break
                }
            }
            
            $match = $videoMatch -or $audioMatch
        }
        
        if ($match) {
            $filteredResults += $result
        }
    }
    
    if ($filteredResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("条件に一致する結果が見つかりませんでした。", "情報")
        return
    }
    
    # 結果一覧を表示
    $resultForm = New-Object System.Windows.Forms.Form
    $resultForm.Text = "絞り込み結果 - $($filteredResults.Count)件"
    $resultForm.Size = New-Object System.Drawing.Size(800, 650)
    $resultForm.StartPosition = "CenterScreen"
    $resultForm.BackColor = $script:bgColor
    $resultForm.ForeColor = $script:fgColor
    
    # ListView
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 10)
    $listView.Size = New-Object System.Drawing.Size(760, 550)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.BackColor = $script:inputBgColor
    $listView.ForeColor = $script:fgColor
    $listView.Anchor = "Top,Bottom,Left,Right"
    
    # 列を追加
    [void]$listView.Columns.Add("ファイル名", 700)
    
    # データを追加
    $rowIndex = 0
    foreach ($result in $filteredResults) {
        $item = New-Object System.Windows.Forms.ListViewItem($result.Title)
        $item.Tag = $result
        
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
    
    # ダブルクリックで詳細表示
    $listView.Add_DoubleClick({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedResult = $listView.SelectedItems[0].Tag
            Show-ResultDetail $selectedResult
        }
    })
    
    $resultForm.Controls.Add($listView)
    
    # 閉じるボタン
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "閉じる"
    $closeButton.Location = New-Object System.Drawing.Point(340, 570)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $closeButton.ForeColor = $script:fgColor
    $closeButton.Anchor = "Bottom"
    $closeButton.Add_Click({
        $resultForm.Close()
    })
    $resultForm.Controls.Add($closeButton)
    
    [void]$resultForm.ShowDialog($form)
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
                    
                    Write-OutputBox("タイトル: $($infoJson.title)")
                    $resultContent += "タイトル: $($infoJson.title)`r`n"
                    Write-OutputBox("アップローダー: $($infoJson.uploader)")
                    $resultContent += "アップローダー: $($infoJson.uploader)`r`n"
                    
                    # 投稿日時を追加
                    if ($infoJson.upload_date) {
                        $dateStr = $infoJson.upload_date
                        # YYYYMMDD形式をYYYY年MM月DD日に変換
                        if ($dateStr -match '^(\d{4})(\d{2})(\d{2})$') {
                            $formattedDate = "$($matches[1])年$($matches[2])月$($matches[3])日"
                            Write-OutputBox("投稿日時: $formattedDate")
                            $resultContent += "投稿日時: $formattedDate`r`n"
                        } else {
                            Write-OutputBox("投稿日時: $dateStr")
                            $resultContent += "投稿日時: $dateStr`r`n"
                        }
                    }
                    
                    $durationText = Format-Time $infoJson.duration
                    Write-OutputBox("再生時間: $durationText")
                    $resultContent += "再生時間: $durationText`r`n"
                    
                    # チャプターの有無をチェック
                    if ($infoJson.chapters -and $infoJson.chapters.Count -gt 0) {
                        Write-OutputBox("✅ チャプターあり ($($infoJson.chapters.Count)個)")
                        $resultContent += "✅ チャプターあり ($($infoJson.chapters.Count)個)`r`n"
                    } else {
                        Write-OutputBox("❌ チャプターなし")
                        $resultContent += "❌ チャプターなし`r`n"
                    }
                    
                    $subs = $infoJson.subtitles.PSObject.Properties.Name
                    if ($subs -match "ja|jpn|Japanese") { 
                        Write-OutputBox("✅ 日本語字幕あり")
                        $resultContent += "✅ 日本語字幕あり`r`n"
                    } else { 
                        Write-OutputBox("❌ 日本語字幕なし")
                        $resultContent += "❌ 日本語字幕なし`r`n"
                    }
                    
                    Write-OutputBox("")
                    $resultContent += "`r`n"
                    Write-OutputBox("--- 利用可能なコーデック一覧 ---")
                    $resultContent += "--- 利用可能なコーデック一覧 ---`r`n"
                    
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
                        
                        if ($videoFormats.Count -eq 0 -and $audioFormats.Count -eq 0) {
                            Write-OutputBox("⚠ フォーマット情報を解析できませんでした。")
                            $resultContent += "⚠ フォーマット情報を解析できませんでした。`r`n"
                        }
                    } else {
                        Write-OutputBox("⚠ フォーマット一覧の取得に失敗しました。")
                        $resultContent += "⚠ フォーマット一覧の取得に失敗しました。`r`n"
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
                    
                    # 動画ファイルを取得
                    $videoExtensions = @('*.mp4', '*.mkv', '*.avi', '*.mov', '*.wmv', '*.flv', '*.webm', '*.m4v', '*.ts', '*.m2ts')
                    
                    if ($script:includeSubfolders) {
                        $files = Get-ChildItem -LiteralPath $input -File -Recurse | Where-Object {
                            $ext = $_.Extension.ToLower()
                            $videoExtensions | ForEach-Object { $_ -replace '\*', '' } | Where-Object { $ext -eq $_ }
                        } | Sort-Object FullName
                    } else {
                        $files = Get-ChildItem -LiteralPath $input -File | Where-Object {
                            $ext = $_.Extension.ToLower()
                            $videoExtensions | ForEach-Object { $_ -replace '\*', '' } | Where-Object { $ext -eq $_ }
                        } | Sort-Object Name
                    }
                    
                    if ($files.Count -eq 0) {
                        Write-OutputBox("⚠ フォルダ内に動画ファイルが見つかりませんでした。")
                        $resultContent += "⚠ フォルダ内に動画ファイルが見つかりませんでした。`r`n"
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
                            # MediaInfo出力をパース（以下、既存のパース処理と同じ）
                            $currentSection = ""
                            $duration = ""
                            $overallBitrate = ""
                            $fileSize = ""
                            $artist = ""
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
                            
                            # パース処理（既存のコードをそのままコピー）
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
                                        $chapterCount++
                                    }
                                    continue
                                }
                                
                                if ($currentSection -eq "Menu" -and $line -match '^Chapters') {
                                    $hasChapters = $true
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
                                        }
                                        "Video" {
                                            if ($key -eq "Format") { $videoInfo["format"] = $value }
                                            if ($key -eq "Width") { $videoInfo["width"] = $value -replace '\D', '' }
                                            if ($key -eq "Height") { $videoInfo["height"] = $value -replace '\D', '' }
                                            if ($key -eq "Frame rate") { $videoInfo["fps"] = $value }
                                            if ($key -eq "Frame rate mode") { $videoInfo["fps_mode"] = $value }
                                            if ($key -eq "Bit rate") { $videoInfo["bitrate"] = $value }
                                            if ($key -eq "Bit rate mode") { $videoInfo["bitrate_mode"] = $value }
                                            if ($key -eq "Stream size") { $videoInfo["stream_size"] = $value }
                                            if ($key -eq "Color primaries") { $videoInfo["color_primaries"] = $value }
                                            if ($key -eq "Transfer characteristics") { $videoInfo["transfer_characteristics"] = $value }
                                            if ($key -eq "Matrix coefficients") { $videoInfo["matrix_coefficients"] = $value }
                                        }
                                        "Audio" {
                                            if ($key -eq "Format") { $audioInfo["format"] = $value }
                                            if ($key -match "Channel") { $audioInfo["channels"] = $value }
                                            if ($key -eq "Sampling rate") { $audioInfo["samplerate"] = $value }
                                            if ($key -eq "Bit rate") { $audioInfo["bitrate"] = $value }
                                            if ($key -eq "Stream size") { $audioInfo["stream_size"] = $value }
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
                            
                            # 情報を表示（既存の表示処理をそのまま使用）
                            $duration = Convert-DurationToJapanese $duration
                            $overallBitrate = Format-Bitrate $overallBitrate
                            
                            Write-OutputBox("  再生時間: $duration")
                            $resultContent += "  再生時間: $duration`r`n"
                            Write-OutputBox("  ビットレート: $overallBitrate")
                            $resultContent += "  ビットレート: $overallBitrate`r`n"
                            
                            if ($artist) {
                                Write-OutputBox("  作成者: $artist")
                                $resultContent += "  作成者: $artist`r`n"
                            }
                            
                            if ($hasChapters) {
                                Write-OutputBox("  ✅ チャプターあり")
                                $resultContent += "  ✅ チャプターあり`r`n"
                            } else {
                                Write-OutputBox("  ❌ チャプターなし")
                                $resultContent += "  ❌ チャプターなし`r`n"
                            }
                            
                            # ストリーム情報表示
                            $videoIndex = 1
                            foreach ($v in $videoStreams) {
                                $format = if ($v["format"]) { $v["format"] } else { "不明" }
                                $res = if ($v["width"] -and $v["height"]) { "$($v['width'])x$($v['height'])" } else { "不明" }
                                $fpsMode = if ($v["fps_mode"]) { "[$($v['fps_mode'])]" } else { "" }
                                $fps = if ($v["fps"]) { ($v["fps"] -replace '\s*FPS', '') + " fps" } else { "不明" }
                                $hdrInfo = Get-HDRInfo $v["color_primaries"] $v["transfer_characteristics"] $v["matrix_coefficients"]
                                $bitrateMode = if ($v["bitrate_mode"]) { "[$($v['bitrate_mode'])]" } else { "" }
                                $bitrate = if ($v["bitrate"]) { (Format-Bitrate $v["bitrate"]) } else { "不明" }
                                $streamSize = if ($v["stream_size"]) { $v["stream_size"] } else { "" }
                                
                                $videoLine = "  映像${videoIndex}: $format $res | $fpsMode $fps | $hdrInfo | $bitrateMode $bitrate"
                                if ($streamSize) { $videoLine += " | $streamSize" }
                                Write-OutputBox($videoLine)
                                $resultContent += $videoLine + "`r`n"
                                $videoIndex++
                            }
                            
                            $audioIndex = 1
                            foreach ($a in $audioStreams) {
                                $format = if ($a["format"]) { $a["format"] } else { "不明" }
                                $samplerate = if ($a["samplerate"]) { $a["samplerate"] } else { "不明" }
                                $bitrate = if ($a["bitrate"]) { (Format-Bitrate $a["bitrate"]) } else { "不明" }
                                $streamSize = if ($a["stream_size"]) { $a["stream_size"] } else { "" }
                                
                                $audioLine = "  音声${audioIndex}: $format | $samplerate | $bitrate"
                                if ($streamSize) { $audioLine += " | $streamSize" }
                                Write-OutputBox($audioLine)
                                $resultContent += $audioLine + "`r`n"
                                $audioIndex++
                            }
                            
                            $imageIndex = 1
                            foreach ($img in $imageStreams) {
                                $format = if ($img["format"]) { $img["format"] } else { "不明" }
                                $res = if ($img["width"] -and $img["height"]) { "$($img['width'])x$($img['height'])" } else { "不明" }
                                $streamSize = if ($img["stream_size"]) { $img["stream_size"] } else { "" }
                                
                                $imageLine = "  カバー画像${imageIndex}: $format $res"
                                if ($streamSize) { $imageLine += " | $streamSize" }
                                Write-OutputBox($imageLine)
                                $resultContent += $imageLine + "`r`n"
                                $imageIndex++
                            }
                            
                            $textIndex = 1
                            foreach ($txt in $textStreams) {
                                $language = if ($txt["language"]) { $txt["language"] } else { "不明" }
                                $default = if ($txt["default"] -eq "Yes") { "はい" } else { "いいえ" }
                                $forced = if ($txt["forced"] -eq "Yes") { "はい" } else { "いいえ" }
                                
                                $textLine = "  テキスト${textIndex}: $language | Default - $default | Forced - $forced"
                                Write-OutputBox($textLine)
                                $resultContent += $textLine + "`r`n"
                                $textIndex++
                            }
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
                    
                    # MediaInfo出力をパース
                    $currentSection = ""
                    $duration = ""
                    $overallBitrate = ""
                    $fileSize = ""
                    $artist = ""
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
                            
                            # Menuセクションがある場合はチャプターありと判断
                            if ($currentSection -eq "Menu") {
                                $hasChapters = $true
                                $chapterCount++
                            }
                            continue
                        }
                        
                        # チャプター数をカウント（Menuセクション内のエントリ数）
                        if ($currentSection -eq "Menu" -and $line -match '^Chapters') {
                            $hasChapters = $true
                        }
                        
                        # 情報抽出
                        if ($line -match '^(.+?)\s*:\s*(.+)$') {
                            $key = $matches[1].Trim()
                            $value = $matches[2].Trim()
                            
                            switch ($currentSection) {
                                "General" {
                                    if ($key -eq "Duration") { $duration = $value }
                                    if ($key -eq "Overall bit rate") { $overallBitrate = $value }
                                    if ($key -eq "File size") { $fileSize = $value }
                                    if ($key -match "^(Performer|Artist)$") { $artist = $value }
                                }
                                "Video" {
                                    if ($key -eq "Format") { $videoInfo["format"] = $value }
                                    if ($key -eq "Width") { $videoInfo["width"] = $value -replace '\D', '' }
                                    if ($key -eq "Height") { $videoInfo["height"] = $value -replace '\D', '' }
                                    if ($key -eq "Frame rate") { $videoInfo["fps"] = $value }
                                    if ($key -eq "Frame rate mode") { $videoInfo["fps_mode"] = $value }
                                    if ($key -eq "Bit rate") { $videoInfo["bitrate"] = $value }
                                    if ($key -eq "Bit rate mode") { $videoInfo["bitrate_mode"] = $value }
                                    if ($key -eq "Stream size") { $videoInfo["stream_size"] = $value }
                                    if ($key -eq "Color primaries") { $videoInfo["color_primaries"] = $value }
                                    if ($key -eq "Transfer characteristics") { $videoInfo["transfer_characteristics"] = $value }
                                    if ($key -eq "Matrix coefficients") { $videoInfo["matrix_coefficients"] = $value }
                                }
                                "Audio" {
                                    if ($key -eq "Format") { $audioInfo["format"] = $value }
                                    if ($key -match "Channel") { $audioInfo["channels"] = $value }
                                    if ($key -eq "Sampling rate") { $audioInfo["samplerate"] = $value }
                                    if ($key -eq "Bit rate") { $audioInfo["bitrate"] = $value }
                                    if ($key -eq "Stream size") { $audioInfo["stream_size"] = $value }
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
                    
                    # 再生時間の日本語表記に変換
                    $duration = Convert-DurationToJapanese $duration
                    
                    # ビットレートのフォーマット（修正版）
                    $overallBitrate = Format-Bitrate $overallBitrate
                    
                    # 基本情報を表示
                    Write-OutputBox("再生時間: $duration")
                    $resultContent += "再生時間: $duration`r`n"
                    Write-OutputBox("ビットレート: $overallBitrate")
                    $resultContent += "ビットレート: $overallBitrate`r`n"
                    
                    # 作成者情報を表示
                    if ($artist) {
                        Write-OutputBox("作成者: $artist")
                        $resultContent += "作成者: $artist`r`n"
                    }
                    
                    # チャプター情報を表示（ローカルファイル用）
                    if ($hasChapters) {
                        Write-OutputBox("✅ チャプターあり")
                        $resultContent += "✅ チャプターあり`r`n"
                    } else {
                        Write-OutputBox("❌ チャプターなし")
                        $resultContent += "❌ チャプターなし`r`n"
                    }
                    
                    # 映像ストリーム情報（番号付き）
                    $videoIndex = 1
                    foreach ($v in $videoStreams) {
                        $format = if ($v["format"]) { $v["format"] } else { "不明" }
                        $res = if ($v["width"] -and $v["height"]) { "$($v['width'])x$($v['height'])" } else { "不明" }
                        $fpsMode = if ($v["fps_mode"]) { "[$($v['fps_mode'])]" } else { "" }
                        
                        # FPSの二重表示を修正
                        $fps = if ($v["fps"]) { 
                            ($v["fps"] -replace '\s*FPS', '') + " fps"
                        } else { 
                            "不明" 
                        }
                        
                        # HDR/SDR情報を取得
                        $hdrInfo = Get-HDRInfo $v["color_primaries"] $v["transfer_characteristics"] $v["matrix_coefficients"]
                        
                        $bitrateMode = if ($v["bitrate_mode"]) { "[$($v['bitrate_mode'])]" } else { "" }
                        $bitrate = if ($v["bitrate"]) { (Format-Bitrate $v["bitrate"]) } else { "不明" }
                        $streamSize = if ($v["stream_size"]) { $v["stream_size"] } else { "" }
                        
                        $videoLine = "映像${videoIndex}: $format $res | $fpsMode $fps | $hdrInfo | $bitrateMode $bitrate"
                        if ($streamSize) { $videoLine += " | $streamSize" }
                        Write-OutputBox($videoLine)
                        $resultContent += $videoLine + "`r`n"
                        $videoIndex++
                    }
                    
                    # 音声ストリーム情報（番号付き）
                    $audioIndex = 1
                    foreach ($a in $audioStreams) {
                        $format = if ($a["format"]) { $a["format"] } else { "不明" }
                        $samplerate = if ($a["samplerate"]) { $a["samplerate"] } else { "不明" }
                        $bitrate = if ($a["bitrate"]) { (Format-Bitrate $a["bitrate"]) } else { "不明" }
                        $streamSize = if ($a["stream_size"]) { $a["stream_size"] } else { "" }
                        
                        $audioLine = "音声${audioIndex}: $format | $samplerate | $bitrate"
                        if ($streamSize) { $audioLine += " | $streamSize" }
                        Write-OutputBox($audioLine)
                        $resultContent += $audioLine + "`r`n"
                        $audioIndex++
                    }
                    
                    # 画像ストリーム情報（番号付き）
                    $imageIndex = 1
                    foreach ($img in $imageStreams) {
                        $format = if ($img["format"]) { $img["format"] } else { "不明" }
                        $res = if ($img["width"] -and $img["height"]) { "$($img['width'])x$($img['height'])" } else { "不明" }
                        $streamSize = if ($img["stream_size"]) { $img["stream_size"] } else { "" }
                        
                        $imageLine = "カバー画像${imageIndex}: $format $res"
                        if ($streamSize) { $imageLine += " | $streamSize" }
                        Write-OutputBox($imageLine)
                        $resultContent += $imageLine + "`r`n"
                        $imageIndex++
                    }
                    
                    # テキストストリーム情報（番号付き）
                    $textIndex = 1
                    foreach ($txt in $textStreams) {
                        $language = if ($txt["language"]) { $txt["language"] } else { "不明" }
                        $default = if ($txt["default"] -eq "Yes") { "はい" } else { "いいえ" }
                        $forced = if ($txt["forced"] -eq "Yes") { "はい" } else { "いいえ" }
                        
                        $textLine = "テキスト${textIndex}: $language | Default - $default | Forced - $forced"
                        Write-OutputBox($textLine)
                        $resultContent += $textLine + "`r`n"
                        $textIndex++
                    }
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
        }
    }

    Set-Progress(100)
    Write-OutputBox("=== 全ファイル解析完了 ===")
    
    # 履歴を保存
    $currentHistory = Load-History
    $newHistory = @($inputs) + $currentHistory
    Save-History $newHistory
    
    # 解析完了後にボタンを有効化
    $showWindowButton.Enabled = $true
}

$button.Add_Click({ Analyze-Video })
[void]$form.ShowDialog()