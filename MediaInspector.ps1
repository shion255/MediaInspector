Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
chcp 65001 > $null
$ErrorActionPreference = "Stop"

# --- ツールパス ---
$yt = "C:\encode\tools\yt-dlp.exe"
$mediainfo = "C:\DTV\tools\MediaInfo_CLI\MediaInfo.exe"

foreach ($tool in @($yt, $mediainfo)) {
    if (-not (Test-Path $tool)) {
        [System.Windows.Forms.MessageBox]::Show("$tool が見つかりません。")
        exit
    }
}

# 設定ファイルのパス
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configDir = Join-Path $scriptDir "ini"
$configFile = Join-Path $configDir "MediaInspector.ini"

# 設定フォルダがなければ作成
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
}

# 設定を読み込む関数
function Load-Config {
    $config = @{
        Theme = "Dark"
        FontName = "Consolas"
        FontSize = 10
        WindowOpacity = 1.0
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

# 設定を保存する関数
function Save-Config {
    $content = @"
Theme=$($script:currentTheme)
FontName=$($script:currentFontName)
FontSize=$($script:currentFontSize)
WindowOpacity=$($script:windowOpacity)
"@
    Set-Content -Path $configFile -Value $content -Encoding UTF8
}

# 設定を読み込み
$config = Load-Config
$script:currentTheme = $config.Theme
$script:currentFontName = $config.FontName
$script:currentFontSize = [int]$config.FontSize
$script:windowOpacity = [double]$config.WindowOpacity

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

# 「ツール」メニュー
$toolMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$toolMenu.Text = "ツール(&T)"

# オプション
$optionsItem = New-Object System.Windows.Forms.ToolStripMenuItem
$optionsItem.Text = "オプション(&O)..."
$optionsItem.Add_Click({
    # オプションダイアログを作成
    $optionsForm = New-Object System.Windows.Forms.Form
    $optionsForm.Text = "オプション"
    $optionsForm.Size = New-Object System.Drawing.Size(450, 350)
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
    
    # 透明度設定
    $opacityLabel = New-Object System.Windows.Forms.Label
    $opacityLabel.Text = "ウィンドウの透明度："
    $opacityLabel.Location = New-Object System.Drawing.Point(20, 140)
    $opacityLabel.Size = New-Object System.Drawing.Size(150, 20)
    $opacityLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($opacityLabel)
    
    $opacityTrackBar = New-Object System.Windows.Forms.TrackBar
    $opacityTrackBar.Location = New-Object System.Drawing.Point(20, 165)
    $opacityTrackBar.Size = New-Object System.Drawing.Size(300, 45)
    $opacityTrackBar.Minimum = 50
    $opacityTrackBar.Maximum = 100
    $opacityTrackBar.TickFrequency = 10
    $opacityTrackBar.Value = [int]($script:windowOpacity * 100)
    $optionsForm.Controls.Add($opacityTrackBar)
    
    $opacityValueLabel = New-Object System.Windows.Forms.Label
    $opacityValueLabel.Location = New-Object System.Drawing.Point(330, 170)
    $opacityValueLabel.Size = New-Object System.Drawing.Size(80, 20)
    $opacityValueLabel.Text = "$([int]($script:windowOpacity * 100))%"
    $opacityValueLabel.ForeColor = $script:fgColor
    $optionsForm.Controls.Add($opacityValueLabel)
    
    $opacityTrackBar.Add_ValueChanged({
        $opacityValueLabel.Text = "$($opacityTrackBar.Value)%"
        $form.Opacity = $opacityTrackBar.Value / 100.0
    })
    
    # OKボタン
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(240, 270)
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
        
        # 透明度適用
        $script:windowOpacity = $opacityTrackBar.Value / 100.0
        $form.Opacity = $script:windowOpacity
        
        # 設定を保存
        Save-Config
    })
    $optionsForm.Controls.Add($okButton)
    
    # キャンセルボタン
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "キャンセル"
    $cancelButton.Location = New-Object System.Drawing.Point(330, 270)
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

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(490, 150)
$progress.Size = New-Object System.Drawing.Size(250, 30)
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
            $resultForm.BackColor = $bgColor
            $resultForm.ForeColor = $fgColor
            $resultForm.Font = New-Object System.Drawing.Font("Meiryo UI", 9)
            
            $resultTextBox = New-Object System.Windows.Forms.TextBox
            $resultTextBox.Multiline = $true
            $resultTextBox.ScrollBars = "Vertical"
            $resultTextBox.ReadOnly = $true
            $resultTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
            $resultTextBox.Dock = "Fill"
            $resultTextBox.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
            $resultTextBox.ForeColor = $fgColor
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
            $resultForm.BackColor = $bgColor
            $resultForm.ForeColor = $fgColor
            $resultForm.Font = New-Object System.Drawing.Font("Meiryo UI", 9)
            
            $resultTextBox = New-Object System.Windows.Forms.TextBox
            $resultTextBox.Multiline = $true
            $resultTextBox.ScrollBars = "Vertical"
            $resultTextBox.ReadOnly = $true
            $resultTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
            $resultTextBox.Dock = "Fill"
            $resultTextBox.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
            $resultTextBox.ForeColor = $fgColor
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

$showWindowButton.Add_Click({ Show-ResultWindows })
$closeAllWindowsButton.Add_Click({ Close-AllResultWindows })

# --- MediaInfo 呼び出し ---
function Invoke-MediaInfo($filePath) {
    try {
        # 一時ファイルに出力してUTF-8で読み込む
        $tempFile = [System.IO.Path]::GetTempFileName()
        $null = & $mediainfo --Output=Text --LogFile="$tempFile" "$filePath" 2>$null
        
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
                $infoJson = & $yt --dump-json --no-warnings "$input" 2>$null | ConvertFrom-Json
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
                    $formatOutput = & $yt -F "$input" 2>&1
                    
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

                    $target = & $yt -f best -g "$input" 2>$null | Select-Object -First 1
                    if (-not $target) { Write-OutputBox("") }
                } else { 
                    Write-OutputBox("yt-dlp で情報取得に失敗")
                    $resultContent += "yt-dlp で情報取得に失敗`r`n"
                }
            } else {
                if (-not (Test-Path -LiteralPath $input)) {
                    Write-OutputBox("ファイルが存在しません: $input")
                    $resultContent += "ファイルが存在しません: $input`r`n"
                    continue
                }
                Write-OutputBox("ローカルファイルとして解析します。")
                $resultContent += "ローカルファイルとして解析します。`r`n"
                $target = $input
                $resultTitle = [System.IO.Path]::GetFileName($input)
            }

            if ($target -and -not $isUrl) {
                Set-Progress(50 + [math]::Round($count / $total * 50))
                $mediaInfoOutput = Invoke-MediaInfo "$target"

                if ($mediaInfoOutput) {
                    Write-OutputBox("--- 詳細情報 ---")
                    $resultContent += "--- 詳細情報 ---`r`n"
                    
                    # MediaInfo出力をパース
                    $currentSection = ""
                    $duration = ""
                    $overallBitrate = ""
                    $fileSize = ""
                    $videoStreams = @()
                    $audioStreams = @()
                    $textStreams = @()
                    $imageStreams = @()
                    
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
                            continue
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
                } else {
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
    
    # 解析完了後にボタンを有効化
    $showWindowButton.Enabled = $true
}

$button.Add_Click({ Analyze-Video })
[void]$form.ShowDialog()