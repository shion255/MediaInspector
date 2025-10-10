Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
chcp 65001 > $null
$ErrorActionPreference = "Stop"

# --- ツールパス ---
$yt = "C:\encode\tools\yt-dlp.exe"
$ffprobe = "C:\encode\tools\ffprobe.exe"

foreach ($tool in @($yt, $ffprobe)) {
    if (-not (Test-Path $tool)) {
        [System.Windows.Forms.MessageBox]::Show("$tool が見つかりません。")
        exit
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

# --- ダークテーマ ---
$bgColor   = [System.Drawing.Color]::FromArgb(28, 28, 28)
$fgColor   = [System.Drawing.Color]::WhiteSmoke
$accent    = [System.Drawing.Color]::FromArgb(70, 130, 180)

# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "動画・音声解析ツール v7"
$form.Size = New-Object System.Drawing.Size(720, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = $bgColor
$form.ForeColor = $fgColor
$form.Font = New-Object System.Drawing.Font("Meiryo UI", 9)

$label = New-Object System.Windows.Forms.Label
$label.Text = "URL または ローカルファイル（複数可、改行で区切る）："
$label.Location = New-Object System.Drawing.Point(10, 15)
$label.AutoSize = $true
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 40)
$textBox.Size = New-Object System.Drawing.Size(680, 80)
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$textBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$textBox.ForeColor = $fgColor
$form.Controls.Add($textBox)

$button = New-Object System.Windows.Forms.Button
$button.Text = "解析開始"
$button.Location = New-Object System.Drawing.Point(10, 130)
$button.Size = New-Object System.Drawing.Size(120, 30)
$button.BackColor = $accent
$button.ForeColor = $fgColor
$form.Controls.Add($button)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(140, 130)
$progress.Size = New-Object System.Drawing.Size(550, 30)
$progress.Style = 'Continuous'
$form.Controls.Add($progress)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputBox.Location = New-Object System.Drawing.Point(10, 170)
$outputBox.Size = New-Object System.Drawing.Size(680, 380)
$outputBox.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$outputBox.ForeColor = $fgColor
$form.Controls.Add($outputBox)

function Write-OutputBox($msg) {
    $outputBox.AppendText($msg + "`r`n")
    $outputBox.ScrollToCaret()
}
function Set-Progress($v) {
    if ($v -gt 100) { $v = 100 }
    $progress.Value = $v
    $form.Refresh()
}

# --- ffprobe 呼び出し ---
function Invoke-FFProbe($filePath) {
    try {
        $jsonRaw = & $ffprobe -v error -print_format json -show_format -show_streams -i "$filePath" 2>$null
        if (-not $jsonRaw) { return $null }
        
        # JSON文字列を結合
        $jsonStr = $jsonRaw -join ""
        
        # JSONパース試行
        try {
            return $jsonStr | ConvertFrom-Json
        } catch {
            # メタデータが原因でエラーの場合、ストリーム情報のみ取得
            Write-OutputBox("⚠ メタデータにパース不可能な文字が含まれています。ストリーム情報のみ取得します。")
            $jsonRaw2 = & $ffprobe -v error -print_format json -show_streams -i "$filePath" 2>$null
            if ($jsonRaw2) {
                $jsonStr2 = $jsonRaw2 -join ""
                return $jsonStr2 | ConvertFrom-Json
            }
            return $null
        }
    } catch {
        Write-OutputBox("⚠ ffprobe エラー: $($_.Exception.Message)")
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

    $inputs = $inputsRaw -split "`r?`n"
    $outputBox.Clear()

    $total = $inputs.Count
    $count = 0

    foreach ($inputRaw in $inputs) {
        $input = $inputRaw.Trim() -replace '^"|"$',''
        $count++
        Write-OutputBox("--------------------------------------------------")
        Write-OutputBox("入力: $input")
        Write-OutputBox("解析開始... 少々お待ちください。")
        Set-Progress([math]::Round($count / $total * 100 / 2))

        $isUrl = $input -match '^https?://'
        $target = $null

        try {
            if ($isUrl) {
                # yt-dlp で情報取得
                $infoJson = & $yt --dump-json --no-warnings "$input" 2>$null | ConvertFrom-Json
                if ($infoJson) {
                    Write-OutputBox("タイトル: $($infoJson.title)")
                    Write-OutputBox("アップローダー: $($infoJson.uploader)")
                    Write-OutputBox("再生時間: " + (Format-Time $infoJson.duration))
                    $subs = $infoJson.subtitles.PSObject.Properties.Name
                    if ($subs -match "ja|jpn|Japanese") { Write-OutputBox("✅ 日本語字幕あり") }
                    else { Write-OutputBox("❌ 日本語字幕なし") }

                    # 実際の再生URL取得
                    $target = & $yt -f best -g "$input" 2>$null | Select-Object -First 1
                    if (-not $target) { Write-OutputBox("⚠ URL直接解析に失敗しました。") }
                } else { Write-OutputBox("yt-dlp で情報取得に失敗") }
            } else {
                if (-not (Test-Path -LiteralPath $input)) {
                    Write-OutputBox("ファイルが存在しません: $input")
                    continue
                }
                Write-OutputBox("ローカルファイルとして解析します。")
                $target = $input
            }

            # ffprobe 解析
            if ($target) {
                Set-Progress(50 + [math]::Round($count / $total * 50))
                $json = Invoke-FFProbe "$target"

                Write-OutputBox("--- ストリーム情報 ---")
                if ($json -and $json.format) {
                    if ($json.format.duration) { Write-OutputBox("再生時間(ffprobe): " + (Format-Time $json.format.duration)) }
                    if ($json.format.bit_rate) { Write-OutputBox("ビットレート: " + ([math]::Round($json.format.bit_rate/1000)) + " kbps") }
                }

                $hasStream = $false
                if ($json -and $json.streams) {
                    foreach ($s in $json.streams) {
                        $hasStream = $true
                        if ($s.codec_type -eq "video") {
                            Write-OutputBox("映像: $($s.codec_name) ${($s.width)}x$($s.height) $($s.avg_frame_rate)fps")
                        } elseif ($s.codec_type -eq "audio") {
                            $channels = if ($s.channels) { "$($s.channels)ch" } else { "" }
                            $sr = if ($s.sample_rate) { "$($s.sample_rate)Hz" } else { "" }
                            $br = if ($s.bit_rate) { "$([math]::Round($s.bit_rate/1000)) kbps" } else { "" }
                            Write-OutputBox("音声: $($s.codec_name) $channels $sr $br")
                        }
                    }
                }

                if (-not $hasStream) { Write-OutputBox("⚠ ストリーム情報を取得できませんでした。") }
            }

        } catch { Write-OutputBox("エラー: $_") }

        Write-OutputBox("解析完了。`r`n")
    }

    Set-Progress(100)
    Write-OutputBox("=== 全ファイル解析完了 ===")
}

$button.Add_Click({ Analyze-Video })
[void]$form.ShowDialog()