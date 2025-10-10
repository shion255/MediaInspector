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
$form.Text = "MediaInspector"
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
        
        $jsonStr = $jsonRaw -join ""
        
        try {
            return $jsonStr | ConvertFrom-Json
        } catch {
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
        Write-OutputBox("--------------------------------------------------")
        Write-OutputBox("入力: $input")
        Write-OutputBox("解析開始... 少々お待ちください。")
        Set-Progress([math]::Round($count / $total * 100 / 2))

        $isUrl = $input -match '^https?://'
        $target = $null

        try {
            if ($isUrl) {
                $infoJson = & $yt --dump-json --no-warnings "$input" 2>$null | ConvertFrom-Json
                if ($infoJson) {
                    Write-OutputBox("タイトル: $($infoJson.title)")
                    Write-OutputBox("アップローダー: $($infoJson.uploader)")
                    Write-OutputBox("再生時間: " + (Format-Time $infoJson.duration))
                    $subs = $infoJson.subtitles.PSObject.Properties.Name
                    if ($subs -match "ja|jpn|Japanese") { Write-OutputBox("✅ 日本語字幕あり") }
                    else { Write-OutputBox("❌ 日本語字幕なし") }

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
                            $w = $s.width
                            $h = $s.height
                            $resolution = if ($w -and $h) { "$w" + "x" + "$h" } else { "不明" }
                            
                            $fps = "不明"
                            $fpsStr = $s.avg_frame_rate
                            if (-not $fpsStr) { $fpsStr = $s.r_frame_rate }
                            
                            if ($fpsStr -and $fpsStr -match '^(\d+)/(\d+)$') {
                                $num = [double]$Matches[1]
                                $den = [double]$Matches[2]
                                if ($den -ne 0) {
                                    $fpsValue = [math]::Round($num / $den, 3)
                                    $fps = "$fpsValue fps"
                                }
                            }
                            
                            $isAttached = $false
                            if ($s.disposition) {
                                if ($s.disposition.attached_pic -eq 1) {
                                    $isAttached = $true
                                }
                            }
                            
                            if ($isAttached) {
                                Write-OutputBox("カバー画像: $($s.codec_name) $resolution")
                            } else {
                                Write-OutputBox("映像: $($s.codec_name) $resolution $fps")
                            }
                        } elseif ($s.codec_type -eq "audio") {
                            $channels = if ($s.channels) { "$($s.channels)ch" } else { "" }
                            $sr = if ($s.sample_rate) { "$([math]::Round($s.sample_rate/1000, 1)) kHz" } else { "" }
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