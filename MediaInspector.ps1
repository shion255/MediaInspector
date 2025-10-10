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
    # "8 503 kb/s" のような形式を処理
    if ($bitrate -match '([\d\s,]+)\s*kb/s') {
        $value = $matches[1] -replace '\s+', '' -replace ',', ''
        $numericValue = [double]$value
        return "{0:N0} kb/s" -f $numericValue
    }
    return $bitrate
}

function Convert-DurationToJapanese($duration) {
    if (-not $duration) { return "不明" }
    
    # "2 h 6 min" → "2時間6分"
    if ($duration -match '(\d+)\s*h\s*(\d+)\s*min') {
        $hours = $matches[1]
        $minutes = $matches[2]
        return "${hours}時間${minutes}分"
    }
    # "50 min 20 s" → "50分20秒"
    elseif ($duration -match '(\d+)\s*min\s*(\d+)\s*s') {
        $minutes = $matches[1]
        $seconds = $matches[2]
        return "${minutes}分${seconds}秒"
    }
    # "2 h" → "2時間"
    elseif ($duration -match '(\d+)\s*h') {
        $hours = $matches[1]
        return "${hours}時間"
    }
    # "1 h 55 min" → "1時間55分"
    elseif ($duration -match '(\d+)\s*h\s*(\d+)\s*min') {
        $hours = $matches[1]
        $minutes = $matches[2]
        return "${hours}時間${minutes}分"
    }
    
    return $duration
}

function Get-HDRInfo($colorPrimaries, $transferCharacteristics, $matrixCoefficients) {
    $hdrType = ""
    $colorSpace = ""
    
    # HDR判定
    if ($transferCharacteristics -match "PQ|SMPTE ST 2084") {
        $hdrType = "HDR10"
    } elseif ($transferCharacteristics -match "HLG") {
        $hdrType = "HLG"
    } else {
        $hdrType = "SDR"
    }
    
    # 色空間判定
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

# --- ダークテーマ ---
$bgColor   = [System.Drawing.Color]::FromArgb(28, 28, 28)
$fgColor   = [System.Drawing.Color]::WhiteSmoke
$accent    = [System.Drawing.Color]::FromArgb(70, 130, 180)

# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "MediaInspector"
$form.Size = New-Object System.Drawing.Size(850, 600)
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
$textBox.Size = New-Object System.Drawing.Size(810, 80)
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
$progress.Size = New-Object System.Drawing.Size(680, 30)
$progress.Style = 'Continuous'
$form.Controls.Add($progress)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputBox.Location = New-Object System.Drawing.Point(10, 170)
$outputBox.Size = New-Object System.Drawing.Size(810, 380)
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

            if ($target -and -not $isUrl) {
                Set-Progress(50 + [math]::Round($count / $total * 50))
                $mediaInfoOutput = Invoke-MediaInfo "$target"

                if ($mediaInfoOutput) {
                    Write-OutputBox("--- 詳細情報 ---")
                    
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
                    Write-OutputBox("ビットレート: $overallBitrate")
                    
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
                        $imageIndex++
                    }
                    
                    # テキストストリーム情報（番号付き）
                    $textIndex = 1
                    foreach ($txt in $textStreams) {
                        $language = if ($txt["language"]) { $txt["language"] } else { "不明" }
                        $default = if ($txt["default"] -eq "Yes") { "はい" } else { "いいえ" }
                        $forced = if ($txt["forced"] -eq "Yes") { "はい" } else { "いいえ" }
                        
                        Write-OutputBox("テキスト${textIndex}: $language | Default - $default | Forced - $forced")
                        $textIndex++
                    }
                } else {
                    Write-OutputBox("⚠ MediaInfo で情報を取得できませんでした。")
                }
            } elseif ($target -and $isUrl) {
                Write-OutputBox("")
                Write-OutputBox("※ URL動画の詳細情報はyt-dlpの情報を参照してください。")
            }

        } catch { Write-OutputBox("エラー: $_") }

        Write-OutputBox("")
        Write-OutputBox("解析完了。`r`n")
    }

    Set-Progress(100)
    Write-OutputBox("=== 全ファイル解析完了 ===")
}

$button.Add_Click({ Analyze-Video })
[void]$form.ShowDialog()