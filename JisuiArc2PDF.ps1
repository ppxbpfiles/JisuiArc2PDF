<#
.SYNOPSIS
    書庫(7-zip対応形式)を、画像を最適化しながら高品質なPDFに変換します。

.DESCRIPTION
    このスクリプトは、指定された書庫またはワイルドカードに一致する複数の書庫を読み込み、以下の処理を自動で行います。
    1. 書庫内の画像を一時フォルダに展開します。
    2. ファイル名を数字の順（自然順）に並べ替えます。
    3. 画像の彩度を自動判別し、「カラー」か「グレースケール」かを判断します。
    4. カラー/グレースケールそれぞれに最適化された設定で、画像をリサイズ・変換します。
    5. 変換された画像を結合し、元の書庫ファイルと同じ名前のPDFを作成します。
    6. PDFのタイムスタンプを元の書庫に合わせて設定します。

.NOTES
    --------------------------------------------------------------------------------
    ■ 必要なソフトウェア (必須)
    --------------------------------------------------------------------------------
    このスクリプトを実行するには、以下の3つのツールを事前にダウンロードし、
    スクリプトと同じフォルダに置くか、PCの「環境変数PATH」に登録しておく必要があります。
    スクリプトは、まず自身と同じフォルダに各ツールの実行ファイルが存在するかを確認し、
    存在しない場合は環境変数PATHから検索します。

    PATH登録が完了しているかは、PowerShellで `ツールの名前 -version` を実行して確認できます。

    1. ImageMagick (magick.exe)
       - 役割: 画像のリサイズ、カラー/グレースケール変換など、画像処理全般を担当します。
       - 入手先: https://imagemagick.org/script/download.php

    2. 7-Zip (7z.exe)
       - 役割: ZIPおよびRAR書庫を展開するために使用します。
       - 入手先: https://www.7-zip.org/

    3. PDFCPU (pdfcpu.exe)
       - 役割: 画像ファイル群からPDFを生成し、最適化を行います。
       - 入手先: https://pdfcpu.io/download

    --------------------------------------------------------------------------------
    ■ スクリプトの使い方
    --------------------------------------------------------------------------------
    PowerShellを開き、引数に処理したいファイルパス（またはワイルドカード）を指定して実行します。
    ファイル名に日本語やスペースなどの特殊文字が含まれる場合は、パスを引用符で囲んでください。

    (例1: ファイルを個別に指定)
    pwsh -File .\JisuiArc2PDF.ps1 "MyBook.zip"

    (例2: ワイルドカードでまとめて指定)
    pwsh -File .\JisuiArc2PDF.ps1 "*.rar"

    詳細なヘルプを見るには:
    pwsh -Command "Get-Help .\JisuiArc2PDF.ps1 -Full"

.PARAMETER ArchiveFilePaths
    処理対象となる書庫のファイルパス、またはワイルドカードを含むパターン。

.PARAMETER SevenZipPath
    7z.exe への絶対パスを明示的に指定します。
    指定しない場合、環境変数PATHから自動的に検索されます。

.PARAMETER MagickPath
    magick.exe への絶対パスを明示的に指定します。
    指定しない場合、環境変数PATHから自動的に検索されます。

.PARAMETER PdfCpuPath
    pdfcpu.exe への絶対パスを明示的に指定します。
    指定しない場合、環境変数PATHから自動的に検索されます。

.PARAMETER Quality
    変換するJPEG画像の品質を1から100の整数で指定します。
    デフォルト値は 85 です。
    エイリアス: -q

.PARAMETER SaturationThreshold
    画像をグレースケールと判断する彩度のしきい値を0.0から1.0の小数で指定します。
    デフォルト値は 0.05 です。
    エイリアス: -s, -sat

.PARAMETER TotalCompressionThreshold
    圧縮後ファイルセットを使用するかの判断に使われる、圧縮率のしきい値(パーセント)です。
    この値が指定されると、「変換後の合計サイズが、元の合計サイズのXX%未満の場合」にのみ、変換後のファイルが使用されます。
    エイリアス: -tcr

.PARAMETER Height
    画像の高さをピクセル単位で指定します。
    エイリアス: -h

.PARAMETER Dpi
    画像のDPI（解像度）を指定します。
    エイリアス: -d

.PARAMETER PaperSize
    用紙サイズを指定して高さを自動計算します（Dpiと併用）。
    指定可能な値: A4, A5, B5, B6
    エイリアス: -p

.PARAMETER Verbose
    詳細な診断メッセージを表示します。
    エイリアス: -v

.EXAMPLE
    # PowerShellターミナルから直接スクリプトを実行する
    .\JisuiArc2PDF.ps1 "MyBook.zip"

    # ワイルドカードを使い、カレントディレクトリの全てのrarファイルを処理
    .\JisuiArc2PDF.ps1 *.rar

    # 異なるフォルダの全てのzipファイルを処理 (パスにスペース等が含まれる場合はダブルクォートで囲む)
    .\JisuiArc2PDF.ps1 "C:\Path\To\Archives\*.zip"
    .\JisuiArc2PDF.ps1 "..\OtherFolder\*.rar"

    # 高さを2000pxに指定
    .\JisuiArc2PDF.ps1 *.rar -Height 2000

    # B5サイズ、300dpiで高さを自動計算
    .\JisuiArc2PDF.ps1 *.rar -PaperSize B5 -Dpi 300
#>
[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [string[]]$ArchiveFilePaths,

    [Parameter(Mandatory=$false)]
    [string]$SevenZipPath,

    [Parameter(Mandatory=$false)]
    [string]$MagickPath,

    [Parameter(Mandatory=$false)]
    [string]$PdfCpuPath,

    [Parameter(Mandatory=$false)]
    [Alias('q')]
    [int]$Quality = 85,

    [Parameter(Mandatory=$false)]
    [Alias('s', 'sat')]
    [double]$SaturationThreshold = 0.05,

    [Parameter(Mandatory=$false)]
    [Alias('tcr')]
    [ValidateRange(0.0, 100.0)]
    [double]$TotalCompressionThreshold,

    [Parameter(Mandatory=$false)]
    [Alias('h')]
    [int]$Height,

    [Parameter(Mandatory=$false)]
    [Alias('d')]
    [int]$Dpi,

    [Parameter(Mandatory=$false)]
    [Alias('p')]
    [ValidateSet('A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'B0', 'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7')]
    [string]$PaperSize,

    [Parameter(Mandatory=$false)]
    [Alias('sc')]
    [switch]$SkipCompression,

    [Parameter(Mandatory=$false)]
    [switch]$Trim,

    [Parameter(Mandatory=$false)]
    [string]$Fuzz = "1%",

    [Parameter(Mandatory=$false)]
    [Alias('lin')]
    [switch]$Linearize,

    [Parameter(Mandatory=$false)]
    [Alias('ds')]
    [switch]$Deskew,

    [Parameter(Mandatory=$false)]
    [Alias('gl')]
    [string]$GrayscaleLevel,

    [Parameter(Mandatory=$false)]
    [Alias('cc')]
    [string]$ColorContrast,

    [Parameter(Mandatory=$false)]
    [Alias('ac')]
    [switch]$AutoContrast,

    [Parameter(Mandatory=$false)]
    [string]$LogPath,

    [Parameter(Mandatory=$false)]
    [Alias('sp')]
    [switch]$SplitPages,

    [Parameter(Mandatory=$false)]
    [Alias('b')]
    [ValidateSet('Right', 'Left')]
    [string]$Binding = 'Right'
)

# -SetPageSizeのデフォルト挙動を決定
if (-not $PSBoundParameters.ContainsKey('SetPageSize')) {
    if ($PSBoundParameters.ContainsKey('PaperSize')) {
        # -PaperSizeが指定されたら、ページサイズ設定をデフォルトで有効にする
        $SetPageSize = $true
    }
    # -Heightのみ、または何も指定されない場合は、デフォルトで無効のまま ($SetPageSizeは$falseのまま)
}

# ==============================================================================
# 対話モードでのパラメータ入力
# ==============================================================================
# === 修正箇所: 対話モードのパス入力を簡素化し、後の処理に任せる ===
if (-not $PSBoundParameters.ContainsKey('ArchiveFilePaths')) {
    while (-not $ArchiveFilePaths) {
        $prompt = "処理する書庫のパスを入力してください (例: Y:\books\*.zip)"
        $inputPath = Read-Host -Prompt $prompt

        $promptWithColon = $prompt + ": "
        if ($inputPath.StartsWith($promptWithColon)) {
            $inputPath = $inputPath.Substring($promptWithColon.Length)
        }

        $inputPath = $inputPath.Trim().Trim('"')

        if (-not ([string]::IsNullOrWhiteSpace($inputPath))) {
            # ここではパスの検証を行わず、生の文字列をそのまま配列に入れる
            # 検証と解決は後続のメインロジックに任せることで、二重処理を防ぐ
            $ArchiveFilePaths = @($inputPath)
        } else {
             Write-Warning "パスが入力されていません。"
        }
    }
}

# コマンドラインで詳細なオプションが指定されていない場合、対話モードを開始
if ($PSBoundParameters.Count -le 1 -and ($PSBoundParameters.Count -eq 0 -or $PSBoundParameters.ContainsKey('ArchiveFilePaths'))) {
    Write-Host "`n----------------------------------------" -ForegroundColor Green
    Write-Host "対話モードで変換設定を行います。" -ForegroundColor Green
    Write-Host "角括弧 [] 内のデフォルト値を使用する場合は、何も入力せずにEnterキーを押してください。"
    Write-Host "----------------------------------------" -ForegroundColor Green

    $skip_in = Read-Host "圧縮をスキップしますか (y/n)? [n]"
    if ($skip_in.ToLower() -eq 'y') { $SkipCompression = $true }

    if (-not $SkipCompression.IsPresent) {
        $quality_in = Read-Host "品質 (1-100) [85]"
        if (-not ([string]::IsNullOrWhiteSpace($quality_in))) { $Quality = [int]$quality_in }

        $sat_in = Read-Host "彩度のしきい値 [0.05]"
        if (-not ([string]::IsNullOrWhiteSpace($sat_in))) { $SaturationThreshold = [double]$sat_in }

        $tcr_in = Read-Host "合計圧縮率のしきい値 (0-100, optional)"
        if (-not ([string]::IsNullOrWhiteSpace($tcr_in))) { $TotalCompressionThreshold = [double]$tcr_in }

        $deskew_in = Read-Host "傾き補正 (y/n)? [n]"
        if ($deskew_in.ToLower() -eq 'y') { $Deskew = $true }

        $gray_contrast_in = Read-Host "グレースケールコントラスト調整 (y/n)? [n]"
        if ($gray_contrast_in.ToLower() -eq 'y') {
            $level_in = Read-Host "グレースケール Level値 [10%,90%]"
            $GrayscaleLevel = if (-not [string]::IsNullOrWhiteSpace($level_in)) { $level_in } else { "10%,90%" }
        }

        $auto_contrast_in = Read-Host "カラーコントラスト自動調整 (y/n)? [n]"
        if ($auto_contrast_in.ToLower() -eq 'y') {
            $AutoContrast = $true
        } else {
            $color_contrast_in = Read-Host "カラーコントラスト手動調整 (y/n)? [n]"
            if ($color_contrast_in.ToLower() -eq 'y') {
                $bright_in = Read-Host "Brightness-Contrast値 [0x25]"
                $ColorContrast = if (-not ([string]::IsNullOrWhiteSpace($bright_in))) { $bright_in } else { "0x25" }
            }
        }

        $trim_in = Read-Host "余白除去 (y/n)? [n]"
        if ($trim_in.ToLower() -eq 'y') {
            $Trim = $true
            $fuzz_in = Read-Host "Fuzz係数 (例: 1%) [1% ]"
            if (-not ([string]::IsNullOrWhiteSpace($fuzz_in))) { $Fuzz = $fuzz_in }
        }

        $split_in = Read-Host "見開きページを分割しますか (y/n)? [n]"
        if ($split_in.ToLower() -eq 'y') {
            $SplitPages = $true
            $binding_in = Read-Host "本の綴じ方向 (1=右綴じ(漫画など), 2=左綴じ) [1]"
            if ($binding_in -eq '2') { $Binding = 'Left' }
            else { $Binding = 'Right' }
        }

        $linearize_in = Read-Host "PDFをリニアライズしますか (y/n)? [n]"
        if ($linearize_in.ToLower() -eq 'y') { $Linearize = $true }

        # 解像度設定を先に質問
        $res_choice = Read-Host "解像度設定: 1=高さ指定, 2=用紙サイズ+DPI [2]"
        if ($res_choice -eq '1') {
            $height_in = Read-Host "高さ (ピクセル)"
            if (-not ([string]::IsNullOrWhiteSpace($height_in))) { $Height = [int]$height_in }
            $dpi_for_h_in = Read-Host "DPI [144]"
            if (-not ([string]::IsNullOrWhiteSpace($dpi_for_h_in))) { $Dpi = [int]$dpi_for_h_in }
            $PaperSize = "Custom"

            # 高さ指定の場合、デフォルトは「自動」
            $ps_prompt = "PDFページサイズ設定 (1:自動, 2:縦向きで設定, 3:横向きで設定) [1]"
            $ps_choice = Read-Host -Prompt $ps_prompt
            switch ($ps_choice) {
                '2' { $SetPageSize = $true }
                '3' { $SetPageSize = $true; $Landscape = $true }
                default {} # Default is 1 (auto), so do nothing
            }
        } else {
            $paper_in = Read-Host "用紙サイズ [A4]"
            $PaperSize = if (-not ([string]::IsNullOrWhiteSpace($paper_in))) { $paper_in } else { "A4" }
            $dpi_in = Read-Host "DPI [144]"
            $Dpi = if (-not ([string]::IsNullOrWhiteSpace($dpi_in))) { [int]$dpi_in } else { 144 }

            # 用紙サイズ指定の場合、デフォルトは「縦向きで設定」
            $ps_prompt = "PDFページサイズ設定 (1:縦向きで設定, 2:横向きで設定, 3:自動) [1]"
            $ps_choice = Read-Host -Prompt $ps_prompt
            switch ($ps_choice) {
                '2' { $SetPageSize = $true; $Landscape = $true }
                '3' { $SetPageSize = $false }
                default { $SetPageSize = $true } # Default is 1 (set portrait)
            }
        }
    }
    Write-Host "----------------------------------------`n" -ForegroundColor Green
}


# ==============================================================================
# ログファイルパス設定
# ==============================================================================
$logFilePath = ""
if ($PSBoundParameters.ContainsKey('LogPath')) {
    $resolvedLogPath = $LogPath
    if (-not ([System.IO.Path]::IsPathRooted($resolvedLogPath))) {
        $resolvedLogPath = Join-Path $PSScriptRoot $resolvedLogPath
    }

    if ((Test-Path -Path $resolvedLogPath -PathType Container) -or ($resolvedLogPath.EndsWith('\') -or $resolvedLogPath.EndsWith('/'))) {
        $logFilePath = Join-Path $resolvedLogPath "JisuiArc2PDF_log.txt"
    } else {
        $logFilePath = $resolvedLogPath
    }
} else {
    $logFilePath = Join-Path $PSScriptRoot "JisuiArc2PDF_log.txt"
}

try {
    $logDirectory = Split-Path -Path $logFilePath -Parent -ErrorAction Stop
    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force -ErrorAction Stop | Out-Null
    }
} catch {
    Write-Error "ログファイルのパスまたはディレクトリの作成に失敗しました: $logFilePath - $($_.Exception.Message)"
    exit 1
}


# ==============================================================================
# 解像度設定の計算
# ==============================================================================
$targetHeight = 0
$targetDpi = 0
$targetPaperSizeForPdfCpu = "A4"

if ($Height) {
    $targetHeight = $Height
    $targetDpi = if ($Dpi) { $Dpi } else { 144 }
    $targetPaperSizeForPdfCpu = "auto"
    Write-Host "[情報] 高さ指定: $targetHeight px, DPI: $targetDpi dpi, PDFページ: 自動"
}
elseif ($Dpi -and $PaperSize) {
    $targetDpi = $Dpi
    $targetPaperSizeForPdfCpu = $PaperSize
    $paperHeightMm = 0
    switch ($PaperSize) {
        'A0' { $paperHeightMm = 1189 } 'A1' { $paperHeightMm = 841 } 'A2' { $paperHeightMm = 594 }
        'A3' { $paperHeightMm = 420 } 'A4' { $paperHeightMm = 297 } 'A5' { $paperHeightMm = 210 }
        'A6' { $paperHeightMm = 148 } 'A7' { $paperHeightMm = 105 } 'B0' { $paperHeightMm = 1414 }
        'B1' { $paperHeightMm = 1000 } 'B2' { $paperHeightMm = 707 } 'B3' { $paperHeightMm = 500 }
        'B4' { $paperHeightMm = 364 } 'B5' { $paperHeightMm = 257 } 'B6' { $paperHeightMm = 182 }
        'B7' { $paperHeightMm = 128 }
    }
    $targetHeight = [math]::Round(($paperHeightMm / 25.4) * $targetDpi)
    Write-Host "[情報] 計算設定: $PaperSize ($paperHeightMm mm) @ ${targetDpi} dpi -> $targetHeight px"
}
elseif ($Dpi -or $PaperSize) {
    if ($Dpi -and -not $PaperSize) { Write-Error "-Dpi が指定されましたが、-PaperSize が指定されていません。両方を指定してください。" }
    else { Write-Error "-PaperSize が指定されましたが、-Dpi が指定されていません。両方を指定してください。" }
    exit 1
}
else {
    $targetDpi = 144
    $paperHeightMm = 297 # A4
    $targetHeight = [math]::Round(($paperHeightMm / 25.4) * $targetDpi)
    $targetPaperSizeForPdfCpu = "A4"
    Write-Host "[情報] デフォルト設定: A4 @ ${targetDpi} dpi -> $targetHeight px"
}
# ==============================================================================

# ==============================================================================
# 引数チェック
# ==============================================================================
if ($PSBoundParameters.ContainsKey('ArchiveFilePaths') -eq $false -and -not $ArchiveFilePaths) {
    $helpMessage = @"
--------------------------------------------------------------------------------
JisuiArc2PDF: 書庫(7-zip対応形式)を高品質なPDFに変換します。
--------------------------------------------------------------------------------

使用法 (簡単・推奨):
  以下のコマンドでスクリプトを起動すると、対話モードで設定を聞かれます。
  pwsh -File .\JisuiArc2PDF.ps1

または、処理したいファイルを直接指定して対話モードを開始することもできます:
  # 単一のファイルを指定
  pwsh -File .\JisuiArc2PDF.ps1 "MyBook.zip"

  # ワイルドカードでまとめて指定
  pwsh -File .\JisuiArc2PDF.ps1 *.rar

詳細なヘルプ:
  pwsh -Command "Get-Help .\JisuiArc2PDF.ps1 -Full"
"@
    Write-Host $helpMessage
    exit 0
}


# ==============================================================================
Write-Verbose "[診断] スクリプト開始。受信した引数: $($ArchiveFilePaths -join ', ')"

# === 修正箇所: このパス解決ブロックが、対話モードからの入力もコマンドライン引数も同様に処理する ===
$resolvedFilePaths = @()
foreach ($rawPath in $ArchiveFilePaths) {
    # パスに含まれる角括弧をエスケープして、ワイルドカードとして誤認識されるのを防ぐ
    $escapedPath = $rawPath -replace '\[', '`[' -replace '\]', '`]'
    try {
        $items = Resolve-Path -Path $escapedPath -ErrorAction Stop
        foreach ($item in $items) {
            $fileInfo = Get-Item -LiteralPath $item.Path
            if ($fileInfo.PSIsContainer) {
                $childFiles = Get-ChildItem -Path $fileInfo.FullName -Recurse -File
                $resolvedFilePaths += $childFiles.FullName
            } else {
                $resolvedFilePaths += $fileInfo.FullName
            }
        }
    } catch {
        Write-Warning "指定されたパスまたはパターンに一致するファイルが見つかりません: '$rawPath'"
    }
}
$ArchiveFilePaths = $resolvedFilePaths | Sort-Object -Unique

if ($ArchiveFilePaths.Count -eq 0) {
    Write-Error "処理対象の書庫ファイルが一つも見つかりませんでした。パスを確認してください。"
    exit 1
}
# ==============================================================================

# ==============================================================================
# 0. 前提ツールのパス解決
# ==============================================================================
# スクリプトの実行に必要な外部ツール（7z, magick, pdfcpu, qpdf）の実行ファイルパスを解決します。
# 検索の優先順位:
# 1. ユーザーが引数で明示的に指定したパス (例: -SevenZipPath "C:\...")
# 2. スクリプト(.ps1)と同一ディレクトリに配置されている実行ファイル
# 3. 環境変数PATHが通っている場所
# 4. (最終手段) Get-Command コマンドレットでの検索

# 0. 前提ツールのパス解決と存在チェック
$sevenzip_exe = $null
if ($PSBoundParameters.ContainsKey('SevenZipPath') -and (Test-Path -LiteralPath $SevenZipPath -PathType Leaf)) {
    $sevenzip_exe = $SevenZipPath
} else {
    $localExePath = Join-Path $PSScriptRoot "7z.exe"
    if (Test-Path -LiteralPath $localExePath -PathType Leaf) {
        $sevenzip_exe = $localExePath
    } else {
        $foundPath = (cmd.exe /c "where.exe 7z.exe" 2>$null).Split([System.Environment]::NewLine) | Select-Object -First 1
        if ($foundPath) {
            $trimmedPath = $foundPath.Trim()
            if ($trimmedPath -and (Test-Path -LiteralPath $trimmedPath -PathType Leaf)) {
                $sevenzip_exe = $trimmedPath
            }
        }
        if (-not $sevenzip_exe) {
            $sevenzip_exe = (Get-Command '7z' -ErrorAction SilentlyContinue).Source
        }
    }
}

$magick_exe = $null
if ($PSBoundParameters.ContainsKey('MagickPath') -and (Test-Path -LiteralPath $MagickPath -PathType Leaf)) {
    $magick_exe = $MagickPath
} else {
    $localExePath = Join-Path $PSScriptRoot "magick.exe"
    if (Test-Path -LiteralPath $localExePath -PathType Leaf) {
        $magick_exe = $localExePath
    } else {
        $foundPath = (cmd.exe /c "where.exe magick.exe" 2>$null).Split([System.Environment]::NewLine) | Select-Object -First 1
        if ($foundPath) {
            $trimmedPath = $foundPath.Trim()
            if ($trimmedPath -and (Test-Path -LiteralPath $trimmedPath -PathType Leaf)) {
                $magick_exe = $trimmedPath
            }
        }
        if (-not $magick_exe) {
            $magick_exe = (Get-Command 'magick' -ErrorAction SilentlyContinue).Source
        }
    }
}

$pdfcpu_exe = $null
if ($PSBoundParameters.ContainsKey('PdfCpuPath') -and (Test-Path -LiteralPath $PdfCpuPath -PathType Leaf)) {
    $pdfcpu_exe = $PdfCpuPath
} else {
    $localExePath = Join-Path $PSScriptRoot "pdfcpu.exe"
    if (Test-Path -LiteralPath $localExePath -PathType Leaf) {
        $pdfcpu_exe = $localExePath
    } else {
        $foundPath = (cmd.exe /c "where.exe pdfcpu.exe" 2>$null).Split([System.Environment]::NewLine) | Select-Object -First 1
        if ($foundPath) {
            $trimmedPath = $foundPath.Trim()
            if ($trimmedPath -and (Test-Path -LiteralPath $trimmedPath -PathType Leaf)) {
                $pdfcpu_exe = $trimmedPath
            }
        }
        if (-not $pdfcpu_exe) {
            $pdfcpu_exe = (Get-Command 'pdfcpu' -ErrorAction SilentlyContinue).Source
        }
    }
}

$qpdf_exe = $null
$localExePath = Join-Path $PSScriptRoot "qpdf.exe"
if (Test-Path -LiteralPath $localExePath -PathType Leaf) {
    $qpdf_exe = $localExePath
} else {
    $foundPath = (cmd.exe /c "where.exe qpdf.exe" 2>$null).Split([System.Environment]::NewLine) | Select-Object -First 1
    if ($foundPath) {
        $trimmedPath = $foundPath.Trim()
        if ($trimmedPath -and (Test-Path -LiteralPath $trimmedPath -PathType Leaf)) {
            $qpdf_exe = $trimmedPath
        }
    }
    if (-not $qpdf_exe) {
        $qpdf_exe = (Get-Command 'qpdf' -ErrorAction SilentlyContinue).Source
    }
}

if (-not $magick_exe) { Write-Error "ImageMagick (magick.exe) が見つかりません。パスを指定するか、環境変数PATHに登録してください。"; exit 1 }
if (-not $sevenzip_exe) { Write-Error "7-Zip (7z.exe) が見つかりません。パスを指定するか、環境変数PATHに登録してください。"; exit 1 }
if (-not $pdfcpu_exe) { Write-Error "PDFCPU (pdfcpu.exe) が見つかりません。パスを指定するか、環境変数PATHに登録してください。"; exit 1 }

foreach ($ArchiveFilePath in $ArchiveFilePaths) {
    Write-Output "[診断] メインループ開始: $ArchiveFilePath"
    $archiveFileInfo = Get-Item -LiteralPath $ArchiveFilePath -ErrorAction SilentlyContinue
    if (-not $archiveFileInfo) {
        Write-Warning "指定された書庫ファイルが見つかりません: $ArchiveFilePath";
        continue
    }

    Write-Host "========================================="
    Write-Host "処理中: $($archiveFileInfo.Name)"
    Write-Host "========================================="

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
    $tempDirConverted = Join-Path $tempDir "converted"
    $tempDirConvertedColor = Join-Path $tempDirConverted "color"
    $tempDirConvertedGray = Join-Path $tempDirConverted "gray"

    New-Item -ItemType Directory -Path $tempDir, $tempDirConverted, $tempDirConvertedColor, $tempDirConvertedGray | Out-Null
    Write-Verbose "一時フォルダを作成しました: $tempDir"

    try {
        $logDetails = @()
        $convertedCount = 0
        $originalCount = 0
        $skippedFiles = @()

        Write-Host "アーカイブを展開しています: $($archiveFileInfo.Name)"
        & $sevenzip_exe e "$($archiveFileInfo.FullName)" "-o$tempDir" -y *>$null

        Write-Host "画像ファイルを検索し、並べ替えています..."
        $allFiles = Get-ChildItem -Path $tempDir -Recurse -File
        $imageFiles = @()
        foreach ($file in $allFiles) {
            & $magick_exe identify "$($file.FullName)" *>$null
            if ($LASTEXITCODE -eq 0) {
                $imageFiles += $file
                Write-Verbose "画像として認識: $($file.Name)"
            } else {
                Write-Verbose "画像ではないためスキップ: $($file.Name)"
            }
        }

        $imageFiles = $imageFiles | Sort-Object -Property @{Expression={ 
            $mangledName = ''
            $parts = $_.Name -split '(\d+)'
            foreach ($part in $parts) {
                if ([string]::IsNullOrEmpty($part)) { continue }
                if ($part -match '^\d+$') { $mangledName += $part.PadLeft(20, '0') }
                else { $mangledName += $part }
            }
            $mangledName
        }}

        # --------------------------------------------------------------------------
        # 2. 見開きページの分割 (オプション)
        # --------------------------------------------------------------------------
        # -SplitPages スイッチが指定されている場合、横長の画像を見開きページと見なして分割する。
        if ($SplitPages.IsPresent) {
            Write-Host "[情報] 見開きページを分割します (綴じ方向: $Binding)..."
            $processedFiles = @()
            $splitTempDir = Join-Path $tempDir "split_pages"
            New-Item -ItemType Directory -Path $splitTempDir | Out-Null
            
            $pageCounter = 0
            foreach ($file in $imageFiles) {
                try {
                    # 画像の幅と高さを取得
                    $dimensions = & $magick_exe identify -format "%w %h" "$($file.FullName)"
                    $dimArray = $dimensions -split ' '
                    $width = [int]$dimArray[0]
                    $height = [int]$dimArray[1]

                    # 幅が高さの1.2倍より大きい場合、見開きページと判断して分割
                    if ($width -gt ($height * 1.2)) {
                        Write-Host "  -> 分割中: $($file.Name) (${width}x${height})"
                        
                        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                        
                        # 分割後の左右のページファイルパスを定義
                        $leftPagePath = Join-Path $splitTempDir ("{0:D4}_{1}_1_left.jpg" -f $pageCounter, $baseName)
                        $rightPagePath = Join-Path $splitTempDir ("{0:D4}_{1}_2_right.jpg" -f $pageCounter, $baseName)
                        
                        # ImageMagickのcrop機能で画像を左右に50%ずつ分割
                        $cropArgsLeft = @("$($file.FullName)", "-crop", "50%x100%+0+0", "+repage", "$leftPagePath")
                        & $magick_exe @cropArgsLeft
                        $cropArgsRight = @("$($file.FullName)", "-crop", "50%x100%+$([int]($width/2))+0", "+repage", "$rightPagePath")
                        & $magick_exe @cropArgsRight

                        if (Test-Path $leftPagePath) {
                            if (Test-Path $rightPagePath) {
                                # -Binding の値に応じて、PDFに格納するページの順序を決定
                                if ($Binding -eq 'Right') { # 右綴じ(漫画など): 右ページが先
                                    $processedFiles += Get-Item -LiteralPath $rightPagePath
                                    $processedFiles += Get-Item -LiteralPath $leftPagePath
                                } else { # 左綴じ: 左ページが先
                                    $processedFiles += Get-Item -LiteralPath $leftPagePath
                                    $processedFiles += Get-Item -LiteralPath $rightPagePath
                                }
                            } else {
                                Write-Warning "ページの分割に失敗しました: $($file.Name)。このページはスキップされます。"
                            }
                        } else {
                            Write-Warning "ページの分割に失敗しました: $($file.Name)。このページはスキップされます。"
                        }
                    } else {
                        # 見開きではない場合、元のファイルをそのままリストに追加
                        $processedFiles += $file
                    }
                } catch {
                    Write-Warning "画像の寸法取得または分割中にエラーが発生しました: $($file.Name) - $($_.Exception.Message)。このページはスキップされます。"
                }
                $pageCounter++
            }
            # 元の画像リストを、分割処理後の新しいリストで上書きする
            $imageFiles = $processedFiles
        }

        if ($imageFiles.Count -eq 0) { throw "アーカイブ内に処理可能な画像ファイルが見つかりませんでした。" }
        
        # --------------------------------------------------------------------------
        # 3. 画像ファイルの処理とPDF化
        # --------------------------------------------------------------------------
        $filesForPdf = @()

        # A. 圧縮スキップモード
        # -SkipCompression が指定された場合、最適化（リサイズ、品質調整等）を行わず、
        #   非JPEGファイルをJPEGに変換するだけの最小限の処理を行う。
        if ($SkipCompression.IsPresent) {
            Write-Host "[情報] -SkipCompression: 最適化をスキップし、非JPEGファイルをJPEGに変換します。"
            $jpegExtensions = @('.jpg', '.jpeg', '.jfif', '.jpe')
            $filesForPdf = @()
            $fileCounter = 0
            $tempDirScPassthrough = Join-Path $tempDir "sc_passthrough"

            foreach ($file in $imageFiles) {
                $extension = [System.IO.Path]::GetExtension($file.FullName).ToLower()
                if ($jpegExtensions -contains $extension) {
                    # 元からJPEGの場合は何もしない
                    $filesForPdf += $file.FullName
                    $logDetails += "    - $($file.Name): Original (JPEG passthrough)"
                    Write-Verbose "  -> $($file.Name): JPEGのため変換不要。"
                } else {
                    # 非JPEGファイルはJPEGに変換する
                    if (-not (Test-Path $tempDirScPassthrough)) { New-Item -ItemType Directory -Path $tempDirScPassthrough | Out-Null }
                    $newFileName = "{0:D4}.jpg" -f $fileCounter
                    $passthroughPath = Join-Path $tempDirScPassthrough $newFileName
                    Write-Verbose "  -> $($file.Name): 非JPEGのため変換。"
                    & $magick_exe "$($file.FullName)" "$passthroughPath"

                    if (-not (Test-Path $passthroughPath)) {
                        Write-Warning "JPEGへの形式変換に失敗しました: $($file.FullName)"
                        $skippedFiles += $file
                        $logDetails += "    - $($file.Name): SKIPPED (File conversion failed)"
                    } else {
                        $filesForPdf += $passthroughPath
                        $logDetails += "    - $($file.Name): Converted (to JPEG)"
                    }
                }
                $fileCounter++
            }
            $originalCount = $imageFiles.Count - $skippedFiles.Count
            $convertedCount = 0
        }
        # B. 通常の変換モード
        else {
            Write-Host "画像を変換し、一時ファイルに保存しています..."
            $fileCounter = 0
            # 変換結果（元ファイルと変換後ファイルのサイズなど）を格納する配列
            $conversionResults = @()

            # 全ての画像ファイルをループして変換処理を行う
            foreach ($file in $imageFiles) {
                Write-Host "ファイル処理中: $($file.Name)"

                # 画像の彩度を取得し、グレースケールかカラーかを判断
                $SATURATION_STR = & $magick_exe "$($file.FullName)" -colorspace HSL -channel G -separate +channel -format "%[mean]" info:
                $SATURATION = [double]$SATURATION_STR / 65535.0
                $newFileName = "{0:D4}.jpg" -f $fileCounter

                try {
                    # 画像の高さを取得
                    $originalHeightStr = & $magick_exe identify -format "%h" "$($file.FullName)"
                    if ($originalHeightStr -match '^\d+') {
                        $originalHeight = [int]$originalHeightStr
                    } else {
                        Write-Warning "画像の高さの取得に失敗しました: $($file.Name)"; $skippedFiles += $file; $fileCounter++; continue
                    }
                } catch {
                    Write-Warning "identifyの実行中にエラーが発生しました: $($file.Name) - $($_.Exception.Message)"; $skippedFiles += $file; $fileCounter++; continue
                }

                # ImageMagickに渡す引数を動的に構築していく
                $magickArgs = @("$($file.FullName)")
                if ($Deskew.IsPresent) { $magickArgs += "-deskew", "40%"; Write-Host "  -> 傾きを補正します。" }
                if ($Trim.IsPresent) { $magickArgs += "-fuzz", $Fuzz, "-trim", "+repage"; Write-Host "  -> 余白を除去します (Fuzz: $Fuzz)。" }

                # 目標の高さより大きい画像のみリサイズする
                if ($targetHeight -gt 0 -and $originalHeight -ge $targetHeight) {
                    $magickArgs += "-resize", "x$targetHeight"
                    Write-Host "  -> 画像をリサイズします (${originalHeight}px -> ${targetHeight}px)。"
                } else {
                    Write-Host "  -> 画像のリサイズをスキップします (高さ: ${originalHeight}px)。"
                }

                $magickArgs += "-density", $targetDpi, "-quality", $Quality

                # 彩度のしきい値に基づいて、グレースケールかカラーかを判断し、それぞれの処理を追加
                if ($SATURATION -lt $SaturationThreshold) {
                    Write-Host "  -> グレースケールを検出しました。ファイル名: $newFileName";
                    $destinationPath = Join-Path $tempDirConvertedGray $newFileName
                    $magickArgs += "-colorspace", "Gray"
                    if ($PSBoundParameters.ContainsKey('GrayscaleLevel')) { $magickArgs += "-level", $GrayscaleLevel; Write-Host "  -> グレースケールコントラストを調整します (Level: $GrayscaleLevel)。" }
                }
                else {
                    Write-Host "  -> カラーを検出しました。ファイル名: $newFileName";
                    $destinationPath = Join-Path $tempDirConvertedColor $newFileName
                    if ($AutoContrast.IsPresent) { $magickArgs += "-normalize"; Write-Host "  -> カラーコントラストを自動調整します (-normalize)。" }
                    elseif ($PSBoundParameters.ContainsKey('ColorContrast')) { $magickArgs += "-brightness-contrast", $ColorContrast; Write-Host "  -> カラーコントラストを調整します (Value: $ColorContrast)。" }
                }
                $magickArgs += "$destinationPath"

                # ImageMagickコマンドを実行
                & $magick_exe @magickArgs
                
                if (-not (Test-Path $destinationPath)) {
                    Write-Warning "変換後ファイルが見つかりません: $destinationPath。このページはスキップされます。"; $skippedFiles += $file; $fileCounter++; continue
                }

                # 後のファイルサイズ比較のために、変換結果をオブジェクトとして保存
                $conversionResults += [pscustomobject]@{ 
                    FileObject = $file; FileName = $file.Name; OriginalPath = $file.FullName; OriginalSize = $file.Length
                    ConvertedPath = $destinationPath; ConvertedSize = (Get-Item -Path $destinationPath).Length; Saturation = $SATURATION
                }
                $fileCounter++
            }

            if ($conversionResults.Count -eq 0 -and $skippedFiles.Count -gt 0) { throw "全ての画像ファイルの変換に失敗しました。" }
            if ($conversionResults.Count -eq 0) { throw "処理可能な画像ファイルが変換されませんでした。" }

            # --------------------------------------------------------------------------
            # 3.1. ファイルサイズの比較と採用判断
            # --------------------------------------------------------------------------
            # 元の画像セットと変換後の画像セットの合計ファイルサイズを比較し、どちらをPDFに使用するかを決定する。
            $totalOriginalSize = ($conversionResults | Measure-Object -Property OriginalSize -Sum).Sum
            $totalConvertedSize = ($conversionResults | Measure-Object -Property ConvertedSize -Sum).Sum
            $totalOriginalSizeMB = [math]::Round($totalOriginalSize / 1MB, 2)
            $totalConvertedSizeMB = [math]::Round($totalConvertedSize / 1MB, 2)
            Write-Host "[比較] 元ファイル合計サイズ: $totalOriginalSize bytes (${totalOriginalSizeMB} MB)"
            Write-Host "[比較] 変換後ファイル合計サイズ: $totalConvertedSize bytes (${totalConvertedSizeMB} MB)"

            $useConvertedFiles = $false
            # -TotalCompressionThreshold が指定されている場合、そのしきい値に基づいて判断
            if ($PSBoundParameters.ContainsKey('TotalCompressionThreshold')) {
                if ($totalOriginalSize -gt 0) {
                    $ratio = ($totalConvertedSize / $totalOriginalSize) * 100
                    Write-Host "[比較] 圧縮率: $($ratio.ToString("F2"))% (しきい値: $TotalCompressionThreshold%)"
                    if ($ratio -lt $TotalCompressionThreshold) { $useConvertedFiles = $true }
                }
            } else {
                # デフォルトの動作: 変換によってファイルサイズが2%以上増加しない限り、変換後のファイルを使用する。
                # これにより、リサイズ等のメリットを享受しつつ、ファイルサイズが極端に増えるのを防ぐ。
                if ($totalOriginalSize -eq 0) { $useConvertedFiles = $true } # ゼロ除算を避ける
                elseif (($totalConvertedSize / [double]$totalOriginalSize) -lt 1.02) { $useConvertedFiles = $true }
            }

            if ($useConvertedFiles) {
                Write-Host "[判断] 変換後ファイルの方が小さいため、変換後の画像を使用します。"
                $filesForPdf = $conversionResults.ConvertedPath; $convertedCount = $conversionResults.Count; $originalCount = 0
            } else {
                # 元ファイルを使用する場合でも、PDF非互換形式(PNG, WEBP等)や、
                # 補正(-Deskew, -Trim等)を適用するために、JPEGとして再エンコード（パススルー処理）を行う。
                Write-Host "[判断] 元ファイルの方が小さいか、圧縮率がしきい値に満たなかったため、元の画像を使用します。"
                Write-Host "[情報] 元ファイルをPDF互換のJPEG形式に変換しています..."
                $tempDirOriginalsPassthrough = Join-Path $tempDir "originals_passthrough"
                New-Item -ItemType Directory -Path $tempDirOriginalsPassthrough | Out-Null
                $filesForPdf = @()
                $fileCounter = 0
                foreach ($result in $conversionResults) {
                    $newFileName = "{0:D4}.jpg" -f $fileCounter
                    $passthroughPath = Join-Path $tempDirOriginalsPassthrough $newFileName
                    $passthroughArgs = @("$($result.OriginalPath)")
                    if ($Deskew.IsPresent) { $passthroughArgs += "-deskew", "40%" }
                    if ($Trim.IsPresent) { $passthroughArgs += "-fuzz", $Fuzz, "-trim", "+repage" }
                    if ($result.Saturation -lt $SaturationThreshold) {
                        $passthroughArgs += "-colorspace", "Gray"
                        if ($PSBoundParameters.ContainsKey('GrayscaleLevel')) { $passthroughArgs += "-level", $GrayscaleLevel }
                    } else {
                        if ($AutoContrast.IsPresent) { $passthroughArgs += "-normalize" }
                        elseif ($PSBoundParameters.ContainsKey('ColorContrast')) { $passthroughArgs += "-brightness-contrast", $ColorContrast }
                    }
                    $passthroughArgs += "-quality", $Quality; $passthroughArgs += "$passthroughPath"
                    & $magick_exe @passthroughArgs
                    if (-not (Test-Path $passthroughPath)) {
                        Write-Warning "元ファイルのJPEG変換に失敗しました: $($result.OriginalPath)"; $skippedFiles += $result.FileObject
                    } else { $filesForPdf += $passthroughPath }
                    $fileCounter++
                }
                $convertedCount = 0; $originalCount = $conversionResults.Count
                Remove-Item -Path $tempDirConverted -Recurse -Force
            }

            # ログ用の詳細メッセージを作成
            foreach ($result in $conversionResults) {
                if ($skippedFiles.FullName -contains $result.OriginalPath) { continue }
                $logMessageDetail = ""
                if ($result.OriginalSize -gt 0) {
                    $actualRatio = ($result.ConvertedSize / $result.OriginalSize) * 100
                    $logMessageDetail = "(Ratio: $("{0:N2}" -f $actualRatio) %)"
                }
                if ($useConvertedFiles) { $logDetails += "    - $($result.FileName): Converted $logMessageDetail" }
                else { $logDetails += "    - $($result.FileName): Original $logMessageDetail" }
            }
            foreach ($file in $skippedFiles) { $logDetails += "    - $($file.Name): SKIPPED (File conversion failed)" }
        }

        if ($filesForPdf.Count -eq 0) { throw "PDFに変換する画像ファイルが見つかりませんでした。全てのページの処理に失敗した可能性があります。" }
        
        # --------------------------------------------------------------------------
        # 3.2. PDFの作成と最適化
        # --------------------------------------------------------------------------
        $pdfName = $archiveFileInfo.BaseName + ".pdf"
        $tempPdfOutputPath = Join-Path $tempDir "temp.pdf"
        $pdfOutputPath = Join-Path $archiveFileInfo.DirectoryName $pdfName
        Write-Host "PDFを作成しています: $pdfOutputPath";

        # pdfcpuに渡す引数を構築。ページサイズ設定などをここで行う。
        $importArgs = @('import', '--')

        if ($SetPageSize.IsPresent) {
            $pageSizeString = $targetPaperSizeForPdfCpu
            if ($Landscape.IsPresent -and $targetPaperSizeForPdfCpu -ne 'auto') {
                $pageSizeString += "L"
            }

            $pdfcpuPageConf = ""
            if ($targetPaperSizeForPdfCpu -eq 'auto') {
                # 画像の寸法に合わせてページサイズを自動調整
                $pdfcpuPageConf = "dim:auto"
            } else {
                # 指定された用紙サイズに、中央寄せで画像を配置
                $pdfcpuPageConf = "f:$pageSizeString, pos:c, sc:1 rel"
            }
            Write-Host "[情報] PDFページ設定を試行: $pdfcpuPageConf"
            $importArgs += $pdfcpuPageConf
        } else {
            Write-Host "[情報] PDFページ設定: pdfcpu自動設定"
        }

        $importArgs += $tempPdfOutputPath
        $importArgs += $filesForPdf
        & $pdfcpu_exe @importArgs
        if (-not (Test-Path $tempPdfOutputPath)) { throw "一時PDFファイルが作成されませんでした: $tempPdfOutputPath" }

        # PDFのビューア設定（ファイルを開いたときにウィンドウタイトルにファイル名を表示する）
        $jsonContent = '{"DisplayDocTitle": true}'
        $tempJsonPath = Join-Path $tempDir "viewerpref.json"
        $jsonContent | Out-File -FilePath $tempJsonPath -Encoding utf8
        & $pdfcpu_exe viewerpref set $tempPdfOutputPath $tempJsonPath *>$null

        # PDFのファイルサイズを最適化
        Write-Host "[情報] PDFを最適化しています..."
        & $pdfcpu_exe optimize $tempPdfOutputPath *>$null

        $finalTempPdfPath = $tempPdfOutputPath
        # -Linearize が指定され、かつqpdf.exeが見つかった場合、ウェブ表示用に最適化（リニアライズ）する
        if ($Linearize.IsPresent) {
            if ($qpdf_exe) {
                Write-Host "[情報] PDFをリニアライズ（ウェブ最適化）しています（QPDFを使用）..."
                $tempLinearizedPdfPath = Join-Path $tempDir "linearized.pdf"
                & $qpdf_exe --linearize $finalTempPdfPath $tempLinearizedPdfPath
                if ($LASTEXITCODE -ne 0) { Write-Warning "qpdf linearize の実行に失敗しました。終了コード: $LASTEXITCODE" }
                else { $finalTempPdfPath = $tempLinearizedPdfPath }
            } else {
                Write-Warning "リニアライズが指定されましたが、QPDF (qpdf.exe) が見つかりませんでした。リニアライズ処理をスキップします。"
            }
        }

        try {
            # 完成した一時PDFファイルを、最終的な出力先に移動
            Move-Item -Path $finalTempPdfPath -Destination $pdfOutputPath -Force -ErrorAction Stop
        } catch [System.IO.IOException] {
            Write-Warning "PDFファイル '$pdfOutputPath' の書き込みに失敗しました。"
            Write-Warning "ファイルが他のプログラム（PDFビューアなど）で開かれている可能性があります。"
            Write-Warning "プログラムを閉じてから、再度スクリプトを実行してください。"
            
            # ログに失敗を記録
            try {
                $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $scriptPathForLog = $MyInvocation.MyCommand.Path
                $argList = New-Object System.Collections.Generic.List[string]
                $PSBoundParameters['ArchiveFilePaths'] | ForEach-Object { $argList.Add(('"' + $_ + '"')) }
                $PSBoundParameters.GetEnumerator() | Sort-Object Key | ForEach-Object {
                    if ($_.Key -ne 'ArchiveFilePaths') {
                        if ($_.Value -is [switch]) { if ($_.Value.IsPresent) { $argList.Add("-$($_.Key)") } }
                        else { $argList.Add("-$($_.Key)"); $argList.Add(('"' + $_.Value + '"')) }
                    }
                }
                if (-not $PSBoundParameters.ContainsKey('Quality')) { $argList.Add("-Quality `"$Quality`"") }
                if (-not $PSBoundParameters.ContainsKey('SaturationThreshold')) { $argList.Add("-SaturationThreshold `"$SaturationThreshold`"") }
                if (-not $PSBoundParameters.ContainsKey('Dpi') -and $Dpi) { $argList.Add("-Dpi `"$Dpi`"") }
                if (-not $PSBoundParameters.ContainsKey('PaperSize') -and $PaperSize -and $PaperSize -ne "Custom") { $argList.Add("-PaperSize `"$PaperSize`"") }
                if (-not $PSBoundParameters.ContainsKey('Height') -and $Height) { $argList.Add("-Height `"$Height`"") }
                $commandLine = "pwsh -File `"$scriptPathForLog`" $($argList -join ' ')"

                $errorMessage = "出力先PDFファイルが使用中などの理由で、書き込みに失敗しました。"
                $logMessage = "Timestamp=`"$logTimestamp`" Status=`"Failed`" Source=`"$($archiveFileInfo.Name)`" Output=`"$pdfOutputPath`" Error=`"$errorMessage`""
                $logBlock = @"
Command: $commandLine
$logMessage
"@
                $logBlock | Add-Content -Path $logFilePath -Encoding utf8
                "`n" | Add-Content -Path $logFilePath -Encoding utf8
            } catch {
                Write-Warning "失敗ログの書き込み中にエラーが発生しました: $($_.Exception.Message)"
            }
            continue 
        } catch {
            Write-Error "PDFファイルの移動中に予期せぬエラーが発生しました: $($_.Exception.Message)"
            continue
        }

        # 元の書庫ファイルのタイムスタンプを、作成したPDFに継承させる
        $archiveLastWriteTime = $archiveFileInfo.LastWriteTime
        $archiveCreationTime = $archiveFileInfo.CreationTime
        [System.IO.File]::SetLastWriteTime($pdfOutputPath, $archiveLastWriteTime)
        [System.IO.File]::SetCreationTime($pdfOutputPath, $archiveCreationTime)

        Write-Host "----------------------------------------"
        Write-Host "成功: PDFが作成されました。" -ForegroundColor Green
        Write-Host $pdfOutputPath
        Write-Host "----------------------------------------"

        # --------------------------------------------------------------------------
        # 3.3. ログの記録
        # --------------------------------------------------------------------------
        try {
            # 実行時のコマンドラインを再現して記録
            $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $scriptPathForLog = $MyInvocation.MyCommand.Path
            $argList = New-Object System.Collections.Generic.List[string]
            $PSBoundParameters['ArchiveFilePaths'] | ForEach-Object { $argList.Add(('"' + $_ + '"')) }
            $PSBoundParameters.GetEnumerator() | Sort-Object Key | ForEach-Object {
                if ($_.Key -ne 'ArchiveFilePaths') {
                    if ($_.Value -is [switch]) { if ($_.Value.IsPresent) { $argList.Add("-$($_.Key)") } }
                    else { $argList.Add("-$($_.Key)"); $argList.Add(('"' + $_.Value + '"')) }
                }
            }
            if (-not $PSBoundParameters.ContainsKey('Quality')) { $argList.Add("-Quality `"$Quality`"") }
            if (-not $PSBoundParameters.ContainsKey('SaturationThreshold')) { $argList.Add("-SaturationThreshold `"$SaturationThreshold`"") }
            if (-not $PSBoundParameters.ContainsKey('Dpi') -and $Dpi) { $argList.Add("-Dpi `"$Dpi`"") }
            if (-not $PSBoundParameters.ContainsKey('PaperSize') -and $PaperSize -and $PaperSize -ne "Custom") { $argList.Add("-PaperSize `"$PaperSize`"") }
            if (-not $PSBoundParameters.ContainsKey('Height') -and $Height) { $argList.Add("-Height `"$Height`"") }
            $commandLine = "pwsh -File `"$scriptPathForLog`" $($argList -join ' ')"
            
            # 実行結果のサマリーを作成
            $skippedCount = if($skippedFiles) { $skippedFiles.Count } else { 0 }
            $status = if ($skippedCount -gt 0) { "Success with pages skipped" } else { "Success" }
            $settingsParts = @()
            if ($PSBoundParameters.ContainsKey('PaperSize') -and $PaperSize -ne "Custom") { $settingsParts += "PaperSize:$($PaperSize)" }
            if ($PSBoundParameters.ContainsKey('TotalCompressionThreshold')) { $settingsParts += "TCR:$($TotalCompressionThreshold)" }
            if ($Trim.IsPresent) { $settingsParts += "Trim:$($Trim.IsPresent)"; $settingsParts += "Fuzz:$($Fuzz)" }
            if ($Deskew.IsPresent) { $settingsParts += "Deskew:$($Deskew.IsPresent)" }
            if ($SplitPages.IsPresent) { $settingsParts += "SplitPages:True"; $settingsParts += "Binding:$($Binding)" }
            if ($AutoContrast.IsPresent) { $settingsParts += "AutoContrast:True" }
            elseif ($PSBoundParameters.ContainsKey('ColorContrast')) { $settingsParts += "ColorContrast:$($ColorContrast)" }
            if ($PSBoundParameters.ContainsKey('GrayscaleLevel')) { $settingsParts += "GrayscaleLevel:$($GrayscaleLevel)" }
            if ($Linearize.IsPresent) { $settingsParts += "Linearize:True" }
            if ($SetPageSize) { $settingsParts += "SetPageSize:True"; if ($Landscape) { $settingsParts += "Landscape:True" } }
            if ($SkipCompression.IsPresent) { $settingsParts += "SkipCompression:True" }
            $settingsParts += "Height:${targetHeight}px"; $settingsParts += "DPI:${targetDpi}"; $settingsParts += "Quality:${Quality}"; $settingsParts += "Saturation:${SaturationThreshold}"
            $settingsString = $settingsParts -join ', '

            # 最終的なログメッセージを組み立ててファイルに追記
            $logMessage = "Timestamp=`"$logTimestamp`" Status=`"$status`" Source=`"$($archiveFileInfo.Name)`" Output=`"$pdfOutputPath`" Images=$($imageFiles.Count) Converted=$convertedCount Originals=$originalCount Skipped=$skippedCount Settings=`"$settingsString`""
            $logBlock = @"
Command: $commandLine
$logMessage
"@
            $logBlock | Add-Content -Path $logFilePath -Encoding utf8
            if ($logDetails.Count -gt 0) { $logDetails | Add-Content -Path $logFilePath -Encoding utf8 }
            "`n" | Add-Content -Path $logFilePath -Encoding utf8
        } catch {
            Write-Warning "ログの書き込み中にエラーが発生しました: $($_.Exception.Message)"
        }
    } catch {
        Write-Error "エラー: $($_.Exception.Message)"
        Write-Error "スタックトレース: $($_.ScriptStackTrace)"
    } finally {
        if (Test-Path -Path $tempDir) {
            Write-Verbose "一時フォルダを削除しています: $tempDir"
            Remove-Item -Path $tempDir -Recurse -Force
        }
    }
}