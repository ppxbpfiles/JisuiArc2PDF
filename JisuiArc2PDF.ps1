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
       - 役割: RARおよびZIP書庫を展開するために使用します。
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
    [Alias('h')]
    [int]$Height,

    [Parameter(Mandatory=$false)]
    [Alias('d')]
    [int]$Dpi,

    [Parameter(Mandatory=$false)]
    [Alias('p')]
    [ValidateSet('A3', 'A4', 'A5', 'B4', 'B5', 'B6')]
    [string]$PaperSize
)

# ==============================================================================
# ログファイルパス設定
# ==============================================================================
$logFilePath = Join-Path $PSScriptRoot "JisuiArc2PDF_log.txt"

# ==============================================================================
# 解像度設定の計算
# ==============================================================================
$targetHeight = 0
$targetDpi = 0

if ($PSBoundParameters.ContainsKey('Height')) {
    # -Height が指定されている場合、それを最優先する
    $targetHeight = $Height
    $targetDpi = if ($PSBoundParameters.ContainsKey('Dpi')) { $Dpi } else { 144 }
    Write-Host "[情報] 高さ指定: $targetHeight px, DPI: $targetDpi dpi"
}
elseif ($PSBoundParameters.ContainsKey('Dpi') -and $PSBoundParameters.ContainsKey('PaperSize')) {
    # -Dpi と -PaperSize が指定されている場合、高さを計算する
    $targetDpi = $Dpi
    $paperHeightMm = 0
    switch ($PaperSize) {
        'A3' { $paperHeightMm = 420 }
        'A4' { $paperHeightMm = 297 }
        'A5' { $paperHeightMm = 210 }
        'B4' { $paperHeightMm = 364 }
        'B5' { $paperHeightMm = 257 }
        'B6' { $paperHeightMm = 182 }
    }
    $targetHeight = [math]::Round(($paperHeightMm / 25.4) * $targetDpi)
    Write-Host "[情報] 計算設定: $PaperSize ($paperHeightMm mm) @ ${targetDpi} dpi -> $targetHeight px"
}
elseif ($PSBoundParameters.ContainsKey('Dpi') -or $PSBoundParameters.ContainsKey('PaperSize')) {
    # -Dpi または -PaperSize のみが指定されている場合、エラーを出力
    if ($PSBoundParameters.ContainsKey('Dpi')) {
        Write-Error "-Dpi が指定されましたが、-PaperSize が指定されていません。両方を指定してください。"
    }
    else {
        Write-Error "-PaperSize が指定されましたが、-Dpi が指定されていません。両方を指定してください。"
    }
    exit 1
}
else {
    # デフォルト設定 (A4 @ 144dpi)
    $targetDpi = 144
    $paperHeightMm = 297 # A4
    $targetHeight = [math]::Round(($paperHeightMm / 25.4) * $targetDpi)
    Write-Host "[情報] デフォルト設定: A4 @ ${targetDpi} dpi -> $targetHeight px"
}
# ==============================================================================

# ==============================================================================
# 引数チェック
# ==============================================================================
if ($PSBoundParameters.ContainsKey('ArchiveFilePaths') -eq $false -or -not $ArchiveFilePaths) {
    $helpMessage = @"
--------------------------------------------------------------------------------
JisuiArc2PDF: 書庫(7-zip対応形式)を高品質なPDFに変換します。
--------------------------------------------------------------------------------

使用法:
  pwsh -File .\JisuiArc2PDF.ps1 <対象ファイル/パターン> [オプション]

説明:
  処理対象となる書庫ファイルのパス、またはワイルドカード
  （例: '*.rar'）を一つ以上指定してください。

使用例:
  # 単一のファイルを処理
  .\JisuiArc2PDF.ps1 "MyBook.zip"

  # カレントディレクトリの全てのrarファイルを処理
  .\JisuiArc2PDF.ps1 *.rar

  # 異なるフォルダの全てのzipファイルを処理 (パスにスペース等が含まれる場合はダブルクォートで囲む)
  .\JisuiArc2PDF.ps1 "C:\Path\To\Archives\*.zip"
  .\JisuiArc2PDF.ps1 "..\OtherFolder\*.rar"

  # B5サイズ, 300dpiで高さを自動計算
  .\JisuiArc2PDF.ps1 *.rar -PaperSize B5 -Dpi 300

詳細なヘルプ:
  pwsh -Command "Get-Help .\JisuiArc2PDF.ps1 -Full"
"@
    Write-Host $helpMessage
    exit 0
}

# ==============================================================================
Write-Verbose "[診断] スクリプト開始。受信した引数: $($ArchiveFilePaths -join ', ')"

# 引数解決処理: ワイルドカード(*)を含むパスを展開し、有効なファイルパスのリストを生成する
$resolvedFilePaths = @()
foreach ($rawPath in $ArchiveFilePaths) {
    $foundPaths = Resolve-Path -Path $rawPath -ErrorAction SilentlyContinue
    if ($foundPaths) {
        foreach ($foundPath in $foundPaths) {
            $resolvedFilePaths += $foundPath.ProviderPath
        }
    } else {
        # Resolve-Path で見つからなかった場合、Get-Item -LiteralPath を試す
        try {
            $item = Get-Item -LiteralPath $rawPath -ErrorAction Stop
            if ($item.PSIsContainer) {
                # フォルダの場合は、その中のファイルを取得
                $items = Get-ChildItem -LiteralPath $rawPath -File -Recurse -ErrorAction Stop
                $resolvedFilePaths += $items.FullName
            } else {
                # ファイルの場合は、そのまま追加
                $resolvedFilePaths += $item.FullName
            }
        } catch {
            Write-Warning "指定されたパスまたはパターンに一致するファイルが見つかりません: $rawPath"
        }
    }
}
# パスリストの重複を排除し、元の変数に上書きする
$ArchiveFilePaths = $resolvedFilePaths | Sort-Object -Unique

# 処理対象のファイルが一つも見つからなかった場合はエラー終了する
if ($ArchiveFilePaths.Count -eq 0) {
    Write-Error "処理対象の書庫ファイルが一つも見つかりませんでした。パスを確認してください。"
    exit 1
}
# ==============================================================================




# 0. 前提ツールのパス解決と存在チェック
$sevenzip_exe = $null
if ($PSBoundParameters.ContainsKey('SevenZipPath') -and (Test-Path -LiteralPath $SevenZipPath -PathType Leaf)) {
    $sevenzip_exe = $SevenZipPath
} else {
    # スクリプトと同じディレクトリを検索 (OSに応じて.exeの有無を考慮)
    $cmdName = '7z'
    $localExePath = Join-Path $PSScriptRoot "$cmdName.exe"
    $localCmdPath = Join-Path $PSScriptRoot $cmdName
    if (Test-Path -LiteralPath $localExePath -PathType Leaf) {
        $sevenzip_exe = $localExePath
    } elseif (Test-Path -LiteralPath $localCmdPath -PathType Leaf) {
        $sevenzip_exe = $localCmdPath
    } else {
        $sevenzip_exe = (Get-Command $cmdName -ErrorAction SilentlyContinue).Source
    }
}

$magick_exe = $null
if ($PSBoundParameters.ContainsKey('MagickPath') -and (Test-Path -LiteralPath $MagickPath -PathType Leaf)) {
    $magick_exe = $MagickPath
} else {
    # スクリプトと同じディレクトリを検索 (OSに応じて.exeの有無を考慮)
    $cmdName = 'magick'
    $localExePath = Join-Path $PSScriptRoot "$cmdName.exe"
    $localCmdPath = Join-Path $PSScriptRoot $cmdName
    if (Test-Path -LiteralPath $localExePath -PathType Leaf) {
        $magick_exe = $localExePath
    } elseif (Test-Path -LiteralPath $localCmdPath -PathType Leaf) {
        $magick_exe = $localCmdPath
    } else {
        $magick_exe = (Get-Command $cmdName -ErrorAction SilentlyContinue).Source
    }
}

$pdfcpu_exe = $null
if ($PSBoundParameters.ContainsKey('PdfCpuPath') -and (Test-Path -LiteralPath $PdfCpuPath -PathType Leaf)) {
    $pdfcpu_exe = $PdfCpuPath
} else {
    # スクリプトと同じディレクトリを検索 (OSに応じて.exeの有無を考慮)
    $cmdName = 'pdfcpu'
    $localExePath = Join-Path $PSScriptRoot "$cmdName.exe"
    $localCmdPath = Join-Path $PSScriptRoot $cmdName
    if (Test-Path -LiteralPath $localExePath -PathType Leaf) {
        $pdfcpu_exe = $localExePath
    } elseif (Test-Path -LiteralPath $localCmdPath -PathType Leaf) {
        $pdfcpu_exe = $localCmdPath
    } else {
        $pdfcpu_exe = (Get-Command $cmdName -ErrorAction SilentlyContinue).Source
    }
}

if (-not $magick_exe) {
    Write-Error "ImageMagick (magick.exe) が見つかりません。パスを指定するか、環境変数PATHに登録してください。";
    exit 1
}
if (-not $sevenzip_exe) {
    Write-Error "7-Zip (7z.exe) が見つかりません。パスを指定するか、環境変数PATHに登録してください。";
    exit 1
}
if (-not $pdfcpu_exe) {
    Write-Error "PDFCPU (pdfcpu.exe) が見つかりません。パスを指定するか、環境変数PATHに登録してください。";
    exit 1
}

foreach ($ArchiveFilePath in $ArchiveFilePaths) {
    Write-Output "[診断] メインループ開始: $ArchiveFilePath"
    # 1. 入力ファイルの検証
    $archiveFileInfo = Get-Item -LiteralPath $ArchiveFilePath -ErrorAction SilentlyContinue
    if (-not $archiveFileInfo) {
        Write-Warning "指定された書庫ファイルが見つかりません: $ArchiveFilePath";
        continue
    }

    Write-Host "========================================";
    Write-Host "処理中: $($archiveFileInfo.Name)";
    Write-Host "========================================";

    # 2. 一時フォルダの作成
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
    $tempDirConverted = Join-Path $tempDir "converted"
    $tempDirConvertedColor = Join-Path $tempDirConverted "color"
    $tempDirConvertedGray = Join-Path $tempDirConverted "gray"

    New-Item -ItemType Directory -Path $tempDir | Out-Null
    New-Item -ItemType Directory -Path $tempDirConverted | Out-Null
    New-Item -ItemType Directory -Path $tempDirConvertedColor | Out-Null
    New-Item -ItemType Directory -Path $tempDirConvertedGray | Out-Null
    Write-Verbose "一時フォルダを作成しました: $tempDir"
    Write-Verbose "変換用フォルダを作成しました: $tempDirConverted"

    try {
        # 3. アーカイブの拡張子に応じて展開
        Write-Host "アーカイブを展開しています: $($archiveFileInfo.Name)";
        # Call Operator (&) を使用して、特殊文字を含むファイルパスを正しく処理する
        & $sevenzip_exe e "$($archiveFileInfo.FullName)" "-o$tempDir" -y

        # 4. 画像ファイルを検索し、自然順で並べ替え
        Write-Host "画像ファイルを検索し、並べ替えています...";
        $allFiles = Get-ChildItem -Path $tempDir -Recurse -File
        $imageFiles = @()
        foreach ($file in $allFiles) {
            # magick identify で画像ファイルかを確認。エラー出力を抑制し、終了コードで判断する。
            & $magick_exe identify "$($file.FullName)" *>$null
            if ($LASTEXITCODE -eq 0) {
                $imageFiles += $file
                Write-Verbose "画像として認識: $($file.Name)"
            } else {
                Write-Verbose "画像ではないためスキップ: $($file.Name)"
            }
        }

        # 自然順でソート
        $imageFiles = $imageFiles | Sort-Object @{Expression={
            [regex]::Replace($_.Name, '\d+', { param($match) $match.Value.PadLeft(20) })
        }}

        if ($imageFiles.Count -eq 0) {
            throw "アーカイブ内に処理可能な画像ファイルが見つかりませんでした。";
        }
        
        # 5. 画像を連番ファイル名で変換・保存
        Write-Host "画像を変換し、一時ファイルに保存しています...";
        $convertedImagePaths = @()
        $fileCounter = 0
        foreach ($file in $imageFiles) {
            Write-Host "Processing $($file.Name)...";
            
            # 画像の平均彩度を計算
            $SATURATION_STR = & $magick_exe "$($file.FullName)" -colorspace HSL -channel G -separate +channel -format "%[mean]" info:
            $SATURATION = [double]$SATURATION_STR / 65535.0

            # 連番のファイル名を生成 (例: 0000.jpg)
            $newFileName = "{0:D4}.jpg" -f $fileCounter
            
            if ($SATURATION -lt $SaturationThreshold) {
                # グレースケール画像として処理
                Write-Host "  -> Grayscale detected. Saving as $newFileName";
                $destinationPath = Join-Path $tempDirConvertedGray $newFileName
                & $magick_exe "$($file.FullName)" -resize "x$targetHeight" -density $targetDpi -quality $Quality -colorspace Gray "$destinationPath"
            }
            else {
                # カラー画像として処理
                Write-Host "  -> Color detected. Saving as $newFileName";
                $destinationPath = Join-Path $tempDirConvertedColor $newFileName
                & $magick_exe "$($file.FullName)" -resize "x$targetHeight" -density $targetDpi -quality $Quality "$destinationPath"
            }

            if (Test-Path $destinationPath) {
                $convertedImagePaths += $destinationPath
            }
            $fileCounter++
        }

        # 6. PDFの作成 (PDFCPU を使用)
        $pdfName = $archiveFileInfo.BaseName + ".pdf"
        $tempPdfName = "temp_" + [System.Guid]::NewGuid().ToString() + ".pdf"
        
        $tempPdfOutputPath = Join-Path $tempDir $tempPdfName
        $pdfOutputPath = Join-Path $archiveFileInfo.DirectoryName $pdfName
        Write-Host "PDFを作成しています: $pdfOutputPath";

        # PDFCPU で画像からPDFを作成
        $importArgs = @('import', '--', $tempPdfOutputPath) + $convertedImagePaths
        $stderrPath = Join-Path $tempDir "pdfcpu_import_stderr.txt"
        $stdoutPath = Join-Path $tempDir "pdfcpu_import_stdout.txt"
        & $pdfcpu_exe @importArgs 2> $stderrPath 1> $stdoutPath
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "pdfcpu import の実行に失敗しました。終了コード: $LASTEXITCODE"
        }

        # エラー出力があれば表示
        if (Test-Path $stderrPath) {
            $importStderr = Get-Content $stderrPath -Raw
            if ($importStderr -and $importStderr.Trim() -ne "") {
                Write-Host "pdfcpu import の stderr:";
                Write-Host $importStderr;
            }
        }
        
        # 標準出力があれば表示
        if (Test-Path $stdoutPath) {
            $importStdout = Get-Content $stdoutPath -Raw
            if ($importStdout -and $importStdout.Trim() -ne "") {
                Write-Host "pdfcpu import の stdout:";
                Write-Host $importStdout;
            }
        }

        if (-not (Test-Path $tempPdfOutputPath)) {
            throw "一時PDFファイルが作成されませんでした: $tempPdfOutputPath"
        }

        # PDFCPU でビューアの設定を一時PDFに対して変更
        $jsonContent = '{"DisplayDocTitle": true}'
        $tempJsonPath = Join-Path $tempDir "viewerpref.json"
        $jsonContent | Out-File -FilePath $tempJsonPath -Encoding utf8
        $viewerPrefArgs = @('viewerpref', 'set', $tempPdfOutputPath, $tempJsonPath)
        & $pdfcpu_exe @viewerPrefArgs
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "pdfcpu viewerpref set の実行に失敗しました。終了コード: $LASTEXITCODE"
        }
        Remove-Item -Path $tempJsonPath -Force

        # PDFCPU で最適化を一時PDFに対して実行
        $optimizeArgs = @('optimize', $tempPdfOutputPath)
        & $pdfcpu_exe @optimizeArgs
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "pdfcpu optimize の実行に失敗しました。終了コード: $LASTEXITCODE"
        }

        # 処理済みの一次PDFファイルを最終的な場所に移動
        Move-Item -Path $tempPdfOutputPath -Destination $pdfOutputPath -Force

        # タイムスタンプをアーカイブファイルに合わせる
        $archiveLastWriteTime = $archiveFileInfo.LastWriteTime
        $archiveCreationTime = $archiveFileInfo.CreationTime
        [System.IO.File]::SetLastWriteTime($pdfOutputPath, $archiveLastWriteTime)
        [System.IO.File]::SetCreationTime($pdfOutputPath, $archiveCreationTime)

        Write-Host "----------------------------------------";
        Write-Host "成功: PDFが作成されました。";
        Write-Host $pdfOutputPath
        Write-Host "----------------------------------------";

        # ログへの書き込み
        try {
            $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            
            # 基本のログメッセージを作成
            $logParts = @(
                "$logTimestamp",
                "Source: $($archiveFileInfo.Name)"
            )

            # PaperSizeが指定されていた場合、ログに追加
            if ($PSBoundParameters.ContainsKey('PaperSize')) {
                $logParts += "PaperSize: $PaperSize"
            }

            # 残りの設定を追加
            $logParts += @(
                "Height: ${targetHeight}px",
                "DPI: ${targetDpi}",
                "Quality: ${Quality}",
                "Saturation: ${SaturationThreshold}",
                "Output: $pdfOutputPath"
            )
            
            # 配列をカンマ区切りで結合
            $logMessage = $logParts -join ', '

            Add-Content -Path $logFilePath -Value $logMessage
            Write-Verbose "設定をログファイルに書き込みました: $logFilePath"
        } catch {
            Write-Warning "ログファイルへの書き込みに失敗しました: $($_.Exception.Message)"
        }

    }
    catch {
        Write-Error "エラーが発生しました: $($_.Exception.Message)";
    }
    finally {
        # 7. 一時フォルダのクリーンアップ
        if (Test-Path $tempDir) {
            Write-Verbose "一時フォルダをクリーンアップしています: $tempDir"
            Remove-Item -Path $tempDir -Recurse -Force
        }
    }
}
Write-Host "========================================";
Write-Host "すべての処理が完了しました。";
Write-Host "========================================";