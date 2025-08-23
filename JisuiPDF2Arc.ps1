<#
.SYNOPSIS
    PDFファイルから画像を一括抽出し、ZIP書庫に変換します。

.DESCRIPTION
    JisuiArc2PDF.ps1の逆変換を行うスクリプトです。
    指定されたPDFファイルから画像を抽出し、元のPDFファイルと同じ名前のZIP書庫を作成します。

.PARAMETER PdfFilePaths
    処理対象となるPDFのファイルパス、またはワイルドカードを含むパターン。

.PARAMETER SevenZipPath
    7z.exe への絶対パスを明示的に指定します。

.PARAMETER PdfCpuPath
    pdfcpu.exe への絶対パスを明示的に指定します。

.PARAMETER LogPath
    ログファイルの出力先を指定します。デフォルトではスクリプトと同じフォルダに作成されます。

.EXAMPLE
    # 単一のPDFをZIPに変換
    .\JisuiPDF2Arc.ps1 "MyBook.pdf"

.EXAMPLE
    # ワイルドカードを使い、カレントディレクトリの全てのPDFを変換
    .\JisuiPDF2Arc.ps1 *.pdf
#>
[CmdletBinding()]
param(
    [Parameter(Position=0, Mandatory=$true)]
    [string[]]$PdfFilePaths,

    [Parameter(Mandatory=$false)]
    [string]$SevenZipPath,

    [Parameter(Mandatory=$false)]
    [string]$PdfCpuPath,

    [Parameter(Mandatory=$false)]
    [string]$LogPath
)

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
        $logFilePath = Join-Path $resolvedLogPath "JisuiPDF2Arc_log.txt"
    } else {
        $logFilePath = $resolvedLogPath
    }
} else {
    $logFilePath = Join-Path $PSScriptRoot "JisuiPDF2Arc_log.txt"
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
# 前提ツールのパス解決と存在チェック (JisuiArc2PDF.ps1から流用)
# ==============================================================================
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

if (-not $sevenzip_exe) {
    Write-Error "7-Zip (7z.exe) が見つかりません。パスを指定するか、環境変数PATHに登録してください。"
    exit 1
}
if (-not $pdfcpu_exe) {
    Write-Error "PDFCPU (pdfcpu.exe) が見つかりません。パスを指定するか、環境変数PATHに登録してください。"
    exit 1
}

# ==============================================================================
# 引数解決処理 (JisuiArc2PDF.ps1から流用)
# ==============================================================================
$resolvedFilePaths = @()
foreach ($rawPath in $PdfFilePaths) {
    $foundPaths = Resolve-Path -Path $rawPath -ErrorAction SilentlyContinue
    if ($foundPaths) {
        foreach ($foundPath in $foundPaths) {
            if ($foundPath.ProviderPath.ToLower().EndsWith(".pdf")) {
                $resolvedFilePaths += $foundPath.ProviderPath
            }
        }
    } else {
        try {
            $item = Get-Item -LiteralPath $rawPath -ErrorAction Stop
            if ($item.PSIsContainer) {
                $items = Get-ChildItem -LiteralPath $rawPath -File -Recurse -Filter *.pdf -ErrorAction Stop
                $resolvedFilePaths += $items.FullName
            } elseif ($item.Name.ToLower().EndsWith(".pdf")) {
                $resolvedFilePaths += $item.FullName
            }
        } catch {
            Write-Warning "指定されたパスまたはパターンに一致するPDFファイルが見つかりません: $rawPath"
        }
    }
}
$resolvedFilePaths = $resolvedFilePaths | Sort-Object -Unique

if ($resolvedFilePaths.Count -eq 0) {
    Write-Error "処理対象のPDFファイルが一つも見つかりませんでした。パスを確認してください。"
    exit 1
}

# ==============================================================================
# メイン処理
# ==============================================================================
foreach ($pdfPath in $resolvedFilePaths) {
    $pdfInfo = Get-Item -LiteralPath $pdfPath
    Write-Host "========================================"
    Write-Host "Processing: $($pdfInfo.Name)"
    Write-Host "========================================"

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    try {
        # 1. PDFから画像を抽出
        Write-Host "  -> Extracting images from PDF..."
        & $pdfcpu_exe extract -mode image "$($pdfInfo.FullName)" "$tempDir"
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to extract images from $($pdfInfo.Name)"
        }

        # 2. 抽出されたファイルを確認
        $extractedFiles = Get-ChildItem -Path $tempDir -Recurse -File
        if ($extractedFiles.Count -eq 0) {
            Write-Warning "No images found in or extracted from $($pdfInfo.Name). Skipping."
            continue
        }

        # 3. ZIP書庫を作成 (出力先: ./pdf2arc_converted/)
        $convertedOutputDir = Join-Path (Get-Location) "pdf2arc_converted"
        if (-not (Test-Path -Path $convertedOutputDir -PathType Container)) {
            New-Item -ItemType Directory -Path $convertedOutputDir | Out-Null
        }
        $zipFileName = $pdfInfo.BaseName + ".zip"
        $zipOutputPath = Join-Path $convertedOutputDir $zipFileName
        Write-Host "  -> Creating ZIP archive: $zipOutputPath"
        
        $sourceForZip = Join-Path $tempDir "*"
        & $sevenzip_exe a -tzip "$zipOutputPath" "$sourceForZip"
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create ZIP archive for $($pdfInfo.Name)"
        }

        Write-Host "  -> Successfully created $zipFileName"

        # 4. ログへの書き込み
        try {
            $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $commandLine = $MyInvocation.Line
            $status = "Success"
            $imageCount = $extractedFiles.Count

            $logMessage = "Timestamp=`"$logTimestamp`" Status=`"$status`" Source=`"$($pdfInfo.Name)`" Output=`"$zipOutputPath`" Images=$imageCount"

            $fullLogContent = @("Command: $commandLine", $logMessage)
            Add-Content -Path $logFilePath -Value ($fullLogContent -join "`n") -NoNewline
            Add-Content -Path $logFilePath -Value "`n`n" # Add two newlines for spacing

            Write-Verbose "設定をログファイルに書き込みました: $logFilePath"
        } catch {
            Write-Warning "ログファイルへの書き込みに失敗しました: $($_.Exception.Message)"
        }

    } catch {
        Write-Error $_.Exception.Message
    } finally {
        # 5. 一時フォルダをクリーンアップ
        if (Test-Path $tempDir) {
            Write-Verbose "Cleaning up temporary directory: $tempDir"
            Remove-Item -Path $tempDir -Recurse -Force
        }
    }
}

Write-Host "========================================"
Write-Host "All processing complete."
Write-Host "========================================"