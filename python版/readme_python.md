# JisuiArc2PDF.py

**コマンドラインが苦手な方は、こちらの詳細なガイドをご覧ください: [使い方ガイド_python.md](使い方ガイド_python.md)**

複数の画像書庫ファイルを、画像の並び順を維持したまま高品質なPDFに変換するPythonスクリプトです。
7-zipで解凍できる形式の書庫（ZIP, RARなど）をサポートしており、本をスキャンしてアーカイブ形式でまとめた自炊書籍をPDFに変換することを想定しています。

## 主な機能

（省略）

## クロスプラットフォーム対応

このPythonスクリプトは、Windows, macOS, Linuxで動作するように設計されています。

-   **スクリプト本体 (`JisuiArc2PDF.py`)**: OSに依存しないため、どの環境でも同じように動作します。
-   **バッチファイル (`JisuiArc2PDF_py.bat`)**: **Windows専用**です。macOSやLinuxでは、ターミナルから直接 `python3 JisuiArc2PDF.py ...` のようにコマンドを実行してください。

## 前提条件

### Python バージョン
- **Python 3.8 以上** の実行環境が必要です。
- OSに付属の`python`コマンドではなく、[Python公式サイト](https://www.python.org/)からダウンロードした`python3`の使用を推奨します。

### 必要な外部ツール
このスクリプトの動作には、以下の外部ツールが必要です。

-   **7-Zip**: 書庫ファイルの展開に使用 (macOS/Linuxでは `p7zip`)
-   **ImageMagick**: 画像処理（リサイズ、品質調整、色空間変換）に使用
-   **PDFCPU**: 画像ファイルからPDFを生成・最適化するために使用
-   **QPDF** (任意): PDFのリニアライズ（ウェブ最適化）処理に使用

--- 

## 環境構築

### Windows
- **方法1 (推奨: ポータブル実行):** スクリプトと同じフォルダに必要な実行ファイルを配置します。PC環境に依存せず、このフォルダだけで完結するのが利点です。
    - **Python**: [Python公式サイト](https://www.python.org/downloads/windows/)から `Windows embeddable package` をダウンロードして展開します。
    - **ImageMagick**: 公式サイトから **static 版** (`...-static-x64.zip`など) をダウンロードしてください。`magick.exe` という単一のファイルで動作し、DLLは不要なため最もシンプルです。
    - **7-Zip**: `7z.exe` と `7z.dll` を配置します。
    - **PDFCPU**: `pdfcpu.exe` を配置します。
    - **QPDF (任意)**: `--Linearize`機能を使う場合は `qpdf.exe` と関連DLLを配置します。
- **方法2 (インストールして使用):** 各ツールをPCにインストールし、実行ファイルへのPATHを通します。

### macOS (Homebrewを使用)
[Homebrew](https://brew.sh/index_ja)がインストールされている場合、ターミナルで以下のコマンドを実行して必要なツールをインストールできます。

```bash
# p7zip(7-Zip), imagemagick, qpdfをインストール
brew install p7zip imagemagick qpdf

# pdfcpuをインストール
brew install pdfcpu
```

### Linux (Debian / Ubuntu の場合)
ターミナルで以下のコマンドを実行して必要なツールをインストールできます。

```bash
# p7zip-full(7-Zip), imagemagick, qpdfをインストール
sudo apt update
sudo apt install p7zip-full imagemagick qpdf -y
```

**LinuxでのPDFCPUのインストール:**
`pdfcpu`はAPTリポジトリにないため、[公式サイトのリリース](https://pdfcpu.io/download)から最新のLinux用バイナリをダウンロードし、パスの通ったディレクトリ（例: `/usr/local/bin`）に配置してください。

--- 

## 使用方法

ターミナル（コマンドプロンプト, PowerShell, bashなど）を開き、`python3 JisuiArc2PDF.py` コマンドに処理したい書庫ファイルのパスを渡します。

```bash
# 現在のフォルダにあるすべてのZIPファイルを処理
python3 JisuiArc2PDF.py *.zip

# 特定のサブフォルダにあるすべてのRARファイルを処理
python3 JisuiArc2PDF.py "/path/to/collections/*.rar"
```

### 設定オプション

（省略）

### 総合的な使用例

```bash
python3 JisuiArc2PDF.py "C:\Scans\*.zip" -p B5 -d 300 -q 92 --Deskew --Trim
```

## バッチファイルによる対話的な実行 (JisuiArc2PDF_py.bat)

**Windowsユーザー向け**の機能です。コマンドラインに慣れていない方向けに、対話形式で設定を行えるバッチファイル `JisuiArc2PDF_py.bat` を同梱しています。

-   **使い方**: 処理したい書庫ファイル（`.rar`, `.zip`など）を、`JisuiArc2PDF_py.bat` のアイコン上にドラッグ＆ドロップします。
-   **前提条件**: この機能を利用するには、PythonがPCにインストールされ、PATHが通っている必要があります。