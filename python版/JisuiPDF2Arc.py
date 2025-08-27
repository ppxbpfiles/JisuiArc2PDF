import argparse
import sys
import os
import shutil
import glob
import subprocess
import tempfile
import datetime
from pathlib import Path

# --- Helper Functions ---
def find_tool(tool_key: str, specified_path: str | None) -> str:
    """必要な外部ツールを、指定されたパス、スクリプトディレクトリ、環境変数PATHの順で検索します。"""
    is_windows = sys.platform == "win32"
    tool_filenames = {
        'sevenzip': '7z.exe' if is_windows else '7z',
        'pdfcpu': 'pdfcpu.exe' if is_windows else 'pdfcpu',
    }
    tool_name = tool_filenames[tool_key]

    if specified_path and os.path.exists(specified_path):
        return specified_path
    try:
        script_dir = os.path.dirname(os.path.realpath(__file__))
        local_path = os.path.join(script_dir, tool_name)
        if os.path.exists(local_path):
            return local_path
    except NameError:
        pass
    system_path = shutil.which(tool_name)
    if system_path:
        return system_path
    raise FileNotFoundError(
        f"{tool_name} not found. Please provide it via arguments, place it in the script's directory, or add it to your system's PATH."
    )

def process_pdf(pdf_path: str, args: argparse.Namespace, tools: dict, log_file_path: Path | None):
    """単一のPDFファイルを処理します。"""
    pdf_path_obj = Path(pdf_path)
    print("\n" + "=" * 40)
    print(f"Processing: {pdf_path_obj.name}")
    print("=" * 40)

    # 一時的な作業ディレクトリを生成
    temp_dir = Path(tempfile.mkdtemp(prefix="JisuiPDF2Arc_"))
    if args.Verbose:
        print(f"Created temporary directory: {temp_dir}")

    try:
        # 1. PDFから画像を抽出
        # pdfcpu extract コマンドを使い、PDF内の全ての画像を一時ディレクトリに書き出す。
        print("  -> Extracting images from PDF...")
        subprocess.run(
            [tools['pdfcpu'], "extract", "-mode", "image", str(pdf_path_obj), str(temp_dir)],
            check=True, capture_output=True, timeout=300
        )

        # 2. 抽出されたファイルを確認
        # 画像が1枚も抽出されなかった場合は、警告を表示して次のファイルの処理に進む。
        extracted_files = list(temp_dir.glob('**/*'))
        # ディレクトリを除外し、ファイルのみをリストアップする
        extracted_files = [f for f in extracted_files if f.is_file()]
        
        if not extracted_files:
            print(f"Warning: No images found in or extracted from {pdf_path_obj.name}. Skipping.")
            return # このPDFの処理を中断し、次のファイルへ

        # 3. 出力ディレクトリを作成
        # 出力先として、スクリプトの実行場所に `pdf2arc_converted` フォルダがなければ作成する。
        converted_output_dir = Path.cwd() / "pdf2arc_converted"
        converted_output_dir.mkdir(parents=True, exist_ok=True)

        # 4. ZIP書庫のパスを定義
        zip_filename = pdf_path_obj.stem + ".zip"
        zip_output_path = converted_output_dir / zip_filename
        print(f"  -> Creating ZIP archive: {zip_output_path}")

        # 5. 7-ZipでZIP書庫を作成
        # 7z a コマンドで、一時ディレクトリ内のすべてのファイルを圧縮対象にする
        source_for_zip = temp_dir / '*'        
        # 7zコマンドを実行
        subprocess.run(
            [tools['sevenzip'], "a", "-tzip", str(zip_output_path), str(source_for_zip)],
            check=True, capture_output=True
        )

        print(f"  -> Successfully created {zip_filename}")

        # 6. ログへの書き込み
        # 処理結果をログファイルに追記する。
        if log_file_path:
            try:
                log_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                command_line = f"python {' '.join(sys.argv)}"
                status = "Success"
                image_count = len(extracted_files)

                log_message = (
                    f'Timestamp="{log_timestamp}" '
                    f'Status="{status}" '
                    f'Source="{pdf_path_obj.name}" '
                    f'Output="{zip_output_path}" '
                    f'Images={image_count}'
                )
                full_log_content = [f"Command: {command_line}", log_message]
                # 追記モード、UTF-8でファイルを開く
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write("\n".join(full_log_content) + "\n\n")
                if args.Verbose:
                    print(f"Wrote settings to log file: {log_file_path}")
            except Exception as e:
                print(f"Warning: Failed to write to log file: {e}", file=sys.stderr)
                
    except subprocess.CalledProcessError as e:
        error_msg = f"Error during subprocess execution for: {' '.join(e.cmd)}"
        print(error_msg, file=sys.stderr)
        if e.stderr:
            print(f"Stderr: {e.stderr.decode(errors='ignore')}", file=sys.stderr)
        # エラーをログに記録
        if log_file_path:
            try:
                log_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                status = "Failed"
                log_message = (
                    f'Timestamp="{log_timestamp}" '
                    f'Status="{status}" '
                    f'Source="{pdf_path_obj.name}" '
                    f'Error="{error_msg}"'
                )
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write(f"Command: python {' '.join(sys.argv)}\n{log_message}\n\n")
            except Exception as log_e:
                print(f"Warning: Failed to write error to log file: {log_e}", file=sys.stderr)
    except Exception as e:
        error_msg = f"An unexpected error occurred: {e}"
        print(error_msg, file=sys.stderr)
        # エラーをログに記録
        if log_file_path:
             try:
                log_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                status = "Failed"
                log_message = (
                    f'Timestamp="{log_timestamp}" '
                    f'Status="{status}" '
                    f'Source="{pdf_path_obj.name}" '
                    f'Error="{error_msg}"'
                )
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write(f"Command: python {' '.join(sys.argv)}\n{log_message}\n\n")
             except Exception as log_e:
                print(f"Warning: Failed to write error to log file: {log_e}", file=sys.stderr)
    finally:
        # 7. 一時フォルダをクリーンアップ
        # 処理が成功しても失敗しても、必ず一時ファイルを削除する。
        if args.Verbose:
            print(f"Cleaning up temporary directory: {temp_dir}")
        # shutil.rmtree を使ってディレクトリとその中身を削除する
        shutil.rmtree(temp_dir, ignore_errors=True)

def main():
    """引数を解析し、変換処理を実行するメイン関数です。"""
    parser = argparse.ArgumentParser(
        description="Extracts images from a PDF and creates a ZIP archive in a 'pdf2arc_converted' folder.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("pdf_file_paths", metavar="PDF_PATHS", nargs='+',
                        help="Paths or glob patterns for PDF files.")
    parser.add_argument("--SevenZipPath", help="Path to 7z executable.")
    parser.add_argument("--PdfCpuPath", help="Path to pdfcpu executable.")
    parser.add_argument("--LogPath", help="Path for logging.")
    parser.add_argument("-v", "--Verbose", action="store_true", help="Enable verbose output.")
    args = parser.parse_args()

    # --- 外部ツールのパス解決 ---
    try:
        tools = {
            'sevenzip': find_tool('sevenzip', args.SevenZipPath),
            'pdfcpu': find_tool('pdfcpu', args.PdfCpuPath),
        }
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    # --- ログファイルパスの処理 ---
    log_file_path = None
    if args.LogPath:
        log_path = Path(args.LogPath).resolve()
        # 指定されたLogPathがディレクトリの場合、その中にデフォルトのログファイル名で作成する
        if log_path.is_dir():
            log_file_path = log_path / "JisuiPDF2Arc_py_log.txt"
        else:
            # ファイルパスとして指定された場合は、それを直接使用する
            log_file_path = log_path
    else:
        # デフォルトのログファイル場所: スクリプトと同じディレクトリ、または現在の作業ディレクトリ
        try:
            script_dir = Path(os.path.dirname(os.path.realpath(__file__)))
            log_file_path = script_dir / "JisuiPDF2Arc_py_log.txt"
        except NameError:
            # __file__が利用できない場合 (例: 対話モードでの実行時)
            log_file_path = Path.cwd() / "JisuiPDF2Arc_py_log.txt"

    # ログファイルのためのディレクトリが存在することを確認
    if log_file_path:
        try:
            log_file_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"Error: Could not create log directory for {log_file_path}. {e}", file=sys.stderr)
            log_file_path = None # ディレクトリを作成できない場合はログ機能を無効にする

    # --- ツールのパスを詳細表示（オプション） ---
    if args.Verbose:
        print("--- Tool Paths ---")
        for tool, path in tools.items():
            if path: print(f"{tool}: {path}")
        print("--------------------")

    # --- 入力ファイルパスの解決 ---
    resolved_files = []
    for pattern in args.pdf_file_paths:
        # glob.glob を使ってワイルドカードを展開し、一致するファイルを見つける
        # `recursive=True` は `**` のような再帰的なパターンを有効にする
        matched_files = glob.glob(pattern, recursive=True)
        for file_path in matched_files:
            # 解決されたパスがファイルであり、かつ拡張子が .pdf であることを確認（大文字小文字を区別しない）
            if os.path.isfile(file_path) and file_path.lower().endswith('.pdf'):
                resolved_files.append(file_path)
    
    # 重複を削除し、処理順序を安定させるためにソートする
    unique_files = sorted(list(set(resolved_files)))

    if not unique_files:
        print("Error: No input PDF files found.", file=sys.stderr)
        sys.exit(1)

    # --- 処理対象ファイルを詳細表示（オプション） ---
    if args.Verbose:
        print(f"--- Found {len(unique_files)} PDFs to process ---")
        for f in unique_files:
            print(f"  - {f}")
        print("--------------------")

    # --- 各PDFファイルを処理 ---
    for pdf_path in unique_files:
        process_pdf(pdf_path, args, tools, log_file_path)

    # --- 最終メッセージ ---
    print("\n" + "=" * 40)
    print("All processing complete.")
    print("=" * 40)

if __name__ == "__main__":
    main()