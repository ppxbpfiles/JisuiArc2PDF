import argparse
import sys
import os
import shutil
import glob
import subprocess
import tempfile
import re
import datetime
from pathlib import Path
import json

# --- ヘルパー関数 ---

def find_tool(tool_key: str, specified_path: str | None) -> str:
    """
    必要な外部ツールの実行可能ファイルパスを検索します。
    
    検索順序:
    1. ユーザーが引数で指定したパス
    2. スクリプトと同じディレクトリ
    3. システムのPATH環境変数
    
    見つからない場合は FileNotFoundError を送出します。
    """
    is_windows = sys.platform == "win32"
    tool_filenames = {
        'sevenzip': '7z.exe' if is_windows else '7z',
        'magick': 'magick.exe' if is_windows else 'magick',
        'pdfcpu': 'pdfcpu.exe' if is_windows else 'pdfcpu',
        'qpdf': 'qpdf.exe' if is_windows else 'qpdf',
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

def natural_sort_key(s: str) -> list:
    """
    ファイル名の自然順ソート用のキーを生成します。
    数字を数値として比較し、一般的な区切り文字を無視します。
    例: ['1.jpg', '2.jpg', '10.jpg'] -> 正しい順序でソート
    """
    def convert(text):
        # 数字の場合は、0埋めして文字列に変換することで、文字列比較でも自然な順序になるようにする
        if text.isdigit():
            return text.zfill(10) # 10桁に0埋め (必要に応じて桁数を調整)
        else:
            # 数字以外の文字は、区切り文字を除去して小文字に変換
            return re.sub(r'[-_@#]', '', text.lower())
    # 分割によって生じる空文字列をフィルタリング
    return [convert(c) for c in re.split(r'(\d+)', s) if c]

def is_image(file_path: str, magick_exe: str) -> bool:
    """
    ImageMagick の 'identify' コマンドを使用して、ファイルが画像かどうかを判定します。
    """
    try:
        subprocess.run([magick_exe, "identify", str(file_path)], check=True, capture_output=True, timeout=15)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False

def get_image_height(file_path: str, magick_exe: str) -> int | None:
    """
    ImageMagick を使用して画像の高さ（ピクセル）を取得します。
    取得に失敗した場合は None を返します。
    """
    try:
        result = subprocess.run([magick_exe, "identify", "-format", "%h", str(file_path)], check=True, capture_output=True, text=True)
        return int(result.stdout)
    except (subprocess.CalledProcessError, FileNotFoundError, ValueError) as e:
        print(f"Warning: Could not get height for {os.path.basename(file_path)}. Error: {e}", file=sys.stderr)
        return None

def get_image_saturation(file_path: str, magick_exe: str) -> float:
    """
    画像の平均彩度を取得します。
    HSL色空間の緑チャンネルの平均値を0.0(無彩色)〜1.0(鮮やか)の範囲で返します。
    失敗した場合はデフォルトの0.5（カラーとみなす）を返します。
    """
    try:
        result = subprocess.run(
            [magick_exe, str(file_path), "-colorspace", "HSL", "-channel", "G", "-separate", "+channel", "-format", "%[mean]", "info:"],
            check=True, capture_output=True, text=True
        )
        return float(result.stdout) / 65535.0
    except (subprocess.CalledProcessError, FileNotFoundError, ValueError):
        return 0.5 # Default to color if saturation check fails

def calculate_target_height(args: argparse.Namespace) -> tuple[int, int]:
    """
    ユーザーの引数に基づいて、変換後の目標画像の高さ（ピクセル）とDPIを計算します。
    """
    paper_heights_mm = {
        'A0': 1189, 'A1': 841, 'A2': 594, 'A3': 420, 'A4': 297, 'A5': 210, 'A6': 148, 'A7': 105,
        'B0': 1414, 'B1': 1000, 'B2': 707, 'B3': 500, 'B4': 364, 'B5': 257, 'B6': 182, 'B7': 128
    }
    if args.Height:
        target_height = args.Height
        target_dpi = args.Dpi if args.Dpi else 144
        print(f"[Info] Height specified: {target_height}px, DPI: {target_dpi}dpi")
        return target_height, target_dpi
    if args.PaperSize and args.Dpi:
        paper_height = paper_heights_mm[args.PaperSize]
        target_dpi = args.Dpi
        target_height = int(round((paper_height / 25.4) * target_dpi))
        print(f"[Info] Paper size specified: {args.PaperSize} @ {target_dpi}dpi -> {target_height}px")
        return target_height, target_dpi
    if args.PaperSize or args.Dpi:
        print("Error: --PaperSize and --Dpi must be used together.", file=sys.stderr)
        sys.exit(1)
    # Default
    target_dpi = args.Dpi if args.Dpi else 144
    paper_size = args.PaperSize if args.PaperSize else 'A4'
    paper_height = paper_heights_mm[paper_size]
    target_height = int(round((paper_height / 25.4) * target_dpi))
    print(f"[Info] Using default: {paper_size} @ {target_dpi}dpi -> {target_height}px")
    return target_height, target_dpi

def process_archive(archive_path: str, args: argparse.Namespace, tools: dict, log_file_path: Path | None, target_height: int, target_dpi: int):
    """単一のアーカイブファイルを処理します。"""
    print("\n" + "=" * 40)
    print(f"Processing: {os.path.basename(archive_path)}")
    print("=" * 40)

    # 一時作業用ディレクトリを作成します。処理が完了すると自動的に削除されます。
    temp_dir = Path(tempfile.mkdtemp(prefix="JisuiArc2PDF_"))
    
    try:
        if args.Verbose:
            print(f"Created temporary directory: {temp_dir}")

        # --- アーカイブの展開 ---
        # 7-Zip を使用してアーカイブを一時ディレクトリに展開します。
        print("Extracting archive...")
        subprocess.run(
            [tools['sevenzip'], "e", archive_path, f"-o{temp_dir}", "-y"],
            check=True, capture_output=True, timeout=300
        )

        # --- 画像ファイルの特定と並び替え ---
        # 一時ディレクトリ内を再帰的に探索し、ImageMagickで画像ファイルを識別します。
        # 非画像ファイルは自動的に無視されます。
        # ファイル名に含まれる数字を数値として解釈して自然順ソートを行い、正しいページ順序を維持します。
        print("Finding and sorting image files...")
        image_files = []
        # Walk through all files and identify images
        for root, _, files in os.walk(temp_dir):
            # Do not search in our own conversion output directory
            if "converted" in Path(root).parts or "passthrough" in Path(root).parts:
                continue
            for file in files:
                full_path = Path(root) / file
                if is_image(str(full_path), tools['magick']):
                    image_files.append(full_path)
                elif args.Verbose:
                    print(f"  Skipping non-image file: {file}")
        
        image_files.sort(key=lambda f: natural_sort_key(f.name))

        if not image_files:
            raise ValueError("No image files found in the archive.")

        if args.Verbose:
            print(f"Found {len(image_files)} images.")

        files_for_pdf = []
        conversion_results = []
        skipped_files = []
        original_count = 0
        converted_count = 0
        use_converted = False

        # --- 画像変換処理 ---
        if args.SkipCompression:
            # --- スキップ圧縮モード ---
            # リサイズや品質調整などの最適化処理をスキップします。
            # JPEG以外のファイルは、画質劣化を最小限に抑えるためJPEGに変換されます。
            # JPEGファイルはそのままPDF化されます。
            print("[Info] --SkipCompression: Skipping optimization, converting non-JPEGs to JPEG.")
            jpeg_extensions = {'.jpg', '.jpeg', '.jfif', '.jpe'}
            passthrough_dir = temp_dir / "sc_passthrough"
            passthrough_dir.mkdir()
            
            for i, img_path in enumerate(image_files):
                if img_path.suffix.lower() in jpeg_extensions:
                    files_for_pdf.append(str(img_path))
                    if args.Verbose: print(f"  -> {img_path.name}: Is JPEG, no conversion needed.")
                else:
                    passthrough_path = passthrough_dir / f"{i:04d}.jpg"
                    if args.Verbose: print(f"  -> {img_path.name}: Is not JPEG, converting.")
                    try:
                        subprocess.run([tools['magick'], str(img_path), str(passthrough_path)], check=True, capture_output=True)
                        files_for_pdf.append(str(passthrough_path))
                    except subprocess.CalledProcessError as e:
                        print(f"Warning: Failed to convert {img_path.name} to JPEG. It will be skipped. Error: {e}", file=sys.stderr)
                        skipped_files.append(img_path)
            original_count = len(image_files)
        else:
            # --- 通常の変換ロジック ---
            # 画像の画質とファイルサイズのバランスを最適化します。
            # 各画像は、設定に基づいてリサイズ、傾き補正、余白除去、コントラスト調整などが行われます。
            # グレースケールとカラー画像は自動判別され、それぞれに最適な処理が適用されます。
            print("Converting images...")
            converted_dir = temp_dir / "converted"
            converted_dir.mkdir()

            # --- 個別画像の変換 ---
            total_images = len(image_files)
            for i, img_path in enumerate(image_files):
                print(f"\r[ {i+1:3d}/{total_images:3d} ] 処理中: {img_path.name}", end='', flush=True)
                if args.Verbose: print() # Verboseモードでは改行して詳細を表示
                
                # --- 画像情報の取得 ---
                # 画像の高さと彩度を取得し、リサイズの必要性とグレースケール判定を行います。
                original_height = get_image_height(str(img_path), tools['magick'])
                if original_height is None:
                    skipped_files.append(img_path)
                    if args.Verbose: print(f"  -> 高さの取得に失敗したためスキップ")
                    continue

                # --- ImageMagickコマンドの構築 ---
                # ImageMagick (`magick`) のコマンドラインを組み立てます。
                # このコマンドには、ユーザーが指定した各種オプション（傾き補正、余白除去、リサイズ、品質、コントラスト等）が含まれます。
                magick_cmd = [tools['magick'], str(img_path)]
                if args.Deskew: 
                    magick_cmd.extend(["-deskew", "40%"]) # 傾き補正
                    if args.Verbose: print(f"          -> 傾き補正")
                if args.Trim: 
                    magick_cmd.extend(["-fuzz", args.Fuzz, "-trim", "+repage"]) # 余白除去
                    if args.Verbose: print(f"          -> 余白除去 (Fuzz: {args.Fuzz})")
                
                # 目標の高さが設定されており、かつ元画像がそれより大きい場合のみリサイズします。
                # これにより、小さな画像が不必要に拡大されることを防ぎます。
                if target_height > 0 and original_height > target_height: 
                    magick_cmd.extend(["-resize", f"x{target_height}"])
                    if args.Verbose: print(f"          -> リサイズ: {target_height}px")
                else:
                    if args.Verbose: print(f"          -> 画像が小さいためリサイズはスキップ")

                # 解像度と品質の設定
                magick_cmd.extend(["-density", str(target_dpi), "-quality", str(args.Quality)])
                if args.Verbose: print(f"          -> 解像度設定: {target_dpi}dpi")
                if args.Verbose: print(f"          -> 画質設定: {args.Quality}")
                
                # カラー画像とグレースケール画像の判別と処理
                saturation = get_image_saturation(str(img_path), tools['magick'])
                if saturation < args.SaturationThreshold:
                    # グレースケール画像: グレースケール化とオプションのレベル補正
                    magick_cmd.extend(["-colorspace", "Gray"])
                    if args.Verbose: print(f"          -> グレースケール変換 (彩度: {saturation:.4f})")
                    if args.GrayscaleLevel: 
                        magick_cmd.extend(["-level", args.GrayscaleLevel])
                        if args.Verbose: print(f"          -> グレースケールレベル補正: {args.GrayscaleLevel}")
                else:
                    # カラー画像: コントラスト調整
                    if args.Verbose: print(f"          -> カラー画像として処理 (彩度: {saturation:.4f})")
                    if args.AutoContrast: 
                        magick_cmd.append("-normalize") # 自動コントラスト
                        if args.Verbose: print(f"          -> 自動コントラスト調整")
                    elif args.ColorContrast: 
                        magick_cmd.extend(["-brightness-contrast", args.ColorContrast]) # 手動コントラスト
                        if args.Verbose: print(f"          -> 手動コントラスト調整: {args.ColorContrast}")

                # 変換後のファイルパスを設定
                converted_path = converted_dir / f"{i:04d}.jpg"
                magick_cmd.append(str(converted_path))
                
                # --- 画像変換の実行 ---
                try:
                    subprocess.run(magick_cmd, check=True, capture_output=True)
                    if not converted_path.exists(): raise FileNotFoundError("Magick command succeeded but output file not found.")
                    
                    # 変換結果を記録（ファイルサイズ比較用）
                    conversion_results.append({
                        "original_path": img_path, "converted_path": converted_path,
                        "original_size": img_path.stat().st_size, "converted_size": converted_path.stat().st_size,
                        "saturation": saturation
                    })
                    if args.Verbose: print(f"          -> JPEGに変換: {converted_path.name}")
                except (subprocess.CalledProcessError, FileNotFoundError) as e:
                    print(f"Warning: Failed to convert {img_path.name}. It will be skipped. Error: {e}", file=sys.stderr)
                    skipped_files.append(img_path)

            if not conversion_results: raise ValueError("All image files failed to convert.")

            # --- ファイルサイズの比較と採用判断 ---
            # 変換後の画像セットと元の画像セットの合計ファイルサイズを比較します。
            # ユーザーが指定したしきい値、またはデフォルトのルール（2%増まで許容）に基づいて、
            # どちらのセットをPDFに採用するかを決定します。
            # これにより、画質を保ちつつファイルサイズを最適化します。
            total_original_size = sum(r['original_size'] for r in conversion_results)
            total_converted_size = sum(r['converted_size'] for r in conversion_results)
            
            print(f"\r[Info] 変換が完了しました。{total_images} ファイル中 {len(skipped_files)} ファイルがスキップされました。")
            print(f"[Compare] Original total size: {total_original_size / 1_048_576:.2f} MB")
            print(f"[Compare] Converted total size: {total_converted_size / 1_048_576:.2f} MB")

            use_converted = False
            if args.TotalCompressionThreshold is not None:
                # ユーザー指定のしきい値を使用
                if total_original_size > 0:
                    ratio = (total_converted_size / total_original_size) * 100
                    print(f"[Compare] Compression ratio: {ratio:.2f}% (Threshold: {args.TotalCompressionThreshold}%)")
                    if ratio < args.TotalCompressionThreshold:
                        use_converted = True
            else:
                # デフォルト動作: 変換後が元より2%以上大きい場合は元ファイルを使う
                if total_original_size == 0: # Prevent division by zero
                    use_converted = True
                elif (total_converted_size / total_original_size) < 1.02:
                    use_converted = True

            # --- PDF作成用ファイルセットの準備 ---
            if use_converted:
                print("[Decision] 変換後のファイルセットを採用します。")
                files_for_pdf = [str(r['converted_path']) for r in conversion_results]
                converted_count = len(files_for_pdf)
            else:
                print("[Decision] 元のファイルセットを採用します。(変換されたファイルは一時ファイルとして使用されます)")
                # 元の画像セットを使用する場合でも、PDF互換形式（JPEG）に再エンコードします。
                # このとき、傾き補正、余白除去、グレースケール判定とレベル補正が再適用されます。
                passthrough_dir = temp_dir / "passthrough"
                passthrough_dir.mkdir()
                for i, r in enumerate(conversion_results):
                    passthrough_path = passthrough_dir / f"{i:04d}.jpg"
                    magick_cmd = [tools['magick'], str(r['original_path'])]
                    if args.Deskew: magick_cmd.extend(["-deskew", "40%"])
                    if args.Trim: magick_cmd.extend(["-fuzz", args.Fuzz, "-trim", "+repage"])
                    
                    if r['saturation'] < args.SaturationThreshold:
                        magick_cmd.extend(["-colorspace", "Gray"])
                        if args.GrayscaleLevel: magick_cmd.extend(["-level", args.GrayscaleLevel])
                    else:
                        if args.AutoContrast: magick_cmd.append("-normalize")
                        elif args.ColorContrast: magick_cmd.extend(["-brightness-contrast", args.ColorContrast])

                    magick_cmd.extend(["-quality", str(args.Quality), str(passthrough_path)])
                    try:
                        subprocess.run(magick_cmd, check=True, capture_output=True)
                        files_for_pdf.append(str(passthrough_path))
                        if args.Verbose: print(f"  -> {r['original_path'].name} を再エンコード: {passthrough_path.name}")
                    except subprocess.CalledProcessError as e:
                        print(f"Warning: Failed to passthrough convert {r['original_path'].name}. It will be skipped. Error: {e}", file=sys.stderr)
                        skipped_files.append(r['original_path'])
                original_count = len(files_for_pdf)

        if not files_for_pdf:
            raise ValueError("No image files were successfully prepared for the PDF.")

        # --- PDFの作成 ---
        # pdfcpu を使用して、選択された画像ファイルセットからPDFを作成します。
        # 必要に応じてページサイズを設定し、PDFの最適化とビューア設定を行います。
        print("Creating PDF...")
        archive_name = Path(archive_path).stem
        output_pdf_path = Path(archive_path).parent / f"{archive_name}.pdf"
        temp_pdf_path = temp_dir / "temp.pdf"

        # pdfcpu import コマンドの構築
        pdfcpu_import_cmd = [tools['pdfcpu'], "import", "--"]

        # SetPageSize 属性が存在するか確認
        if hasattr(args, 'SetPageSize') and args.SetPageSize:
            # Determine the paper size string for pdfcpu.
            # This logic mirrors the PowerShell version's $targetPaperSizeForPdfCpu
            if args.Height:
                paper_size_for_cpu = "auto"
            else:
                paper_size_for_cpu = args.PaperSize or "A4"

            page_size_string = paper_size_for_cpu
            if args.Landscape and paper_size_for_cpu != 'auto':
                page_size_string += "L"

            if paper_size_for_cpu == 'auto':
                pdfcpu_page_conf = "dim:auto"
            else:
                pdfcpu_page_conf = f"f:{page_size_string}, pos:c, sc:1 rel"

            print(f"[Info] Attempting to set PDF page size: {pdfcpu_page_conf}")
            pdfcpu_import_cmd.append(pdfcpu_page_conf)
        else:
            print("[Info] PDF page size: auto-determined by pdfcpu")

        pdfcpu_import_cmd.append(str(temp_pdf_path))
        pdfcpu_import_cmd.extend(files_for_pdf)

        subprocess.run(pdfcpu_import_cmd, check=True, capture_output=True)
        
        # Set viewer preferences
        viewer_pref_json_path = temp_dir / "viewerpref.json"
        viewer_pref_json_path.write_text('{"DisplayDocTitle": true}', encoding="utf-8")
        try:
            subprocess.run([tools['pdfcpu'], "viewerpref", "set", str(temp_pdf_path), str(viewer_pref_json_path)], check=True, capture_output=True)
        except subprocess.CalledProcessError as e:
            print(f"Warning: pdfcpu viewerpref set failed. Stderr: {e.stderr.decode(errors='ignore')}", file=sys.stderr)

        print("Optimizing PDF...")
        subprocess.run([tools['pdfcpu'], "optimize", str(temp_pdf_path)], check=True, capture_output=True)
        
        final_pdf_path = temp_pdf_path
        if args.Linearize and tools.get('qpdf'):
            print("Linearizing PDF with QPDF...")
            linearized_pdf_path = temp_dir / "linearized.pdf"
            subprocess.run([tools['qpdf'], "--linearize", str(temp_pdf_path), str(linearized_pdf_path)], check=True, capture_output=True)
            final_pdf_path = linearized_pdf_path
        elif args.Linearize:
             print("Warning: --Linearize specified but QPDF not found. Skipping.", file=sys.stderr)
        
        shutil.move(str(final_pdf_path), str(output_pdf_path))
        
        source_stat = Path(archive_path).stat()
        os.utime(output_pdf_path, (source_stat.st_atime, source_stat.st_mtime))
        
        print(f"\nSuccess! PDF created at: {output_pdf_path}")

        # --- ログの出力 ---
        # 処理結果、使用した設定、スキップされたページなどの情報をログファイルに記録します。
        # これにより、後からどのファイルがどの設定で変換されたかを確認できます。
        if log_file_path:
            try:
                log_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                import shlex
                command_line = "python " + " ".join(shlex.quote(arg) for arg in sys.argv[1:])
                skipped_count = len(skipped_files)
                status = "Success with pages skipped" if skipped_count > 0 else "Success"

                settings_parts = []
                if args.PaperSize: settings_parts.append(f"PaperSize:{args.PaperSize}")
                if args.TotalCompressionThreshold is not None: settings_parts.append(f"TCR:{args.TotalCompressionThreshold}")
                if args.Trim: settings_parts.append(f"Trim:True"); settings_parts.append(f"Fuzz:{args.Fuzz}")
                if args.Deskew: settings_parts.append(f"Deskew:True")
                if args.AutoContrast: settings_parts.append(f"AutoContrast:True")
                elif args.ColorContrast: settings_parts.append(f"ColorContrast:{args.ColorContrast}")
                if args.GrayscaleLevel: settings_parts.append(f"GrayscaleLevel:{args.GrayscaleLevel}")
                if args.Linearize: settings_parts.append(f"Linearize:True")
                if hasattr(args, 'SetPageSize') and args.SetPageSize:
                    settings_parts.append(f"SetPageSize:True")
                    if args.Landscape:
                        settings_parts.append(f"Landscape:True")
                settings_parts.append(f"Height:{target_height}px")
                settings_parts.append(f"DPI:{target_dpi}")
                settings_parts.append(f"Quality:{args.Quality}")
                settings_parts.append(f"Saturation:{args.SaturationThreshold}")
                settings_string = ", ".join(settings_parts)

                log_message = (
                    f'Timestamp="{log_timestamp}" '
                    f'Status="{status}" '
                    f'Source="{os.path.basename(archive_path)}" '
                    f'Output="{output_pdf_path}" '
                    f'Images={len(image_files)} '
                    f'Converted={converted_count} '
                    f'Originals={original_count} '
                    f'Skipped={skipped_count} '
                    f'Settings="{settings_string}"'
                )

                log_details = []
                if not args.SkipCompression:
                    for res in conversion_results:
                        ratio_str = ""
                        if res['original_size'] > 0:
                            ratio = (res['converted_size'] / res['original_size'] * 100)
                            ratio_str = f" (Ratio: {ratio:.2f} %)"
                        log_status = "Converted" if use_converted else "Original"
                        log_details.append(f"    - {res['original_path'].name}: {log_status}{ratio_str}")
                
                for skipped_file in skipped_files:
                    log_details.append(f"    - {skipped_file.name}: SKIPPED (File conversion failed)")

                full_log_content = [f"Command: {command_line}", log_message] + log_details
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write("\n".join(full_log_content) + "\n\n")

                if args.Verbose: print(f"Wrote log to: {log_file_path}")
            except Exception as e:
                print(f"Warning: Failed to write to log file: {e}", file=sys.stderr)

    except Exception as e:
        print(f"Error processing {os.path.basename(archive_path)}: {e}", file=sys.stderr)
    finally:
        if args.Verbose: print(f"Cleaning up temporary directory: {temp_dir}")
        shutil.rmtree(temp_dir, ignore_errors=True)

def main():
    """
    メイン関数。引数の解析と変換処理の実行を制御します。
    
    この関数は全体のワークフローを制御します。
    1. コマンドライン引数を解析、または対話モードでユーザー入力を取得
    2. 必要な外部ツール（7-Zip, ImageMagick, pdfcpu, QPDF）のパスを特定
    3. 入力アーカイブのパス（ワイルドカード含む）を解決
    4. ユーザー設定に基づいて目標画像サイズを計算
    5. 各アーカイブファイルに対して `process_archive` を呼び出し変換
    6. ログ出力と最終的なクリーンアップを実行
    """
    # --- 引数の解析 ---
    parser = argparse.ArgumentParser(description="Converts archives to high-quality PDFs.", formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("archive_file_paths", metavar="ARCHIVE_PATHS", nargs='*', help="Paths or glob patterns for archives. If empty, enters interactive mode.")
    parser.add_argument("--SevenZipPath", help="Path to 7z executable.")
    parser.add_argument("--MagickPath", help="Path to magick executable.")
    parser.add_argument("--PdfCpuPath", help="Path to pdfcpu executable.")
    parser.add_argument("--QpdfPath", help="Path to qpdf executable.")
    parser.add_argument("-q", "--Quality", type=int, default=85, help="JPEG quality. Default: 85")
    parser.add_argument("-s", "--SaturationThreshold", type=float, default=0.05, help="Saturation threshold for grayscale. Default: 0.05")
    # 高さ指定のショートカットを -H に変更しました
    parser.add_argument("-H", "--Height", type=int, help="Target height in pixels.")
    parser.add_argument("-d", "--Dpi", type=int, help="Image DPI.")
    parser.add_argument("-p", "--PaperSize", choices=['A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'B0', 'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7'], help="Paper size for height calculation.")
    parser.add_argument("-sc", "--SkipCompression", action="store_true", help="Skip optimizations.")
    parser.add_argument("-t", "--Trim", action="store_true", help="Trim margins.")
    parser.add_argument("--Fuzz", default="1%", help="Fuzz factor for --Trim. Default: 1%%") 
    parser.add_argument("-ds", "--Deskew", action="store_true", help="Deskew images.")
    parser.add_argument("-gl", "--GrayscaleLevel", help="Level for grayscale contrast (e.g., '10%%,90%%').")
    parser.add_argument("-cc", "--ColorContrast", help="Value for color contrast (e.g., '0x25').")
    parser.add_argument("-ac", "--AutoContrast", action="store_true", help="Auto-adjust color contrast using normalize.")
    parser.add_argument("-lin", "--Linearize", action="store_true", help="Linearize PDF (requires QPDF).")
    parser.add_argument("-tcr", "--TotalCompressionThreshold", type=float, help="Threshold to decide whether to use converted files.")
    parser.add_argument("--LogPath", help="Path for logging.")
    parser.add_argument("-v", "--Verbose", action="store_true", help="Enable verbose output.")
    
    args = parser.parse_args()

    # --- 条件付きデフォルトロジック ---
    # PaperSizeが指定され、かつSetPageSizeが明示的にFalseでない場合、デフォルトでページサイズを設定する。
    if ('--PaperSize' in sys.argv or '-p' in sys.argv) and '--SetPageSize' not in sys.argv:
        args.SetPageSize = True

    # ユーザーがオプション引数を指定したかどうかを確認します。
    # これにより、対話モードに入るかどうかを判断します。
    optional_args_provided = any(arg.startswith('-') for arg in sys.argv[1:])

    # --- 対話モード ---
    # 最初に、コマンドラインでファイルパスが指定されていなければ、ユーザーに尋ねます。
    if not args.archive_file_paths:
        print("\n" + "-" * 40)
        print(" 対話モード: パス入力")
        print("-" * 40 + "\n")
        prompt = r"アーカイブファイルのパスを入力してください (例: C:\books\*.zip): "
        while not args.archive_file_paths:
            path_in = input(prompt)

            # Handle accidental paste of the prompt
            if path_in.startswith(prompt.strip()):
                path_in = path_in[len(prompt.strip()):].lstrip(": ")

            # Trim whitespace and surrounding quotes
            path_in = path_in.strip().strip('"')

            if path_in:
                # Escape brackets to treat them as literals, not wildcards
                args.archive_file_paths = [path_in]
            else:
                print("エラー: パスを空にはできません。パスまたはパターンを入力してください。", file=sys.stderr)

    # 第二に、オプション設定が指定されていなければ、対話モードで設定を取得します。
    # これは、ユーザーがパスのみを指定した場合、または上記のプロンプトでパスを入力した場合に実行されます。
    if not optional_args_provided:
        print("\n" + "-" * 40)
        print(" 対話モード: 設定入力")
        print(" Enterキーを押すと、[]内のデフォルト値が使用されます。")
        print("-" * 40)
        
        skip_in = input("圧縮をスキップしますか (y/n)? [n]: ").lower()
        if skip_in == 'y': args.SkipCompression = True

        if not args.SkipCompression:
            quality_in = input(f"画質 (1-100) [{args.Quality}]: ")
            if quality_in: args.Quality = int(quality_in)
            
            sat_in = input(f"グレースケールの彩度しきい値 [{args.SaturationThreshold}]: ")
            if sat_in: args.SaturationThreshold = float(sat_in)

            tcr_in = input("合計圧縮率のしきい値 (0-100, 省略可): ")
            if tcr_in: args.TotalCompressionThreshold = float(tcr_in)

            deskew_in = input("傾き補正 (y/n)? [n]: ").lower()
            if deskew_in == 'y': args.Deskew = True

            gray_contrast_in = input("グレースケールのコントラストを調整しますか (y/n)? [n]: ").lower()
            if gray_contrast_in == 'y':
                level_in = input("グレースケールのレベル値 [10%,90%]: ")
                args.GrayscaleLevel = level_in if level_in else "10%,90%"

            auto_contrast_in = input("カラーのコントラストを自動調整しますか (y/n)? [n]: ").lower()
            if auto_contrast_in == 'y':
                args.AutoContrast = True
            else:
                color_contrast_in = input("カラーのコントラストを手動で調整しますか (y/n)? [n]: ").lower()
                if color_contrast_in == 'y':
                    bright_in = input("カラーの明るさ-コントラスト値 [0x25]: ")
                    args.ColorContrast = bright_in if bright_in else "0x25"

            trim_in = input("余白をトリミングしますか (y/n)? [n]: ").lower()
            if trim_in == 'y':
                args.Trim = True
                fuzz_in = input(f"トリミングのFuzz値 [{args.Fuzz}]: ")
                if fuzz_in: args.Fuzz = fuzz_in

            linearize_in = input("PDFをウェブ用に最適化しますか (y/n)? [n]: ").lower()
            if linearize_in == 'y': args.Linearize = True

            res_choice = input("解像度指定方法: 1=高さ指定, 2=用紙サイズ+DPI [2]: ")
            if res_choice == '1':
                height_in = input("高さ (ピクセル): ")
                if height_in: args.Height = int(height_in)
                dpi_for_h_in = input("DPI [144]: ")
                if dpi_for_h_in: args.Dpi = int(dpi_for_h_in)
            else:
                paper_in = input("用紙サイズ [A4]: ")
                args.PaperSize = paper_in if paper_in else "A4"
                dpi_in = input("DPI [144]: ")
                args.Dpi = int(dpi_in) if dpi_in else 144
        print("-" * 40 + "\n")

    try:
        tools = {
            'sevenzip': find_tool('sevenzip', args.SevenZipPath),
            'magick': find_tool('magick', args.MagickPath),
            'pdfcpu': find_tool('pdfcpu', args.PdfCpuPath),
            'qpdf': find_tool('qpdf', args.QpdfPath) if args.Linearize else None
        }
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    log_file_path = None
    if args.LogPath:
        log_path = Path(args.LogPath).resolve()
        if log_path.is_dir():
            log_file_path = log_path / "JisuiArc2PDF_py_log.txt"
        else:
            log_file_path = log_path
    else:
        try:
            script_dir = Path(os.path.dirname(os.path.realpath(__file__)))
            log_file_path = script_dir / "JisuiArc2PDF_py_log.txt"
        except NameError:
            log_file_path = Path.cwd() / "JisuiArc2PDF_py_log.txt"

    # デバッグ: log_file_path の値を表示
    if args.Verbose:
        print(f"[Debug] log_file_path: {log_file_path}")

    if log_file_path:
        try:
            log_file_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"Error: Could not create log directory for {log_file_path}. {e}", file=sys.stderr)
            log_file_path = None

    if args.Verbose:
        print("--- Tool Paths ---")
        for tool, path in tools.items():
            if path: print(f"{tool}: {path}")
        print("--------------------\n")

    resolved_files = []
    for pattern in args.archive_file_paths:
        try:
            expanded = glob.glob(pattern, recursive=True)
            if expanded:
                resolved_files.extend(expanded)
            elif os.path.exists(pattern):
                resolved_files.append(pattern)
        except Exception:
            if os.path.exists(pattern):
                 resolved_files.append(pattern)
            else:
                print(f"Warning: Could not resolve pattern '{pattern}'", file=sys.stderr)

    unique_files = sorted(list(set(resolved_files)))
    if not unique_files:
        print("Error: No input files found.", file=sys.stderr)
        sys.exit(1)

    if args.Verbose:
        print(f"--- Found {len(unique_files)} archives to process ---")
        for f in unique_files:
            print(f"  - {f}")
        print("--------------------\n")

    target_height, target_dpi = calculate_target_height(args)

    for archive_path in unique_files:
        process_archive(archive_path, args, tools, log_file_path, target_height, target_dpi)

    print("\n" + "=" * 40)
    print("All processing complete.")
    print("=" * 40)

if __name__ == "__main__":
    main()