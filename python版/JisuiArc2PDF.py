import argparse
import sys
import os
import shutil
import glob
import subprocess
import tempfile
import re
from pathlib import Path

# --- Helper Functions ---

def find_tool(tool_key: str, specified_path: str | None) -> str:
    """Find a required tool by its key, searching in order: specified path, script dir, system PATH."""
    # Define tool filenames based on OS
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
    except NameError: # __file__ is not defined in interactive environments
        pass
    system_path = shutil.which(tool_name)
    if system_path:
        return system_path
    raise FileNotFoundError(
        f"{tool_name} not found. Please provide it via arguments, place it in the script's directory, or add it to your system's PATH."
    )

def natural_sort_key(s: str) -> list:
    """Create a key for natural sorting (e.g., '2.jpg' before '10.jpg')."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

def is_image(file_path: str, magick_exe: str) -> bool:
    """Check if a file is an image using magick identify."""
    try:
        subprocess.run([magick_exe, "identify", file_path], check=True, capture_output=True, timeout=15)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False

def get_image_height(file_path: str, magick_exe: str) -> int:
    try:
        result = subprocess.run([magick_exe, "identify", "-format", "%h", file_path], check=True, capture_output=True, text=True)
        return int(result.stdout)
    except (subprocess.CalledProcessError, FileNotFoundError, ValueError):
        return 0

def get_image_saturation(file_path: str, magick_exe: str) -> float:
    try:
        result = subprocess.run(
            [magick_exe, file_path, "-colorspace", "HSL", "-channel", "G", "-separate", "+channel", "-format", "%[mean]", "info:"],
            check=True, capture_output=True, text=True
        )
        return float(result.stdout) / 65535.0
    except (subprocess.CalledProcessError, FileNotFoundError, ValueError):
        return 0.5

def calculate_target_height(args: argparse.Namespace) -> tuple[int, int]:
    """Calculate target height and DPI based on user arguments."""
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
        target_height = int((paper_height / 25.4) * target_dpi)
        print(f"[Info] Paper size specified: {args.PaperSize} @ {target_dpi}dpi -> {target_height}px")
        return target_height, target_dpi
    if args.PaperSize or args.Dpi:
        print("Error: --PaperSize and --Dpi must be used together.", file=sys.stderr)
        sys.exit(1)
    target_dpi = 144
    target_height = int((paper_heights_mm['A4'] / 25.4) * target_dpi)
    print(f"[Info] Using default: A4 @ {target_dpi}dpi -> {target_height}px")
    return target_height, target_dpi

def process_archive(archive_path: str, args: argparse.Namespace, tools: dict):
    """Process a single archive file."""
    print("\n" + "=" * 40)
    print(f"Processing: {os.path.basename(archive_path)}")
    print("=" * 40)

    temp_dir = Path(tempfile.mkdtemp(prefix="JisuiArc2PDF_"))
    converted_dir = temp_dir / "converted"
    converted_dir.mkdir()

    if args.Verbose:
        print(f"Created temporary directory: {temp_dir}")

    try:
        print("Extracting archive...")
        subprocess.run(
            [tools['sevenzip'], "e", archive_path, f"-o{temp_dir}", "-y"],
            check=True, capture_output=True, timeout=300
        )

        print("Finding and sorting image files...")
        image_files = []
        for root, _, files in os.walk(temp_dir):
            if Path(root) == converted_dir:
                continue
            for file in files:
                full_path = Path(root) / file
                if is_image(str(full_path), tools['magick']):
                    image_files.append(full_path)
        
        image_files.sort(key=lambda f: natural_sort_key(f.name))

        if not image_files:
            print("Error: No image files found in the archive.", file=sys.stderr)
            return

        if args.Verbose:
            print(f"Found {len(image_files)} images.")

        target_height, target_dpi = calculate_target_height(args)

        print("Converting images...")
        conversion_results = []
        for i, img_path in enumerate(image_files):
            print(f"  Processing {img_path.name}...")
            original_size = img_path.stat().st_size
            magick_cmd = [tools['magick'], str(img_path)]
            if args.Deskew: magick_cmd.extend(["-deskew", "40%"])
            if args.Trim: magick_cmd.extend(["-fuzz", args.Fuzz, "-trim", "+repage"])
            img_height = get_image_height(str(img_path), tools['magick'])
            if target_height > 0 and img_height > target_height: magick_cmd.extend(["-resize", f"x{target_height}"])
            magick_cmd.extend(["-density", str(target_dpi), "-quality", str(args.Quality)])
            
            # Add contrast adjustments based on saturation
            if get_image_saturation(str(img_path), tools['magick']) < args.SaturationThreshold:
                magick_cmd.extend(["-colorspace", "Gray"])
                if args.GrayscaleLevel:
                    magick_cmd.extend(["-level", args.GrayscaleLevel])
            else:
                if args.ColorContrast:
                    magick_cmd.extend(["-brightness-contrast", args.ColorContrast])
            converted_path = converted_dir / f"{i:04d}.jpg"
            magick_cmd.append(str(converted_path))
            subprocess.run(magick_cmd, check=True, capture_output=True)
            conversion_results.append({
                "original_path": img_path, "converted_path": converted_path,
                "original_size": original_size, "converted_size": converted_path.stat().st_size if converted_path.exists() else 0
            })

        total_original_size = sum(r['original_size'] for r in conversion_results)
        total_converted_size = sum(r['converted_size'] for r in conversion_results)
        use_converted = total_converted_size < total_original_size
        if args.TotalCompressionThreshold is not None and total_original_size > 0:
            ratio = (total_converted_size / total_original_size) * 100
            print(f"[Info] Compression ratio: {ratio:.2f}% (Threshold: {args.TotalCompressionThreshold}%)")
            use_converted = ratio < args.TotalCompressionThreshold

        files_for_pdf = []
        if use_converted:
            print("[Info] Using converted (smaller) image set.")
            files_for_pdf = [str(r['converted_path']) for r in conversion_results]
        else:
            print("[Info] Using original image set (re-encoded as JPEG).")
            passthrough_dir = temp_dir / "passthrough"
            passthrough_dir.mkdir()
            for i, r in enumerate(conversion_results):
                passthrough_path = passthrough_dir / f"{i:04d}.jpg"
                magick_cmd = [tools['magick'], str(r['original_path'])]
                if args.Deskew: magick_cmd.extend(["-deskew", "40%"])
                if args.Trim: magick_cmd.extend(["-fuzz", args.Fuzz, "-trim", "+repage"])
                
                # Add contrast adjustments based on saturation for passthrough
                if get_image_saturation(str(r['original_path']), tools['magick']) < args.SaturationThreshold:
                    magick_cmd.extend(["-colorspace", "Gray"])
                    if args.GrayscaleLevel:
                        magick_cmd.extend(["-level", args.GrayscaleLevel])
                else:
                    if args.AutoContrast:
                        magick_cmd.append("-normalize")
                    elif args.ColorContrast:
                        magick_cmd.extend(["-brightness-contrast", args.ColorContrast])
                magick_cmd.extend(["-quality", str(args.Quality), str(passthrough_path)])
                subprocess.run(magick_cmd, check=True, capture_output=True)
                files_for_pdf.append(str(passthrough_path))

        print("Creating PDF...")
        archive_name = Path(archive_path).stem
        output_pdf_path = Path(archive_path).parent / f"{archive_name}.pdf"
        temp_pdf_path = temp_dir / "temp.pdf"
        subprocess.run([tools['pdfcpu'], "import", str(temp_pdf_path)] + files_for_pdf, check=True, capture_output=True)
        print("Optimizing PDF...")
        subprocess.run([tools['pdfcpu'], "optimize", str(temp_pdf_path)], check=True, capture_output=True)
        final_pdf_path = temp_pdf_path
        if args.Linearize and tools['qpdf']:
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

    except FileNotFoundError as e:
        print(f"Error: A tool was not found: {e}", file=sys.stderr)
    except subprocess.CalledProcessError as e:
        print(f"Error during subprocess execution for: {' '.join(e.cmd)}", file=sys.stderr)
        print(f"Stderr: {e.stderr.decode(errors='ignore')}", file=sys.stderr)
    except Exception as e:
        print(f"An unexpected error occurred: {e}", file=sys.stderr)
    finally:
        if args.Verbose:
            print(f"Cleaning up temporary directory: {temp_dir}")
        shutil.rmtree(temp_dir)

def main():
    """Main function to parse arguments and execute the conversion process."""
    parser = argparse.ArgumentParser(description="Converts archives to high-quality PDFs.", formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("archive_file_paths", metavar="ARCHIVE_PATHS", nargs='+', help="Paths or glob patterns for archives.")
    parser.add_argument("--SevenZipPath", help="Path to 7z executable.")
    parser.add_argument("--MagickPath", help="Path to magick executable.")
    parser.add_argument("--PdfCpuPath", help="Path to pdfcpu executable.")
    parser.add_argument("--QpdfPath", help="Path to qpdf executable.")
    parser.add_argument("-q", "--Quality", type=int, default=85, help="JPEG quality. Default: 85")
    parser.add_argument("-s", "--SaturationThreshold", type=float, default=0.05, help="Saturation threshold for grayscale. Default: 0.05")
    parser.add_argument("-h", "--Height", type=int, help="Target height in pixels.")
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

    if args.Verbose:
        print("--- Tool Paths ---")
        for tool, path in tools.items():
            if path: print(f"{tool}: {path}")
        print("--------------------
")

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
        print("--------------------
")

    for archive_path in unique_files:
        process_archive(archive_path, args, tools)

    print("\n" + "=" * 40)
    print("All processing complete.")
    print("=" * 40)

if __name__ == "__main__":
    main()
