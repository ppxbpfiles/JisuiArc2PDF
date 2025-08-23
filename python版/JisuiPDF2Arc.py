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
    """Find a required tool by its key, searching in order: specified path, script dir, system PATH."""
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
    """Process a single PDF file."""
    pdf_path_obj = Path(pdf_path)
    print("\n" + "=" * 40)
    print(f"Processing: {pdf_path_obj.name}")
    print("=" * 40)

    # Create a temporary directory for extraction
    temp_dir = Path(tempfile.mkdtemp(prefix="JisuiPDF2Arc_"))
    if args.Verbose:
        print(f"Created temporary directory: {temp_dir}")

    try:
        # 1. Extract images using pdfcpu
        print("  -> Extracting images from PDF...")
        subprocess.run(
            [tools['pdfcpu'], "extract", "-mode", "image", str(pdf_path_obj), str(temp_dir)],
            check=True, capture_output=True, timeout=300
        )

        # 2. Check if any images were extracted
        # Using glob to find all files, as pdfcpu might create subdirectories or various file types.
        extracted_files = list(temp_dir.glob('**/*'))
        # Filter to only actual files, not directories
        extracted_files = [f for f in extracted_files if f.is_file()]
        
        if not extracted_files:
            print(f"Warning: No images found in or extracted from {pdf_path_obj.name}. Skipping.")
            return # Exit function successfully, but log the skip

        # 3. Create output directory if it doesn't exist
        # Output directory is ./pdf2arc_converted relative to the current working directory
        converted_output_dir = Path.cwd() / "pdf2arc_converted"
        converted_output_dir.mkdir(parents=True, exist_ok=True)

        # 4. Define ZIP output path
        zip_filename = pdf_path_obj.stem + ".zip"
        zip_output_path = converted_output_dir / zip_filename
        print(f"  -> Creating ZIP archive: {zip_output_path}")

        # 5. Create ZIP archive using 7-Zip
        # We need to add files individually or use a method that works with the shell/command line.
        # The '*' pattern might not work reliably in all environments when passed directly.
        # A safer way is to pass the directory containing the files.
        # However, pdfcpu extracts to the temp_dir directly, so we can compress the whole temp_dir content.
        # 7z a archive.zip dir/* is the standard way.
        # Let's pass the temp_dir/* pattern. We need to be careful with shell interpretation.
        # Using shell=False is generally safer. Let's pass the directory content.
        # 7z a archive.zip dir/*
        # This requires the shell to expand '*'. To avoid shell dependency, we can list files.
        # But for simplicity and mirroring the PS1 script, we will use the directory path with '*'.
        # The PS1 script used `Join-Path $tempDir "*"` which works in PowerShell.
        # In Python, we can use `str(temp_dir / '*')` and rely on 7z to handle it,
        # or we can list files. Let's try the former first to keep it simple and similar.
        # Actually, it's safer to list files explicitly to avoid shell interpretation issues.
        # But for now, to mimic PS1 closely: sourceForZip = Join-Path $tempDir "*"
        # So we do: source_for_zip = temp_dir / '*'
        # And pass it as a string. 7z should handle it.
        # Let's do it the PS1 way for now.
        source_for_zip = temp_dir / '*'
        
        # Run 7z command
        subprocess.run(
            [tools['sevenzip'], "a", "-tzip", str(zip_output_path), str(source_for_zip)],
            check=True, capture_output=True
        )

        print(f"  -> Successfully created {zip_filename}")

        # 6. Logging
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
                # Open in append mode with utf-8 encoding
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
        # Log the error
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
        # Log the error
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
        # 7. Cleanup temporary directory
        if args.Verbose:
            print(f"Cleaning up temporary directory: {temp_dir}")
        # Use shutil.rmtree to remove the directory and its contents
        shutil.rmtree(temp_dir, ignore_errors=True)

def main():
    """Main function to parse arguments and execute the conversion process."""
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

    # --- Find Tools ---
    try:
        tools = {
            'sevenzip': find_tool('sevenzip', args.SevenZipPath),
            'pdfcpu': find_tool('pdfcpu', args.PdfCpuPath),
        }
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    # --- Handle Log Path ---
    log_file_path = None
    if args.LogPath:
        log_path = Path(args.LogPath).resolve()
        # If the provided LogPath is a directory, create the default log file name inside it
        if log_path.is_dir():
            log_file_path = log_path / "JisuiPDF2Arc_py_log.txt"
        else:
            # If it's a file path, use it directly
            log_file_path = log_path
    else:
        # Default log file location: script directory or current directory
        try:
            script_dir = Path(os.path.dirname(os.path.realpath(__file__)))
            log_file_path = script_dir / "JisuiPDF2Arc_py_log.txt"
        except NameError:
            # If __file__ is not available (e.g., interactive interpreter)
            log_file_path = Path.cwd() / "JisuiPDF2Arc_py_log.txt"

    # Ensure the directory for the log file exists
    if log_file_path:
        try:
            log_file_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"Error: Could not create log directory for {log_file_path}. {e}", file=sys.stderr)
            log_file_path = None # Disable logging if we can't create the directory

    # --- Verbose Output for Tools (if requested) ---
    if args.Verbose:
        print("--- Tool Paths ---")
        for tool, path in tools.items():
            if path: print(f"{tool}: {path}")
        print("--------------------")

    # --- Resolve Input File Paths ---
    resolved_files = []
    for pattern in args.pdf_file_paths:
        # Use glob.glob to resolve wildcards and find matching files
        # `recursive=True` allows for `**` patterns if needed in the future.
        matched_files = glob.glob(pattern, recursive=True)
        for file_path in matched_files:
            # Check if the resolved path is a file and has a .pdf extension (case-insensitive)
            if os.path.isfile(file_path) and file_path.lower().endswith('.pdf'):
                resolved_files.append(file_path)
    
    # Remove duplicates and sort for consistent processing order
    unique_files = sorted(list(set(resolved_files)))

    if not unique_files:
        print("Error: No input PDF files found.", file=sys.stderr)
        sys.exit(1)

    # --- Verbose Output for Files (if requested) ---
    if args.Verbose:
        print(f"--- Found {len(unique_files)} PDFs to process ---")
        for f in unique_files:
            print(f"  - {f}")
        print("--------------------")

    # --- Process Each PDF ---
    for pdf_path in unique_files:
        process_pdf(pdf_path, args, tools, log_file_path)

    # --- Final Message ---
    print("\n" + "=" * 40)
    print("All processing complete.")
    print("=" * 40)

if __name__ == "__main__":
    main()