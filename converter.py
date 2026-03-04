import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional

# Base directory:
# - dev: folder containing the .py
# - pyinstaller onefile: temporary extraction folder (_MEIPASS)
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(sys.argv[0]).resolve().parent))
EXE_DIR = Path(sys.argv[0]).resolve().parent

DROP_FOLDER = Path(r"C:\WarehouseDrop\PrintToTiff")


def is_pdf(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() == ".pdf"


def _prepend_to_path(p: Path) -> None:
    if p.exists():
        os.environ["PATH"] = str(p) + os.pathsep + os.environ.get("PATH", "")


def configure_portable_deps() -> None:
    """
    If Ghostscript and/or ImageMagick are bundled next to the EXE,
    make sure they're discoverable.
    """
    gs_bin_candidates = [
        EXE_DIR / "gs" / "bin",
        BASE_DIR / "gs" / "bin",
    ]

    for gs_bin in gs_bin_candidates:
        _prepend_to_path(gs_bin)

    magick_home_candidates = [
        EXE_DIR / "magick",
        BASE_DIR / "magick",
    ]

    for mh in magick_home_candidates:
        if (mh / "magick.exe").exists():
            os.environ["MAGICK_HOME"] = str(mh)
            break


def find_magick_exe() -> Optional[str]:
    """
    Locate ImageMagick's magick.exe.

    Priority:
    1) ./magick/magick.exe (bundled alongside exe)
    2) ./magick.exe (same folder as exe)
    3) PATH
    """
    candidates = [
        EXE_DIR / "magick" / "magick.exe",
        EXE_DIR / "magick.exe",
        BASE_DIR / "magick" / "magick.exe",
        BASE_DIR / "magick.exe",
    ]

    for c in candidates:
        if c.exists():
            return str(c)

    return shutil.which("magick")


def find_modi_exe() -> Optional[str]:
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")

    candidates = [
        Path(pf86) / r"Microsoft Office\Office12\MODI.EXE",
        Path(pf) / r"Microsoft Office\Office12\MODI.EXE",
        Path(pf86) / r"Microsoft Office\Office11\MODI.EXE",
        Path(pf) / r"Microsoft Office\Office11\MODI.EXE",
        Path(pf86) / r"Microsoft Office\Office10\MODI.EXE",
        Path(pf) / r"Microsoft Office\Office10\MODI.EXE",
    ]

    for c in candidates:
        if c.exists():
            return str(c)

    return None


def open_tiff(tiff_path: Path) -> None:
    modi = find_modi_exe()

    try:
        if modi:
            subprocess.Popen([modi, str(tiff_path)], close_fds=True)
        else:
            os.startfile(str(tiff_path))
    except Exception as e:
        raise RuntimeError(f"Converted successfully, but failed to open TIFF: {e}") from e


def convert_pdf_to_tiff(pdf_path: Path, magick_exe: str) -> Path:
    if not is_pdf(pdf_path):
        raise ValueError("Input must be an existing .pdf file.")

    out_path = pdf_path.with_suffix(".tif")

    cmd = [
        magick_exe,
        "-density", "300",
        "-units", "PixelsPerInch",
        str(pdf_path),

        "-alpha", "off",
        "-colorspace", "Gray",
        "-threshold", "60%",
        "-type", "bilevel",
        "-depth", "1",

        "-compress", "group4",
        "-define", "tiff:rows-per-strip=0",
        str(out_path),
    ]

    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        )

    except FileNotFoundError:
        raise RuntimeError(
            "ImageMagick (magick.exe) was not found. Install it or bundle it with this tool."
        )

    except Exception as e:
        raise RuntimeError(f"Failed to launch ImageMagick: {e}") from e

    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()

        if not err:
            err = f"ImageMagick failed with exit code {proc.returncode}."

        raise RuntimeError(err)

    if not out_path.exists():
        raise RuntimeError("ImageMagick reported success but the output TIFF was not created.")

    return out_path


def newest_unconverted_pdf(folder: Path) -> Path:
    """
    Return the newest PDF that does not already have a corresponding TIFF.
    """
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"Folder does not exist: {folder}")

    candidates = []

    for p in folder.iterdir():
        if p.is_file() and p.suffix.lower() == ".pdf":
            tif = p.with_suffix(".tif")

            if not tif.exists():
                candidates.append(p)

    if not candidates:
        raise FileNotFoundError("No unprocessed PDFs found.")

    return max(candidates, key=lambda p: p.stat().st_mtime)


def wait_for_file_complete(path: Path, timeout: int = 10, interval: float = 0.5) -> None:
    """
    Wait until a file stops growing in size, indicating it has finished writing.
    Useful for printer-generated PDFs.
    """
    start = time.time()
    last_size = -1

    while True:
        if not path.exists():
            raise FileNotFoundError(f"File disappeared: {path}")

        current_size = path.stat().st_size

        if current_size == last_size and current_size > 0:
            return

        last_size = current_size

        if time.time() - start > timeout:
            raise TimeoutError(f"Timed out waiting for file to finish writing: {path}")

        time.sleep(interval)


def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="converter",
        description="PDF -> multipage TIFF (300dpi, Group4 bilevel) using ImageMagick, then open in MODI/default viewer.",
        add_help=True,
    )

    p.add_argument(
        "pdf",
        nargs="?",
        help="PDF path (drag & drop passes it here).",
    )

    p.add_argument(
        "--latest",
        "-latest",
        action="store_true",
        help=r"Convert newest PDF in C:\WarehouseDrop\PrintToTiff\ and open TIFF.",
    )

    p.add_argument(
        "--drop-folder",
        default=str(DROP_FOLDER),
        help="Override the print drop folder for --latest mode.",
    )

    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    configure_portable_deps()

    args = parse_args(argv)

    if args.latest:
        pdf_path = newest_unconverted_pdf(Path(args.drop_folder))
        wait_for_file_complete(pdf_path)

    elif args.pdf:
        pdf_path = Path(args.pdf)

    else:
        raise ValueError("No input provided. Drag a PDF onto the EXE or run with --latest.")

    if not is_pdf(pdf_path):
        raise ValueError("Rejected: input is not a PDF file.")

    print(f"Processing: {pdf_path}")

    magick = find_magick_exe()

    if not magick:
        raise RuntimeError(
            "ImageMagick not found.\n\n"
            "Install ImageMagick (and ensure 'magick' is on PATH),\n"
            "or bundle magick.exe alongside this EXE."
        )

    tiff_path = convert_pdf_to_tiff(pdf_path, magick)

    open_tiff(tiff_path)

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))

    except Exception as e:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, str(e), "PDF → TIFF Converter", 0x10)
        except Exception:
            pass

        raise SystemExit(1)