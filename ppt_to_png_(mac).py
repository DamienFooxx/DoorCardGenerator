#!/usr/bin/env python3
import os
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Iterable, Optional
from uuid import uuid4

# ----------------------------
# Configuration (edit as needed)
# ----------------------------
INPUT_FOLDER  = Path("./2526_pptx")
OUTPUT_FOLDER = Path("./2526doorcards_png")
RECURSIVE     = False  # set True to scan subfolders of INPUT_FOLDER

# ----------------------------
# Helpers
# ----------------------------
def find_soffice() -> Optional[str]:
    """
    Locate the LibreOffice 'soffice' binary on macOS.
    Tries PATH first, then common macOS app bundle path.
    """
    path = shutil.which("soffice")
    if path:
        return path

    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/Applications/LibreOfficeDev.app/Contents/MacOS/soffice",
        "/usr/local/bin/soffice",
        "/opt/homebrew/bin/soffice",
    ]
    for c in candidates:
        if os.path.isfile(c) and os.access(c, os.X_OK):
            return c
    return None


def iter_powerpoints(root: Path, recursive: bool = False) -> Iterable[Path]:
    patterns = ["*.ppt", "*.pptx"]
    if recursive:
        for pat in patterns:
            yield from root.rglob(pat)
    else:
        for pat in patterns:
            yield from root.glob(pat)


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def convert_one(soffice_bin: str, pptx_path: Path, out_root: Path) -> bool:
    """
    Convert a single PPT/PPTX to PNG using LibreOffice headless mode.
    Exports to a temporary directory first, then moves only *new* PNGs into out_root.
    Existing PNG filenames in out_root are never overwritten (skipped).
    """
    ensure_dir(out_root)

    # Create a unique temp directory inside the output root (keeps same filesystem -> fast moves)
    temp_dir = out_root / f".tmp_{pptx_path.stem}_{uuid4().hex[:8]}"
    ensure_dir(temp_dir)

    cmd = [
        soffice_bin,
        "--headless",
        "--convert-to", "png",
        "--outdir", str(temp_dir),
        str(pptx_path),
    ]

    try:
        completed = subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
    except subprocess.CalledProcessError as e:
        print(f"[FAIL] {pptx_path.name}")
        if e.stdout:
            print("  stdout:", e.stdout.strip())
        if e.stderr:
            print("  stderr:", e.stderr.strip())
        # cleanup temp_dir if empty/failed
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass
        return False

    # Move only files that do not already exist in out_root
    produced = list(temp_dir.glob("*.png"))
    moved, skipped = 0, 0
    for png in produced:
        dest = out_root / png.name
        if dest.exists():
            skipped += 1
        else:
            shutil.move(str(png), str(dest))
            moved += 1

    # Clean up temporary directory (should be empty after moves)
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass

    print(f"[OK] {pptx_path.name} -> {out_root}  (new: {moved}, skipped existing: {skipped})")
    # Uncomment for verbose logs:
    # print("STDOUT:", completed.stdout)
    # print("STDERR:", completed.stderr)
    return True


def ppt_to_png(input_folder: Path, output_folder: Path, recursive: bool = False) -> None:
    if not input_folder.exists():
        print(f"Input folder not found: {input_folder}")
        sys.exit(1)

    ensure_dir(output_folder)

    soffice_bin = find_soffice()
    if not soffice_bin:
        print(
            "Could not find 'soffice'. Please install LibreOffice and ensure "
            "'soffice' is in your PATH or installed at:\n"
            "  /Applications/LibreOffice.app/Contents/MacOS/soffice\n"
            "Homebrew (cask) install example:\n"
            "  brew install --cask libreoffice"
        )
        sys.exit(1)

    files = list(iter_powerpoints(input_folder, recursive))
    if not files:
        print(f"No .ppt or .pptx files found in {input_folder}{' (recursive)' if recursive else ''}.")
        return

    print(f"Found {len(files)} file(s). Converting with: {soffice_bin}\n")
    success = 0
    for i, f in enumerate(sorted(files), start=1):
        print(f"({i}/{len(files)}) Converting: {f}")
        if convert_one(soffice_bin, f, output_folder):
            success += 1

    print(f"\nDone. {success}/{len(files)} file(s) processed without overwriting existing PNGs.")
    print(f"PNG output root: {output_folder.resolve()}")


if __name__ == "__main__":
    # Optional CLI:
    #   python convert_ppt_to_png.py [input_folder] [output_folder] [--recursive]
    in_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else INPUT_FOLDER
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else OUTPUT_FOLDER
    recursive_flag = RECURSIVE or ("--recursive" in sys.argv[3:] or "--recursive" in sys.argv[1:])

    ppt_to_png(in_dir, out_dir, recursive_flag)
