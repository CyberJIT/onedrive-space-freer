#!/usr/bin/env python3
"""
onedrive_space_freer.py

A command-line tool to scan files in a parametrized OneDrive folder,
and mark them for OneDrive uploading and freeing local disk space (marking with the
FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS attribute on Windows / calling attrib +U -P).

Supports:
- Recursive scanning under subfolders.
- Filtering files by name parts / wildcards (include/exclude).
- Filtering files by file extensions (include/exclude).
- Filtering searched folders by name parts / wildcards (include/exclude).
- Handling long paths via extended prefix (\\\\?\\) and short 8.3 path fallbacks (GetShortPathNameW).
- Graceful error recovery: logs any file-level failure and continues processing.
- Detailed final summary of actions (processed, already dehydrated, marked, skipped, errors, space freed).
- Dry-run mode for previewing actions safely.
"""

import argparse
import ctypes
import fnmatch
import os
import subprocess
import sys
from typing import Dict, List, Optional, Set, Tuple


# Windows File Attribute Constants
FILE_ATTRIBUTE_READONLY = 0x00000001
FILE_ATTRIBUTE_HIDDEN = 0x00000002
FILE_ATTRIBUTE_SYSTEM = 0x00000004
FILE_ATTRIBUTE_DIRECTORY = 0x00000010
FILE_ATTRIBUTE_ARCHIVE = 0x00000020
FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400
FILE_ATTRIBUTE_OFFLINE = 0x00001000
# Cloud Files / OneDrive Files On-Demand attributes
FILE_ATTRIBUTE_PINNED = 0x00080000
FILE_ATTRIBUTE_UNPINNED = 0x00100000
FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x00400000

# Kernel32 error codes
INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF


def normalize_long_path(path: str) -> str:
    """
    On Windows, prepend the \\\\?\\ prefix to absolute paths to bypass MAX_PATH (260 chars).
    On non-Windows platforms, returns the path unmodified.
    """
    if os.name != "nt":
        return path
    abs_path = os.path.abspath(path)
    if abs_path.startswith("\\\\?\\"):
        return abs_path
    if abs_path.startswith("\\\\"):
        # UNC path: \\server\share -> \\?\UNC\server\share
        return "\\\\?\\UNC\\" + abs_path[2:]
    return "\\\\?\\" + abs_path


def get_short_path_name(path: str) -> Optional[str]:
    """
    Retrieve Windows 8.3 short path equivalent using GetShortPathNameW API.
    Returns None if not on Windows or if short path generation is disabled/unavailable.
    """
    if os.name != "nt":
        return None
    try:
        buffer_size = 1024
        buffer = ctypes.create_unicode_buffer(buffer_size)
        res = ctypes.windll.kernel32.GetShortPathNameW(path, buffer, buffer_size)
        if 0 < res < buffer_size:
            return buffer.value
    except Exception:
        pass
    return None


def get_file_attributes_win32(path: str) -> int:
    """
    Query Win32 file attributes using GetFileAttributesW, supporting extended paths
    and short 8.3 path fallback.
    """
    if os.name != "nt":
        return 0

    extended = normalize_long_path(path)
    attrs = ctypes.windll.kernel32.GetFileAttributesW(extended)
    if attrs != INVALID_FILE_ATTRIBUTES:
        return attrs

    short_p = get_short_path_name(path)
    if short_p and short_p != path:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(short_p)
        if attrs != INVALID_FILE_ATTRIBUTES:
            return attrs

    return INVALID_FILE_ATTRIBUTES


def set_file_attributes_win32(path: str, attrs: int) -> bool:
    """
    Set Win32 file attributes using SetFileAttributesW, supporting extended paths
    and short 8.3 path fallback.
    """
    if os.name != "nt":
        return True

    extended = normalize_long_path(path)
    if ctypes.windll.kernel32.SetFileAttributesW(extended, attrs):
        return True

    short_p = get_short_path_name(path)
    if short_p and short_p != path:
        if ctypes.windll.kernel32.SetFileAttributesW(short_p, attrs):
            return True

    return False


def is_already_free_space(attrs: int) -> bool:
    """
    Check if the file attributes indicate that the file is already online-only
    (Recall on Data Access / Unpinned / Offline).
    """
    if attrs == INVALID_FILE_ATTRIBUTES:
        return False
    return bool(attrs & (FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS | FILE_ATTRIBUTE_OFFLINE))


def mark_file_free_space(path: str, dry_run: bool = False) -> Tuple[bool, str]:
    """
    Mark a file to upload/sync and free up space locally.
    1. Removes FILE_ATTRIBUTE_PINNED (Always keep on this device).
    2. Adds FILE_ATTRIBUTE_UNPINNED.
    3. Adds FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS (Free up space).
    4. Falls back to executing Windows 'attrib.exe +U -P "<path>"' if direct API fails.

    Returns: (success: bool, status: str)
      status: 'marked', 'already_freed', or error description
    """
    if os.name != "nt":
        # Posix simulation / mock mode
        if dry_run:
            return True, "dry_run"
        return True, "marked"

    attrs = get_file_attributes_win32(path)
    if attrs == INVALID_FILE_ATTRIBUTES:
        err = ctypes.GetLastError()
        return False, f"Failed to get file attributes (Win32 Error: {err})"

    # If it is already online-only (Recall On Data Access or Offline) and not pinned
    if (attrs & FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS) and not (attrs & FILE_ATTRIBUTE_PINNED):
        return True, "already_freed"

    if dry_run:
        return True, "marked"

    # Calculate new attributes
    new_attrs = attrs
    # Remove PINNED (0x00080000)
    new_attrs &= ~FILE_ATTRIBUTE_PINNED
    # Add UNPINNED (0x00100000) and RECALL_ON_DATA_ACCESS (0x00400000)
    new_attrs |= FILE_ATTRIBUTE_UNPINNED | FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS

    if set_file_attributes_win32(path, new_attrs):
        return True, "marked"

    # Fallback to Windows attrib command: attrib.exe +U -P "<path>"
    # +U marks the file as unpinned/cloud-only (free up space)
    # -P removes the pinned attribute
    target_path = path
    if len(target_path) >= 250:
        short_p = get_short_path_name(target_path)
        if short_p:
            target_path = short_p

    try:
        proc = subprocess.run(
            ["attrib.exe", "+U", "-P", target_path],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
        if proc.returncode == 0:
            return True, "marked"
        else:
            return False, f"attrib +U -P failed: {proc.stderr.strip() or proc.stdout.strip()}"
    except Exception as ex:
        return False, f"Failed to set attributes via API and attrib: {ex}"


def matches_patterns(name: str, patterns: Optional[List[str]]) -> bool:
    """
    Check if a name contains any pattern as a substring (case-insensitive)
    or matches a wildcard pattern.
    """
    if not patterns:
        return False
    name_lower = name.lower()
    for pat in patterns:
        pat_lower = pat.lower()
        if any(c in pat_lower for c in ("*", "?", "[", "]")):
            if fnmatch.fnmatch(name_lower, pat_lower) or fnmatch.fnmatch(name_lower, f"*{pat_lower}*"):
                return True
        else:
            if pat_lower in name_lower:
                return True
    return False


def normalize_extensions(extensions: Optional[List[str]]) -> Set[str]:
    """
    Normalize list of file extensions to lowercase leading dot (e.g., 'pdf' -> '.pdf', '.docx' -> '.docx').
    """
    if not extensions:
        return set()
    result = set()
    for ext in extensions:
        cleaned = ext.strip().lower()
        if not cleaned:
            continue
        if not cleaned.startswith("."):
            cleaned = "." + cleaned
        result.add(cleaned)
    return result


def should_include_file(
    filename: str,
    file_include: Optional[List[str]] = None,
    file_exclude: Optional[List[str]] = None,
    ext_include: Optional[List[str]] = None,
    ext_exclude: Optional[List[str]] = None,
) -> bool:
    """
    Evaluate file against name patterns and extensions:
    1. If extension exclusion matches, exclude.
    2. If extension inclusion is specified and file extension does not match, exclude.
    3. If file name exclusion matches, exclude.
    4. If file name inclusion is specified and does not match, exclude.
    5. Otherwise include.
    """
    _, ext = os.path.splitext(filename)
    ext_lower = ext.lower()

    norm_ext_inc = normalize_extensions(ext_include)
    norm_ext_exc = normalize_extensions(ext_exclude)

    if norm_ext_exc and ext_lower in norm_ext_exc:
        return False

    if norm_ext_inc and ext_lower not in norm_ext_inc:
        return False

    if file_exclude and matches_patterns(filename, file_exclude):
        return False

    if file_include and not matches_patterns(filename, file_include):
        return False

    return True


def should_include_folder(
    folder_name: str,
    folder_include: Optional[List[str]] = None,
    folder_exclude: Optional[List[str]] = None,
) -> bool:
    """
    Evaluate whether a directory name should be traversed during recursive search.
    """
    if folder_exclude and matches_patterns(folder_name, folder_exclude):
        return False

    if folder_include and not matches_patterns(folder_name, folder_include):
        return False

    return True


def format_bytes(size_bytes: int) -> str:
    """Format bytes into human-readable string (KB, MB, GB)."""
    if size_bytes < 0:
        return "0 B"
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}" if unit != "B" else f"{size_bytes} B"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} PB"


def get_file_size_safe(filepath: str) -> int:
    """Get file size safely handling long paths."""
    try:
        return os.path.getsize(filepath)
    except OSError:
        if os.name == "nt":
            try:
                return os.path.getsize(normalize_long_path(filepath))
            except OSError:
                short_p = get_short_path_name(filepath)
                if short_p:
                    try:
                        return os.path.getsize(short_p)
                    except OSError:
                        pass
    return 0


def collect_files(
    target_path: str,
    recurse: bool = False,
    file_include: Optional[List[str]] = None,
    file_exclude: Optional[List[str]] = None,
    ext_include: Optional[List[str]] = None,
    ext_exclude: Optional[List[str]] = None,
    folder_include: Optional[List[str]] = None,
    folder_exclude: Optional[List[str]] = None,
) -> List[str]:
    """
    Discover all matching files under target_path respecting recursion and filters.
    """
    matched_files: List[str] = []
    target_path = os.path.abspath(target_path)

    if not os.path.exists(target_path):
        return matched_files

    if os.path.isfile(target_path):
        fname = os.path.basename(target_path)
        if should_include_file(fname, file_include, file_exclude, ext_include, ext_exclude):
            matched_files.append(target_path)
        return matched_files

    if not recurse:
        try:
            entries = os.listdir(target_path)
        except OSError as e:
            print(f"[!] Warning: Cannot read directory '{target_path}': {e}", file=sys.stderr)
            return matched_files

        for entry in sorted(entries):
            full_path = os.path.join(target_path, entry)
            if os.path.isfile(full_path):
                if should_include_file(entry, file_include, file_exclude, ext_include, ext_exclude):
                    matched_files.append(full_path)
        return matched_files

    for root, dirs, files in os.walk(target_path, topdown=True):
        # Prune subdirectories according to folder filter rules
        dirs[:] = [
            d for d in dirs
            if should_include_folder(d, folder_include, folder_exclude)
        ]

        for fname in sorted(files):
            if should_include_file(fname, file_include, file_exclude, ext_include, ext_exclude):
                matched_files.append(os.path.join(root, fname))

    return matched_files


def process_onedrive_files(
    target_dir: str,
    recurse: bool = False,
    file_include: Optional[List[str]] = None,
    file_exclude: Optional[List[str]] = None,
    ext_include: Optional[List[str]] = None,
    ext_exclude: Optional[List[str]] = None,
    folder_include: Optional[List[str]] = None,
    folder_exclude: Optional[List[str]] = None,
    dry_run: bool = False,
) -> Dict[str, any]:
    """
    Orchestrate scanning and marking files to free up space.
    Gracefully catches and logs errors for individual files without failing the run.
    """
    summary = {
        "target_directory": os.path.abspath(target_dir),
        "total_discovered": 0,
        "marked_for_freeing": 0,
        "already_freed": 0,
        "skipped": 0,
        "errors": 0,
        "bytes_freed_candidate": 0,
        "error_details": [],
        "dry_run": dry_run,
    }

    if not os.path.exists(target_dir):
        msg = f"Target path '{target_dir}' does not exist."
        print(f"[ERROR] {msg}", file=sys.stderr)
        summary["errors"] += 1
        summary["error_details"].append({"path": target_dir, "error": msg})
        return summary

    files = collect_files(
        target_path=target_dir,
        recurse=recurse,
        file_include=file_include,
        file_exclude=file_exclude,
        ext_include=ext_include,
        ext_exclude=ext_exclude,
        folder_include=folder_include,
        folder_exclude=folder_exclude,
    )

    summary["total_discovered"] = len(files)
    prefix_tag = "[DRY-RUN] " if dry_run else ""

    print(f"{prefix_tag}Found {len(files)} matching file(s) in '{target_dir}'.")
    if not files:
        return summary

    print("=" * 70)
    for idx, filepath in enumerate(files, start=1):
        rel_path = os.path.relpath(filepath, target_dir)
        file_size = get_file_size_safe(filepath)

        try:
            success, status = mark_file_free_space(filepath, dry_run=dry_run)
            if success:
                if status == "already_freed":
                    summary["already_freed"] += 1
                    print(f"[{idx}/{len(files)}] [ALREADY ONLINE-ONLY] {rel_path} ({format_bytes(file_size)})")
                else:
                    summary["marked_for_freeing"] += 1
                    summary["bytes_freed_candidate"] += file_size
                    action_name = "WOULD MARK TO FREE SPACE" if dry_run else "MARKED TO FREE SPACE"
                    print(f"[{idx}/{len(files)}] [{action_name}] {rel_path} ({format_bytes(file_size)})")
            else:
                summary["errors"] += 1
                summary["error_details"].append({"path": filepath, "error": status})
                print(f"[{idx}/{len(files)}] [FAILED] {rel_path} -> {status}", file=sys.stderr)
        except Exception as ex:
            summary["errors"] += 1
            summary["error_details"].append({"path": filepath, "error": str(ex)})
            print(f"[{idx}/{len(files)}] [ERROR] {rel_path} -> Unexpected error: {ex}", file=sys.stderr)

    return summary


def print_summary_report(summary: Dict[str, any]) -> None:
    """Print clean formatted summary of actions taken."""
    print("=" * 70)
    mode_str = " (DRY-RUN - NO CHANGES APPLIED)" if summary["dry_run"] else ""
    print(f"ONEDRIVE SPACE FREER SUMMARY{mode_str}")
    print("=" * 70)
    print(f"Target Directory     : {summary['target_directory']}")
    print(f"Total Discovered     : {summary['total_discovered']}")
    print(f"Marked to Free Space : {summary['marked_for_freeing']}")
    print(f"Already Online-Only  : {summary['already_freed']}")
    print(f"Failed / Errors      : {summary['errors']}")
    print(f"Local Space Impact   : {format_bytes(summary['bytes_freed_candidate'])}")

    if summary["error_details"]:
        print("\nErrors Encountered:")
        for item in summary["error_details"]:
            print(f"  - {item['path']}: {item['error']}")
    print("=" * 70)


def build_parser() -> argparse.ArgumentParser:
    """Construct CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="onedrive_space_freer.py",
        description="Mark files in a OneDrive directory for uploading and freeing up local disk space.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Scan top-level folder and mark all files to free up space
  python onedrive_space_freer.py -p "C:\\Users\\User\\OneDrive - Org\\Reports"

  # Recursive scan, dry-run preview
  python onedrive_space_freer.py -p "C:\\Users\\User\\OneDrive" -r --dry-run

  # Filter specific extensions (e.g. only .pbix, .zip, .mp4)
  python onedrive_space_freer.py -p "C:\\Users\\User\\OneDrive" -r --ext-include pbix zip mp4

  # Exclude temporary or log files and skip archive folders
  python onedrive_space_freer.py -p "C:\\Users\\User\\OneDrive" -r \\
    --file-exclude "*temp*" "*~*" \\
    --ext-exclude tmp log \\
    --folder-exclude "*Backup*" "*Archive*"
        """,
    )

    parser.add_argument(
        "-p",
        "--path",
        required=True,
        help="Path to the target OneDrive directory to process.",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Recursively scan subfolders.",
    )
    parser.add_argument(
        "--file-include",
        nargs="+",
        metavar="PATTERN",
        help="Substrings or wildcard patterns for file names to include.",
    )
    parser.add_argument(
        "--file-exclude",
        nargs="+",
        metavar="PATTERN",
        help="Substrings or wildcard patterns for file names to exclude.",
    )
    parser.add_argument(
        "--ext-include",
        nargs="+",
        metavar="EXT",
        help="File extensions to include (e.g. pbix, docx, pdf). Leading dot is optional.",
    )
    parser.add_argument(
        "--ext-exclude",
        nargs="+",
        metavar="EXT",
        help="File extensions to exclude (e.g. tmp, log). Leading dot is optional.",
    )
    parser.add_argument(
        "--folder-include",
        nargs="+",
        metavar="PATTERN",
        help="Substrings or wildcard patterns for subdirectories to include during recursion.",
    )
    parser.add_argument(
        "--folder-exclude",
        nargs="+",
        metavar="PATTERN",
        help="Substrings or wildcard patterns for subdirectories to exclude during recursion.",
    )
    parser.add_argument(
        "-n",
        "--dry-run",
        action="store_true",
        help="Simulate execution without modifying file attributes.",
    )

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    summary = process_onedrive_files(
        target_dir=args.path,
        recurse=args.recursive,
        file_include=args.file_include,
        file_exclude=args.file_exclude,
        ext_include=args.ext_include,
        ext_exclude=args.ext_exclude,
        folder_include=args.folder_include,
        folder_exclude=args.folder_exclude,
        dry_run=args.dry_run,
    )

    print_summary_report(summary)
    return 0 if summary["errors"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
