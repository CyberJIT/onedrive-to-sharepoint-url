#!/usr/bin/env python3
"""
onedrive_to_sharepoint_url.py

Resolves local OneDrive / SharePoint synced files to their corresponding
SharePoint / OneDrive Online URLs and exports the results to a CSV file.

Zero external dependencies - uses Python 3 standard library only.
Designed for Windows 11 CLI with cross-platform fallback for testing and CI.
"""

from __future__ import annotations

import argparse
import csv
import os
import posixpath
import sys
import urllib.parse
from dataclasses import dataclass
from typing import Dict, Iterator, List, Optional, Sequence, Tuple

# Try importing winreg on Windows; mock/handle gracefully on other OS
try:
    import winreg
except ImportError:
    winreg = None  # type: ignore


@dataclass
class SyncMapping:
    """Represents a local folder to remote URL namespace mapping."""
    mount_point: str
    url_namespace: str
    scope_type: str = "business"  # "business", "personal", or "manual"
    account_name: Optional[str] = None


def normalize_local_path(path: str) -> str:
    """Normalize local file path for consistent prefix matching.
    
    Converts forward slashes, strips trailing separators (unless root),
    and standardizes casing on Windows.
    """
    abs_path = os.path.abspath(os.path.expanduser(os.path.expandvars(path)))
    # Strip Windows extended-length prefix if present for clean string operations
    if abs_path.startswith("\\\\?\\UNC\\"):
        abs_path = "\\\\" + abs_path[8:]
    elif abs_path.startswith("\\\\?\\"):
        abs_path = abs_path[4:]
    return os.path.normpath(abs_path)


def get_windows_registry_mappings() -> List[SyncMapping]:
    """Inspect Windows Registry for active OneDrive / SharePoint sync scopes.
    
    Checks:
    1. HKCU\\Software\\SyncEngines\\Providers\\OneDrive (All synced libraries & folders)
    2. HKCU\\Software\\Microsoft\\OneDrive\\Accounts (Business1..N, Personal)
    """
    mappings: List[SyncMapping] = []
    if not winreg:
        return mappings

    # 1. Primary SyncEngines provider keys
    sync_engines_path = r"Software\SyncEngines\Providers\OneDrive"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sync_engines_path, 0, winreg.KEY_READ) as base_key:
            subkeys_count, _, _ = winreg.QueryInfoKey(base_key)
            for i in range(subkeys_count):
                try:
                    subkey_name = winreg.EnumKey(base_key, i)
                    sub_path = f"{sync_engines_path}\\{subkey_name}"
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_path, 0, winreg.KEY_READ) as sub_key:
                        mount_point = ""
                        url_namespace = ""
                        try:
                            mount_point, _ = winreg.QueryValueEx(sub_key, "MountPoint")
                        except OSError:
                            pass
                        
                        # UrlNamespace is the canonical SharePoint URL root
                        for url_prop in ("UrlNamespace", "WebUrl", "LibraryScopeUrl"):
                            try:
                                url_val, _ = winreg.QueryValueEx(sub_key, url_prop)
                                if url_val:
                                    url_namespace = url_val
                                    break
                            except OSError:
                                pass

                        if mount_point and url_namespace:
                            mappings.append(SyncMapping(
                                mount_point=normalize_local_path(mount_point),
                                url_namespace=url_namespace.rstrip("/"),
                                scope_type="business",
                                account_name=subkey_name,
                            ))
                except OSError:
                    continue
    except OSError:
        pass

    # 2. Accounts subkeys (Business and Personal roots)
    accounts_path = r"Software\Microsoft\OneDrive\Accounts"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, accounts_path, 0, winreg.KEY_READ) as acc_base_key:
            acc_count, _, _ = winreg.QueryInfoKey(acc_base_key)
            for i in range(acc_count):
                try:
                    acc_name = winreg.EnumKey(acc_base_key, i)
                    acc_sub_path = f"{accounts_path}\\{acc_name}"
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, acc_sub_path, 0, winreg.KEY_READ) as acc_key:
                        user_folder = ""
                        service_endpoint = ""
                        cid = ""

                        try:
                            user_folder, _ = winreg.QueryValueEx(acc_key, "UserFolder")
                        except OSError:
                            pass

                        try:
                            service_endpoint, _ = winreg.QueryValueEx(acc_key, "ServiceEndpointUri")
                        except OSError:
                            pass

                        try:
                            cid, _ = winreg.QueryValueEx(acc_key, "CID")
                        except OSError:
                            pass

                        if user_folder:
                            norm_folder = normalize_local_path(user_folder)
                            # Only add if not already covered by a more specific MountPoint
                            if service_endpoint:
                                # Often points to https://<tenant>-my.sharepoint.com/personal/...
                                mappings.append(SyncMapping(
                                    mount_point=norm_folder,
                                    url_namespace=service_endpoint.rstrip("/"),
                                    scope_type="business",
                                    account_name=acc_name,
                                Kupp = None if False else None))
                            elif acc_name.lower() == "personal" or cid:
                                personal_url = f"https://onedrive.live.com/?cid={cid}" if cid else "https://onedrive.live.com"
                                mappings.append(SyncMapping(
                                    mount_point=norm_folder,
                                    url_namespace=personal_url,
                                    scope_type="personal",
                                    account_name=acc_name,
                                ))
                except OSError:
                    continue
    except OSError:
        pass

    return mappings


def inspect_onedrive_ini_settings() -> List[SyncMapping]:
    """Inspect local OneDrive ClientPolicy.ini or settings .ini files as fallback."""
    mappings: List[SyncMapping] = []
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        return mappings

    settings_dir = os.path.join(local_app_data, "Microsoft", "OneDrive", "settings")
    if not os.path.isdir(settings_dir):
        return mappings

    # Check for Personal or Business directories
    for root, _, files in os.walk(settings_dir):
        for f in files:
            if f.endswith(".ini") and f != "ClientPolicy.ini":
                ini_path = os.path.join(root, f)
                try:
                    with open(ini_path, "r", encoding="utf-8", errors="ignore") as fp:
                        content = fp.read()
                        # Search for DavUrlNamespace or local library mapping lines if available
                        dav_ns = None
                        mount = None
                        for line in content.splitlines():
                            line = line.strip()
                            if line.startswith("DavUrlNamespace"):
                                parts = line.split("=", 1)
                                if len(parts) == 2:
                                    dav_ns = parts[1].strip().strip('"')
                            elif line.startswith("library =") or line.startswith("Folder ="):
                                parts = line.split("=", 1)
                                if len(parts) == 2:
                                    mount = parts[1].strip().strip('"')
                        if dav_ns and mount and os.path.exists(mount):
                            mappings.append(SyncMapping(
                                mount_point=normalize_local_path(mount),
                                url_namespace=dav_ns.rstrip("/"),
                                scope_type="business",
                                account_name="settings_ini",
                            ))
                except Exception:
                    continue

    return mappings


def encode_sharepoint_url_path(rel_path: str) -> str:
    """Encode relative path segment using standard SharePoint URL encoding.
    
    SharePoint expects '/' as directory separator, and spaces/special characters
    percent-encoded (%20, etc.), while preserving slashes.
    """
    clean_parts = [urllib.parse.quote(part) for part in rel_path.replace("\\", "/").strip("/").split("/") if part]
    return "/".join(clean_parts)


def resolve_file_to_urls(
    file_path: str,
    mappings: Sequence[SyncMapping]
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Resolve a local file path to:
    (sharepoint_web_url, direct_url, mount_point_matched)
    
    Uses longest prefix match on MountPoint to properly handle nested synced libraries.
    """
    norm_file = normalize_local_path(file_path)
    file_lower = norm_file.lower()

    best_mapping: Optional[SyncMapping] = None
    best_len = -1

    for m in mappings:
        mount_norm = normalize_local_path(m.mount_point)
        mount_lower = mount_norm.lower()

        # Ensure exact folder prefix match
        if file_lower == mount_lower or file_lower.startswith(mount_lower + os.sep):
            if len(mount_norm) > best_len:
                best_len = len(mount_norm)
                best_mapping = m

    if not best_mapping:
        return None, None, None

    mount_point = best_mapping.mount_point
    url_base = best_mapping.url_namespace

    # Extract relative path inside sync folder
    if len(norm_file) == len(mount_point):
        rel_path = ""
    else:
        rel_path = norm_file[len(mount_point):].lstrip(os.sep)

    encoded_rel = encode_sharepoint_url_path(rel_path) if rel_path else ""

    if best_mapping.scope_type == "personal" and "onedrive.live.com" in url_base:
        # Personal web url
        if encoded_rel:
            web_url = f"{url_base}&id={urllib.parse.quote(rel_path)}"
        else:
            web_url = url_base
        direct_url = web_url
    else:
        # Business / SharePoint Online library
        if encoded_rel:
            web_url = f"{url_base}/{encoded_rel}"
        else:
            web_url = url_base
        direct_url = web_url

    return web_url, direct_url, mount_point


def collect_files_from_folder(
    folder_path: str,
    recursive: bool = False
) -> Iterator[str]:
    """Yield file paths from a folder directory."""
    norm_folder = normalize_local_path(folder_path)
    if not os.path.isdir(norm_folder):
        raise ValueError(f"Path is not a valid directory: {folder_path}")

    if recursive:
        for root, _, files in os.walk(norm_folder):
            for f in sorted(files):
                # Skip desktop.ini and OneDrive hidden sync state files
                if f.lower() in ("desktop.ini", "thumbs.db", ".ds_store"):
                    continue
                yield os.path.join(root, f)
    else:
        for entry in sorted(os.scandir(norm_folder), key=lambda e: e.name):
            if entry.is_file():
                if entry.name.lower() in ("desktop.ini", "thumbs.db", ".ds_store"):
                    continue
                yield entry.path


def collect_files_from_csv(csv_path: str, column_name: Optional[str] = None) -> Iterator[str]:
    """Read local file paths from a CSV file.
    
    If column_name is given, reads that column. Otherwise uses first column.
    """
    norm_csv = normalize_local_path(csv_path)
    if not os.path.isfile(norm_csv):
        raise ValueError(f"CSV file not found: {csv_path}")

    with open(norm_csv, mode="r", encoding="utf-8-sig", errors="replace") as f:
        # Sniff delimiter or fallback to comma
        sample = f.read(2048)
        f.seek(0)
        delimiter = ","
        try:
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(sample)
            delimiter = dialect.delimiter
        except Exception:
            pass

        reader = csv.reader(f, delimiter=delimiter)
        header = next(reader, None)
        if not header:
            return

        col_idx = 0
        has_headers = False

        if column_name:
            for idx, col in enumerate(header):
                if col.strip().lower() == column_name.strip().lower():
                    col_idx = idx
                    has_headers = True
                    break
        else:
            # Check if header looks like an actual column name rather than a path
            first_val = header[0].strip()
            if not os.path.exists(first_val) and not (":" in first_val or "/" in first_val or "\\" in first_val):
                has_headers = True
            else:
                # First line might actually be data!
                yield first_val

        for row in reader:
            if not row or len(row) <= col_idx:
                continue
            val = row[col_idx].strip()
            if val:
                yield val


def process_files(
    file_paths: Iterator[str],
    mappings: Sequence[SyncMapping],
    output_csv_path: str
) -> Dict[str, int]:
    """Process files, resolve their SharePoint URLs, and write output CSV."""
    norm_out = normalize_local_path(output_csv_path)
    out_dir = os.path.dirname(norm_out)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    stats = {
        "total_files": 0,
        "resolved": 0,
        "unmatched": 0,
    }

    fieldnames = [
        "local_file_path",
        "file_name",
        "sharepoint_web_url",
        "sharepoint_mount_point",
        "resolution_status",
    ]

    with open(norm_out, mode="w", newline="", encoding="utf-8-sig") as out_f:
        writer = csv.DictWriter(out_f, fieldnames=fieldnames)
        writer.writeheader()

        for fpath in file_paths:
            stats["total_files"] += 1
            norm_f = normalize_local_path(fpath)
            fname = os.path.basename(norm_f)

            web_url, _, mount_pt = resolve_file_to_urls(norm_f, mappings)

            if web_url:
                status = "RESOLVED"
                stats["resolved"] += 1
            else:
                status = "NOT_IN_ONEDRIVE_MOUNT"
                stats["unmatched"] += 1
                web_url = ""
                mount_pt = ""

            writer.writerow({
                "local_file_path": norm_f,
                "file_name": fname,
                "sharepoint_web_url": web_url,
                "sharepoint_mount_point": mount_pt or "",
                "resolution_status": status,
            })

    return stats


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert local OneDrive / SharePoint file paths to SharePoint Online URLs and export to CSV."
    )
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument(
        "-p", "--path",
        dest="folder_path",
        help="Path to a local OneDrive directory to scan."
    )
    input_group.add_argument(
        "-c", "--csv",
        dest="input_csv",
        help="Path to an input CSV file containing local file paths."
    )

    parser.add_argument(
        "-o", "--output",
        dest="output_csv",
        required=True,
        help="Path for destination CSV file output."
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="Scan directory recursively (only valid with --path)."
    )
    parser.add_argument(
        "--column",
        dest="column_name",
        default=None,
        help="Column header name in input CSV containing file paths (default: first column)."
    )
    parser.add_argument(
        "--mount",
        dest="manual_mounts",
        action="append",
        metavar="LOCAL_PATH=WEB_URL",
        help="Manual mapping override in format: C:\\Local\\Folder=https://tenant.sharepoint.com/sites/library (can specify multiple times)."
    )
    parser.add_argument(
        "--list-mounts",
        action="store_true",
        help="List detected OneDrive / SharePoint mount points and exit."
    )

    return parser.parse_args(argv)


def get_all_sync_mappings(manual_mounts: Optional[List[str]] = None) -> List[SyncMapping]:
    """Retrieve all available sync mappings from registry, fallback configs, and CLI overrides."""
    mappings: List[SyncMapping] = []

    # 1. Manual CLI overrides (highest priority)
    if manual_mounts:
        for m_str in manual_mounts:
            if "=" in m_str:
                lp, url = m_str.split("=", 1)
                mappings.append(SyncMapping(
                    mount_point=normalize_local_path(lp.strip()),
                    url_namespace=url.strip().rstrip("/"),
                    scope_type="manual",
                    account_name="CLI_OVERRIDE",
                ))

    # 2. Windows Registry
    mappings.extend(get_windows_registry_mappings())

    # 3. Settings files fallback
    mappings.extend(inspect_onedrive_ini_settings())

    # Deduplicate while preserving priority (longest mount points first)
    unique_mappings: List[SyncMapping] = []
    seen = set()
    for m in mappings:
        key = (m.mount_point.lower(), m.url_namespace.lower())
        if key not in seen:
            seen.add(key)
            unique_mappings.append(m)

    # Sort descending by mount_point length so longest prefix matches first
    unique_mappings.sort(key=lambda x: len(x.mount_point), reverse=True)
    return unique_mappings


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    mappings = get_all_sync_mappings(args.manual_mounts)

    if args.list_mounts:
        print(f"Discovered {len(mappings)} OneDrive / SharePoint sync mapping(s):")
        for idx, m in enumerate(mappings, 1):
            print(f" [{idx}] [{m.scope_type}] Local: {m.mount_point}")
            print(f"      Remote: {m.url_namespace}")
        return 0

    if not mappings:
        print(
            "WARNING: No active OneDrive / SharePoint sync mappings detected via Windows Registry.\n"
            "If running outside standard sync or in testing, supply manual mappings using:\n"
            "  --mount \"C:\\Users\\User\\OneDrive - Company=https://company.sharepoint.com/sites/MySite\"",
            file=sys.stderr
        )

    # Collect input files
    try:
        if args.folder_path:
            file_iter = collect_files_from_folder(args.folder_path, recursive=args.recursive)
        else:
            file_iter = collect_files_from_csv(args.input_csv, column_name=args.column_name)

        stats = process_files(file_iter, mappings, args.output_csv)

        print("\n=== Resolution Summary ===")
        print(f"Total files checked : {stats['total_files']}")
        print(f"Successfully mapped : {stats['resolved']}")
        print(f"Unmatched (outside) : {stats['unmatched']}")
        print(f"Output exported to  : {args.output_csv}")

        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
