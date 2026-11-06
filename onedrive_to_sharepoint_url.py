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
import logging
import os
import posixpath
import re
import sys
import traceback
import urllib.parse
from dataclasses import dataclass
from typing import Dict, Iterator, List, Optional, Sequence, Tuple

# Try importing winreg on Windows; mock/handle gracefully on other OS
try:
    import winreg
except ImportError:
    winreg = None  # type: ignore


# ==============================================================================
# Path Sanitization & Safe Redaction Engine
# ==============================================================================

def get_redaction_patterns() -> List[Tuple[re.Pattern, str]]:
    """Build dynamic regular expression patterns to redact local usernames and user profile paths."""
    patterns = []
    
    # 1. System user names and home directory paths
    user_names = set()
    user_paths = []
    
    for env_var in ("USERNAME", "USER", "LOGNAME"):
        val = os.environ.get(env_var)
        if val and len(val) > 1:
            user_names.add(val)

    for env_var in ("USERPROFILE", "HOME"):
        val = os.environ.get(env_var)
        if val:
            user_paths.append(val)
            base = os.path.basename(val)
            if base and len(base) > 1:
                user_names.add(base)

    for p in user_paths:
        norm = os.path.normpath(p)
        escaped_norm = re.escape(norm)
        patterns.append((re.compile(escaped_norm, re.IGNORECASE), "<USERPROFILE>"))
        patterns.append((re.compile(re.escape(norm.replace("\\", "/")), re.IGNORECASE), "<USERPROFILE>"))

    # Generic Windows C:\Users\<Username> and C:\Documents and Settings\<Username>
    patterns.append((
        re.compile(r'([A-Za-z]:\\Users\\)[^\s\\\'\"\,]+', re.IGNORECASE),
        r'\1<USER>'
    ))
    patterns.append((
        re.compile(r'([A-Za-z]:/Users/)[^\s/\'\"\,]+', re.IGNORECASE),
        r'\1<USER>'
    ))
    # Unix /Users/<Username> or /home/<Username>
    patterns.append((
        re.compile(r'(/Users/)[^\s/\'\"\,]+', re.IGNORECASE),
        r'\1<USER>'
    ))
    patterns.append((
        re.compile(r'(/home/)[^\s/\'\"\,]+', re.IGNORECASE),
        r'\1<USER>'
    ))

    # Redact explicit username if still present
    for u in user_names:
        if len(u) >= 3:
            patterns.append((re.compile(re.escape(u), re.IGNORECASE), "<USER>"))

    return patterns


_REDACTION_PATTERNS = get_redaction_patterns()


def redact_path(text: str) -> str:
    """Sanitize any string containing local paths or usernames for safe log output."""
    if not isinstance(text, str):
        text = str(text)
    for pattern, replacement in _REDACTION_PATTERNS:
        text = pattern.sub(replacement, text)
    return text


class SanitizedFormatter(logging.Formatter):
    """Logging formatter that intercepts formatted messages and stack traces to redact local paths."""
    def format(self, record: logging.LogRecord) -> str:
        formatted = super().format(record)
        return redact_path(formatted)


def safe_excepthook(exc_type, exc_val, exc_tb):
    """Top-level unhandled exception hook redacting local user paths from tracebacks."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_val, exc_tb)
        return
    tb_text = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
    sanitized = redact_path(tb_text)
    sys.stderr.write(f"\n[CRITICAL UNHANDLED EXCEPTION]\n{sanitized}\n")


sys.excepthook = safe_excepthook

# Configure package logger
logger = logging.getLogger("OneDriveUrlResolver")
_log_handler = logging.StreamHandler(sys.stdout)
_log_handler.setFormatter(SanitizedFormatter("%(message)s"))
logger.addHandler(_log_handler)
logger.setLevel(logging.INFO)


def safe_print(msg: str, file=sys.stdout):
    """Print to console with guaranteed local path redaction."""
    print(redact_path(msg), file=file)


@dataclass
class SyncMapping:
    """Represents a local folder to remote URL namespace mapping."""
    mount_point: str
    url_namespace: str
    scope_type: str = "business"  # "business", "personal", or "manual"
    account_name: Optional[str] = None


def normalize_local_path(path: str) -> str:
    """Normalize local file path for consistent prefix matching.
    
    Handles paths with spaces, brackets, hashes, unicode accents.
    Converts forward slashes, strips trailing separators (unless root),
    and standardizes casing on Windows.
    """
    if not path:
        return ""
    clean = path.strip("\ufeff\u200b\u200e\u200f \t\r\n")
    if (clean.startswith('"') and clean.endswith('"')) or (clean.startswith("'") and clean.endswith("'")):
        clean = clean[1:-1].strip()

    expanded = os.path.expanduser(os.path.expandvars(clean))
    abs_path = os.path.abspath(expanded)
    
    # Strip Windows extended-length prefix if present for clean string operations
    if abs_path.startswith("\\\\?\\UNC\\"):
        abs_path = "\\\\" + abs_path[8:]
    elif abs_path.startswith("\\\\?\\"):
        abs_path = abs_path[4:]
    return os.path.normpath(abs_path)


def _safe_query_reg_value(key, value_names: Sequence[str]) -> Optional[str]:
    """Safely query any of the specified registry value names from an open key."""
    if not winreg or not key:
        return None
    for name in value_names:
        try:
            val, _ = winreg.QueryValueEx(key, name)
            if val and str(val).strip():
                return str(val).strip()
        except OSError:
            continue
    return None


def get_windows_registry_mappings(verbose: bool = False) -> List[SyncMapping]:
    """Inspect Windows Registry for active OneDrive / SharePoint sync scopes.
    
    Checks:
    1. HKCU\\Software\\SyncEngines\\Providers\\OneDrive (All synced libraries & folders)
    2. HKCU\\Software\\Microsoft\\OneDrive\\Accounts (Business1..N, Personal)
    3. HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\*\\ScopeIdToMountPointPathCache & ScopeIdToWebUrlPathCache
    """
    mappings: List[SyncMapping] = []
    if not winreg:
        if verbose:
            logger.debug("[Registry] winreg module not available on this platform.")
        return mappings

    # 1. Primary SyncEngines provider keys
    sync_engines_path = r"Software\SyncEngines\Providers\OneDrive"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sync_engines_path, 0, winreg.KEY_READ) as base_key:
            subkeys_count, _, _ = winreg.QueryInfoKey(base_key)
            if verbose:
                logger.debug(f"[Registry] Found {subkeys_count} subkeys in {sync_engines_path}")
            for i in range(subkeys_count):
                try:
                    subkey_name = winreg.EnumKey(base_key, i)
                    sub_path = f"{sync_engines_path}\\{subkey_name}"
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_path, 0, winreg.KEY_READ) as sub_key:
                        mount_point = _safe_query_reg_value(sub_key, ["MountPoint"])
                        url_namespace = _safe_query_reg_value(
                            sub_key,
                            ["UrlNamespace", "Url", "WebUrl", "LibraryScopeUrl"]
                        )

                        if mount_point and url_namespace:
                            norm_mount = normalize_local_path(mount_point)
                            clean_url = url_namespace.rstrip("/")
                            mappings.append(SyncMapping(
                                mount_point=norm_mount,
                                url_namespace=clean_url,
                                scope_type="business",
                                account_name=subkey_name,
                            ))
                            if verbose:
                                logger.debug(f"[Registry SyncEngines] Mapped: {norm_mount} -> {clean_url}")
                except OSError as e:
                    if verbose:
                        logger.debug(f"[Registry SyncEngines] Skip subkey index {i}: {e}")
                    continue
    except OSError as e:
        if verbose:
            logger.debug(f"[Registry SyncEngines] Key not found or inaccessible: {e}")

    # 2. Accounts subkeys (Business and Personal roots & ScopeId Caches)
    accounts_path = r"Software\Microsoft\OneDrive\Accounts"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, accounts_path, 0, winreg.KEY_READ) as acc_base_key:
            acc_count, _, _ = winreg.QueryInfoKey(acc_base_key)
            if verbose:
                logger.debug(f"[Registry Accounts] Found {acc_count} accounts in {accounts_path}")
            for i in range(acc_count):
                try:
                    acc_name = winreg.EnumKey(acc_base_key, i)
                    acc_sub_path = f"{accounts_path}\\{acc_name}"
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, acc_sub_path, 0, winreg.KEY_READ) as acc_key:
                        user_folder = _safe_query_reg_value(acc_key, ["UserFolder"])
                        service_endpoint = _safe_query_reg_value(
                            acc_key,
                            ["ServiceEndpointUri", "ServiceEndpointUriSharedWithMe", "DavUrlNamespace"]
                        )
                        cid = _safe_query_reg_value(acc_key, ["CID"])

                        if user_folder:
                            norm_folder = normalize_local_path(user_folder)
                            if service_endpoint:
                                clean_endpoint = service_endpoint.rstrip("/")
                                mappings.append(SyncMapping(
                                    mount_point=norm_folder,
                                    url_namespace=clean_endpoint,
                                    scope_type="business",
                                    account_name=acc_name,
                                ))
                                if verbose:
                                    logger.debug(f"[Registry Account Root] Mapped {acc_name}: {norm_folder} -> {clean_endpoint}")
                            elif acc_name.lower() == "personal" or cid:
                                personal_url = f"https://onedrive.live.com/?cid={cid}" if cid else "https://onedrive.live.com"
                                mappings.append(SyncMapping(
                                    mount_point=norm_folder,
                                    url_namespace=personal_url,
                                    scope_type="personal",
                                    account_name=acc_name,
                                ))
                                if verbose:
                                    logger.debug(f"[Registry Personal Root] Mapped {acc_name}: {norm_folder} -> {personal_url}")

                        # Check ScopeId caches inside this account
                        try:
                            scope_mounts: Dict[str, str] = {}
                            scope_urls: Dict[str, str] = {}
                            
                            mount_cache_path = f"{acc_sub_path}\\ScopeIdToMountPointPathCache"
                            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, mount_cache_path, 0, winreg.KEY_READ) as mc_key:
                                mc_values_count, _, _ = winreg.QueryInfoKey(mc_key)
                                for j in range(mc_values_count):
                                    v_name, v_val, _ = winreg.EnumValue(mc_key, j)
                                    if v_val:
                                        scope_mounts[v_name] = str(v_val)

                            url_cache_path = f"{acc_sub_path}\\ScopeIdToWebUrlPathCache"
                            try:
                                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, url_cache_path, 0, winreg.KEY_READ) as uc_key:
                                    uc_values_count, _, _ = winreg.QueryInfoKey(uc_key)
                                    for j in range(uc_values_count):
                                        v_name, v_val, _ = winreg.EnumValue(uc_key, j)
                                        if v_val:
                                            scope_urls[v_name] = str(v_val)
                            except OSError:
                                pass

                            for s_id, m_path in scope_mounts.items():
                                u_target = scope_urls.get(s_id)
                                if u_target:
                                    mappings.append(SyncMapping(
                                        mount_point=normalize_local_path(m_path),
                                        url_namespace=u_target.rstrip("/"),
                                        scope_type="business",
                                        account_name=f"{acc_name}_{s_id[:8]}",
                                    ))
                                    if verbose:
                                        logger.debug(f"[Registry ScopeIdCache] Mapped {s_id}: {m_path} -> {u_target}")
                        except OSError:
                            pass
                except OSError as e:
                    if verbose:
                        logger.debug(f"[Registry Accounts] Skip account index {i}: {e}")
                    continue
    except OSError as e:
        if verbose:
            logger.debug(f"[Registry Accounts] Key not found or inaccessible: {e}")

    return mappings


def inspect_onedrive_ini_settings(verbose: bool = False) -> List[SyncMapping]:
    """Inspect local OneDrive ClientPolicy.ini or settings .ini files as fallback."""
    mappings: List[SyncMapping] = []
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        return mappings

    settings_dir = os.path.join(local_app_data, "Microsoft", "OneDrive", "settings")
    if not os.path.isdir(settings_dir):
        return mappings

    for root, _, files in os.walk(settings_dir):
        for f in files:
            if f.endswith(".ini") and f != "ClientPolicy.ini":
                ini_path = os.path.join(root, f)
                try:
                    with open(ini_path, "r", encoding="utf-8", errors="ignore") as fp:
                        content = fp.read()
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
                        if dav_ns and mount:
                            mappings.append(SyncMapping(
                                mount_point=normalize_local_path(mount),
                                url_namespace=dav_ns.rstrip("/"),
                                scope_type="business",
                                account_name="settings_ini",
                            ))
                            if verbose:
                                logger.debug(f"[Settings INI] Found {mount} -> {dav_ns}")
                except Exception:
                    continue

    return mappings


def encode_sharepoint_url_path(rel_path: str) -> str:
    """Encode relative path segment using standard SharePoint URL encoding.
    
    Preserves directories and handles spaces, symbols, and Unicode while escaping
    special characters (%20, # -> %23, etc.) conforming to SharePoint standards.
    """
    clean_parts = [
        urllib.parse.quote(part) 
        for part in rel_path.replace("\\", "/").strip("/").split("/") 
        if part
    ]
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

        # Match exact root or folder boundary
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
        if encoded_rel:
            web_url = f"{url_base}&id={urllib.parse.quote(rel_path)}"
        else:
            web_url = url_base
        direct_url = web_url
    else:
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
        raise ValueError(f"Path is not a valid directory: {redact_path(folder_path)}")

    if recursive:
        for root, _, files in os.walk(norm_folder):
            for f in sorted(files):
                if f.lower() in ("desktop.ini", "thumbs.db", ".ds_store"):
                    continue
                yield os.path.join(root, f)
    else:
        for entry in sorted(os.scandir(norm_folder), key=lambda e: e.name):
            if entry.is_file():
                if entry.name.lower() in ("desktop.ini", "thumbs.db", ".ds_store"):
                    continue
                yield entry.path


def _is_path_like(text: str) -> bool:
    """Heuristic to check if a string looks like a file/directory path rather than a column header."""
    clean = text.strip("\ufeff\u200b\u200e\u200f \"'\t\r\n")
    if not clean:
        return False
    # Drive letter: C:\ or C:/
    if len(clean) >= 2 and clean[1] == ":" and (clean[0].isalpha()):
        return True
    # UNC or root path
    if clean.startswith("\\\\") or clean.startswith("//") or clean.startswith("/") or clean.startswith("\\"):
        return True
    # Common separators with extension
    if ("\\" in clean or "/" in clean) and "." in os.path.basename(clean):
        return True
    return False


def collect_files_from_csv(csv_path: str, column_name: Optional[str] = None) -> Iterator[str]:
    """Read local file paths from a CSV file.
    
    Guarantees that spaces, parentheses, brackets, hashes, and special characters
    inside file paths are preserved and NEVER truncated by space delimiters.
    """
    norm_csv = normalize_local_path(csv_path)
    if not os.path.isfile(norm_csv):
        raise ValueError(f"CSV file not found: {redact_path(csv_path)}")

    with open(norm_csv, mode="r", encoding="utf-8-sig", errors="replace") as f:
        lines = [line for line in f]

    if not lines:
        return

    # 1. Detect standard CSV delimiter strictly from standard symbols (comma, semicolon, tab, pipe)
    # NEVER include ' ' (space) or ':' (colon) as delimiter candidates!
    first_lines = lines[:20]
    sample_text = "".join(first_lines)

    delimiter = ","
    try:
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(sample_text, delimiters=",\t;|")
        if dialect.delimiter in (",", "\t", ";", "|"):
            delimiter = dialect.delimiter
    except Exception:
        # Fallback heuristic: check which delimiter appears most consistently
        for cand in (",", "\t", ";", "|"):
            if cand in sample_text:
                delimiter = cand
                break

    # 2. Parse using Python's standard csv.reader with the strict non-space delimiter
    reader = csv.reader(lines, delimiter=delimiter)
    all_rows = list(reader)

    if not all_rows:
        return

    first_row = [c.strip() for c in all_rows[0] if c is not None]

    # Determine target column index
    col_idx = 0
    start_row_idx = 0

    if column_name:
        for idx, col in enumerate(first_row):
            if col.strip().lower() == column_name.strip().lower():
                col_idx = idx
                start_row_idx = 1
                break
    else:
        # If no explicit column given, check if first row contains a column header or is already a path
        first_entry = first_row[0] if first_row else ""
        if _is_path_like(first_entry):
            # First row is actual data (e.g. single column of paths without headers)
            start_row_idx = 0
        else:
            # First row is likely header
            # Find candidate column by name if present
            candidate_found = False
            for idx, col in enumerate(first_row):
                if any(k in col.lower() for k in ("path", "file", "location", "full_name", "filepath", "filename", "url")):
                    col_idx = idx
                    candidate_found = True
                    break
            start_row_idx = 1

    for row_idx in range(start_row_idx, len(all_rows)):
        row = all_rows[row_idx]
        if not row:
            continue
        if len(row) <= col_idx:
            # If row was not split by delimiter, check full raw line
            raw_line = lines[row_idx].strip("\r\n")
            if raw_line:
                clean_raw = raw_line.strip("\ufeff\u200b\u200e\u200f \"'")
                if clean_raw:
                    yield clean_raw
            continue

        raw_val = row[col_idx]
        if not raw_val:
            continue

        # Strip surrounding quotes and invisible zero-width spaces without touching internal spaces
        clean_val = raw_val.strip("\ufeff\u200b\u200e\u200f \t\r\n")
        if (clean_val.startswith('"') and clean_val.endswith('"')) or (clean_val.startswith("'") and clean_val.endswith("'")):
            clean_val = clean_val[1:-1].strip()

        if clean_val:
            yield clean_val


def process_files(
    file_paths: Iterator[str],
    mappings: Sequence[SyncMapping],
    output_csv_path: str,
    verbose: bool = False
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
                if verbose:
                    safe_print(f"  [OK] {fname} -> {web_url}")
            else:
                status = "NOT_IN_ONEDRIVE_MOUNT"
                stats["unmatched"] += 1
                web_url = ""
                mount_pt = ""
                if verbose:
                    safe_print(f"  [SKIP] {fname} (Not inside any detected OneDrive mount)")

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
    input_group = parser.add_mutually_exclusive_group(required=False)
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
        default=None,
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
        help="Column header name in input CSV containing file paths (default: auto-detected or first column)."
    )
    parser.add_argument(
        "--mount",
        dest="manual_mounts",
        action="append",
        metavar="LOCAL_PATH=WEB_URL",
        help="Manual mapping override in format: C:\\Local\\Folder=https://tenant.sharepoint.com/sites/library"
    )
    parser.add_argument(
        "--list-mounts",
        action="store_true",
        help="List detected OneDrive / SharePoint mount points and exit."
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose diagnostics logging with automatic local path redaction."
    )

    args = parser.parse_args(argv)
    if not args.list_mounts:
        if not (args.folder_path or args.input_csv):
            parser.error("one of the arguments -p/--path or -c/--csv is required (unless --list-mounts is specified)")
        if not args.output_csv:
            parser.error("the following arguments are required: -o/--output")

    return args


def get_all_sync_mappings(
    manual_mounts: Optional[List[str]] = None,
    verbose: bool = False
) -> List[SyncMapping]:
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
                if verbose:
                    safe_print(f"[CLI Mount Override] {lp.strip()} -> {url.strip()}")

    # 2. Windows Registry
    mappings.extend(get_windows_registry_mappings(verbose=verbose))

    # 3. Settings files fallback
    mappings.extend(inspect_onedrive_ini_settings(verbose=verbose))

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

    if args.verbose:
        logger.setLevel(logging.DEBUG)

    mappings = get_all_sync_mappings(args.manual_mounts, verbose=args.verbose)

    if args.list_mounts:
        safe_print(f"\nDiscovered {len(mappings)} OneDrive / SharePoint sync mapping(s):")
        for idx, m in enumerate(mappings, 1):
            safe_print(f" [{idx}] [{m.scope_type}] Local: {m.mount_point}")
            safe_print(f"      Remote: {m.url_namespace}")
        return 0

    if not mappings:
        safe_print(
            "WARNING: No active OneDrive / SharePoint sync mappings detected via Windows Registry.\n"
            "If running outside standard sync or in testing, supply manual mappings using:\n"
            "  --mount \"C:\\Users\\<USER>\\OneDrive - Company=https://company.sharepoint.com/sites/MySite\"",
            file=sys.stderr
        )

    # Collect input files
    try:
        if args.folder_path:
            file_iter = collect_files_from_folder(args.folder_path, recursive=args.recursive)
        else:
            file_iter = collect_files_from_csv(args.input_csv, column_name=args.column_name)

        stats = process_files(file_iter, mappings, args.output_csv, verbose=args.verbose)

        safe_print("\n=== Resolution Summary ===")
        safe_print(f"Total files checked : {stats['total_files']}")
        safe_print(f"Successfully mapped : {stats['resolved']}")
        safe_print(f"Unmatched (outside) : {stats['unmatched']}")
        safe_print(f"Output exported to  : {args.output_csv}")

        return 0
    except Exception as e:
        safe_print(f"ERROR: {e}", file=sys.stderr)
        if args.verbose:
            tb = "".join(traceback.format_exc())
            safe_print(tb, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
