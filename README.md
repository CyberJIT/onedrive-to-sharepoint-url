# OneDrive to SharePoint Online URL Resolver

A lightweight, zero-dependency dual **Python** and **PowerShell** CLI toolset designed for Windows 11 (and compatible cross-platform).

Given a folder path (with optional recursion) or an input CSV containing local file paths from a synchronized OneDrive or SharePoint Online directory, it auto-discovers active sync mount points from the Windows Registry and maps each file to its corresponding **SharePoint / OneDrive Online web URL**, exporting the results into a CSV file.

---

## Features

- **Dual Implementation**: Run using either pure **Python 3** (`onedrive_to_sharepoint_url.py`) or native **PowerShell 5.1+ / 7+** (`Resolve-OneDriveToSharePointUrl.ps1`).
- **Zero External Dependencies**: Uses only standard libraries (`winreg`, `urllib.parse`, `csv`, `os`, `argparse` in Python; built-in cmdlets and .NET in PowerShell).
- **No Administrator Rights Required**: Reads user-level configuration from `HKCU:\Software\SyncEngines\Providers\OneDrive` and `HKCU:\Software\Microsoft\OneDrive\Accounts`.
- **Intelligent Longest-Prefix Matching**: Automatically handles nested SharePoint synced libraries, shared folders, and Personal OneDrive roots.
- **Flexible Inputs**:
  - Scan a directory (with optional recursive `-r` / `-Recursive` traversal).
  - Or process an existing CSV file containing a list of file paths.
- **URL-Safe Encoding**: Encodes spaces and special characters conforming to SharePoint web standards while preserving directory hierarchies.
- **Fallback / Custom Mapping Support**: Allows `--mount` / `-Mount` CLI parameter overrides for air-gapped systems, custom SharePoint tenants, or testing environments.

---

## Output CSV Schema

The generated CSV contains:
| Column | Description |
| :--- | :--- |
| `local_file_path` | Absolute local path to the target file. |
| `file_name` | Name of the file with extension. |
| `sharepoint_web_url` | Full SharePoint Online / OneDrive browser URL to the file. |
| `sharepoint_mount_point` | The matched OneDrive/SharePoint local sync root. |
| `resolution_status` | `RESOLVED` if mapped successfully, or `NOT_IN_ONEDRIVE_MOUNT`. |

---

## Usage: Python CLI (`onedrive_to_sharepoint_url.py`)

### 1. Scan a folder recursively
```bash
python onedrive_to_sharepoint_url.py --path "C:\Users\John\OneDrive - Contoso\Finance" --recursive --output "finance_urls.csv"
```

### 2. Process an input CSV with file paths
```bash
python onedrive_to_sharepoint_url.py --csv "input_files.csv" --column "FilePath" --output "resolved_urls.csv"
```

### 3. List discovered OneDrive sync points
```bash
python onedrive_to_sharepoint_url.py --list-mounts
```

### 4. Provide manual mount override (useful if registry is unavailable)
```bash
python onedrive_to_sharepoint_url.py --path "C:\Data\Project" --recursive --mount "C:\Data\Project=https://contoso.sharepoint.com/sites/PMO/Shared Documents" --output "output.csv"
```

---

## Usage: PowerShell CLI (`Resolve-OneDriveToSharePointUrl.ps1`)

### 1. Scan a folder recursively
```powershell
.\Resolve-OneDriveToSharePointUrl.ps1 -Path "C:\Users\John\OneDrive - Contoso\Finance" -Recursive -Output ".\finance_urls.csv"
```

### 2. Process an input CSV with file paths
```powershell
.\Resolve-OneDriveToSharePointUrl.ps1 -Csv ".\input_files.csv" -Column "FilePath" -Output ".\resolved_urls.csv"
```

### 3. List discovered OneDrive sync points
```powershell
.\Resolve-OneDriveToSharePointUrl.ps1 -ListMounts
```

### 4. Provide manual mount override
```powershell
.\Resolve-OneDriveToSharePointUrl.ps1 -Path "C:\Data\Project" -Recursive -Mount "C:\Data\Project=https://contoso.sharepoint.com/sites/PMO/Shared Documents" -Output ".\output.csv"
```

---

## Registry Discovery Mechanism

The tool queries user-level Windows registry paths populated by the Microsoft OneDrive sync client:
1. `HKCU\Software\SyncEngines\Providers\OneDrive`: Contains subkeys for every synced SharePoint document library or OneDrive account, exposing:
   - `MountPoint`: Local file directory (e.g., `C:\Users\<user>\<Tenant>\<Site> - Documents`)
   - `UrlNamespace` / `WebUrl` / `LibraryScopeUrl`: Remote SharePoint Online URL root
2. `HKCU\Software\Microsoft\OneDrive\Accounts`: Contains account roots (`UserFolder`, `ServiceEndpointUri`, `CID`).

---

## Running Tests

Run the Python unit test suite:
```bash
python -m unittest test_onedrive_to_sharepoint_url.py
```

---

## License

MIT
