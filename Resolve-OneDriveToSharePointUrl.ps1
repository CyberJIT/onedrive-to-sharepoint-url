<#
.SYNOPSIS
    Resolves local OneDrive and SharePoint synced file paths to SharePoint Online URLs
    and exports them to a CSV file.

.DESCRIPTION
    Resolve-OneDriveToSharePointUrl inspects the Windows Registry for OneDrive / SharePoint
    sync scopes (HKCU:\Software\SyncEngines\Providers\OneDrive and HKCU:\Software\Microsoft\OneDrive\Accounts).
    It scans a local directory (optionally recursive) or reads an input CSV file containing paths,
    matches each file to its corresponding MountPoint, and constructs the user-friendly
    SharePoint Online web URL.

    Pure native PowerShell 5.1+ / PowerShell 7+ implementation with ZERO required external
    modules or administrative rights. All screen logs redact local user profile paths.

.PARAMETER Path
    Path to a local folder synchronized with OneDrive / SharePoint.

.PARAMETER Csv
    Path to an input CSV file containing local file paths.

.PARAMETER Output
    Destination path for the generated CSV report.

.PARAMETER Recursive
    Switch to scan folders and subfolders recursively.

.PARAMETER Column
    Header name of the column in the input CSV containing file paths (default: auto-detected or first column).

.PARAMETER Mount
    Optional manual mapping override(s) formatted as 'LocalPath=RemoteUrl'.
    Example: -Mount "C:\Users\<USER>\OneDrive - Contoso=https://contoso.sharepoint.com/sites/Team"

.PARAMETER ListMounts
    Switch to display all discovered OneDrive / SharePoint mount points and exit.

.PARAMETER Verbose
    Common parameter to output detailed diagnostic mapping steps.
#>

[CmdletBinding(DefaultParameterSetName = "Folder")]
param (
    [Parameter(Mandatory = $true, ParameterSetName = "Folder", Position = 0)]
    [string]$Path,

    [Parameter(Mandatory = $true, ParameterSetName = "CsvInput")]
    [string]$Csv,

    [Parameter(Mandatory = $true, ParameterSetName = "Folder")]
    [Parameter(Mandatory = $true, ParameterSetName = "CsvInput")]
    [string]$Output,

    [Parameter(ParameterSetName = "Folder")]
    [switch]$Recursive,

    [Parameter(ParameterSetName = "CsvInput")]
    [string]$Column,

    [Parameter(Mandatory = $false)]
    [string[]]$Mount,

    [Parameter(Mandatory = $true, ParameterSetName = "List")]
    [switch]$ListMounts
)

# Flexible property handling for registry objects
Set-StrictMode -Off

function Redact-LocalPath {
    param ([string]$InputText)
    if ([string]::IsNullOrWhiteSpace($InputText)) { return "" }
    
    $scrubbed = $InputText

    # 1. Redact specific user profile directory if present
    $userProfile = [System.Environment]::GetEnvironmentVariable("USERPROFILE")
    if (-not [string]::IsNullOrWhiteSpace($userProfile)) {
        $scrubbed = $scrubbed -replace [regex]::Escape($userProfile), "<USERPROFILE>"
        $scrubbed = $scrubbed -replace [regex]::Escape($userProfile.Replace('\', '/')), "<USERPROFILE>"
    }

    $home = [System.Environment]::GetEnvironmentVariable("HOME")
    if (-not [string]::IsNullOrWhiteSpace($home)) {
        $scrubbed = $scrubbed -replace [regex]::Escape($home), "<USERPROFILE>"
    }

    # 2. Redact Windows C:\Users\<Username>
    $scrubbed = [regex]::Replace($scrubbed, '([A-Za-z]:\\Users\\)[^\s\\''"\,]+', '$1<USER>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $scrubbed = [regex]::Replace($scrubbed, '([A-Za-z]:/Users/)[^\s/''"\,]+', '$1<USER>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    # 3. Redact explicit username
    $username = [System.Environment]::GetEnvironmentVariable("USERNAME")
    if (-not [string]::IsNullOrWhiteSpace($username) -and $username.Length -ge 3) {
        $scrubbed = $scrubbed -replace [regex]::Escape($username), "<USER>"
    }

    return $scrubbed
}

function Write-SafeHost {
    param (
        [string]$Message,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::White
    )
    $redacted = Redact-LocalPath $Message
    Write-Host $redacted -ForegroundColor $ForegroundColor
}

function Write-SafeWarning {
    param ([string]$Message)
    $redacted = Redact-LocalPath $Message
    Write-Warning $redacted
}

function Normalize-LocalPath {
    param ([string]$InputPath)
    if ([string]::IsNullOrWhiteSpace($InputPath)) { return "" }
    
    # Strip quotes and zero-width characters while preserving internal spaces
    $cleaned = $InputPath.Trim(" `t`r`n`uFEFF`u200B`u200E`u200F")
    if (($cleaned.StartsWith('"') -and $cleaned.EndsWith('"')) -or ($cleaned.StartsWith("'") -and $cleaned.EndsWith("'"))) {
        $cleaned = $cleaned.Substring(1, $cleaned.Length - 2).Trim()
    }
    
    $expanded = [System.Environment]::ExpandEnvironmentVariables($cleaned)
    $resolved = [System.IO.Path]::GetFullPath($expanded)
    
    # Remove extended length prefix if present
    if ($resolved.StartsWith("\\?\UNC\")) {
        $resolved = "\\" + $resolved.Substring(8)
    } elseif ($resolved.StartsWith("\\?\")) {
        $resolved = $resolved.Substring(4)
    }
    
    return $resolved.TrimEnd('\')
}

function Safe-GetRegistryValue {
    param (
        [Microsoft.Win32.RegistryKey]$Key,
        [string[]]$Names
    )
    if (-not $Key) { return $null }
    foreach ($n in $Names) {
        try {
            $val = $Key.GetValue($n, $null)
            if ($null -ne $val -and -not [string]::IsNullOrWhiteSpace($val.ToString())) {
                return $val.ToString().Trim()
            }
        } catch {
            # Continue checking fallback property names
        }
    }
    return $null
}

function Get-OneDriveSyncMappings {
    param ([string[]]$ManualMounts)

    $mappings = [System.Collections.Generic.List[PSCustomObject]]::new()

    # 1. Manual mount overrides (highest priority)
    if ($ManualMounts) {
        foreach ($m in $ManualMounts) {
            if ($m -like "*=*") {
                $parts = $m -split "=", 2
                $mappings.Add([PSCustomObject]@{
                    MountPoint   = Normalize-LocalPath $parts[0].Trim()
                    UrlNamespace = $parts[1].Trim().TrimEnd('/')
                    ScopeType    = "Manual"
                    AccountName  = "CLI_OVERRIDE"
                })
            }
        }
    }

    # 2. Windows Registry: HKCU:\Software\SyncEngines\Providers\OneDrive
    $syncEnginesPath = "Software\SyncEngines\Providers\OneDrive"
    try {
        $hkcu = [Microsoft.Win32.Registry]::CurrentUser
        $baseKey = $hkcu.OpenSubKey($syncEnginesPath)
        if ($baseKey) {
            $subKeyNames = $baseKey.GetSubKeyNames()
            Write-Verbose "[Registry SyncEngines] Found $($subKeyNames.Count) subkeys."
            foreach ($subName in $subKeyNames) {
                try {
                    $subKey = $baseKey.OpenSubKey($subName)
                    if ($subKey) {
                        $mountPoint = Safe-GetRegistryValue -Key $subKey -Names @("MountPoint")
                        $urlNamespace = Safe-GetRegistryValue -Key $subKey -Names @("UrlNamespace", "Url", "WebUrl", "LibraryScopeUrl")

                        if (-not [string]::IsNullOrWhiteSpace($mountPoint) -and -not [string]::IsNullOrWhiteSpace($urlNamespace)) {
                            $normMount = Normalize-LocalPath $mountPoint
                            $cleanUrl = $urlNamespace.TrimEnd('/')
                            $mappings.Add([PSCustomObject]@{
                                MountPoint   = $normMount
                                UrlNamespace = $cleanUrl
                                ScopeType    = "Business"
                                AccountName  = $subName
                            })
                            Write-Verbose "[Registry SyncEngines] Mapped $normMount -> $cleanUrl"
                        }
                        $subKey.Close()
                    }
                } catch {
                    # Ignore individual inaccessible subkey
                }
            }
            $baseKey.Close()
        }
    } catch {
        Write-SafeWarning "Could not read SyncEngines registry: $_"
    }

    # 3. Windows Registry: HKCU:\Software\Microsoft\OneDrive\Accounts
    $accountsPath = "Software\Microsoft\OneDrive\Accounts"
    try {
        $hkcu = [Microsoft.Win32.Registry]::CurrentUser
        $baseKey = $hkcu.OpenSubKey($accountsPath)
        if ($baseKey) {
            $subKeyNames = $baseKey.GetSubKeyNames()
            Write-Verbose "[Registry Accounts] Found $($subKeyNames.Count) accounts."
            foreach ($subName in $subKeyNames) {
                try {
                    $subKey = $baseKey.OpenSubKey($subName)
                    if ($subKey) {
                        $userFolder = Safe-GetRegistryValue -Key $subKey -Names @("UserFolder")
                        $serviceUri = Safe-GetRegistryValue -Key $subKey -Names @("ServiceEndpointUri", "ServiceEndpointUriSharedWithMe", "DavUrlNamespace")
                        $cid = Safe-GetRegistryValue -Key $subKey -Names @("CID")

                        if (-not [string]::IsNullOrWhiteSpace($userFolder)) {
                            $normFolder = Normalize-LocalPath $userFolder
                            if (-not [string]::IsNullOrWhiteSpace($serviceUri)) {
                                $cleanEndpoint = $serviceUri.TrimEnd('/')
                                $mappings.Add([PSCustomObject]@{
                                    MountPoint   = $normFolder
                                    UrlNamespace = $cleanEndpoint
                                    ScopeType    = "Business"
                                    AccountName  = $subName
                                })
                                Write-Verbose "[Registry Accounts Root] Mapped $normFolder -> $cleanEndpoint"
                            } elseif ($subName -ieq "Personal" -or -not [string]::IsNullOrWhiteSpace($cid)) {
                                $pUrl = if ($cid) { "https://onedrive.live.com/?cid=$cid" } else { "https://onedrive.live.com" }
                                $mappings.Add([PSCustomObject]@{
                                    MountPoint   = $normFolder
                                    UrlNamespace = $pUrl
                                    ScopeType    = "Personal"
                                    AccountName  = $subName
                                })
                                Write-Verbose "[Registry Personal Root] Mapped $normFolder -> $pUrl"
                            }
                        }

                        # Check ScopeId caches
                        try {
                            $mcKey = $subKey.OpenSubKey("ScopeIdToMountPointPathCache")
                            $ucKey = $subKey.OpenSubKey("ScopeIdToWebUrlPathCache")
                            if ($mcKey) {
                                foreach ($valName in $mcKey.GetValueNames()) {
                                    $mPath = $mcKey.GetValue($valName)
                                    $uTarget = if ($ucKey) { $ucKey.GetValue($valName) } else { $null }
                                    if ($mPath -and $uTarget) {
                                        $normM = Normalize-LocalPath $mPath
                                        $cleanU = $uTarget.ToString().Trim().TrimEnd('/')
                                        $mappings.Add([PSCustomObject]@{
                                            MountPoint   = $normM
                                            UrlNamespace = $cleanU
                                            ScopeType    = "Business"
                                            AccountName  = "$($subName)_$($valName.Substring(0, [Math]::Min(8, $valName.Length)))"
                                        })
                                        Write-Verbose "[Registry ScopeIdCache] Mapped $normM -> $cleanU"
                                    }
                                }
                                $mcKey.Close()
                            }
                            if ($ucKey) { $ucKey.Close() }
                        } catch {
                            # Ignore scope cache errors
                        }

                        $subKey.Close()
                    }
                } catch {
                    # Ignore individual account key error
                }
            }
            $baseKey.Close()
        }
    } catch {
        Write-SafeWarning "Could not read OneDrive Accounts registry: $_"
    }

    # Sort descending by MountPoint length (longest matching folder path wins)
    return $mappings | Sort-Object -Property { $_.MountPoint.Length } -Descending
}

function Encode-SharePointUrlPath {
    param ([string]$RelativePath)
    
    $cleanParts = $RelativePath.Replace('\', '/').Trim('/').Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $encodedParts = foreach ($part in $cleanParts) {
        [System.Uri]::EscapeDataString($part)
    }
    return [string]::Join('/', $encodedParts)
}

function Resolve-PathToUrl {
    param (
        [string]$FilePath,
        [System.Collections.Generic.List[PSCustomObject]]$Mappings
    )

    $normFile = Normalize-LocalPath $FilePath
    $bestMatch = $null
    $maxLen = -1

    foreach ($m in $Mappings) {
        $mountPoint = $m.MountPoint
        $prefixWithSlash = $mountPoint + '\'

        if ($normFile -ieq $mountPoint -or $normFile.StartsWith($prefixWithSlash, [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($mountPoint.Length -gt $maxLen) {
                $maxLen = $mountPoint.Length
                $bestMatch = $m
            }
        }
    }

    if (-not $bestMatch) {
        return @{
            WebUrl     = ""
            MountPoint = ""
            Status     = "NOT_IN_ONEDRIVE_MOUNT"
        }
    }

    $relPath = ""
    if ($normFile.Length -gt $bestMatch.MountPoint.Length) {
        $relPath = $normFile.Substring($bestMatch.MountPoint.Length).TrimStart('\')
    }

    $encodedRel = Encode-SharePointUrlPath $relPath
    $baseUrl = $bestMatch.UrlNamespace

    if ($bestMatch.ScopeType -eq "Personal" -and $baseUrl -like "*onedrive.live.com*") {
        $finalUrl = if ($encodedRel) { "$baseUrl&id=$([System.Uri]::EscapeDataString($relPath))" } else { $baseUrl }
    } else {
        $finalUrl = if ($encodedRel) { "$baseUrl/$encodedRel" } else { $baseUrl }
    }

    return @{
        WebUrl     = $finalUrl
        MountPoint = $bestMatch.MountPoint
        Status     = "RESOLVED"
    }
}

# --- Execution ---

$allMappings = Get-OneDriveSyncMappings -ManualMounts $Mount

if ($ListMounts) {
    Write-SafeHost "`nDiscovered $($allMappings.Count) OneDrive / SharePoint sync mapping(s):" -ForegroundColor Cyan
    $i = 1
    foreach ($m in $allMappings) {
        Write-SafeHost " [$i] [$($m.ScopeType)] Local: $($m.MountPoint)" -ForegroundColor Green
        Write-SafeHost "      Remote: $($m.UrlNamespace)" -ForegroundColor Gray
        $i++
    }
    exit 0
}

if ($allMappings.Count -eq 0) {
    Write-SafeWarning "No active OneDrive / SharePoint sync mappings detected in registry."
    Write-SafeWarning "Use -Mount 'C:\Users\<USER>\OneDrive - Company=https://company.sharepoint.com/sites/library' to supply manual mappings."
}

# Collect target files
$targetFiles = [System.Collections.Generic.List[string]]::new()

if ($PSCmdlet.ParameterSetName -eq "Folder") {
    $normFolder = Normalize-LocalPath $Path
    if (-not (Test-Path $normFolder -PathType Container)) {
        throw "Directory not found: $(Redact-LocalPath $normFolder)"
    }

    Write-SafeHost "Scanning folder: $normFolder (Recursive: $Recursive)..." -ForegroundColor Cyan
    $gciParams = @{
        Path        = $normFolder
        File        = $true
        Force       = $false
        ErrorAction = "SilentlyContinue"
    }
    if ($Recursive) {
        $gciParams.Recurse = $true
    }

    $files = Get-ChildItem @gciParams | Where-Object {
        $_.Name -notin @("desktop.ini", "thumbs.db", ".DS_Store")
    }
    foreach ($f in $files) {
        $targetFiles.Add($f.FullName)
    }
} elseif ($PSCmdlet.ParameterSetName -eq "CsvInput") {
    $normCsv = Normalize-LocalPath $Csv
    if (-not (Test-Path $normCsv -PathType Leaf)) {
        throw "CSV file not found: $(Redact-LocalPath $normCsv)"
    }

    Write-SafeHost "Reading file paths from CSV: $normCsv..." -ForegroundColor Cyan

    # Inspect first lines to detect delimiter and presence of headers safely
    $rawLines = Get-Content -Path $normCsv -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($rawLines.Count -eq 0) {
        Write-SafeWarning "CSV file is empty."
        exit 0
    }

    $firstLine = $rawLines[0].Trim(" `t`r`n`uFEFF`u200B`u200E`u200F")
    
    # Check if first line is comma, semicolon, tab, or pipe separated
    $delimiter = ","
    if ($firstLine -match ";") { $delimiter = ";" }
    elseif ($firstLine -match "`t") { $delimiter = "`t" }
    elseif ($firstLine -match "\|") { $delimiter = "|" }

    $csvData = Import-Csv -Path $normCsv -Delimiter $delimiter -Encoding UTF8

    if ($csvData.Count -gt 0) {
        $props = $csvData[0].PSObject.Properties.Name
        $targetCol = $Column

        if ([string]::IsNullOrWhiteSpace($targetCol)) {
            # Heuristic: check if any property matches path/file
            foreach ($p in $props) {
                if ($p -match "(?i)path|file|location|filename|filepath|url") {
                    $targetCol = $p
                    break
                }
            }
            if ([string]::IsNullOrWhiteSpace($targetCol)) {
                $targetCol = $props | Select-Object -First 1
            }
        }

        # Check if first row was actually data rather than header
        $firstVal = $props | Select-Object -First 1
        if ($firstVal -match '^[A-Za-z]:[\\/]' -or $firstVal.StartsWith("\\") -or $firstVal.StartsWith("/")) {
            $targetFiles.Add($firstVal.Trim())
        }

        foreach ($row in $csvData) {
            $val = $row.$targetCol
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $cleanVal = $val.Trim(" `t`r`n`uFEFF`u200B`u200E`u200F")
                if (($cleanVal.StartsWith('"') -and $cleanVal.EndsWith('"')) -or ($cleanVal.StartsWith("'") -and $cleanVal.EndsWith("'"))) {
                    $cleanVal = $cleanVal.Substring(1, $cleanVal.Length - 2).Trim()
                }
                if (-not [string]::IsNullOrWhiteSpace($cleanVal)) {
                    $targetFiles.Add($cleanVal)
                }
            }
        }
    } else {
        # Fallback to plain line reading if Import-Csv failed
        foreach ($line in $rawLines) {
            $cleanLine = $line.Trim(" `t`r`n`uFEFF`u200B`u200E`u200F")
            if (($cleanLine.StartsWith('"') -and $cleanLine.EndsWith('"')) -or ($cleanLine.StartsWith("'") -and $cleanLine.EndsWith("'"))) {
                $cleanLine = $cleanLine.Substring(1, $cleanLine.Length - 2).Trim()
            }
            if (-not [string]::IsNullOrWhiteSpace($cleanLine)) {
                $targetFiles.Add($cleanLine)
            }
        }
    }
}

$outputResolved = Normalize-LocalPath $Output
$outputDir = [System.IO.Path]::GetDirectoryName($outputResolved)
if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$resolvedCount = 0
$unmatchedCount = 0

foreach ($f in $targetFiles) {
    $normF = Normalize-LocalPath $f
    $res = Resolve-PathToUrl -FilePath $normF -Mappings $allMappings
    
    if ($res.Status -eq "RESOLVED") {
        $resolvedCount++
        Write-Verbose "  [OK] $([System.IO.Path]::GetFileName($normF)) -> $($res.WebUrl)"
    } else {
        $unmatchedCount++
        Write-Verbose "  [SKIP] $([System.IO.Path]::GetFileName($normF)) (Not in any mount)"
    }

    $results.Add([PSCustomObject]@{
        local_file_path        = $normF
        file_name              = [System.IO.Path]::GetFileName($normF)
        sharepoint_web_url     = $res.WebUrl
        sharepoint_mount_point = $res.MountPoint
        resolution_status      = $res.Status
    })
}

$results | Export-Csv -Path $outputResolved -NoTypeInformation -Encoding utf8

Write-SafeHost "`n=== Resolution Summary ===" -ForegroundColor Green
Write-SafeHost "Total files checked : $($targetFiles.Count)"
Write-SafeHost "Successfully mapped : $resolvedCount"
Write-SafeHost "Unmatched (outside) : $unmatchedCount"
Write-SafeHost "Output exported to  : $outputResolved`n"
