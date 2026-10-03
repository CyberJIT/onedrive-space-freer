<#
.SYNOPSIS
    Frees up local disk space in a OneDrive folder by marking files for cloud-only storage (Recall On Data Access).

.DESCRIPTION
    Scans files in a parametrized OneDrive folder, and marks them to free up local disk space.
    Removes the Pinned attribute (-P) and adds the Unpinned / Recall On Data Access attributes (+U),
    prompting OneDrive Files On-Demand to synchronize any local changes and hydrate files on demand only.
    
    Supports:
    - Recursive scanning under subfolders.
    - File name filtering (include/exclude substrings or wildcards).
    - File extension filtering (include/exclude).
    - Folder name filtering (include/exclude substrings or wildcards).
    - Long path workarounds using \\?\ extended prefix and Win32 GetShortPathNameW 8.3 short paths.
    - Graceful error recovery: logs any file-level failure and continues processing remaining items.
    - Detailed summary report of actions taken (marked, already dehydrated, skipped, errors, space freed).
    - Dry-run simulation mode.

.PARAMETER Path
    Path to target OneDrive directory.

.PARAMETER Recurse
    Recursively scan subdirectories.

.PARAMETER FileInclude
    Array of substrings or wildcards to include file names.

.PARAMETER FileExclude
    Array of substrings or wildcards to exclude file names.

.PARAMETER ExtInclude
    Array of file extensions to include (e.g. "pbix", ".pdf", "docx").

.PARAMETER ExtExclude
    Array of file extensions to exclude (e.g. "tmp", "log").

.PARAMETER FolderInclude
    Array of substrings or wildcards of folder names to allow during recursion.

.PARAMETER FolderExclude
    Array of substrings or wildcards of folder names to skip during recursion.

.PARAMETER DryRun
    Simulate operations without altering file attributes.

.EXAMPLE
    .\Free-OneDriveSpace.ps1 -Path "C:\Users\Username\OneDrive - Org\Reports"

.EXAMPLE
    .\Free-OneDriveSpace.ps1 -Path "C:\Users\Username\OneDrive" -Recurse -DryRun

.EXAMPLE
    .\Free-OneDriveSpace.ps1 -Path "C:\Users\Username\OneDrive" -Recurse -ExtInclude "pbix", "zip" -FileExclude "*draft*"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [Parameter(Mandatory = $false)]
    [switch]$Recurse,

    [Parameter(Mandatory = $false)]
    [string[]]$FileInclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FileExclude,

    [Parameter(Mandatory = $false)]
    [string[]]$ExtInclude,

    [Parameter(Mandatory = $false)]
    [string[]]$ExtExclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FolderInclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FolderExclude,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Detect operating system
$isWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)

# Cloud Files / OneDrive Windows File Attribute Constants
$FILE_ATTRIBUTE_OFFLINE               = 0x00001000
$FILE_ATTRIBUTE_PINNED                = 0x00080000
$FILE_ATTRIBUTE_UNPINNED              = 0x00100000
$FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x00400000
$INVALID_FILE_ATTRIBUTES              = 0xFFFFFFFF

# Win32 P/Invoke Definition for Long and Short Path Support and Attribute Manipulation
if ($isWindows) {
    $csharpSource = @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class OneDriveWin32Helper {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern uint GetShortPathNameW(string lpszLongPath, [Out] StringBuilder lpszShortPath, uint cchBuffer);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern uint GetFileAttributesW(string lpFileName);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetFileAttributesW(string lpFileName, uint dwFileAttributes);

    public static string GetShortPath(string path) {
        try {
            StringBuilder sb = new StringBuilder(1024);
            uint res = GetShortPathNameW(path, sb, (uint)sb.Capacity);
            if (res > 0 && res < sb.Capacity) {
                return sb.ToString();
            }
        } catch { }
        return null;
    }
}
"@
    try {
        if (-not ([System.Management.Automation.PSTypeName]'OneDriveWin32Helper').Type) {
            Add-Type -TypeDefinition $csharpSource
        }
    } catch {
        # Type might already exist or running in restricted session
    }
}

function Format-ExtendedPath {
    param([string]$FilePath)
    if (-not $isWindows) { return $FilePath }
    if ($FilePath.StartsWith("\\?\")) { return $FilePath }
    if ($FilePath.StartsWith("\\")) {
        return "\\?\UNC\" + $FilePath.Substring(2)
    }
    return "\\?\" + $FilePath
}

function Get-Win32FileAttributesSafe {
    param([string]$FilePath)
    if (-not $isWindows) { return 0 }
    
    $extPath = Format-ExtendedPath $FilePath
    $attrs = [OneDriveWin32Helper]::GetFileAttributesW($extPath)
    if ($attrs -ne $INVALID_FILE_ATTRIBUTES) {
        return $attrs
    }

    $shortPath = [OneDriveWin32Helper]::GetShortPath($FilePath)
    if ($shortPath -and ($shortPath -ne $FilePath)) {
        $attrs = [OneDriveWin32Helper]::GetFileAttributesW($shortPath)
        if ($attrs -ne $INVALID_FILE_ATTRIBUTES) {
            return $attrs
        }
    }

    return $INVALID_FILE_ATTRIBUTES
}

function Set-Win32FileAttributesSafe {
    param(
        [string]$FilePath,
        [uint32]$Attributes
    )
    if (-not $isWindows) { return $true }

    $extPath = Format-ExtendedPath $FilePath
    if ([OneDriveWin32Helper]::SetFileAttributesW($extPath, $Attributes)) {
        return $true
    }

    $shortPath = [OneDriveWin32Helper]::GetShortPath($FilePath)
    if ($shortPath -and ($shortPath -ne $FilePath)) {
        if ([OneDriveWin32Helper]::SetFileAttributesW($shortPath, $Attributes)) {
            return $true
        }
    }

    return $false
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -lt 0) { return "0 B" }
    $units = @("B", "KB", "MB", "GB", "TB")
    $size = [double]$Bytes
    $index = 0
    while ($size -ge 1024 -and $index -lt ($units.Count - 1)) {
        $size /= 1024.0
        $index++
    }
    if ($index -eq 0) {
        return "{0} B" -f [long]$size
    }
    return "{0:N2} {1}" -f $size, $units[$index]
}

function Test-MatchesAnyPattern {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string[]]$Patterns
    )
    if ($null -eq $Patterns) { return $false }
    $patternList = @($Patterns)
    if ($patternList.Count -eq 0) { return $false }

    foreach ($pat in $patternList) {
        if ([string]::IsNullOrWhiteSpace($pat)) { continue }
        if ($pat -match '[\*\?\[\]]') {
            if ($Name -like $pat -or $Name -like "*$pat*") {
                return $true
            }
        } else {
            if ($Name.IndexOf($pat, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                return $true
            }
        }
    }
    return $false
}

function Normalize-ExtensionList {
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$Extensions
    )
    $list = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $Extensions) { return $list }
    foreach ($ext in @($Extensions)) {
        if ([string]::IsNullOrWhiteSpace($ext)) { continue }
        $cleaned = $ext.Trim().ToLowerInvariant()
        if (-not $cleaned.StartsWith(".")) {
            $cleaned = "." + $cleaned
        }
        if (-not $list.Contains($cleaned)) {
            [void]$list.Add($cleaned)
        }
    }
    return $list
}

function Test-FileIncluded {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Parameter(Mandatory = $false)]
        [string[]]$FInclude,

        [Parameter(Mandatory = $false)]
        [string[]]$FExclude,

        [Parameter(Mandatory = $false)]
        $EInclude,

        [Parameter(Mandatory = $false)]
        $EExclude
    )
    $ext = [System.IO.Path]::GetExtension($FileName).ToLowerInvariant()

    if ($null -ne $EExclude -and $EExclude.Count -gt 0 -and $EExclude.Contains($ext)) {
        return $false
    }
    if ($null -ne $EInclude -and $EInclude.Count -gt 0 -and (-not $EInclude.Contains($ext))) {
        return $false
    }
    if ($null -ne $FExclude -and (Test-MatchesAnyPattern -Name $FileName -Patterns $FExclude)) {
        return $false
    }
    if ($null -ne $FInclude -and @($FInclude).Count -gt 0 -and (-not (Test-MatchesAnyPattern -Name $FileName -Patterns $FInclude))) {
        return $false
    }
    return $true
}

function Test-FolderIncluded {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderName,

        [Parameter(Mandatory = $false)]
        [string[]]$FldInclude,

        [Parameter(Mandatory = $false)]
        [string[]]$FldExclude
    )
    if ($null -ne $FldExclude -and (Test-MatchesAnyPattern -Name $FolderName -Patterns $FldExclude)) {
        return $false
    }
    if ($null -ne $FldInclude -and @($FldInclude).Count -gt 0 -and (-not (Test-MatchesAnyPattern -Name $FolderName -Patterns $FldInclude))) {
        return $false
    }
    return $true
}

function Mark-FileFreeSpace {
    param(
        [string]$FilePath,
        [switch]$IsDryRun
    )
    if (-not $isWindows) {
        if ($IsDryRun) { return @{ Success = $true; Status = "dry_run" } }
        return @{ Success = $true; Status = "marked" }
    }

    $attrs = Get-Win32FileAttributesSafe -FilePath $FilePath
    if ($attrs -eq $INVALID_FILE_ATTRIBUTES) {
        $lastErr = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        return @{ Success = $false; Status = "Failed to query attributes (Win32 Error: $lastErr)" }
    }

    # If already marked as recall on data access and not pinned
    $isRecall = ($attrs -band $FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS) -ne 0
    $isPinned = ($attrs -band $FILE_ATTRIBUTE_PINNED) -ne 0
    if ($isRecall -and (-not $isPinned)) {
        return @{ Success = $true; Status = "already_freed" }
    }

    if ($IsDryRun) {
        return @{ Success = $true; Status = "marked" }
    }

    # Calculate new attributes
    $newAttrs = ($attrs -band (-bnot $FILE_ATTRIBUTE_PINNED)) -bor $FILE_ATTRIBUTE_UNPINNED -bor $FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS
    $setResult = Set-Win32FileAttributesSafe -FilePath $FilePath -Attributes ([uint32]$newAttrs)

    if ($setResult) {
        return @{ Success = $true; Status = "marked" }
    }

    # Fallback to attrib.exe +U -P
    $targetPath = $FilePath
    if ($targetPath.Length -ge 250) {
        $short = [OneDriveWin32Helper]::GetShortPath($targetPath)
        if ($short) { $targetPath = $short }
    }

    try {
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = "attrib.exe"
        $pinfo.Arguments = "+U -P `"$targetPath`""
        $pinfo.RedirectStandardOutput = $true
        $pinfo.RedirectStandardError = $true
        $pinfo.UseShellExecute = $false
        $pinfo.CreateNoWindow = $true

        $process = [System.Diagnostics.Process]::Start($pinfo)
        $process.WaitForExit(15000)
        $stderr = $process.StandardError.ReadToEnd()
        $stdout = $process.StandardOutput.ReadToEnd()

        if ($process.ExitCode -eq 0) {
            return @{ Success = $true; Status = "marked" }
        } else {
            $errDetail = if ($stderr) { $stderr.Trim() } else { $stdout.Trim() }
            return @{ Success = $false; Status = "attrib +U -P failed: $errDetail" }
        }
    } catch {
        return @{ Success = $false; Status = "Error running attrib.exe fallback: $($_.Exception.Message)" }
    }
}

# Resolve and validate target path
$targetDir = [System.IO.Path]::GetFullPath($Path)
if (-not (Test-Path -LiteralPath $targetDir)) {
    Write-Error "Target path does not exist: $targetDir"
    exit 1
}

$normExtInc = Normalize-ExtensionList -Extensions $ExtInclude
$normExtExc = Normalize-ExtensionList -Extensions $ExtExclude

# Discover candidate files
$candidateFiles = [System.Collections.Generic.List[string]]::new()

if ((Get-Item -LiteralPath $targetDir) -is [System.IO.FileInfo]) {
    $item = Get-Item -LiteralPath $targetDir
    if (Test-FileIncluded -FileName $item.Name -FInclude $FileInclude -FExclude $FileExclude -EInclude $normExtInc -EExclude $normExtExc) {
        $candidateFiles.Add($item.FullName)
    }
} elseif (-not $Recurse) {
    try {
        $entries = [System.IO.Directory]::GetFiles($targetDir)
        foreach ($file in $entries) {
            $fileName = [System.IO.Path]::GetFileName($file)
            if (Test-FileIncluded -FileName $fileName -FInclude $FileInclude -FExclude $FileExclude -EInclude $normExtInc -EExclude $normExtExc) {
                $candidateFiles.Add($file)
            }
        }
    } catch {
        Write-Warning "Failed reading directory: $($_.Exception.Message)"
    }
} else {
    # Recursive discovery with directory pruning
    $dirQueue = [System.Collections.Generic.Queue[string]]::new()
    $dirQueue.Enqueue($targetDir)

    while ($dirQueue.Count -gt 0) {
        $currentDir = $dirQueue.Dequeue()
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($currentDir)
            foreach ($sd in $subDirs) {
                $dirName = [System.IO.Path]::GetFileName($sd)
                if (Test-FolderIncluded -FolderName $dirName -FldInclude $FolderInclude -FldExclude $FolderExclude) {
                    $dirQueue.Enqueue($sd)
                }
            }
        } catch {
            Write-Warning "Could not list subdirectories of '$currentDir': $($_.Exception.Message)"
        }

        try {
            $filesInDir = [System.IO.Directory]::GetFiles($currentDir)
            foreach ($file in $filesInDir) {
                $fileName = [System.IO.Path]::GetFileName($file)
                if (Test-FileIncluded -FileName $fileName -FInclude $FileInclude -FExclude $FileExclude -EInclude $normExtInc -EExclude $normExtExc) {
                    $candidateFiles.Add($file)
                }
            }
        } catch {
            Write-Warning "Could not list files in '$currentDir': $($_.Exception.Message)"
        }
    }
}

$prefixTag = if ($DryRun) { "[DRY-RUN] " } else { "" }
Write-Host "$prefixTag" -NoNewline
Write-Host "Found $($candidateFiles.Count) matching file(s) in '$targetDir'." -ForegroundColor Cyan

$markedCount = 0
$alreadyFreedCount = 0
$errorCount = 0
$bytesFreedCandidate = 0L
$errorDetails = [System.Collections.Generic.List[PSObject]]::new()

if ($candidateFiles.Count -gt 0) {
    Write-Host ("=" * 70)
    for ($i = 0; $i -lt $candidateFiles.Count; $i++) {
        $file = $candidateFiles[$i]
        $relPath = $file.Substring($targetDir.Length).TrimStart("\", "/")
        if ([string]::IsNullOrEmpty($relPath)) { $relPath = [System.IO.Path]::GetFileName($file) }

        $fileSize = 0L
        try {
            $fileInfo = [System.IO.FileInfo]::new($file)
            $fileSize = $fileInfo.Length
        } catch { }

        $sizeStr = Format-FileSize -Bytes $fileSize
        $idxStr = "[$($i + 1)/$($candidateFiles.Count)]"

        try {
            $res = Mark-FileFreeSpace -FilePath $file -IsDryRun:$DryRun
            if ($res.Success) {
                if ($res.Status -eq "already_freed") {
                    $alreadyFreedCount++
                    Write-Host "$idxStr [ALREADY ONLINE-ONLY] $relPath ($sizeStr)" -ForegroundColor DarkGray
                } else {
                    $markedCount++
                    $bytesFreedCandidate += $fileSize
                    $actionLabel = if ($DryRun) { "WOULD MARK TO FREE SPACE" } else { "MARKED TO FREE SPACE" }
                    Write-Host "$idxStr [$actionLabel] $relPath ($sizeStr)" -ForegroundColor Green
                }
            } else {
                $errorCount++
                $errorDetails.Add([PSCustomObject]@{ Path = $file; Error = $res.Status })
                Write-Host "$idxStr [FAILED] $relPath -> $($res.Status)" -ForegroundColor Red
            }
        } catch {
            $errorCount++
            $errMsg = $_.Exception.Message
            $errorDetails.Add([PSCustomObject]@{ Path = $file; Error = $errMsg })
            Write-Host "$idxStr [ERROR] $relPath -> $errMsg" -ForegroundColor Red
        }
    }
}

# Final summary report
Write-Host ("=" * 70)
$modeTitle = if ($DryRun) { " (DRY-RUN - NO CHANGES APPLIED)" } else { "" }
Write-Host "ONEDRIVE SPACE FREER SUMMARY$modeTitle" -ForegroundColor Yellow
Write-Host ("=" * 70)
Write-Host "Target Directory     : $targetDir"
Write-Host "Total Discovered     : $($candidateFiles.Count)"
Write-Host "Marked to Free Space : $markedCount" -ForegroundColor $(if ($markedCount -gt 0) { "Green" } else { "Gray" })
Write-Host "Already Online-Only  : $alreadyFreedCount"
Write-Host "Failed / Errors      : $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Gray" })
Write-Host "Local Space Impact   : $(Format-FileSize -Bytes $bytesFreedCandidate)" -ForegroundColor Cyan

if ($errorDetails.Count -gt 0) {
    Write-Host "`nErrors Encountered:" -ForegroundColor Red
    foreach ($err in $errorDetails) {
        Write-Host "  - $($err.Path): $($err.Error)" -ForegroundColor DarkRed
    }
}
Write-Host ("=" * 70)

if ($errorCount -gt 0) {
    exit 1
}
exit 0
