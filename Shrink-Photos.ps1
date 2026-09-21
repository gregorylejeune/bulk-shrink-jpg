#Requires -Version 5.1
# Bulk-shrink JPG photos to Windows "Send to Mail" Medium size (fit inside 800 x 600).
#
# Interactive PowerShell script. Prompts for a source folder and an output folder,
# then resizes every *.jpg / *.jpeg file the same way Windows does when you
# right-click photos and choose Send to Mail Recipient -> Medium (800 x 600).
#
# Originals are never overwritten. Failed files are skipped so a 900-photo
# batch can finish even if a few images are corrupt.
#
# Windows only. Uses inbox Microsoft pieces exclusively:
#   - PowerShell cmdlets that ship with Windows
#   - WIA COM  (Wia.ImageFile / Wia.ImageProcess)  -- wiaaut.dll
# No NuGet, no ImageMagick, no System.Drawing, nothing to install.
#
# Author : Gregory LeJeune
# Repo   : https://github.com/gregorylejeune/bulk-shrink-jpg

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$MaxWidth    = 800
$MaxHeight   = 600
$JpegQuality = 85

# WIA JPEG format identifier (Microsoft wiaimgfmt.h / wiaaut.dll).
$WiaFormatJpeg = '{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}'

function Write-Banner {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  BULK SHRINK JPG" -ForegroundColor Cyan
    Write-Host "  Windows Send-to-Mail Medium  |  fit inside 800 x 600" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  Originals are never touched. Output goes to a folder you pick." -ForegroundColor DarkGray
    Write-Host "  Engine: Windows Image Acquisition (WIA) -- built into Windows." -ForegroundColor DarkGray
    Write-Host ""
}

function Write-Info     { param([string]$Message) Write-Host $Message -ForegroundColor Gray }
function Write-Ok       { param([string]$Message) Write-Host $Message -ForegroundColor Green }
function Write-WarnLine { param([string]$Message) Write-Host $Message -ForegroundColor Yellow }
function Write-ErrLine  { param([string]$Message) Write-Host $Message -ForegroundColor Red }

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return ('{0:N0} B' -f $Bytes)
}

function Get-YesNo {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [bool]$Default = $true
    )
    $hint = if ($Default) { 'Y/n' } else { 'y/N' }
    while ($true) {
        $raw = Read-Host "$Prompt [$hint]"
        if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
        switch -Regex ($raw.Trim()) {
            '^[Yy](es)?$' { return $true }
            '^[Nn](o)?$'  { return $false }
            default {
                Write-WarnLine "  Please answer Y or N."
            }
        }
    }
}

function Read-TrimmedPath {
    param([Parameter(Mandatory)][string]$Prompt)
    while ($true) {
        $raw = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($raw)) {
            Write-WarnLine "  A path is required. Try again, or press Ctrl+C to cancel."
            continue
        }
        $cleaned = $raw.Trim().Trim('"').Trim("'")
        if ([string]::IsNullOrWhiteSpace($cleaned)) {
            Write-WarnLine "  A path is required. Try again, or press Ctrl+C to cancel."
            continue
        }
        return $cleaned
    }
}

function Get-FullPathSafe {
    param([Parameter(Mandatory)][string]$Path)
    try {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    }
    catch {
        throw "Could not resolve path '$Path'. $($_.Exception.Message)"
    }
}

function Read-ExistingDirectory {
    param([Parameter(Mandatory)][string]$Prompt)
    while ($true) {
        $entered = Read-TrimmedPath -Prompt $Prompt
        try {
            $full = Get-FullPathSafe -Path $entered
        }
        catch {
            Write-ErrLine "  $($_.Exception.Message)"
            continue
        }

        if (-not (Test-Path -LiteralPath $full)) {
            Write-ErrLine "  Folder does not exist: $full"
            Write-Info    "  Check the path and try again."
            continue
        }
        if (-not (Test-Path -LiteralPath $full -PathType Container)) {
            Write-ErrLine "  That path exists but is not a folder: $full"
            continue
        }
        return $full
    }
}

function Read-OutputDirectory {
    param([Parameter(Mandatory)][string]$Prompt)
    while ($true) {
        $entered = Read-TrimmedPath -Prompt $Prompt
        try {
            $full = Get-FullPathSafe -Path $entered
        }
        catch {
            Write-ErrLine "  $($_.Exception.Message)"
            continue
        }

        if (Test-Path -LiteralPath $full -PathType Container) {
            return $full
        }
        if (Test-Path -LiteralPath $full) {
            Write-ErrLine "  That path exists but is not a folder: $full"
            continue
        }

        Write-WarnLine "  Output folder does not exist yet:"
        Write-Info     "  $full"
        if (Get-YesNo -Prompt "  Create this folder?" -Default $true) {
            try {
                New-Item -ItemType Directory -Path $full -Force | Out-Null
                Write-Ok "  Created: $full"
                return $full
            }
            catch {
                Write-ErrLine "  Could not create folder. $($_.Exception.Message)"
                continue
            }
        }
        Write-Info "  OK, enter a different output path."
    }
}

function Test-SameOrNestedPath {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination
    )
    $src = $Source.TrimEnd('\', '/').ToLowerInvariant()
    $dst = $Destination.TrimEnd('\', '/').ToLowerInvariant()
    if ($src -eq $dst) { return 'same' }
    $prefix = $src + '\'
    if ($dst.StartsWith($prefix)) { return 'nested' }
    return 'ok'
}

function Clear-ComObject {
    param($ComObject)
    if ($null -eq $ComObject) { return }
    try {
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
    catch {
        # Already released or not a COM RCW.
    }
}

function Test-WiaAvailable {
    $probe = $null
    try {
        $probe = New-Object -ComObject Wia.ImageFile
        return $true
    }
    catch {
        throw "Windows Image Acquisition (WIA) is not available. It ships with Windows as wiaaut.dll (Wia.ImageFile). $($_.Exception.Message)"
    }
    finally {
        Clear-ComObject $probe
    }
}

function Get-WiaOrientation {
    param($Image)
    try {
        $count = [int]$Image.Properties.Count
        for ($i = 1; $i -le $count; $i++) {
            $prop = $null
            try {
                $prop = $Image.Properties.Item($i)
                $id = [int]$prop.PropertyID
                if ($id -eq 274) {
                    return [int]$prop.Value
                }
            }
            catch {
                # Skip unreadable EXIF entries.
            }
        }
    }
    catch {
        # No EXIF bag on this file.
    }
    return 1
}

function Set-WiaFilterProperty {
    param(
        $Filter,
        [Parameter(Mandatory)][string]$Name,
        $Value
    )
    try {
        $Filter.Properties.Item($Name).Value = $Value
    }
    catch {
        throw "WIA filter property '$Name' could not be set. $($_.Exception.Message)"
    }
}

function Add-WiaFilter {
    param(
        $Process,
        [Parameter(Mandatory)][string]$Name
    )
    try {
        $Process.Filters.Add($Process.FilterInfos.Item($Name).FilterID)
    }
    catch {
        throw "WIA filter '$Name' is not available on this Windows install. $($_.Exception.Message)"
    }
    return [int]$Process.Filters.Count
}

function Convert-Photo {
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $image   = $null
    $process = $null
    $result  = $null

    try {
        $image = New-Object -ComObject Wia.ImageFile
        $image.LoadFile($InputPath)

        $srcW = [int]$image.Width
        $srcH = [int]$image.Height
        if ($srcW -lt 1 -or $srcH -lt 1) {
            throw "Image has invalid dimensions ($srcW x $srcH)."
        }

        $orientation = Get-WiaOrientation -Image $image
        $needsRotate = ($orientation -ge 2 -and $orientation -le 8)

        $dispW = $srcW
        $dispH = $srcH
        if ($orientation -eq 5 -or $orientation -eq 6 -or $orientation -eq 7 -or $orientation -eq 8) {
            $dispW = $srcH
            $dispH = $srcW
        }
        $needsScale = ($dispW -gt $MaxWidth) -or ($dispH -gt $MaxHeight)

        $outDir = Split-Path -Parent $OutputPath
        if (-not (Test-Path -LiteralPath $outDir -PathType Container)) {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        }

        if (-not $needsRotate -and -not $needsScale) {
            Clear-ComObject $image
            $image = $null
            Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force
            return [pscustomobject]@{
                Width     = $srcW
                Height    = $srcH
                NewWidth  = $srcW
                NewHeight = $srcH
                Action    = 'copied'
            }
        }

        $process = New-Object -ComObject Wia.ImageProcess

        if ($needsRotate) {
            $idx = Add-WiaFilter -Process $process -Name 'RotateFlip'
            $filter = $process.Filters.Item($idx)
            $angle = 0
            $flipH = 0
            switch ($orientation) {
                2 { $flipH = 1 }
                3 { $angle = 180 }
                4 { $angle = 180; $flipH = 1 }
                5 { $angle = 90;  $flipH = 1 }
                6 { $angle = 90 }
                7 { $angle = 270; $flipH = 1 }
                8 { $angle = 270 }
            }
            Set-WiaFilterProperty -Filter $filter -Name 'RotationAngle'  -Value ([int]$angle)
            Set-WiaFilterProperty -Filter $filter -Name 'FlipHorizontal' -Value ([int]$flipH)
            Set-WiaFilterProperty -Filter $filter -Name 'FlipVertical'   -Value 0
        }

        if ($needsScale) {
            $idx = Add-WiaFilter -Process $process -Name 'Scale'
            $filter = $process.Filters.Item($idx)
            Set-WiaFilterProperty -Filter $filter -Name 'MaximumWidth'  -Value ([int]$MaxWidth)
            Set-WiaFilterProperty -Filter $filter -Name 'MaximumHeight' -Value ([int]$MaxHeight)
            try {
                $filter.Properties.Item('PreserveAspectRatio').Value = 1
            }
            catch {
                # Property exists on every current Windows WIA; ignore if an old build lacks it.
            }
        }

        $idx = Add-WiaFilter -Process $process -Name 'Convert'
        $filter = $process.Filters.Item($idx)
        Set-WiaFilterProperty -Filter $filter -Name 'FormatID' -Value $WiaFormatJpeg
        try {
            $filter.Properties.Item('Quality').Value = [int]$JpegQuality
        }
        catch {
            # Quality only applies to JPEG convert; if missing, WIA uses its default.
        }

        $result = $process.Apply($image)
        if ($null -eq $result) {
            throw "WIA returned no image after Apply."
        }

        if (Test-Path -LiteralPath $OutputPath) {
            Remove-Item -LiteralPath $OutputPath -Force
        }
        $result.SaveFile($OutputPath)

        return [pscustomobject]@{
            Width     = $dispW
            Height    = $dispH
            NewWidth  = [int]$result.Width
            NewHeight = [int]$result.Height
            Action    = 'resized'
        }
    }
    finally {
        Clear-ComObject $result
        Clear-ComObject $process
        Clear-ComObject $image
    }
}

# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
$exitCode = 0

try {
    Write-Banner

    if ($env:OS -ne 'Windows_NT') {
        throw "This script uses Windows Image Acquisition (WIA) and must be run on Windows PowerShell or PowerShell 7 for Windows."
    }

    Test-WiaAvailable | Out-Null

    $sourceDir = Read-ExistingDirectory -Prompt "Source folder (where the original JPGs are)"
    $outputDir = Read-OutputDirectory   -Prompt "Output folder (where shrunk JPGs will be written)"

    $relation = Test-SameOrNestedPath -Source $sourceDir -Destination $outputDir
    if ($relation -eq 'same') {
        throw "Source and output folders are the same. Pick a different output folder so originals are never overwritten."
    }
    if ($relation -eq 'nested') {
        throw "Output folder sits inside the source folder. Pick an output path outside the source so we do not mix originals with shrunk copies."
    }

    $includeSubfolders = Get-YesNo -Prompt "Include subfolders?" -Default $true
    $overwriteExisting = Get-YesNo -Prompt "Overwrite files that already exist in the output folder?" -Default $false

    Write-Host ""
    Write-Info "Scanning for JPG files..."

    try {
        $gciParams = @{
            LiteralPath = $sourceDir
            File        = $true
            ErrorAction = 'Stop'
        }
        if ($includeSubfolders) { $gciParams['Recurse'] = $true }

        $files = @(
            Get-ChildItem @gciParams -Filter '*.jpg'
            Get-ChildItem @gciParams -Filter '*.jpeg'
        ) | Sort-Object -Property FullName -Unique
    }
    catch {
        throw "Could not list files in '$sourceDir'. $($_.Exception.Message)"
    }

    $total = $files.Count
    if ($total -eq 0) {
        Write-WarnLine "No *.jpg or *.jpeg files found in:"
        Write-WarnLine "  $sourceDir"
        if (-not $includeSubfolders) {
            Write-Info "Tip: re-run and answer Y to 'Include subfolders?' if photos sit in nested folders."
        }
        return
    }

    Write-Host ""
    Write-Host "  Source : $sourceDir" -ForegroundColor White
    Write-Host "  Output : $outputDir" -ForegroundColor White
    Write-Host "  Files  : $total JPG$(if ($total -ne 1) { 's' })" -ForegroundColor White
    Write-Host "  Size   : fit inside ${MaxWidth} x ${MaxHeight}  (aspect ratio kept)" -ForegroundColor White
    Write-Host "  JPEG   : quality $JpegQuality" -ForegroundColor White
    Write-Host "  Engine : WIA (Wia.ImageFile) -- inbox Windows COM" -ForegroundColor White
    Write-Host ""

    if (-not (Get-YesNo -Prompt "Start shrinking now?" -Default $true)) {
        Write-WarnLine "Cancelled. No files were written."
        return
    }

    $ok      = 0
    $copied  = 0
    $skipped = 0
    $failed  = 0
    $bytesIn  = [long]0
    $bytesOut = [long]0
    $failures = @()
    $index = 0
    $sourcePrefix = $sourceDir.TrimEnd('\', '/') + '\'

    Write-Host ""
    foreach ($file in $files) {
        $index++
        $name = $file.Name
        $percent = [int](($index / $total) * 100)
        Write-Progress -Activity "Shrinking photos to 800 x 600" -Status "[$index / $total] $name" -PercentComplete $percent

        $relative = if ($file.FullName.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $file.FullName.Substring($sourcePrefix.Length)
        } else {
            $file.Name
        }
        $destPath = Join-Path -Path $outputDir -ChildPath $relative

        try {
            if ((Test-Path -LiteralPath $destPath) -and -not $overwriteExisting) {
                $skipped++
                Write-WarnLine ("[{0}/{1}] SKIP  {2}  (already exists)" -f $index, $total, $relative)
                continue
            }

            $result = Convert-Photo -InputPath $file.FullName -OutputPath $destPath

            try {
                $outItem = Get-Item -LiteralPath $destPath
                $outItem.CreationTime  = $file.CreationTime
                $outItem.LastWriteTime = $file.LastWriteTime
            }
            catch {
                # Timestamp copy is best-effort.
            }

            $inSize  = $file.Length
            $outSize = (Get-Item -LiteralPath $destPath).Length
            $bytesIn  += $inSize
            $bytesOut += $outSize
            $ok++
            if ($result.Action -eq 'copied') { $copied++ }

            $verb = if ($result.Action -eq 'copied') { 'COPY' } else { 'OK  ' }
            $line = "[{0}/{1}] {2}  {3}  {4}x{5} -> {6}x{7}  {8} -> {9}" -f `
                $index, $total, $verb, $relative, `
                $result.Width, $result.Height, $result.NewWidth, $result.NewHeight, `
                (Format-Bytes $inSize), (Format-Bytes $outSize)
            Write-Ok $line
        }
        catch {
            $failed++
            $reason = $_.Exception.Message
            $failures += "$relative  --  $reason"
            Write-ErrLine ("[{0}/{1}] FAIL  {2}  {3}" -f $index, $total, $relative, $reason)
        }
    }

    Write-Progress -Activity "Shrinking photos to 800 x 600" -Completed

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  DONE" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ("  Succeeded : {0}" -f $ok) -ForegroundColor Green
    if ($copied  -gt 0) { Write-Host ("  Copied    : {0}  (already within 800 x 600)" -f $copied) -ForegroundColor Green }
    if ($skipped -gt 0) { Write-Host ("  Skipped   : {0}  (already in output)" -f $skipped) -ForegroundColor Yellow }
    if ($failed  -gt 0) { Write-Host ("  Failed    : {0}" -f $failed) -ForegroundColor Red }
    Write-Host ("  Input     : {0}" -f (Format-Bytes $bytesIn))
    Write-Host ("  Output    : {0}" -f (Format-Bytes $bytesOut))
    if ($bytesIn -gt 0 -and $bytesOut -gt 0 -and $bytesOut -lt $bytesIn) {
        $savedPct = [Math]::Round((1 - ($bytesOut / $bytesIn)) * 100, 1)
        Write-Host ("  Saved     : {0}  ({1}%)" -f (Format-Bytes ($bytesIn - $bytesOut)), $savedPct) -ForegroundColor Green
    }
    Write-Host "  Folder    : $outputDir"
    Write-Host ""

    if ($failures.Count -gt 0) {
        Write-ErrLine "Failures:"
        foreach ($item in $failures) {
            Write-ErrLine "  - $item"
        }
        Write-Host ""
        $exitCode = 1
    }
}
catch {
    Write-Host ""
    Write-ErrLine "Stopped: $($_.Exception.Message)"
    Write-Host ""
    $exitCode = 1
}
finally {
    Write-Progress -Activity "Shrinking photos to 800 x 600" -Completed -ErrorAction SilentlyContinue
}

if ($Host.Name -eq 'ConsoleHost') {
    Write-Host "Press Enter to exit" -ForegroundColor DarkGray
    try { [void](Read-Host) } catch { }
}

exit $exitCode
