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
# Windows only. Uses built-in System.Drawing (GDI+). No extra software required.
# Author : Gregory LeJeune
# Repo   : https://github.com/gregorylejeune/bulk-shrink-jpg

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$MaxWidth    = 800
$MaxHeight   = 600
$JpegQuality = 85

function Write-Banner {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  BULK SHRINK JPG" -ForegroundColor Cyan
    Write-Host "  Windows Send-to-Mail Medium  |  fit inside 800 x 600" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  Originals are never touched. Output goes to a folder you pick." -ForegroundColor DarkGray
    Write-Host ""
}

function Write-Info    { param([string]$Message) Write-Host $Message -ForegroundColor Gray }
function Write-Ok      { param([string]$Message) Write-Host $Message -ForegroundColor Green }
function Write-WarnLine{ param([string]$Message) Write-Host $Message -ForegroundColor Yellow }
function Write-ErrLine { param([string]$Message) Write-Host $Message -ForegroundColor Red }

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
        if ([System.IO.Path]::IsPathRooted($Path)) {
            return [System.IO.Path]::GetFullPath($Path)
        }
        return [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location).Path -ChildPath $Path))
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
    $prefix = $src + [System.IO.Path]::DirectorySeparatorChar
    if ($dst.StartsWith($prefix)) { return 'nested' }
    return 'ok'
}

function Get-JpegCodec {
    $codecs = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()
    foreach ($codec in $codecs) {
        if ($codec.MimeType -eq 'image/jpeg') { return $codec }
    }
    throw "This Windows install has no JPEG encoder (System.Drawing)."
}

function Apply-ExifOrientation {
    param([Parameter(Mandatory)][System.Drawing.Image]$Image)
    $orientationId = 0x0112
    try {
        if ($Image.PropertyIdList -notcontains $orientationId) { return }
        $value = $Image.GetPropertyItem($orientationId).Value[0]
        switch ($value) {
            2 { $Image.RotateFlip([System.Drawing.RotateFlipType]::RotateNoneFlipX) }
            3 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipNone) }
            4 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipX) }
            5 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipX) }
            6 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone) }
            7 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipX) }
            8 { $Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone) }
        }
        try { [void]$Image.RemovePropertyItem($orientationId) } catch { }
    }
    catch {
        # EXIF is optional. Keep going with the pixels as stored.
    }
}

function Save-JpegImage {
    param(
        [Parameter(Mandatory)][System.Drawing.Image]$Image,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][long]$Quality
    )
    $codec = Get-JpegCodec
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    try {
        $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
            [System.Drawing.Imaging.Encoder]::Quality,
            $Quality
        )
        $Image.Save($Path, $codec, $encoderParams)
    }
    finally {
        $encoderParams.Dispose()
    }
}

function Convert-Photo {
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $src = $null
    $destBmp = $null
    $graphics = $null
    $fileStream = $null

    try {
        $fileStream = [System.IO.File]::Open(
            $InputPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::Read
        )
        $src = [System.Drawing.Image]::FromStream($fileStream)
        Apply-ExifOrientation -Image $src

        $srcW = [double]$src.Width
        $srcH = [double]$src.Height
        if ($srcW -lt 1 -or $srcH -lt 1) {
            throw "Image has invalid dimensions ($($src.Width) x $($src.Height))."
        }

        $scale = [Math]::Min(1.0, [Math]::Min(($MaxWidth / $srcW), ($MaxHeight / $srcH)))
        $newW  = [Math]::Max(1, [int][Math]::Round($srcW * $scale))
        $newH  = [Math]::Max(1, [int][Math]::Round($srcH * $scale))

        $outDir = Split-Path -Parent $OutputPath
        if (-not (Test-Path -LiteralPath $outDir -PathType Container)) {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        }

        if ($scale -ge 1.0) {
            $copyW = $src.Width
            $copyH = $src.Height
            # Release the read lock before copying the original bytes.
            $src.Dispose(); $src = $null
            $fileStream.Dispose(); $fileStream = $null
            [System.IO.File]::Copy($InputPath, $OutputPath, $true)
            return [pscustomobject]@{
                Width     = $copyW
                Height    = $copyH
                NewWidth  = $copyW
                NewHeight = $copyH
                Action    = 'copied'
            }
        }

        $destBmp = New-Object System.Drawing.Bitmap $newW, $newH
        $destBmp.SetResolution($src.HorizontalResolution, $src.VerticalResolution)
        $graphics = [System.Drawing.Graphics]::FromImage($destBmp)
        $graphics.CompositingMode    = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        $destRect = New-Object System.Drawing.Rectangle 0, 0, $newW, $newH
        $graphics.DrawImage($src, $destRect, 0, 0, $src.Width, $src.Height, [System.Drawing.GraphicsUnit]::Pixel)

        Save-JpegImage -Image $destBmp -Path $OutputPath -Quality ([long]$JpegQuality)

        return [pscustomobject]@{
            Width     = $src.Width
            Height    = $src.Height
            NewWidth  = $newW
            NewHeight = $newH
            Action    = 'resized'
        }
    }
    finally {
        if ($null -ne $graphics)   { $graphics.Dispose() }
        if ($null -ne $destBmp)    { $destBmp.Dispose() }
        if ($null -ne $src)        { $src.Dispose() }
        if ($null -ne $fileStream) { $fileStream.Dispose() }
    }
}

# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
$exitCode = 0

try {
    Write-Banner

    $onWindows = [System.Environment]::OSVersion.Platform -eq 'Win32NT'
    if (-not $onWindows) {
        throw "This script uses Windows GDI+ and must be run on Windows PowerShell or PowerShell 7 for Windows."
    }

    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    }
    catch {
        throw "Could not load System.Drawing. $($_.Exception.Message)"
    }

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
    $failures = New-Object System.Collections.Generic.List[string]
    $index = 0
    $sourcePrefix = $sourceDir.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar

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
            $failures.Add("$relative  --  $reason") | Out-Null
            Write-ErrLine ("[{0}/{1}] FAIL  {2}  {3}" -f $index, $total, $relative, $reason)
        }
    }

    Write-Progress -Activity "Shrinking photos to 800 x 600" -Completed

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  DONE" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ("  Succeeded : {0}" -f $ok)      -ForegroundColor Green
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

if ([Environment]::UserInteractive) {
    Write-Host "Press Enter to exit" -ForegroundColor DarkGray
    try { [void](Read-Host) } catch { }
}

exit $exitCode
