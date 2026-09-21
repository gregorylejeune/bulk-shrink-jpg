# Bulk Shrink JPG

**Windows "Send to Mail" shrinking — for a whole folder of photos.**  
One PowerShell script. Two folder prompts. Originals never touched.

![PowerShell](https://img.shields.io/badge/Windows-PowerShell_5.1%2B-5391FE?style=flat-square)
![Target](https://img.shields.io/badge/Target-800%20x%20600-22c55e?style=flat-square)
![Format](https://img.shields.io/badge/Format-.jpg%20%2F%20.jpeg-f59e0b?style=flat-square)
![Dependencies](https://img.shields.io/badge/Dependencies-None-94a3b8?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-38bdf8?style=flat-square)


---

Right-click one photo, **Send to Mail Recipient**, pick **Medium (800 x 600)** — that is the Windows trick for making a picture small enough to email. It is perfect for one file. It is miserable for nine hundred.

`Shrink-Photos.ps1` is that same shrink, in bulk.

| Windows Send to Mail (one file) | This script (~900 files) |
| --- | --- |
| Right-click, Send to Mail Recipient | Run one `.ps1`, answer a few prompts |
| Medium size = fit inside **800 x 600** | Same target, same aspect-ratio rule |
| Writes a temporary smaller copy | Writes a full output folder you choose |
| Easy to miss files | Loops every `*.jpg` / `*.jpeg` |

A typical 1800 x 1200 photo (3:2) becomes **800 x 533**. Windows does not stretch 3:2 into 4:3, and neither does this script.

---

## What you get

- **One file.** `Shrink-Photos.ps1` is the whole tool.
- **Prompts, not parameters.** Source folder, output folder, subfolders, overwrite, then a final "go".
- **Existence checks first.** Missing source? It asks again. Missing output? It offers to create it.
- **Try / catch on every photo.** One corrupt JPG does not kill a 900-file run.
- **Originals are sacred.** Source and output must be different folders. Output cannot sit inside source.
- **Resumable.** Already-shrunk files in the output folder are skipped unless you say overwrite.
- **No extra software.** Uses Windows Image Acquisition (WIA) COM -- `Wia.ImageFile` / `Wia.ImageProcess`. That DLL already lives on Windows (`wiaaut.dll`). Nothing to install, no NuGet, no System.Drawing.

The shrink itself is Microsoft's WIA **Scale** filter: `MaximumWidth = 800`, `MaximumHeight = 600`, `PreserveAspectRatio = yes`. Same bounding box Windows uses for Send to Mail Medium.


---

## Quick start

1. Download [`Shrink-Photos.ps1`](./Shrink-Photos.ps1) (or clone this repo).
2. Right-click the script, then **Run with PowerShell**.

   If Windows blocks it, open PowerShell in that folder and run:

   ```powershell
   Set-Location $env:USERPROFILE\Downloads\bulk-shrink-jpg
   powershell -ExecutionPolicy Bypass -File .\Shrink-Photos.ps1
   ```

3. Answer the prompts.

```text
============================================================
  BULK SHRINK JPG
  Windows Send-to-Mail Medium  |  fit inside 800 x 600
============================================================

Source folder (where the original JPGs are): C:\Photos\Trip
Output folder (where shrunk JPGs will be written): C:\Photos\Trip-Email
Include subfolders? [Y/n]: Y
Overwrite files that already exist in the output folder? [y/N]: N

  Source : C:\Photos\Trip
  Output : C:\Photos\Trip-Email
  Files  : 912 JPGs
  Size   : fit inside 800 x 600  (aspect ratio kept)
  JPEG   : quality 85

Start shrinking now? [Y/n]: Y

[1/912] OK    DSC_0001.jpg  1800x1200 -> 800x533  2.41 MB -> 186 KB
[2/912] OK    DSC_0002.jpg  1800x1200 -> 800x533  2.38 MB -> 179 KB
...
============================================================
  DONE
============================================================
  Succeeded : 912
  Input     : 2.12 GB
  Output    : 168 MB
  Saved     : 1.95 GB  (92.1%)
```

Paste paths with or without quotes — Explorer "Copy path" works.

---

## How the size works

Windows Mail **Medium** is a bounding box, not a crop and not a stretch.

```text
                  800 px
          +------------------+
          |                  |
    600   |    photo fits    |
     px   |    inside this   |
          |                  |
          +------------------+
```

| Original | Result | Why |
| --- | --- | --- |
| 1800 x 1200 (3:2) | **800 x 533** | Width hits 800 first |
| 1200 x 1800 (portrait) | **400 x 600** | Height hits 600 first |
| 1920 x 1080 (16:9) | **800 x 450** | Width hits 800 first |
| 640 x 480 (already small) | **copied as-is** | Never upscaled, never re-encoded |

JPEG quality is **85** — close to the Windows email wizard, sharp enough to view, small enough to send.

Phone photos that are stored sideways are rotated from EXIF orientation before shrinking, so they do not come out on their side.

Folder structure is preserved:

```text
C:\Photos\Trip\                ->  C:\Photos\Trip-Email\
  beach\IMG_01.jpg             ->    beach\IMG_01.jpg
  dinner\IMG_02.jpg            ->    dinner\IMG_02.jpg
```

---

## Prompts and safety

Every path is checked before a single file is written.

| Prompt | What happens |
| --- | --- |
| **Source folder** | Must already exist and must be a folder. Wrong path? The script says so and asks again. |
| **Output folder** | If it does not exist, you are asked whether to create it. |
| **Include subfolders?** | Default **Yes**. Nested trip folders stay nested in the output. |
| **Overwrite existing?** | Default **No** — so you can re-run a failed batch and only the missing files get processed. |
| **Start shrinking now?** | Last chance to bail. Nothing is written until you confirm. |

Hard stops (the script will not continue):

- Source folder missing or not a directory
- Source and output are the same path
- Output is inside the source folder (would mix originals with copies)
- Not running on Windows / WIA COM (`Wia.ImageFile`) cannot load

Soft failures (the batch keeps going):

- Unreadable / corrupt JPG
- File locked by another program
- Individual write errors

Each failure is printed in red and listed again in the summary.

---

## Requirements

| Item | Detail |
| --- | --- |
| OS | Windows 10 / 11 (or Windows Server) |
| Runtime | Windows PowerShell 5.1, or PowerShell 7+ for Windows |
| Software | None. WIA COM ships with Windows. No ImageMagick, no .NET SDK, no Photoshop, no extra modules. |
| Files | `*.jpg` and `*.jpeg` (any case -- `.JPG` counts) |

What the script actually calls -- all Microsoft, all already on the box:

| Piece | Source |
| --- | --- |
| `Get-ChildItem`, `Copy-Item`, `Test-Path`, `New-Item`, `Read-Host` | Inbox PowerShell modules (`Microsoft.PowerShell.Management`, `Microsoft.PowerShell.Utility`) |
| `Wia.ImageFile` / `Wia.ImageProcess` | Windows Image Acquisition Automation (`wiaaut.dll`) |
| Scale / Convert / RotateFlip filters | Same WIA COM library |

Nothing is downloaded. Nothing is registered. If PowerShell opens, this runs.

macOS and Linux will refuse to run -- WIA is a Windows component.


---

## Troubleshooting

**"Running scripts is disabled on this system."**

```powershell
powershell -ExecutionPolicy Bypass -File .\Shrink-Photos.ps1
```

That bypass applies only to this one launch. It does not change the machine policy.

**The window flashes and closes.**

Start it from a PowerShell window instead of double-clicking, or just press Enter at the "Press Enter to exit" pause — the summary stays on screen.

**No files found.**

- Confirm the folder actually has `.jpg` / `.jpeg` files (`.png` and `.heic` are ignored on purpose).
- Answer **Y** to "Include subfolders?" if the photos live one level down.

**A few files failed, the rest worked.**

Re-run against the same output folder with overwrite **N**. Successful files are skipped; only the failures are retried. Open those files in Photos first — they are usually already corrupt.

**I wanted exact 800 x 600, even if it stretches.**

This script matches Windows: it *fits inside* 800 x 600. Stretching 1800 x 1200 into 800 x 600 would squash every 3:2 photo. If you truly need a padded or cropped 800 x 600, that is a different operation.

---

## Repo layout

```text
bulk-shrink-jpg/
|-- Shrink-Photos.ps1   -- the tool (the only script)
|-- README.md           -- this file
```

That is the entire project.

---

## License

MIT. Use it, fork it, send nine hundred vacation photos without opening Mail once.
