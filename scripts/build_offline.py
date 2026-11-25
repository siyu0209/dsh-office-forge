#!/usr/bin/env python3
"""Build a dependency-honest offline DSH plugin zip from an npm tarball."""
from __future__ import annotations

import argparse
import hashlib
import zipfile
from pathlib import Path

INSTALL_PS1 = r'''param(
  [string]$DshHome = $env:DSH_HOME,
  [string]$Profile = "desktop",
  [string]$Tarball = ""
)
$ErrorActionPreference = "Stop"
if (-not $DshHome) { throw "Set DSH_HOME or pass -DshHome" }
if (-not $Tarball) { $Tarball = Join-Path $PSScriptRoot "dsh-office-forge.tgz" }
$profileDir = Join-Path $DshHome "profiles\$Profile"
$pnpm = Join-Path $DshHome "dsh-runtimes\dsh-primary-runtime\dependencies\pnpm\bin\pnpm.mjs"
if (-not (Test-Path $pnpm)) { throw "pnpm not found: $pnpm" }
Push-Location $profileDir
try { node $pnpm add $Tarball } finally { Pop-Location }
$manifestPath = Join-Path $profileDir "package.json"
$manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
if ($manifest.dsh.profile.bundles -notcontains "dsh-office-forge") {
  $manifest.dsh.profile.bundles += "dsh-office-forge"
  $manifest | ConvertTo-Json -Depth 100 | Set-Content $manifestPath -Encoding UTF8
  Write-Host "Added dsh-office-forge to dsh.profile.bundles."
}
Write-Host "Installed."
'''

CHECK_PS1 = r'''param([string]$DshHome = $env:DSH_HOME, [string]$Profile = "desktop")
$ErrorActionPreference = "Stop"
$json = Join-Path $DshHome "profiles\$Profile\package.json"
if (-not (Test-Path $json)) { throw "Profile manifest not found: $json" }
$p = Get-Content $json -Raw | ConvertFrom-Json
"dsh-office-forge: " + ($p.dependencies.'dsh-office-forge' ?? "not installed")
"bundle listed: " + ($p.dsh.profile.bundles -contains "dsh-office-forge")
'''

README = """Office Forge offline package

1. Extract this zip.
2. Run install.ps1 with DSH_HOME set, or pass -DshHome.
3. Run check.ps1 to confirm the dependency and bundle entry.
4. Restart DSH, then run officeforge_doctor.

OCR needs PyMuPDF + pytesseract + Tesseract; LibreOffice is optional.
"""


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--tarball", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    tarball = Path(args.tarball).resolve()
    output = Path(args.output).resolve()
    if not tarball.is_file():
        raise SystemExit(f"tarball not found: {tarball}")
    output.parent.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256(tarball.read_bytes()).hexdigest()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(tarball, "dsh-office-forge.tgz")
        archive.writestr("install.ps1", INSTALL_PS1)
        archive.writestr("check.ps1", CHECK_PS1)
        archive.writestr("README.txt", README)
        archive.writestr("SHA256SUMS.txt", f"{digest}  dsh-office-forge.tgz\n")
    print(f"OFFLINE_ZIP_OK {output} bytes={output.stat().st_size} sha256={digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
