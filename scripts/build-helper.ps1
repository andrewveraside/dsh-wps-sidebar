param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# Compile the office-helper from a WORKING COPY of helper/OfficeHelper.cs.
#
# The plugin picks its helper by the sha256 of the source file, so the exe is
# written into the cache folder named after that hash:
#     %LOCALAPPDATA%\dsh-office\helper-<hash12>\office-helper.exe
#
# Run this after ANY change to OfficeHelper.cs, then run `npm run smoke`.
# (Editing the installed copy under node_modules is a trap: the hash follows the
#  WORKING COPY, so the old exe keeps being used and the change looks ignored.)
#
# Usage:
#     powershell -ExecutionPolicy Bypass -File scripts/build-helper.ps1
#     powershell -ExecutionPolicy Bypass -File scripts/build-helper.ps1 -RepoRoot E:\other\checkout
# ─────────────────────────────────────────────────────────────────────────────

$src = Join-Path $RepoRoot 'helper\OfficeHelper.cs'
if (-not (Test-Path $src)) {
    # `Join-Path x 'sub\name'` stops at the first backslash, so a missing subfolder
    # shows up as a bare file name. Fall back to an explicitly built path.
    $src = [System.IO.Path]::Combine($RepoRoot, 'helper', 'OfficeHelper.cs')
}
if (-not (Test-Path $src)) { throw "source not found: $src (RepoRoot=$RepoRoot)" }

# The in-box .NET Framework compiler: no SDK, no NuGet, no admin rights needed.
$fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$csc = Join-Path $fw 'csc.exe'
if (-not (Test-Path $csc)) { throw "csc.exe not found at $csc (.NET Framework 4 is required)" }

$hash = (Get-FileHash $src -Algorithm SHA256).Hash.Substring(0, 12).ToLower()
$dir = Join-Path $env:LOCALAPPDATA "dsh-office\helper-$hash"
New-Item -ItemType Directory -Force -Path $dir | Out-Null

# A running helper holds its exe open; csc then fails with CS0016.
Get-Process office-helper -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 800

$exe = Join-Path $dir 'office-helper.exe'
$tmp = Join-Path $dir ("build-" + $PID + ".exe")

# NOTE for the next maintainer: `dynamic` needs the Microsoft.CSharp COMPILER
# assembly plus the Microsoft.CSharp.dll FRAMEWORK assembly supplied by
# System.Core.dll's type forwarder. If a build ever fails with CS0656
# ("Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo.Create"), that pair is
# what went missing -- check Microsoft.CSharp.dll also exists in
# $fw, and pass its path with -r:.
# NOTE: csc is run with -noconfig. Without it, csc silently adds csc.rsp (which
# references Microsoft.CSharp.dll) and fails with CS1567 "error creating Win32
# resource ... Access is denied", because the Windows compile token cannot write
# the resource .TMP next to the -out target. The references we need are passed
# explicitly below, so the default response file buys nothing.
$refs = @(
    '-r:mscorlib.dll', '-r:System.dll', '-r:System.Core.dll', '-r:Microsoft.CSharp.dll',
    '-r:System.Drawing.dll', '-r:System.Windows.Forms.dll', '-r:System.Web.Extensions.dll',
    '-r:System.IO.Compression.dll', '-r:System.IO.Compression.FileSystem.dll'
)
# -nostdlib: only the references listed above; keeps csc from auto-adding .NET 4.5+ SDK facades.
$common = @('-nologo', '-noconfig', '-nostdlib', '-optimize+', '-target:exe') + $refs

# Two-stage compile: write into the cache folder first (exactly like the plugin
# does at run time). If that folder refuses writes -- under a sandboxed agent the
# cache folder and %TEMP% are often read-only, and csc then dies with CS1567 --
# fall back to a fresh temp dir and copy the result in. The fallback keeps the
# build working without weakening anything the plugin does when it compiles on
# demand.
$built = $null
& $csc @common "-out:$tmp" $src
if ($LASTEXITCODE -eq 0) {
    $built = $tmp
} else {
    $scratch = Join-Path ([System.IO.Path]::GetTempPath()) ("dsh-helper-build-" + $PID)
    New-Item -ItemType Directory -Force -Path $scratch | Out-Null
    $alt = Join-Path $scratch 'office-helper.exe'
    Write-Warning "csc could not write into $dir (usually a sandbox); retrying in $scratch"
    & $csc @common "-out:$alt" $src
    if ($LASTEXITCODE -ne 0) { throw "csc failed with exit $LASTEXITCODE" }
    $built = $alt
}

Move-Item -Force $built $exe
$item = Get-Item $exe
Write-Output "SRC  = $src"
Write-Output "HASH = $hash"
Write-Output "EXE  = $($item.FullName)"
Write-Output "SIZE = $($item.Length)"
Write-Output ""
Write-Output "Next: npm run smoke"
