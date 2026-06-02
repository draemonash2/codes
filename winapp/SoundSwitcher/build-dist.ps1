# Builds the distributable into the output directory (default: C:\prg_exe\SoundSwitcher).
#   <out>\SoundSwitcher.exe      - self-contained launcher (runs without .NET; checks
#                                  the runtime and installs it via winget if missing)
#   <out>\app\SoundSwitcher.exe  - the framework-dependent app (tiny, fast)
#   <out>\README.md              - user documentation
#
# Usage:  powershell -ExecutionPolicy Bypass -File build-dist.ps1
#         powershell -ExecutionPolicy Bypass -File build-dist.ps1 -OutDir D:\somewhere

param(
    [string]$OutDir = 'C:\prg_exe\SoundSwitcher'
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

Write-Host "Output: $OutDir"
Write-Host 'Cleaning output...'
Remove-Item $OutDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path (Join-Path $OutDir 'app') -Force | Out-Null

Write-Host 'Publishing app (framework-dependent, single file)...'
dotnet publish (Join-Path $root 'SoundSwitcher.csproj') -c Release -r win-x64 `
    --self-contained false -p:PublishSingleFile=true -o (Join-Path $OutDir 'app') | Out-Null

Write-Host 'Publishing launcher (self-contained, trimmed)...'
dotnet publish (Join-Path $root 'Launcher\Launcher.csproj') -c Release -o (Join-Path $root 'Launcher\pub') | Out-Null
Copy-Item (Join-Path $root 'Launcher\pub\SoundSwitcher.exe') (Join-Path $OutDir 'SoundSwitcher.exe') -Force

# Ship the user documentation alongside the binaries.
Copy-Item (Join-Path $root 'README.md') (Join-Path $OutDir 'README.md') -Force

# Tidy: drop pdbs from the distribution
Get-ChildItem $OutDir -Recurse -Filter *.pdb | Remove-Item -Force

$total = (Get-ChildItem $OutDir -Recurse -File | Measure-Object Length -Sum).Sum
Write-Host ("Done. Output total: {0:N2} MB" -f ($total / 1MB))
Get-ChildItem $OutDir -Recurse -File | ForEach-Object {
    Write-Host ("  {0,8:N2} MB  {1}" -f ($_.Length / 1MB), $_.FullName.Substring($OutDir.Length))
}
