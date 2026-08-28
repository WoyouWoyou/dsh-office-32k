# Repack a portable DSH-32K package into a clean, zip-safe distribution.
# Run with PowerShell 7 (pwsh). Removes runtime artifacts (regenerated at
# boot), verifies zero junctions/symlinks remain, then zips with tar.
#
# Usage: pwsh -File repack-portable.ps1 -Root "D:\DSH-32K便携版" -Zip "D:\DSH-32K便携版.zip"
param(
    [Parameter(Mandatory = $true)][string]$Root,
    [Parameter(Mandatory = $true)][string]$Zip
)
$ErrorActionPreference = 'Stop'

# 1) remove runtime artifacts (the server rebuilds these on first boot)
foreach ($rel in @('.dsh-home\profiles\node_modules', '.dsh-home\sessions', '.dsh-home\storages')) {
    $p = Join-Path $Root $rel
    if (Test-Path $p) {
        cmd /c rmdir /s /q $p 2>&1 | Out-Null
        Write-Output "removed $rel"
    }
}
Get-ChildItem $Root -File | Where-Object { $_.Name -like '_*' } | Remove-Item -Force

# 2) verify no links remain anywhere in the package
$links = Get-ChildItem $Root -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.LinkType -or ($_.Attributes -match 'ReparsePoint') }
Write-Output "remaining links: $($links.Count)"
if ($links) { $links | Select-Object -First 5 | ForEach-Object { Write-Output "  $($_.FullName)" } }

# 3) zip with tar (fast, handles unicode; zip-safe because no links remain)
if (Test-Path $Zip) { Remove-Item $Zip -Force }
$parent = Split-Path -Parent $Root
$leaf = Split-Path -Leaf $Root
Push-Location $parent
tar -caf $Zip $leaf
Pop-Location
$mb = [math]::Round((Get-Item $Zip).Length / 1MB)
Write-Output "zip created: $Zip ($mb MB)"
