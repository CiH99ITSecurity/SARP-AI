$ErrorActionPreference = "SilentlyContinue"
$folder = "C:\SARP\etc\Wallpapers"
$extensions = @(".png", ".jpg", ".jpeg", ".bmp", ".gif")
$files = @(Get-ChildItem -LiteralPath $folder -File -Recurse | Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } | Sort-Object FullName)
if ($files.Count -eq 0) { exit 0 }
$stateFile = Join-Path $folder ".SARP_AI_WallpaperIndex.txt"
$index = -1
if (Test-Path -LiteralPath $stateFile) {
  $saved = Get-Content -LiteralPath $stateFile -Raw
  $parsed = 0
  if ([int]::TryParse($saved, [ref]$parsed)) { $index = $parsed }
}
$index = ($index + 1) % $files.Count
Set-Content -LiteralPath $stateFile -Value $index -Encoding ASCII
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class SarpWallpaperNative {
  [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  public static extern bool SystemParametersInfo(int action, int param, string path, int flags);
}
"@
[SarpWallpaperNative]::SystemParametersInfo(0x0014, 0, $files[$index].FullName, 0x01 -bor 0x02) | Out-Null
