Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class W {
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
}
"@

function Get-Text([IntPtr]$h) {
    $sb = New-Object System.Text.StringBuilder 512
    [void][W]::GetWindowTextW($h, $sb, 512)
    return $sb.ToString()
}

function Get-Class([IntPtr]$h) {
    $sb = New-Object System.Text.StringBuilder 256
    [void][W]::GetClassNameW($h, $sb, 256)
    return $sb.ToString()
}

$excelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
if ($excelPids.Count -eq 0) {
    Write-Output "no EXCEL process running"
    exit 0
}

$handles = New-Object System.Collections.ArrayList
$cb = [W+EnumProc]{ param($h, $l)
    $ownerId = 0
    [void][W]::GetWindowThreadProcessId($h, [ref]$ownerId)
    if ($excelPids -contains $ownerId) { [void]$script:handles.Add($h) }
    return $true
}
# capture the list in script scope for the delegate
$script:handles = $handles
[void][W]::EnumWindows($cb, [IntPtr]::Zero)

function Dump-Tree([IntPtr]$h, [int]$depth) {
    $indent = ' ' * ($depth * 2)
    $class = Get-Class $h
    $text = Get-Text $h
    $visible = if ([W]::IsWindowVisible($h)) { 'visible' } else { 'hidden' }
    Write-Output "$indent[$class] $visible '$text'"
    if ($depth -ge 3) { return }
    $child = [W]::GetWindow($h, 5)
    while ($child -ne [IntPtr]::Zero) {
        Dump-Tree $child ($depth + 1)
        $child = [W]::GetWindow($child, 2)
    }
}

foreach ($h in $handles) {
    Dump-Tree $h 0
}
