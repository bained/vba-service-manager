# Locates the line of a VBA compile error without any UI automation tool.
#
#     powershell -NoProfile -ExecutionPolicy Bypass -File tools\vbe-compile-locate.ps1
#     powershell -NoProfile -ExecutionPolicy Bypass -File tools\vbe-compile-locate.ps1 -All
#
# How it works: Excel is started with the generated workbook (by default
# build\Service_Manager.test.xlsm, because it carries the test modules as well;
# use -Target to point at another file), the VBE command
# "Debug -> Compile VBAProject" (CommandBar control ID 578) is executed through
# COM, and the modal "Compile error" dialog is detected by enumerating the top
# level windows. The error message is read from the dialog, the dialog is closed
# with a deterministic PostMessage (no SendKeys, no AutoIt), and the module plus
# the marked line are read from Application.VBE.ActiveCodePane.
#
# Without -All only the first error is reported. With -All the offending line is
# commented out in the in-memory copy and Compile runs again, so every compile
# error of the project is listed in one pass. The workbook is never saved, so
# the files on disk stay untouched.
#
# Output: build\compile-location.txt (and the console).

param(
    [switch]$All,
    [string]$Root = '',
    [string]$Target = '',
    [int]$DialogTimeoutSeconds = 25,
    [int]$MaxErrors = 25
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

$buildDir = Join-Path $Root 'build'
$reportPath = Join-Path $buildDir 'compile-location.txt'
$report = New-Object System.Collections.Generic.List[string]

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class VbeWin32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxCount);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    public static string Title(IntPtr hWnd) {
        StringBuilder buffer = new StringBuilder(512);
        GetWindowText(hWnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    public static string ClassName(IntPtr hWnd) {
        StringBuilder buffer = new StringBuilder(256);
        GetClassName(hWnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    // Returns the first visible top level dialog (class #32770) whose title
    // equals or starts with the given text; IntPtr.Zero when there is none.
    public static IntPtr FindDialog(string title) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            if (!IsWindowVisible(hWnd)) return true;
            if (ClassName(hWnd) != "#32770") return true;
            string caption = Title(hWnd);
            if (caption.Equals(title, StringComparison.OrdinalIgnoreCase) ||
                caption.StartsWith(title, StringComparison.OrdinalIgnoreCase)) {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    // All static texts of a dialog, joined with " | ".
    public static string DialogTexts(IntPtr dialog) {
        StringBuilder joined = new StringBuilder();
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) { return true; }, IntPtr.Zero);
        EnumChildWindows(dialog, delegate(IntPtr hWnd, IntPtr lParam) {
            if (ClassName(hWnd) == "Static") {
                string text = Title(hWnd).Trim();
                if (text.Length > 0) {
                    if (joined.Length > 0) joined.Append(" | ");
                    joined.Append(text);
                }
            }
            return true;
        }, IntPtr.Zero);
        return joined.ToString();
    }

    public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumChildProc callback, IntPtr lParam);

    public static void CloseDialog(IntPtr dialog) {
        // IDOK of a message box, then a plain close as a fallback.
        PostMessage(dialog, 0x0111, (IntPtr)1, IntPtr.Zero);
        System.Threading.Thread.Sleep(600);
        if (IsWindowVisible(dialog)) SendMessage(dialog, 0x0010, IntPtr.Zero, IntPtr.Zero);
    }
}
"@

# ---- VBE helpers --------------------------------------------------------

function Find-VbeControl($controls, [int]$controlId) {
    foreach ($control in $controls) {
        if ($control.Id -eq $controlId) { return $control }
        try {
            $child = Find-VbeControl $control.Controls $controlId
            if ($null -ne $child) { return $child }
        } catch {
            # not a popup control, so it has no children
        }
    }
    return $null
}

# Reads the module and the marked line from the VBE. The caret stays on the
# offending line after the compile error dialog was closed.
function Get-VbeSelection($xl) {
    $result = New-Object PSObject -Property @{
        Module = ''
        Line   = 0
        Column = 0
        Text   = ''
        Note   = ''
    }
    try {
        $pane = $xl.VBE.ActiveCodePane
        if ($null -eq $pane) {
            $result.Note = 'no active code pane'
            return $result
        }
        $code = $pane.CodeModule
        $result.Module = $code.Name
        $startLine = [int]0; $startColumn = [int]0; $endLine = [int]0; $endColumn = [int]0
        $pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
        $result.Line = [int]$startLine
        $result.Column = [int]$startColumn
        if ($result.Line -gt 0) {
            $result.Text = $code.Lines([int]$startLine, 1)
        }
    } catch {
        $result.Note = 'GetSelection failed: ' + $_.Exception.Message
    }
    return $result
}

# ---- main ---------------------------------------------------------------

$xl = $null
$wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $xl.Interactive = $true
    $xl.AutomationSecurity = 1          # msoAutomationSecurityLow
    $xl.ScreenUpdating = $false

    $target = $Target
    if ([string]::IsNullOrWhiteSpace($target)) {
        # The test workbook carries the application sources plus the test
        # modules, so it is the one that has to compile for a test run. The
        # plain build is the fallback.
        $target = Join-Path $buildDir 'Service_Manager.test.xlsm'
        if (-not (Test-Path $target)) { $target = Join-Path $buildDir 'Service_Manager.xlsm' }
    }
    if (-not (Test-Path $target)) {
        throw "no generated workbook in $buildDir - run 'cscript //nologo tools\build.vbs test' first"
    }

    $wb = $xl.Workbooks.Open($target)

    $report.Add("VBA compile error location")
    $report.Add("generated : " + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
    $report.Add("workbook  : " + $target)
    $report.Add("excel     : " + $xl.Version)
    if ($All) { $report.Add("mode      : enumerate all errors") } else { $report.Add("mode      : first error") }
    $report.Add("")

    $compileControl = Find-VbeControl $xl.VBE.CommandBars 578
    if ($null -eq $compileControl) { throw "the VBE Compile command (CommandBar control id 578) was not found" }

    $seen = @{}
    $found = $false

    for ($run = 1; $run -le $MaxErrors; $run++) {
        $compileControl.Execute()

        $dialog = $null
        $waited = 0.0
        while ($null -eq $dialog -and $waited -lt $DialogTimeoutSeconds) {
            $handle = [VbeWin32]::FindDialog('Microsoft Visual Basic for Applications')
            if ($handle -ne [IntPtr]::Zero) { $dialog = $handle } else { Start-Sleep -Milliseconds 500; $waited += 0.5 }
        }

        if ($null -eq $dialog) {
            if ($found) { $report.Add("no further compile error dialog after $waited s") }
            else { $report.Add("no compile error dialog after $waited s - the project compiles") }
            break
        }

        $found = $true
        $dialogText = [VbeWin32]::DialogTexts($dialog)
        $before = Get-VbeSelection $xl
        [VbeWin32]::CloseDialog($dialog)
        Start-Sleep -Milliseconds 900
        $after = Get-VbeSelection $xl

        $location = $after
        $source = 'after the dialog was closed'
        if ($null -eq $location -or $location.Line -le 0) {
            $location = $before
            $source = 'while the dialog was open'
        }

        $report.Add("--- compile error $run ---")
        $report.Add("dialog    : " + $dialogText)

        if ($null -eq $location -or $location.Module -eq '' -or $location.Line -le 0) {
            $report.Add("module    : could not be read from the VBE")
            if ($null -ne $before -and $before.Note -ne '') { $report.Add("note      : " + $before.Note) }
            if ($null -ne $after -and $after.Note -ne '') { $report.Add("note      : " + $after.Note) }
            break
        }

        $report.Add("module    : " + $location.Module)
        $report.Add("line      : " + $location.Line + " (column " + $location.Column + ", read " + $source + ")")
        $report.Add("code      : " + $location.Text.Trim())
        if ($location.Note -ne '') { $report.Add("note      : " + $location.Note) }

        $key = $location.Module + ':' + $location.Line
        if ($seen.ContainsKey($key)) {
            $report.Add("(the same location came up again - it is not fixed yet)")
            break
        }
        $seen[$key] = $true

        if (-not $All) { break }

        # Comment the offending line out in the in memory copy and compile
        # again, so the next error surfaces. The workbook is never saved.
        $pane = $xl.VBE.ActiveCodePane
        $codeLine = $pane.CodeModule.Lines([int]$location.Line, 1)
        $pane.CodeModule.ReplaceLine([int]$location.Line, "' [vbe-compile-locate] " + $codeLine)
        $report.Add("action    : the line was commented out in memory (not saved)")
        $report.Add("")
    }

    $report.Add("")
    $report.Add("done")
}
catch {
    $report.Add("FAILED: " + $_.Exception.Message)
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    if ($null -ne $xl) { try { $xl.Quit() } catch { } }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$report -join "`r`n" | Set-Content -Path $reportPath -Encoding UTF8
Write-Output ("compile location report written: " + $reportPath)
$report | ForEach-Object { Write-Output $_ }


