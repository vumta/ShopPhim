# AutoSavePAD_Final_FullFixed_v2.ps1
# Auto-Save Power Automate Desktop (PAD) — tray app
# Fixes:
# - Correct Add-Type C# here-strings (single-quoted) to avoid escaping issues
# - Native types added early
# - Show-Toast defined early, script-scoped UI, form created before handlers
# - Safe-BringPADToFront + AppActivate fallback
# - Prompt before each save (Yes/No). If No, skip cycle.
# - Robust logging and trap
# Save as UTF-8 with BOM. Run with:
#   powershell.exe -STA -ExecutionPolicy Bypass -File "C:\Scripts\AutoSavePAD_Final_FullFixed_v2.ps1"

# ---------------- Relaunch as STA if needed ----------------
if ([System.Threading.Thread]::CurrentThread.GetApartmentState().ToString() -ne 'STA') {
    $psExe = (Get-Command powershell).Source
    $scriptPath = $MyInvocation.MyCommand.Path
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $psExe
    $psi.Arguments = "-STA -ExecutionPolicy Bypass -File `"$scriptPath`""
    $psi.UseShellExecute = $true
    try { [System.Diagnostics.Process]::Start($psi) | Out-Null } catch {}
    exit
}

# ---------------- Basic assemblies ----------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- Logging helpers ----------------
$script:LogFolder = Join-Path $env:USERPROFILE "Documents"
$script:LogPrefix = "AutoSavePAD_"
$script:CriticalLogPath = Join-Path $script:LogFolder "AutoSavePAD_CriticalErrors.log"
function Write-CriticalLog { param([string]$m) try { $ts=Get-Date -Format "yyyy-MM-dd HH:mm:ss"; Add-Content -Path $script:CriticalLogPath -Value ("{0} - {1}" -f $ts,$m) -ErrorAction SilentlyContinue } catch {} }
function Write-Log { param([string]$m) try { $ts=Get-Date -Format "yyyy-MM-dd HH:mm:ss"; $p = Join-Path $script:LogFolder ("{0}{1}.log" -f $script:LogPrefix,(Get-Date -Format 'yyyy-MM-dd')); Add-Content -Path $p -Value ("{0} - {1}" -f $ts,$m) -ErrorAction SilentlyContinue } catch {}; Write-Host $m }

# ---------------- Trap ----------------
trap {
    try {
        $err = $_.Exception
        $inv = $_.InvocationInfo
        Write-CriticalLog (("TRAP: {0}: {1}" -f $err.GetType().FullName, $err.Message))
        if ($inv -and $inv.ScriptName) {
            Write-CriticalLog (("Location: {0} line {1}" -f $inv.ScriptName, $inv.ScriptLineNumber))
            try {
                $start = [math]::Max(1, $inv.ScriptLineNumber - 8)
                $end = $inv.ScriptLineNumber + 8
                $ctx = (Get-Content $inv.ScriptName)[$start-1..($end-1)]
                $i = $start
                foreach ($l in $ctx) { Write-CriticalLog (("{0,4}: {1}" -f $i,$l)); $i++ }
            } catch { Write-CriticalLog ("Failed to read context: {0}" -f $_.Exception.Message) }
        }
        Write-CriticalLog (("StackTrace:`n{0}" -f $err.StackTrace))
    } catch {
        try { Write-CriticalLog ("Trap handler failure: {0}" -f $_.Exception.Message) } catch {}
    }
    continue
}

# ---------------- Safe Add-Type helper ----------------
function SafeAddType {
    param([string]$TypeName,[string]$SourceCode)
    try { if ([type]::GetType($TypeName,$false,$false)) { Write-Log ("SafeAddType: {0} exists" -f $TypeName); return $true } } catch {}
    try {
        Add-Type -TypeDefinition $SourceCode -ErrorAction Stop
        Write-Log ("SafeAddType: added {0}" -f $TypeName)
        return $true
    } catch {
        Write-CriticalLog (("Add-Type failed for {0}: {1}" -f $TypeName, $_.Exception.Message))
        try {
            $prov = New-Object Microsoft.CSharp.CSharpCodeProvider
            $cp = New-Object System.CodeDom.Compiler.CompilerParameters
            $cp.GenerateExecutable = $false; $cp.GenerateInMemory = $true
            $cp.ReferencedAssemblies.Add("System.dll") | Out-Null
            $cp.ReferencedAssemblies.Add("System.Core.dll") | Out-Null
            $res = $prov.CompileAssemblyFromSource($cp, $SourceCode)
            if ($res.Errors.HasErrors) {
                $sb = New-Object System.Text.StringBuilder
                foreach ($e in $res.Errors) { $sb.AppendLine(("{0} at {1},{2}: {3}" -f (if ($e.IsWarning) {"WARN"} else {"ERR"}), $e.Line, $e.Column, $e.ErrorText)) | Out-Null }
                Write-CriticalLog (("Compiler details for {0}:`n{1}" -f $TypeName, $sb.ToString()))
            }
        } catch { Write-CriticalLog ("SafeAddType helper failed: {0}" -f $_.Exception.Message) }
        return $false
    }
}

# ---------------- Embedded native helpers (use single-quoted here-strings) ----------------
$idleCode = @'
using System;
using System.Runtime.InteropServices;
public static class IdleTimeHelper {
    [StructLayout(LayoutKind.Sequential)]
    public struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
    [DllImport("user32.dll")] public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    public static TimeSpan GetIdleTime() {
        LASTINPUTINFO lii = new LASTINPUTINFO(); lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
        if (!GetLastInputInfo(ref lii)) return TimeSpan.Zero;
        return TimeSpan.FromMilliseconds(Environment.TickCount - lii.dwTime);
    }
}
'@

$winApiCode = @'
using System;
using System.Runtime.InteropServices;
public static class WinAPI_User32 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("kernel32.dll")] public static extern uint GetLastError();
}
'@

# Add types early and exit if not possible
if (-not (SafeAddType -TypeName "IdleTimeHelper" -SourceCode $idleCode)) { Write-CriticalLog "Cannot add IdleTimeHelper type; aborting"; exit 1 }
if (-not (SafeAddType -TypeName "WinAPI_User32" -SourceCode $winApiCode)) { Write-CriticalLog "Cannot add WinAPI_User32 type; aborting"; exit 1 }

function Get-IdleTime { try { if ([type]::GetType("IdleTimeHelper",$false,$false)) { return [IdleTimeHelper]::GetIdleTime() }; return [TimeSpan]::Zero } catch { Write-CriticalLog ("Get-IdleTime error: {0}" -f $_.Exception.Message); return [TimeSpan]::Zero } }

# ---------------- Config / State ----------------
$script:StatePath = Join-Path $script:LogFolder "AutoSavePAD_State.json"
$script:IdleLimitMs = 600000
$script:DefaultIntervalMs = 180000
$script:Settings = @{ IntervalMs = $script:DefaultIntervalMs; MaxRetries = 3; InitialBackoffMs = 300; MaxBackoffMs = 3000; ScheduleStart="00:00"; ScheduleEnd="23:59" }
$script:AutoSaveEnabled = $false
$script:InSleepMode = $false
$script:ConsecutiveFailures = 0
$script:SuccessCount = 0
$script:FailureCount = 0
$script:LastSuccessTime = $null

function Save-State {
    try {
        $obj = [PSCustomObject]@{
            AutoSaveEnabled = $script:AutoSaveEnabled
            InSleepMode = $script:InSleepMode
            ConsecutiveFailures = $script:ConsecutiveFailures
            SuccessCount = $script:SuccessCount
            FailureCount = $script:FailureCount
            LastSuccessTime = if ($script:LastSuccessTime) { $script:LastSuccessTime.ToString("o") } else { $null }
            Settings = $script:Settings
        }
        $json = $obj | ConvertTo-Json -Depth 6
        $dir = Split-Path $script:StatePath
        if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
        Set-Content -Path $script:StatePath -Value $json -Encoding UTF8
        Write-Log "State saved"
    } catch { Write-CriticalLog ("Save-State error: {0}" -f $_.Exception.Message) }
}

function Load-State {
    try {
        if (-not (Test-Path $script:StatePath)) { return }
        $obj = Get-Content -Path $script:StatePath -Raw | ConvertFrom-Json
        if ($obj) {
            $script:AutoSaveEnabled = $obj.AutoSaveEnabled
            $script:InSleepMode = $obj.InSleepMode
            $script:ConsecutiveFailures = $obj.ConsecutiveFailures
            $script:SuccessCount = $obj.SuccessCount
            $script:FailureCount = $obj.FailureCount
            if ($obj.LastSuccessTime) { $script:LastSuccessTime = [datetime]::Parse($obj.LastSuccessTime) }
            if ($obj.Settings) { foreach ($k in $obj.Settings.PSObject.Properties.Name) { $script:Settings[$k] = $obj.Settings.$k } }
            Write-Log "State loaded"
        }
    } catch { Write-CriticalLog ("Load-State error: {0}" -f $_.Exception.Message) }
}

# ---------------- Robust Show-Toast (declare early) ----------------
if (-not (Get-Command -Name Show-Toast -CommandType Function -ErrorAction SilentlyContinue)) {
    function Show-Toast {
        param([string]$message)
        try {
            if (Get-Module -Name BurntToast -ListAvailable) {
                try { Import-Module BurntToast -ErrorAction SilentlyContinue; New-BurntToastNotification -Text "Auto-Save PAD", $message -ErrorAction Stop; return } catch { Write-Log ("BurntToast invocation failed: {0}" -f $_.Exception.Message) }
            }
        } catch {}
        try {
            if ($script:NotifyIcon) {
                $script:NotifyIcon.BalloonTipTitle = "Auto-Save PAD"
                $script:NotifyIcon.BalloonTipText = $message
                $script:NotifyIcon.ShowBalloonTip(3000)
                return
            }
            $tmp = New-Object System.Windows.Forms.NotifyIcon
            $tmp.Icon = [System.Drawing.SystemIcons]::Information
            $tmp.Visible = $true
            $tmp.BalloonTipTitle = "Auto-Save PAD"
            $tmp.BalloonTipText = $message
            try { $tmp.ShowBalloonTip(3000) } catch {}
            Start-Sleep -Seconds 3
            try { $tmp.Dispose() } catch {}
        } catch { Write-Log ("Show-Toast final fallback error: {0}" -f $_.Exception.Message) }
    }
}

# ---------------- UI (create form first, then register events) ----------------
$script:NotifyIcon = $null
$script:Menu = $null
$script:MiStatus = $null
$script:MiStartStop = $null
$script:MiForceSave = $null
$script:MiOpenLog = $null
$script:MiSettings = $null
$script:MiExit = $null
$script:Form = $null

try {
    $script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $script:NotifyIcon.Icon = [System.Drawing.SystemIcons]::Application
    $script:NotifyIcon.Text = "Auto-Save PAD"
    $script:NotifyIcon.Visible = $true

    $script:Menu = New-Object System.Windows.Forms.ContextMenuStrip
    $script:MiStatus = New-Object System.Windows.Forms.ToolStripMenuItem "Auto-Save: OFF"
    $script:MiStartStop = New-Object System.Windows.Forms.ToolStripMenuItem "Bật Auto-Save"
    $script:MiForceSave = New-Object System.Windows.Forms.ToolStripMenuItem "Force Save Now"
    $script:MiOpenLog = New-Object System.Windows.Forms.ToolStripMenuItem "Mở log"
    $script:MiSettings = New-Object System.Windows.Forms.ToolStripMenuItem "Cấu hình..."
    $script:MiExit = New-Object System.Windows.Forms.ToolStripMenuItem "Thoát"

    $script:Menu.Items.AddRange(@($script:MiStatus,$script:MiStartStop,$script:MiForceSave,$script:MiOpenLog,$script:MiSettings,$script:MiExit))
    $script:NotifyIcon.ContextMenuStrip = $script:Menu
} catch { Write-CriticalLog ("UI init failed: {0}" -f $_.Exception.Message) }

try {
    $script:Form = New-Object System.Windows.Forms.Form
    $script:Form.Size = New-Object System.Drawing.Size(0,0)
    $script:Form.ShowInTaskbar = $false
    $script:Form.WindowState = "Minimized"
} catch { Write-CriticalLog ("Form creation failed: {0}" -f $_.Exception.Message); $script:Form = $null }

function Update-Tooltip {
    $status = if ($script:AutoSaveEnabled) {"ON"} elseif ($script:InSleepMode) {"SLEEP"} else {"OFF"}
    $tooltip = ("Auto-Save: {0}`nSuccess: {1}; Fail: {2}`nConsecFailures: {3}" -f $status, $script:SuccessCount, $script:FailureCount, $script:ConsecutiveFailures)
    try {
        if ($script:NotifyIcon) {
            $script:NotifyIcon.Text = $tooltip.Substring(0,[math]::Min(63,$tooltip.Length))
            $script:NotifyIcon.BalloonTipTitle = "Auto-Save PAD"
            $script:NotifyIcon.BalloonTipText = ("Status: {0}`nSuccess: {1}`nFail: {2}" -f $status, $script:SuccessCount, $script:FailureCount)
        } else { Write-Log "Update-Tooltip: NotifyIcon is null" }
    } catch { Write-Log ("Update-Tooltip error: {0}" -f $_.Exception.Message) }
}
Update-Tooltip

# ---------------- Timers ----------------
$script:Timer = New-Object System.Windows.Forms.Timer
$script:Timer.Interval = $script:Settings.IntervalMs
$script:IdleWatcher = New-Object System.Windows.Forms.Timer
$script:IdleWatcher.Interval = 5000

# ---------------- PAD helpers ----------------
function Get-PADMainProcess { try { Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*Power Automate*" } | Select-Object -First 1 } catch { return $null } }
function Is-PADRunning { return (Get-PADMainProcess) -ne $null }
function Is-PADResponsive { $p = Get-PADMainProcess; if (-not $p) { return $false }; return $p.Responding }
function Can-SendKeysToPAD { $p = Get-PADMainProcess; if (-not $p) { return $false }; if (-not $p.Responding) { return $false }; if ([string]::IsNullOrWhiteSpace($p.MainWindowTitle)) { return $false }; return $true }

# ---------------- Improved Safe-BringPADToFront ----------------
function Safe-BringPADToFront {
    param($proc)
    try {
        if (-not $proc) { Write-Log "Safe-BringPADToFront: proc is null"; return $false }
        $hWnd = $proc.MainWindowHandle
        if (-not $hWnd -or $hWnd -eq [IntPtr]::Zero) { Write-Log "Safe-BringPADToFront: invalid MainWindowHandle"; return $false }
        Write-Log ("Safe-BringPADToFront: Attempting focus: Handle={0}" -f $hWnd)

        try {
            $isIconic = [WinAPI_User32]::IsIconic($hWnd)
            if ($isIconic) {
                [WinAPI_User32]::ShowWindow($hWnd,9) | Out-Null
                Start-Sleep -Milliseconds 250
            }
        } catch { Write-Log ("Safe-BringPADToFront: IsIconic/ShowWindow failed: {0}" -f $_.Exception.Message) }

        try {
            if ([WinAPI_User32]::SetForegroundWindow($hWnd)) { Write-Log "Safe-BringPADToFront: SetForegroundWindow succeeded"; return $true }
            else { Write-Log "Safe-BringPADToFront: SetForegroundWindow returned false" }
        } catch { Write-Log ("Safe-BringPADToFront: SetForegroundWindow exception: {0}" -f $_.Exception.Message) }

        try { [WinAPI_User32]::BringWindowToTop($hWnd) | Out-Null; Start-Sleep -Milliseconds 120 } catch {}

        try {
            $fg = [WinAPI_User32]::GetForegroundWindow()
            $currentThread = [WinAPI_User32]::GetCurrentThreadId()
            $out1 = 0
            $winThread = [WinAPI_User32]::GetWindowThreadProcessId($fg, [ref]$out1)
            $out2 = 0
            $targetThread = [WinAPI_User32]::GetWindowThreadProcessId($hWnd, [ref]$out2)
            Write-Log ("Safe-BringPADToFront: threads: current={0} fg={1} target={2}" -f $currentThread, $winThread, $targetThread)

            if ($targetThread -ne 0 -and $winThread -ne 0) {
                if ([WinAPI_User32]::AttachThreadInput($currentThread, $targetThread, $true)) {
                    Start-Sleep -Milliseconds 100
                    [WinAPI_User32]::SetForegroundWindow($hWnd) | Out-Null
                    [WinAPI_User32]::AttachThreadInput($currentThread, $targetThread, $false) | Out-Null
                    Start-Sleep -Milliseconds 100
                    if ([WinAPI_User32]::GetForegroundWindow() -eq $hWnd) { Write-Log "Safe-BringPADToFront: success via AttachThreadInput"; return $true } else { Write-Log "Safe-BringPADToFront: AttachThreadInput path did not set foreground" }
                } else { Write-Log "Safe-BringPADToFront: AttachThreadInput failed" }
            } else { Write-Log "Safe-BringPADToFront: Could not obtain thread IDs" }
        } catch { Write-Log ("Safe-BringPADToFront: AttachThreadInput exception: {0}" -f $_.Exception.Message) }

        try {
            [WinAPI_User32]::ShowWindow($hWnd,9) | Out-Null
            Start-Sleep -Milliseconds 150
            [WinAPI_User32]::BringWindowToTop($hWnd) | Out-Null
            Start-Sleep -Milliseconds 150
            if ([WinAPI_User32]::SetForegroundWindow($hWnd)) { Write-Log "Safe-BringPADToFront: final SetForegroundWindow succeeded"; return $true }
        } catch { Write-Log ("Safe-BringPADToFront: final attempt exception: {0}" -f $_.Exception.Message) }

        $err = [WinAPI_User32]::GetLastError()
        Write-Log ("Safe-BringPADToFront: giving up; Win32 GetLastError={0}" -f $err)
        return $false
    } catch {
        Write-Log ("Safe-BringPADToFront: unexpected error: {0}" -f $_.Exception.Message)
        return $false
    }
}

# ---------------- Show-SettingsWindow ----------------
function Show-SettingsWindow {
    try {
        $script:SettingsForm = New-Object System.Windows.Forms.Form
        $script:SettingsForm.Text = "Cấu hình Auto-Save PAD"
        $script:SettingsForm.Size = New-Object System.Drawing.Size(420,380)
        $script:SettingsForm.StartPosition = "CenterScreen"
        $script:SettingsForm.TopMost = $true

        $lblInterval = New-Object System.Windows.Forms.Label
        $lblInterval.Text = "Chu kỳ lưu (phút):"
        $lblInterval.Location = New-Object System.Drawing.Point(20,20)
        $lblInterval.AutoSize = $true
        $script:SettingsForm.Controls.Add($lblInterval)

        $numInterval = New-Object System.Windows.Forms.NumericUpDown
        $numInterval.Minimum = 1
        $numInterval.Maximum = 1440
        $numInterval.Value = [int]($script:Settings.IntervalMs/60000)
        $numInterval.Location = New-Object System.Drawing.Point(160,18)
        $script:SettingsForm.Controls.Add($numInterval)

        $lblRetries = New-Object System.Windows.Forms.Label
        $lblRetries.Text = "Max retry khi thất bại:"
        $lblRetries.Location = New-Object System.Drawing.Point(20,60)
        $lblRetries.AutoSize = $true
        $script:SettingsForm.Controls.Add($lblRetries)

        $numRetries = New-Object System.Windows.Forms.NumericUpDown
        $numRetries.Minimum = 0
        $numRetries.Maximum = 20
        $numRetries.Value = [int]$script:Settings.MaxRetries
        $numRetries.Location = New-Object System.Drawing.Point(160,58)
        $script:SettingsForm.Controls.Add($numRetries)

        $lblInitBack = New-Object System.Windows.Forms.Label
        $lblInitBack.Text = "Initial backoff (ms):"
        $lblInitBack.Location = New-Object System.Drawing.Point(20,100)
        $lblInitBack.AutoSize = $true
        $script:SettingsForm.Controls.Add($lblInitBack)

        $txtInitBack = New-Object System.Windows.Forms.TextBox
        $txtInitBack.Text = $script:Settings.InitialBackoffMs.ToString()
        $txtInitBack.Location = New-Object System.Drawing.Point(160,98)
        $script:SettingsForm.Controls.Add($txtInitBack)

        $lblMaxBack = New-Object System.Windows.Forms.Label
        $lblMaxBack.Text = "Max backoff (ms):"
        $lblMaxBack.Location = New-Object System.Drawing.Point(20,140)
        $lblMaxBack.AutoSize = $true
        $script:SettingsForm.Controls.Add($lblMaxBack)

        $txtMaxBack = New-Object System.Windows.Forms.TextBox
        $txtMaxBack.Text = $script:Settings.MaxBackoffMs.ToString()
        $txtMaxBack.Location = New-Object System.Drawing.Point(160,138)
        $script:SettingsForm.Controls.Add($txtMaxBack)

        $lblSchedule = New-Object System.Windows.Forms.Label
        $lblSchedule.Text = "Lịch lưu (HH:mm):"
        $lblSchedule.Location = New-Object System.Drawing.Point(20,180)
        $lblSchedule.AutoSize = $true
        $script:SettingsForm.Controls.Add($lblSchedule)

        $txtStart = New-Object System.Windows.Forms.TextBox
        $txtStart.Text = $script:Settings.ScheduleStart
        $txtStart.Location = New-Object System.Drawing.Point(160,178)
        $txtStart.Width = 80
        $script:SettingsForm.Controls.Add($txtStart)

        $lblTo = New-Object System.Windows.Forms.Label
        $lblTo.Text = "đến"
        $lblTo.Location = New-Object System.Drawing.Point(250,180)
        $lblTo.AutoSize = $true
        $script:SettingsForm.Controls.Add($lblTo)

        $txtEnd = New-Object System.Windows.Forms.TextBox
        $txtEnd.Text = $script:Settings.ScheduleEnd
        $txtEnd.Location = New-Object System.Drawing.Point(290,178)
        $txtEnd.Width = 80
        $script:SettingsForm.Controls.Add($txtEnd)

        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Text = "Lưu"
        $btnSave.Location = New-Object System.Drawing.Point(160,300)
        $btnSave.Size = New-Object System.Drawing.Size(80,30)
        $script:SettingsForm.Controls.Add($btnSave)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Hủy"
        $btnCancel.Location = New-Object System.Drawing.Point(260,300)
        $btnCancel.Size = New-Object System.Drawing.Size(80,30)
        $script:SettingsForm.Controls.Add($btnCancel)

        $btnSave.Add_Click({
            try {
                $newIntervalMin = [int]$numInterval.Value
                $script:Settings.IntervalMs = $newIntervalMin * 60000
                if ($script:Timer) { $script:Timer.Interval = $script:Settings.IntervalMs }

                $script:Settings.MaxRetries = [int]$numRetries.Value
                $script:Settings.InitialBackoffMs = [int]$txtInitBack.Text
                $script:Settings.MaxBackoffMs = [int]$txtMaxBack.Text

                [void][datetime]::ParseExact($txtStart.Text,'HH:mm',$null)
                [void][datetime]::ParseExact($txtEnd.Text,'HH:mm',$null)
                $script:Settings.ScheduleStart = $txtStart.Text
                $script:Settings.ScheduleEnd = $txtEnd.Text

                Write-Log ("Cập nhật cấu hình: Interval={0}ms, Retries={1}, Backoff={2}-{3}ms, Schedule={4}-{5}" -f $script:Settings.IntervalMs, $script:Settings.MaxRetries, $script:Settings.InitialBackoffMs, $script:Settings.MaxBackoffMs, $script:Settings.ScheduleStart, $script:Settings.ScheduleEnd)
                Show-Toast "Đã lưu cấu hình Auto-Save"
                Update-Tooltip
                Save-State
                try { $script:SettingsForm.Close() } catch {}
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Giá trị nhập không hợp lệ. Vui lòng kiểm tra lại.","Lỗi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

        $btnCancel.Add_Click({ try { $script:SettingsForm.Close() } catch {} })
        $script:SettingsForm.ShowDialog() | Out-Null
    } catch {
        Write-CriticalLog (("Show-SettingsWindow error: {0}`n{1}" -f $_.Exception.Message, $_.Exception.StackTrace))
        try { [System.Windows.Forms.MessageBox]::Show("Không thể mở cửa sổ Cấu hình. Xem log để biết chi tiết.","Lỗi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null } catch {}
    }
}

# ---------------- Try-SendSave (AppActivate + SendKeys, prompt before save) ----------------
function Try-SendSave {
    param([int]$maxAttempts = $script:Settings.MaxRetries, [int]$initialBackoff = $script:Settings.InitialBackoffMs, [int]$maxBackoff = $script:Settings.MaxBackoffMs)

    Write-Log ("Try-SendSave started (maxAttempts={0})" -f $maxAttempts)
    if (-not (Get-PADMainProcess)) { Write-Log "Try-SendSave: PAD not running"; return $false }
    if (-not (Is-PADResponsive)) { Write-Log "Try-SendSave: PAD not responding"; return $false }
    if (-not (Is-InSchedule $script:Settings.ScheduleStart $script:Settings.ScheduleEnd)) { Write-Log "Try-SendSave: Outside schedule"; return $false }
    if (Is-FlowRunning) { Write-Log "Try-SendSave: Flow running; skip"; return $false }

    $attempt = 0
    $delay = [int]$initialBackoff

    while ($attempt -lt $maxAttempts) {
        $attempt++
        Write-Log ("Try-SendSave: attempt {0}" -f $attempt)
        $pad = Get-PADMainProcess
        if (-not $pad) { Write-Log "Try-SendSave: PAD window missing"; break }

        try { Write-Log ("PAD info: Id={0}; Title=`"{1}`"; Responding={2}; Handle={3}" -f $pad.Id, $pad.MainWindowTitle, $pad.Responding, $pad.MainWindowHandle) } catch {}

        # Prefer bringing via WinAPI; if fails, fallback to AppActivate
        $focused = Safe-BringPADToFront $pad
        Write-Log ("Safe-BringPADToFront returned: {0}" -f $focused)
        Start-Sleep -Milliseconds 300

        if (-not $focused) {
            try {
                $wsh = New-Object -ComObject WScript.Shell
                $ok = $false
                try { $ok = $wsh.AppActivate($pad.Id); Write-Log ("AppActivate by PID returned: {0}" -f $ok) } catch { Write-Log ("AppActivate by PID exception: {0}" -f $_.Exception.Message) }
                if (-not $ok -and $pad.MainWindowTitle) {
                    try { $ok = $wsh.AppActivate($pad.MainWindowTitle); Write-Log ("AppActivate by Title returned: {0}" -f $ok) } catch { Write-Log ("AppActivate by Title exception: {0}" -f $_.Exception.Message) }
                }
                Start-Sleep -Milliseconds 350
                if ($ok) { $focused = $true }
            } catch { Write-Log ("AppActivate overall exception: {0}" -f $_.Exception.Message) }
        }

        if (-not $focused) {
            Write-Log "Could not focus PAD; will retry after backoff"
            Start-Sleep -Milliseconds $delay
            $delay = [math]::Min($maxBackoff, $delay * 2)
            continue
        }

        # Prompt user before sending save
        try {
            $caption = "Auto-Save PAD"
            $message = "Auto-Save định gửi lệnh Lưu (Ctrl+S) cho Power Automate Desktop ngay bây giờ.`nBạn có muốn tiếp tục?"
            $res = [System.Windows.Forms.MessageBox]::Show($message, $caption, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button1)
            if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log "User chose NOT to save this cycle (skipping)."
                Show-Toast "Bỏ qua lưu lần này"
                return $false
            }
        } catch { Write-Log ("Prompt for save failed: {0}" -f $_.Exception.Message) }

        # Try sending Ctrl+S
        $sentOk = $false
        try {
            if (-not $wsh) { $wsh = New-Object -ComObject WScript.Shell }
            try {
                $wsh.SendKeys('^s')
                Start-Sleep -Milliseconds 500
                $sentOk = $true
                Write-Log "Sent Ctrl+S via WScript.Shell.SendKeys"
            } catch { Write-Log ("Wsh.SendKeys failed: {0}" -f $_.Exception.Message) }
        } catch {}

        if (-not $sentOk) {
            try {
                [System.Windows.Forms.SendKeys]::SendWait('^s')
                Start-Sleep -Milliseconds 500
                $sentOk = $true
                Write-Log "Sent Ctrl+S via SendKeys.SendWait"
            } catch { Write-Log ("SendKeys.SendWait failed: {0}" -f $_.Exception.Message) }
        }

        if ($sentOk) {
            $script:ConsecutiveFailures = 0
            $script:SuccessCount++
            $script:LastSuccessTime = Get-Date
            Update-Tooltip; Save-State
            Show-Toast ("Đã lưu flow lúc {0}" -f (Get-Date -Format 'HH:mm:ss'))
            return $true
        } else {
            Write-Log "Send attempt failed; will retry"
        }

        Start-Sleep -Milliseconds $delay
        $delay = [math]::Min($maxBackoff, $delay * 2)
    }

    $script:ConsecutiveFailures++
    $script:FailureCount++
    Update-Tooltip; Save-State
    Write-Log ("Try-SendSave failed after {0} attempts" -f $attempt)
    Show-Toast ("Auto-Save: không lưu được (lần thử {0})" -f $attempt)
    return $false
}

# ---------------- Helper schedule/flow checks ----------------
function Is-InSchedule { param($startTime,$endTime) $now=Get-Date; $fmt='HH:mm'; try { $s=[datetime]::ParseExact($startTime,$fmt,$null); $e=[datetime]::ParseExact($endTime,$fmt,$null) } catch { return $true }; $todayS=Get-Date -Hour $s.Hour -Minute $s.Minute -Second 0; $todayE=Get-Date -Hour $e.Hour -Minute $e.Minute -Second 0; if ($todayS -le $todayE) { return ($now -ge $todayS -and $now -le $todayE) } else { return ($now -ge $todayS -or $now -le $todayE) } }
function Is-FlowRunning { try { $host = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq "PAD.Console.Host" -or $_.ProcessName -like "PAD*" }; return ($host -ne $null -and $host.Count -gt 0) } catch { return $false } }

# ---------------- Timer handlers ----------------
$script:Timer.Add_Tick({
    try {
        if (-not $script:AutoSaveEnabled) { return }
        if ($script:InSleepMode) { return }

        $idle = Get-IdleTime
        if ($idle.TotalMilliseconds -ge $script:IdleLimitMs) {
            Write-Log ("Idle {0:N2} phút -> Sleep" -f $idle.TotalMinutes)
            Show-Toast "Auto-Save tạm dừng do idle"
            $script:InSleepMode = $true
            if ($script:MiStatus) { $script:MiStatus.Text = "Auto-Save: SLEEP" }
            if ($script:MiStartStop) { $script:MiStartStop.Text = "Bật Auto-Save" }
            try { $script:Timer.Stop() } catch {}
            Update-Tooltip; Save-State
            return
        }

        if (-not (Is-InSchedule $script:Settings.ScheduleStart $script:Settings.ScheduleEnd)) { Write-Log "Outside schedule"; return }
        if (-not (Is-PADRunning)) { Write-Log "PAD not running"; return }
        if (-not (Is-PADResponsive)) { Write-Log "PAD not responding"; return }

        Try-SendSave -maxAttempts $script:Settings.MaxRetries -initialBackoff $script:Settings.InitialBackoffMs -maxBackoff $script:Settings.MaxBackoffMs | Out-Null
    } catch { Write-CriticalLog ("Timer error: {0}" -f $_.Exception.Message) }
})

$script:IdleWatcher.Add_Tick({
    try {
        if (-not $script:InSleepMode) { return }
        if (Is-PADRunning) {
            Write-Log "PAD started -> resume autosave"
            Show-Toast "PAD chạy lại — Auto-Save resume"
            $script:InSleepMode = $false; $script:AutoSaveEnabled = $true
            if ($script:MiStatus) { $script:MiStatus.Text = "Auto-Save: ON" }
            if ($script:MiStartStop) { $script:MiStartStop.Text = "Tắt Auto-Save" }
            try { $script:Timer.Interval = $script:Settings.IntervalMs; $script:Timer.Start() } catch {}
            Update-Tooltip; Save-State
            return
        }
        $idle = Get-IdleTime
        if ($idle.TotalMilliseconds -lt $script:IdleLimitMs) {
            Write-Log ("User active (idle {0:N1}s) -> resume" -f $idle.TotalSeconds)
            Show-Toast "Phát hiện tương tác — Auto-Save resume"
            $script:InSleepMode = $false; $script:AutoSaveEnabled = $true
            if ($script:MiStatus) { $script:MiStatus.Text = "Auto-Save: ON" }
            if ($script:MiStartStop) { $script:MiStartStop.Text = "Tắt Auto-Save" }
            try { $script:Timer.Interval = $script:Settings.IntervalMs; $script:Timer.Start() } catch {}
            Update-Tooltip; Save-State
            return
        }
    } catch { Write-CriticalLog ("IdleWatcher error: {0}" -f $_.Exception.Message) }
})

# ---------------- Menu handlers ----------------
$script:MiStartStop.Add_Click({
    try {
        $script:AutoSaveEnabled = -not $script:AutoSaveEnabled
        if ($script:AutoSaveEnabled) {
            if ($script:InSleepMode) { Show-Toast "Ứng dụng đang ở chế độ Sleep"; $script:AutoSaveEnabled = $false; return }
            if ($script:MiStatus) { $script:MiStatus.Text = "Auto-Save: ON" }
            if ($script:MiStartStop) { $script:MiStartStop.Text = "Tắt Auto-Save" }
            try { $script:Timer.Interval = $script:Settings.IntervalMs; $script:Timer.Start() } catch {}
            Write-Log "Auto-Save ON"; Show-Toast "Auto-Save bật"
        } else {
            if ($script:MiStatus) { $script:MiStatus.Text = "Auto-Save: OFF" }
            if ($script:MiStartStop) { $script:MiStartStop.Text = "Bật Auto-Save" }
            try { $script:Timer.Stop() } catch {}
            Write-Log "Auto-Save OFF"; Show-Toast "Auto-Save tắt"
        }
        Update-Tooltip; Save-State
    } catch { Write-CriticalLog ("MiStartStop handler error: {0}`n{1}" -f $_.Exception.Message, $_.Exception.StackTrace) }
})

$script:MiForceSave.Add_Click({
    try { Write-Log "Force save clicked"; Try-SendSave -maxAttempts $script:Settings.MaxRetries -initialBackoff $script:Settings.InitialBackoffMs -maxBackoff $script:Settings.MaxBackoffMs | Out-Null } catch { Write-CriticalLog ("ForceSave handler error: {0}" -f $_.Exception.Message) }
})

$script:MiOpenLog.Add_Click({
    try {
        $lp = Join-Path $env:USERPROFILE ("Documents\{0}{1}.log" -f $script:LogPrefix,(Get-Date -Format 'yyyy-MM-dd'))
        if (-not (Test-Path $lp)) { New-Item -Path $lp -ItemType File -Force | Out-Null }
        Start-Process -FilePath "notepad.exe" -ArgumentList $lp
    } catch { Write-CriticalLog ("OpenLog handler error: {0}" -f $_.Exception.Message) }
})

$script:MiSettings.Add_Click({ Show-SettingsWindow })
$script:MiExit.Add_Click({
    try {
        Write-Log "Exit clicked"
        if ($script:NotifyIcon) { try { $script:NotifyIcon.Visible = $false } catch {} }
        try { $script:Timer.Stop() } catch {}
        try { $script:IdleWatcher.Stop() } catch {}
        Save-State
        try { [System.Windows.Forms.Application]::Exit() } catch {}
    } catch { Write-CriticalLog ("Exit handler error: {0}" -f $_.Exception.Message) }
})

# ---------------- Form events ----------------
try {
    if ($script:Form -and ($script:Form -is [System.Windows.Forms.Form])) {
        $script:Form.Add_Shown({ try { $script:Form.Hide() } catch { Write-Log ("Form Shown handler failed: {0}" -f $_.Exception.Message) } })
        $script:Form.Load.Add({
            try {
                $script:Form.Hide()
                Update-Tooltip
                try { $script:IdleWatcher.Start() } catch { Write-Log ("IdleWatcher start failed: {0}" -f $_.Exception.Message) }
                Write-Log "AutoSavePAD started"
            } catch { Write-CriticalLog ("Form load error: {0}" -f $_.Exception.Message) }
        })
    } else { Write-Log "Warning: form not created; skipping form event registration" }
} catch { Write-CriticalLog ("Registering form handlers failed: {0}" -f $_.Exception.Message) }

# ---------------- Restore state and start ----------------
Load-State
try { if ($script:AutoSaveEnabled -and -not $script:InSleepMode) { $script:Timer.Interval = $script:Settings.IntervalMs; $script:Timer.Start(); Write-Log "Auto-Save resumed" } } catch {}

# ---------------- Message loop ----------------
try {
    if ($script:Form -and ($script:Form -is [System.Windows.Forms.Form])) {
        [System.Windows.Forms.Application]::Run($script:Form)
    } else {
        Write-Log "No WinForms form available; starting fallback wait loop to keep script alive"
        while ($true) { Start-Sleep -Seconds 60 }
    }
} catch { Write-CriticalLog ("Application loop error: {0}`n{1}" -f $_.Exception.Message, $_.Exception.StackTrace) }
