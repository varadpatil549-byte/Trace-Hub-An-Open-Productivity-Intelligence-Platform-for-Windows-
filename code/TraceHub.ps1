Add-Type -AssemblyName System.Windows.Forms

# ==========================
# WINDOWS API
# ==========================

if (-not ("Win32App" -as [type])) {

$signature = @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class Win32App
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern int GetWindowText(
        IntPtr hWnd,
        StringBuilder text,
        int count
    );

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId
    );

    [StructLayout(LayoutKind.Sequential)]
    public struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }

    [DllImport("user32.dll")]
    public static extern bool GetLastInputInfo(
        ref LASTINPUTINFO plii
    );
}
"@

Add-Type -TypeDefinition $signature
}

# ==========================
# SETTINGS
# ==========================

$folderPath =
"$env:USERPROFILE\Documents\Trace HUB CSV"

if (!(Test-Path $folderPath)) {
    New-Item `
    -Path $folderPath `
    -ItemType Directory | Out-Null
}

$csvPath =
Join-Path $folderPath "Trace hub.csv"

$username = $env:USERNAME
$computer = $env:COMPUTERNAME

# Create CSV Header
if (!(Test-Path $csvPath)) {

"Timestamp,Username,Computer,App,WindowTitle,Status,SessionDurationSec" |
Out-File `
-FilePath $csvPath `
-Encoding UTF8
}

# ==========================
# TRACKING VARIABLES
# ==========================

$lastTimestamp = Get-Date

# ==========================
# MAIN LOOP
# ==========================

while ($true) {

    $currentTime = Get-Date

    # ----------------------
    # WINDOWS IDLE DETECTION
    # ----------------------

    $lastInput =
    New-Object Win32App+LASTINPUTINFO

    $lastInput.cbSize =
    [System.Runtime.InteropServices.Marshal]::SizeOf($lastInput)

    [Win32App]::GetLastInputInfo(
        [ref]$lastInput
    ) | Out-Null

    $idleMilliseconds =
    [Environment]::TickCount -
    $lastInput.dwTime

    $idleSeconds =
    [Math]::Floor(
        $idleMilliseconds / 1000
    )

    if ($idleSeconds -ge 60) {
        $status = "IDLE"
    }
    else {
        $status = "ACTIVE"
    }

    # ----------------------
    # ACTIVE WINDOW
    # ----------------------

    $handle =
    [Win32App]::GetForegroundWindow()

    $title =
    New-Object System.Text.StringBuilder 256

    [Win32App]::GetWindowText(
        $handle,
        $title,
        $title.Capacity
    ) | Out-Null

    $windowTitle =
    $title.ToString()

    if (
        [string]::IsNullOrWhiteSpace(
            $windowTitle
        )
    ) {
        $windowTitle = "Unknown"
    }

    # ----------------------
    # PROCESS NAME
    # ----------------------

    $procId = 0

    [Win32App]::GetWindowThreadProcessId(
        $handle,
        [ref]$procId
    ) | Out-Null

    $processName = "Unknown"

    try {
        $processName =
        (Get-Process -Id $procId).ProcessName
    }
    catch {}

    # ----------------------
    # SESSION TIME
    # ----------------------

    $sessionSeconds =
    [Math]::Round(
        (
            $currentTime -
            $lastTimestamp
        ).TotalSeconds,
        0
    )

    # ----------------------
    # EXPORT CSV
    # ----------------------

    [PSCustomObject]@{

        Timestamp =
        $currentTime.ToString(
            "yyyy-MM-dd HH:mm:ss"
        )

        Username =
        $username

        Computer =
        $computer

        App =
        $processName

        WindowTitle =
        $windowTitle

        Status =
        $status

        SessionDurationSec =
        $sessionSeconds

    } |
    Export-Csv `
    -Path $csvPath `
    -Append `
    -NoTypeInformation

    $lastTimestamp =
    $currentTime

    Start-Sleep -Seconds 5
}