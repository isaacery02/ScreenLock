<#
.SYNOPSIS
    Prevents the system from sleeping and the screen from going black by simulating more human-like interactions, including randomized mouse movements, key presses, and timing variations, while also using system calls to prevent screen timeout.

.PARAMETER baseSleepSeconds
    Specifies the base interval (in seconds) between iterations. The actual interval will vary randomly around this. Default is 30 seconds.

.PARAMETER sleepRandomnessSeconds
    Specifies the maximum amount (in seconds) to randomly add or subtract from the baseSleepSeconds. Default is 10 seconds. (e.g., 30 +/- 10 means sleep between 20 and 40 seconds).

.PARAMETER mouseJiggleRange
    Specifies the maximum number of pixels (positive or negative) to move the mouse randomly in the X and Y directions. Default is 5 pixels.

.PARAMETER minActiveTime
    Specifies the earliest time (HH:mm format) the script should be active. Default is '08:00'.

.PARAMETER maxActiveTime
    Specifies the latest time (HH:mm format) the script should stop being active. Default is '22:00'.
#>

param(
    [int]$baseSleepSeconds = 30,          # Base interval in seconds
    [int]$sleepRandomnessSeconds = 10,   # Random variation +/- in seconds
    [int]$mouseJiggleRange = 5,          # Max pixels +/- for mouse jiggle
    [string]$minActiveTime = '08:00',     # Start time for the routine
    [string]$maxActiveTime = '22:00'      # End time for the routine
)

# --- Input Validation ---
if ($baseSleepSeconds -le $sleepRandomnessSeconds) {
    Write-Warning "baseSleepSeconds ($baseSleepSeconds) should ideally be greater than sleepRandomnessSeconds ($sleepRandomnessSeconds) to avoid very short or zero sleep intervals. Adjusting minimum sleep to 5 seconds."
}
if ($mouseJiggleRange -lt 1) {
    Write-Warning "mouseJiggleRange must be at least 1. Setting to 1."
    $mouseJiggleRange = 1
}

# --- Setup Power Management ---
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class PowerHelper
{
    // Prevent sleep state
    public const uint ES_CONTINUOUS = 0x80000000;
    public const uint ES_SYSTEM_REQUIRED = 0x00000001;
    // Prevent display from turning off
    public const uint ES_DISPLAY_REQUIRED = 0x00000002;

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@ -Language CSharp

# Prevent the display from turning off and system from sleeping
$powerStateFlags = [PowerHelper]::ES_DISPLAY_REQUIRED -bor [PowerHelper]::ES_SYSTEM_REQUIRED -bor [PowerHelper]::ES_CONTINUOUS
$initialPowerState = [PowerHelper]::SetThreadExecutionState($powerStateFlags)

if ($initialPowerState -ne 0) {
    Write-Host "System and Display sleep prevention activated." -ForegroundColor Green
} else {
    Write-Host "Failed to set power state. Error code: $([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())" -ForegroundColor Red
    # Consider exiting if this fails, depending on requirements
}

# --- Initialization ---
$announcementInterval = 6 # Number of loops before runtime announcement (adjusted for potentially shorter sleep)
try {
    $minTime = (Get-Date $minActiveTime).TimeOfDay
    $maxTime = (Get-Date $maxActiveTime).TimeOfDay
} catch {
    Write-Error "Invalid time format for minActiveTime or maxActiveTime. Please use HH:mm format (e.g., '08:00', '22:30')."
    exit 1
}
$now = Get-Date

# Add necessary assembly for mouse control
Add-Type -AssemblyName System.Windows.Forms

# Initialize WScript Shell for key simulation
$WShell = New-Object -ComObject Wscript.Shell

# Define a set of relatively harmless keys to press randomly
$harmlessKeys = @("{SCROLLLOCK}", "{NUMLOCK}", "{CAPSLOCK}", "{SHIFT}", "{CTRL}", "{ALT}") # Shift/Ctrl/Alt alone do nothing in most apps

# Format and display the start time
$date = Get-Date -Format "dddd MM/dd HH:mm"
Write-Host "Executing Enhanced Human-Like NoSleep routine."
Write-Host "Routine active between $minActiveTime and $maxActiveTime."
Write-Host "Base Sleep: $baseSleepSeconds seconds (+/- $sleepRandomnessSeconds seconds)."
Write-Host "Mouse Jiggle: +/- $mouseJiggleRange pixels."
Write-Host "Start time: $date"
Write-Host "<start3" -ForegroundColor Cyan

# Stopwatch for runtime announcements
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# --- Main Loop ---
$index = 0
while ($minTime -le $now.TimeOfDay -and $maxTime -ge $now.TimeOfDay) {
    $index++
    $actionDescription = ""

    # Randomly decide which action(s) to take
    $actionType = Get-Random -Minimum 1 -Maximum 5 # 1: Key, 2: Mouse, 3: Both, 4: Key (different timing), 5: Mouse (different timing)

    # --- Action: Simulate Key Press ---
    if ($actionType -eq 1 -or $actionType -eq 3 -or $actionType -eq 4) {
        $keyToPress = $harmlessKeys | Get-Random
        try {
            $WShell.SendKeys($keyToPress)
            $actionDescription += "Pressed '$keyToPress'. "
            # Small delay after key press to seem less instant
            Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 150)
            # If it's a toggle key, press it again to revert state (optional, but often desired)
            if ($keyToPress -in @("{SCROLLLOCK}", "{NUMLOCK}", "{CAPSLOCK}")) {
                $WShell.SendKeys($keyToPress)
                # $actionDescription += "(Toggled back). " # Uncomment if you want this logged
            }
        } catch {
            Write-Warning "Failed to send key '$keyToPress': $($_.Exception.Message)"
        }
    }

    # --- Action: Simulate Mouse Movement ---
    if ($actionType -eq 2 -or $actionType -eq 3 -or $actionType -eq 5) {
        try {
            $currentPosition = [System.Windows.Forms.Cursor]::Position
            $offsetX = Get-Random -Minimum (-$mouseJiggleRange) -Maximum ($mouseJiggleRange + 1) # Max is exclusive
            $offsetY = Get-Random -Minimum (-$mouseJiggleRange) -Maximum ($mouseJiggleRange + 1)

            # Ensure we actually move if we intended to (offsetX/Y might both be 0)
            if ($offsetX -eq 0 -and $offsetY -eq 0) {
                $offsetX = (Get-Random -Minimum 0 -Maximum 2) * 2 - 1 # Get -1 or 1
            }

            $newX = $currentPosition.X + $offsetX
            $newY = $currentPosition.Y + $offsetY

            # Basic check to keep cursor roughly on screen (optional but safer)
            # $screenBounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
            # $newX = [Math]::Max($screenBounds.Left, [Math]::Min($screenBounds.Right - 1, $newX))
            # $newY = [Math]::Max($screenBounds.Top, [Math]::Min($screenBounds.Bottom - 1, $newY))

            $newPosition = [System.Drawing.Point]::new($newX, $newY)

            [System.Windows.Forms.Cursor]::Position = $newPosition
            # Wait a tiny bit before moving back - more human than instant snap back
            Start-Sleep -Milliseconds (Get-Random -Minimum 80 -Maximum 250)
            [System.Windows.Forms.Cursor]::Position = $currentPosition # Move back to original position
            $actionDescription += "Jiggled mouse ($offsetX, $offsetY)."
        } catch {
            Write-Warning "Failed to move mouse: $($_.Exception.Message)"
        }
    }

    # Log the action taken this iteration
     Write-Host "$(Get-Date -Format 'HH:mm:ss') - Action: $actionDescription" -ForegroundColor Gray

    # --- Randomized Sleep Interval ---
    $minSleep = [Math]::Max(5, $baseSleepSeconds - $sleepRandomnessSeconds) # Ensure minimum sleep > 0 (e.g., at least 5 sec)
    $maxSleep = $baseSleepSeconds + $sleepRandomnessSeconds
    $currentSleepSeconds = Get-Random -Minimum $minSleep -Maximum ($maxSleep + 1) # Max is exclusive for Get-Random

    #Write-Host "Sleeping for $currentSleepSeconds seconds..." # Verbose logging if needed
    Start-Sleep -Seconds $currentSleepSeconds

    # Update the current time
    $now = Get-Date

    # Announce runtime at the specified interval
    if ($stopwatch.IsRunning -and ($index % $announcementInterval) -eq 0) {
        Write-Host "--- Runtime: $($stopwatch.Elapsed.ToString('dd\.hh\:mm\:ss')) --- Current Time: $($now.ToString('HH:mm:ss')) ---" -ForegroundColor Yellow
    }
}

# --- Completion ---
# Reset the thread execution state to allow system/display to sleep normally again
$finalPowerState = [PowerHelper]::SetThreadExecutionState([PowerHelper]::ES_CONTINUOUS)
if ($finalPowerState -eq 0) {
     # This means the previous state setting was cleared successfully.
     Write-Host "System and Display sleep prevention deactivated." -ForegroundColor Green
} else {
     # This is unexpected, means the state might still be set.
     Write-Warning "Could not reset thread execution state. System/Display sleep might still be prevented."
}

$stopwatch.Stop()
Write-Host "NoSleep routine completed (or time window ended) at: $(Get-Date -Format "dddd MM/dd HH:mm (K)")"
Write-Host "Total runtime: $($stopwatch.Elapsed.ToString('dd\.hh\:mm\:ss'))"