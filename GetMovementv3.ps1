<#
.SYNOPSIS
    Prevents the system from sleeping and the screen from going black by simulating human-like interactions, including randomized mouse movements, harmless key combinations (like Ctrl+C, Shift+F10), toggle key presses, and timing variations, while also using system calls to prevent screen timeout.

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
# Define the C# code for PowerHelper
$cSharpCode = @"
using System;
using System.Runtime.InteropServices;

public class PowerHelper
{
    // Prevent sleep state: ES_CONTINUOUS keeps the state active until explicitly cleared
    public const uint ES_CONTINUOUS = 0x80000000;
    // Prevent system sleep (keeps the system running)
    public const uint ES_SYSTEM_REQUIRED = 0x00000001;
    // Prevent display from turning off
    public const uint ES_DISPLAY_REQUIRED = 0x00000002;

    // Import the SetThreadExecutionState function from kernel32.dll
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@
# Add the C# type definition to the PowerShell session
Add-Type -TypeDefinition $cSharpCode -Language CSharp

# Combine flags to prevent both system sleep and display timeout, continuously
$powerStateFlags = [PowerHelper]::ES_DISPLAY_REQUIRED -bor [PowerHelper]::ES_SYSTEM_REQUIRED -bor [PowerHelper]::ES_CONTINUOUS
# Attempt to set the power state
$initialPowerState = [PowerHelper]::SetThreadExecutionState($powerStateFlags)

# Check if setting the power state was successful (non-zero return value indicates success)
if ($initialPowerState -ne 0) {
    Write-Host "System and Display sleep prevention activated." -ForegroundColor Green
} else {
    # Get the last Win32 error code if it failed
    $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    Write-Host "Failed to set power state. Error code: $errorCode" -ForegroundColor Red
    # Consider exiting if this fails, depending on requirements
    # exit 1
}

# --- Initialization ---
$announcementInterval = 6 # Number of loops before runtime announcement
# Try parsing the start and end times provided as parameters
try {
    $minTime = (Get-Date $minActiveTime).TimeOfDay
    $maxTime = (Get-Date $maxActiveTime).TimeOfDay
} catch {
    # Error if the time format is invalid
    Write-Error "Invalid time format for minActiveTime or maxActiveTime. Please use HH:mm format (e.g., '08:00', '22:30')."
    exit 1 # Exit the script if times are invalid
}
# Get the current date and time
$now = Get-Date

# Add the necessary .NET assembly for controlling the mouse cursor
Add-Type -AssemblyName System.Windows.Forms

# Create a COM object for simulating key presses
$WShell = New-Object -ComObject Wscript.Shell

# Define a set of relatively harmless keys/combinations to press randomly
# ^c   = Ctrl+C (Copy)
# +({F10}) = Shift+F10 (Context Menu / Right-click simulation)
$harmlessKeys = @("{SCROLLLOCK}", "{NUMLOCK}", "{CAPSLOCK}", "^c", "+({F10})")
# Define which keys from the list are toggle keys that should be pressed twice
$toggleKeys = @("{SCROLLLOCK}", "{NUMLOCK}", "{CAPSLOCK}")

# Format and display the start time and configuration
$date = Get-Date -Format "dddd MM/dd HH:mm"
Write-Host "Executing Enhanced Human-Like NoSleep routine." -ForegroundColor Cyan
Write-Host "Routine active between $minActiveTime and $maxActiveTime."
Write-Host "Base Sleep: $baseSleepSeconds seconds (+/- $sleepRandomnessSeconds seconds)."
Write-Host "Mouse Jiggle: +/- $mouseJiggleRange pixels."
Write-Host "Keys/Combinations to send: $($harmlessKeys -join ', ')" # Show which keys might be pressed
Write-Host "Start time: $date"
Write-Host "--- Starting Loop ---" -ForegroundColor Cyan

# Start a stopwatch to track the total runtime
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# --- Main Loop ---
$index = 0 # Loop counter
# Continue looping while the current time is within the specified active window
while ($minTime -le $now.TimeOfDay -and $maxTime -ge $now.TimeOfDay) {
    $index++
    $actionDescription = "" # String to describe actions taken in this iteration

    # Randomly decide which action(s) to take this cycle
    # 1: Press Key/Combination, 2: Move Mouse, 3: Do Both
    $actionType = Get-Random -Minimum 1 -Maximum 4 # Max is exclusive, so this gives 1, 2, or 3

    # --- Action: Simulate Key Press/Combination ---
    # Perform if action type is 1 or 3
    if ($actionType -eq 1 -or $actionType -eq 3) {
        # Only proceed if there are keys defined in the array
        if ($harmlessKeys.Count -gt 0) {
            # Select a random key/combination from the list
            $keyToSend = $harmlessKeys | Get-Random
            try {
                # Send the key press/combination
                $WShell.SendKeys($keyToSend)
                $actionDescription += "Sent '$keyToSend'. "
                # Wait a very short, random time to make it seem less instantaneous
                Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 150)

                # *** UPDATED LOGIC: Only press toggle keys twice ***
                if ($keyToSend -in $toggleKeys) {
                    # Send the key press again to toggle it back to its original state
                    $WShell.SendKeys($keyToSend)
                    $actionDescription += "(Toggled back). "
                }

            } catch {
                # Log a warning if sending the key fails
                Write-Warning "Failed to send key '$keyToSend': $($_.Exception.Message)"
            }
        } else {
             # Note if no keys are available to press
             $actionDescription += "No harmless keys defined to send. "
        }
    }

    # --- Action: Simulate Mouse Movement ---
    # Perform if action type is 2 or 3
    if ($actionType -eq 2 -or $actionType -eq 3) {
        try {
            # Get the current mouse cursor position
            $currentPosition = [System.Windows.Forms.Cursor]::Position
            # Calculate random offsets for X and Y within the specified range
            $offsetX = Get-Random -Minimum (-$mouseJiggleRange) -Maximum ($mouseJiggleRange + 1) # Max is exclusive
            $offsetY = Get-Random -Minimum (-$mouseJiggleRange) -Maximum ($mouseJiggleRange + 1)

            # Ensure we actually move if we intended to (offsetX/Y might both be 0 randomly)
            if ($offsetX -eq 0 -and $offsetY -eq 0) {
                # Try to force a move of at least 1 pixel in one direction if both offsets were 0
                 $offsetX = Get-Random -Minimum -1 -Maximum 2 # Result: -1, 0, or 1
                 if ($offsetX -eq 0) { $offsetY = (Get-Random -Minimum 0 -Maximum 2) * 2 - 1 } # Result: -1 or 1 if X is still 0
            }

            # Check if the final offset is still (0,0) - if so, skip moving
            if ($offsetX -eq 0 -and $offsetY -eq 0) {
                 $actionDescription += "Skipped zero mouse move. "
            } else {
                # Calculate the new X and Y coordinates
                $newX = $currentPosition.X + $offsetX
                $newY = $currentPosition.Y + $offsetY

                # Optional: Basic check to keep cursor roughly on the primary screen
                # $screenBounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
                # $newX = [Math]::Max($screenBounds.Left, [Math]::Min($screenBounds.Right - 1, $newX))
                # $newY = [Math]::Max($screenBounds.Top, [Math]::Min($screenBounds.Bottom - 1, $newY))

                # Create a new Point object for the target position
                $newPosition = [System.Drawing.Point]::new($newX, $newY)

                # Move the cursor to the new position
                [System.Windows.Forms.Cursor]::Position = $newPosition
                # Wait a short, random time before moving back - more human-like than an instant snap
                Start-Sleep -Milliseconds (Get-Random -Minimum 80 -Maximum 250)
                # Move the cursor back to its original position
                [System.Windows.Forms.Cursor]::Position = $currentPosition
                $actionDescription += "Jiggled mouse ($offsetX, $offsetY)."
            }
        } catch {
            # Log a warning if moving the mouse fails
            Write-Warning "Failed to move mouse: $($_.Exception.Message)"
        }
    }

    # Log the action(s) taken in this iteration if any occurred
    if ($actionDescription.Trim().Length -gt 0) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Action: $actionDescription" -ForegroundColor Gray
    } else {
        # Log if no specific input simulation happened this cycle
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Action: No input simulated this cycle." -ForegroundColor DarkGray
    }


    # --- Randomized Sleep Interval ---
    # Calculate the minimum sleep duration, ensuring it's at least 5 seconds
    $minSleep = [Math]::Max(5, $baseSleepSeconds - $sleepRandomnessSeconds)
    # Calculate the maximum sleep duration
    $maxSleep = $baseSleepSeconds + $sleepRandomnessSeconds
    # Get a random sleep duration within the calculated range
    $currentSleepSeconds = Get-Random -Minimum $minSleep -Maximum ($maxSleep + 1) # Max is exclusive for Get-Random

    # Verbose logging for sleep duration (optional)
    # Write-Host "Sleeping for $currentSleepSeconds seconds..." -ForegroundColor DarkGray
    # Pause the script execution for the calculated duration
    Start-Sleep -Seconds $currentSleepSeconds

    # Update the current time for the next loop check
    $now = Get-Date

    # Announce total runtime periodically based on the announcement interval
    if ($stopwatch.IsRunning -and ($index % $announcementInterval) -eq 0) {
        Write-Host "--- Runtime: $($stopwatch.Elapsed.ToString('dd\.hh\:mm\:ss')) --- Current Time: $($now.ToString('HH:mm:ss')) ---" -ForegroundColor Yellow
    }
} # End of the main while loop

# --- Completion ---
Write-Host "--- Exiting Loop ---" -ForegroundColor Cyan
# Reset the thread execution state to allow system/display to sleep normally again
# Passing ES_CONTINUOUS alone effectively clears the previous ES_SYSTEM_REQUIRED and ES_DISPLAY_REQUIRED flags
$finalPowerState = [PowerHelper]::SetThreadExecutionState([PowerHelper]::ES_CONTINUOUS)
# Check if the state was reset successfully (0 means the previous state was cleared)
if ($finalPowerState -eq 0) {
     Write-Host "System and Display sleep prevention deactivated." -ForegroundColor Green
} else {
     # This is unexpected, the state might still be preventing sleep.
     Write-Warning "Could not reset thread execution state. System/Display sleep might still be prevented."
}

# Stop the stopwatch
$stopwatch.Stop()
# Log the completion time and total runtime
Write-Host "NoSleep routine completed (or time window ended) at: $(Get-Date -Format "dddd MM/dd HH:mm (K)")"
Write-Host "Total runtime: $($stopwatch.Elapsed.ToString('dd\.hh\:mm\:ss'))"
