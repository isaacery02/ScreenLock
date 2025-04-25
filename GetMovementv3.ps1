<#
.SYNOPSIS
    Prevents the system from sleeping and the screen from going black by simulating human-like interactions (mouse/keyboard), timing variations, and using system calls (via Add-Type and a helper function to find the type). Includes ASCII art display.

.PARAMETER baseSleepSeconds
    Specifies the base interval (in seconds) between main loop iterations (mouse/key simulation). Default is 30 seconds.

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

# --- Helper function to find a type in loaded assemblies ---
# This is necessary because [System.Type]::GetType() sometimes fails to find types added via Add-Type.
function Find-Type {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TypeName
    )

    # Iterate through all assemblies currently loaded in the application domain.
    foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
        # Try to get the type from the current assembly.
        $type = $assembly.GetType($TypeName, $false) # $false = do not throw error if not found
        if ($type) {
            # If found, return the type object and exit the function.
            return $type
        }
    }
    # If the type wasn't found in any assembly, return null.
    return $null
}


# --- Input Validation ---
# Warn if base sleep is less than or equal to randomness, could lead to zero/negative sleep.
if ($baseSleepSeconds -le $sleepRandomnessSeconds) {
    Write-Warning "baseSleepSeconds ($baseSleepSeconds) should ideally be greater than sleepRandomnessSeconds ($sleepRandomnessSeconds) to avoid very short or zero sleep intervals. Adjusting minimum sleep to 5 seconds."
}
# Ensure mouse jiggle range is at least 1 pixel.
if ($mouseJiggleRange -lt 1) {
    Write-Warning "mouseJiggleRange must be at least 1. Setting to 1."
    $mouseJiggleRange = 1
}

# --- Setup Power Management ---
# Define the C# code WITH a namespace and the class name that worked in the test.
$cSharpCode = @"
using System;
using System.Runtime.InteropServices;

namespace NoSleepUtils
{
    public class MyPowerUtil
    {
        // Flags for SetThreadExecutionState:
        public const uint ES_CONTINUOUS = 0x80000000;
        public const uint ES_SYSTEM_REQUIRED = 0x00000001;
        public const uint ES_DISPLAY_REQUIRED = 0x00000002;

        // Import the SetThreadExecutionState function from kernel32.dll.
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern uint SetThreadExecutionState(uint esFlags);
    }
}
"@
# Define the fully qualified name we expect
$powerUtilTypeName = 'NoSleepUtils.MyPowerUtil'

# Check if the type exists using our helper function BEFORE trying to add
$powerUtilType = Find-Type -TypeName $powerUtilTypeName
if (-not $powerUtilType) {
    Write-Verbose "$powerUtilTypeName type not found, attempting Add-Type..."
    try {
        # Attempt to add the type. Use -ErrorAction Stop to catch compilation errors.
        Add-Type -TypeDefinition $cSharpCode -Language CSharp -ErrorAction Stop
        # Verify again immediately after Add-Type using the helper function.
        $powerUtilType = Find-Type -TypeName $powerUtilTypeName
        if ($powerUtilType) {
            Write-Host "Successfully added type $powerUtilTypeName." -ForegroundColor Green
        } else {
            # This case should be rare now, but indicates Add-Type completed but type still not findable.
            Write-Warning "Add-Type command completed but type '$powerUtilTypeName' still not found by Find-Type function."
        }
    } catch {
        # Catch any errors during Add-Type compilation.
        Write-Error "Error occurred during Add-Type for ${powerUtilTypeName}: $($_.ToString())"
        # $powerUtilType will remain $null if Add-Type failed.
    }
} else {
    Write-Verbose "$powerUtilTypeName type already exists in this session."
}


# --- Activate Power Management (if type exists) ---
# Combine the flags to request that both the system and display stay active continuously.
# Only attempt if the type was successfully found or added.
if ($powerUtilType) {
    # Use the type object found by Find-Type to access static members.
    $powerStateFlags = $powerUtilType::ES_DISPLAY_REQUIRED -bor $powerUtilType::ES_SYSTEM_REQUIRED -bor $powerUtilType::ES_CONTINUOUS
    # Attempt to set the required execution state.
    $initialPowerState = $powerUtilType::SetThreadExecutionState($powerStateFlags)

    # Check if the call was successful. A non-zero return value indicates success.
    if ($initialPowerState -ne 0) {
        Write-Host "System and Display sleep prevention activated using $powerUtilTypeName." -ForegroundColor Green
    } else {
        # If the call failed, retrieve and display the Windows error code.
        $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Host "Failed to set power state using $powerUtilTypeName. Error code: $errorCode" -ForegroundColor Red
    }
} else {
     Write-Warning "$powerUtilTypeName type not loaded. Cannot prevent system/display sleep via API."
}


# --- Initialization ---
$announcementInterval = 6 # How many loops before printing runtime status.
# Try parsing the start and end times provided as parameters using Get-Date.
try {
    $minTime = (Get-Date $minActiveTime).TimeOfDay
    $maxTime = (Get-Date $maxActiveTime).TimeOfDay
} catch {
    # If parsing fails, output an error.
    Write-Error "Invalid time format for minActiveTime or maxActiveTime. Please use HH:mm format (e.g., '08:00', '22:30')."
    # Attempt to set default times if parsing failed.
    Write-Warning "Using default active times (08:00-22:00) due to parsing error."
    $minTime = (Get-Date '08:00').TimeOfDay
    $maxTime = (Get-Date '22:00').TimeOfDay
}
# Get the current date and time.
$now = Get-Date

# Add the necessary .NET assembly for controlling the mouse cursor (System.Windows.Forms).
# Use SilentlyContinue for this too, as it can also report "already loaded".
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue


# Create a COM object for simulating key presses (Wscript.Shell).
# Check if it already exists (e.g., if dot-sourced).
if (-not $Global:WShell) { # Check global scope in case it was dot-sourced
    try {
        # Create the COM object if it doesn't exist. Store globally for potential reuse.
        $Global:WShell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
    } catch {
         # Output an error if the COM object cannot be created.
         Write-Error "Failed to create Wscript.Shell COM object: $($_.Exception.Message)"
         $Global:WShell = $null # Ensure it's null if creation failed
    }
}


# Define a list of relatively harmless keys and key combinations to send randomly.
# {F15} - Function key 15, rarely used, unlikely to do anything.
# +({F10}) = Shift+F10 (Context Menu / Right-click simulation)
$harmlessKeys = @("{SCROLLLOCK}", "{NUMLOCK}", "{CAPSLOCK}", "{F15}", "+({F10})")
# Define which keys from the list are toggle keys that should be pressed twice to revert state.
$toggleKeys = @("{SCROLLLOCK}", "{NUMLOCK}", "{CAPSLOCK}")

# --- ASCII Art Definition ---
$coffeeArt = @(
    '       ( (',
    '        ) )',
    '     .........',
    '     |       |]',
    '     \       /',
    '      `-----`'
)
$artWidth = ($coffeeArt | Measure-Object -Maximum -Property Length).Maximum
$artHeight = $coffeeArt.Count
$lastArtCursorPosition = $null # To store where the art was last drawn

# --- Function to Draw ASCII Art Slowly ---
Function Draw-AsciiArtSlowly {
    param(
        [string[]]$ArtLines,
        [int]$DurationSeconds,
        [int]$ArtHeight,
        [int]$ArtWidth
    )

    # Minimum duration to attempt drawing, otherwise just sleep.
    $minDrawDuration = 1
    if ($DurationSeconds -lt $minDrawDuration) {
        Start-Sleep -Seconds $DurationSeconds
        return
    }

    # Calculate delay per line in milliseconds. Ensure minimum 1ms delay.
    $lines = $ArtLines.Count
    $delayMs = [Math]::Max(1, [int](($DurationSeconds * 1000) / $lines))

    # Get current cursor position ONCE before loop, if not already stored
    if ($script:lastArtCursorPosition -eq $null) {
        try {
            $script:lastArtCursorPosition = $Host.UI.RawUI.CursorPosition
            # If starting near bottom, scroll up slightly to make space
            if (($script:lastArtCursorPosition.Y + $ArtHeight) -ge $Host.UI.RawUI.WindowSize.Height) {
                 $script:lastArtCursorPosition = New-Object System.Management.Automation.Host.Coordinates (0, ($Host.UI.RawUI.CursorPosition.Y - $ArtHeight -1))
                 if ($script:lastArtCursorPosition.Y -lt 0) {$script:lastArtCursorPosition.Y = 0} # Don't go negative
                 $Host.UI.RawUI.CursorPosition = $script:lastArtCursorPosition
            }
        } catch {
             Write-Warning "Could not get/set initial cursor position for drawing."
             $script:lastArtCursorPosition = New-Object System.Management.Automation.Host.Coordinates (0,0) # Default to top-left
        }
    }

    # Store starting position for this draw cycle
    $startPos = $script:lastArtCursorPosition

    # Loop through each line of the art
    for ($i = 0; $i -lt $lines; $i++) {
        # Set cursor position for the current line
        $currentPos = New-Object System.Management.Automation.Host.Coordinates $startPos.X, ($startPos.Y + $i)
        # Avoid error if console resized making position invalid
        try { $Host.UI.RawUI.CursorPosition = $currentPos } catch { Write-Warning "Could not set cursor position (console resized?)."; break }


        # Write the line, padding with spaces to overwrite previous longer lines/artifacts
        Write-Host ($ArtLines[$i].PadRight($ArtWidth + 2)) # Pad slightly wider

        # Pause for the calculated delay
        Start-Sleep -Milliseconds $delayMs
    }
     # Leave cursor after the art
     try { $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates $startPos.X, ($startPos.Y + $lines) } catch {}
}


# --- Display Initial Configuration ---
$date = Get-Date -Format "dddd MM/dd HH:mm"
Write-Host "Executing Enhanced Human-Like NoSleep routine." -ForegroundColor Cyan
Write-Host "Routine active between $minActiveTime and $maxActiveTime."
Write-Host "Base Sleep: $baseSleepSeconds seconds (+/- $sleepRandomnessSeconds seconds)."
Write-Host "Mouse Jiggle: +/- $mouseJiggleRange pixels."
Write-Host "Keys/Combinations to send: $($harmlessKeys -join ', ')"
Write-Host "Start time: $date"
Write-Host "--- Starting Loop ---" -ForegroundColor Cyan

# Start a stopwatch to track the total runtime of the main loop.
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# --- Main Loop ---
$index = 0 # Initialize loop counter.
# Continue looping as long as the current time is within the specified active window.
while ($minTime -le $now.TimeOfDay -and $maxTime -ge $now.TimeOfDay) {
    $index++ # Increment loop counter.
    $actionDescription = "" # Initialize string to describe actions taken in this iteration.

    # --- Check for WShell object before attempting actions ---
    if (-not $Global:WShell) {
        Write-Warning "Wscript.Shell object not available. Skipping key/mouse simulation."
        # Still need to wait
        $currentSleepSeconds = Get-Random -Minimum ([Math]::Max(5, $baseSleepSeconds - $sleepRandomnessSeconds)) -Maximum ($baseSleepSeconds + $sleepRandomnessSeconds + 1)
        Draw-AsciiArtSlowly -ArtLines $coffeeArt -DurationSeconds $currentSleepSeconds -ArtHeight $artHeight -ArtWidth $artWidth
        $now = Get-Date # Update time
        continue # Skip to next loop iteration
    }

    # Randomly decide which action(s) to take this cycle.
    # 1: Press Key/Combination, 2: Move Mouse, 3: Do Both
    $actionType = Get-Random -Minimum 1 -Maximum 4 # Max is exclusive, so this gives 1, 2, or 3.

    # --- Action: Simulate Key Press/Combination ---
    # Perform if action type is 1 (Key) or 3 (Both).
    if ($actionType -eq 1 -or $actionType -eq 3) {
        # Only proceed if there are keys defined in the array.
        if ($harmlessKeys.Count -gt 0) {
            # Select a random key or combination from the list.
            $keyToSend = $harmlessKeys | Get-Random
            try {
                # Send the selected key press/combination using Wscript.Shell.
                $Global:WShell.SendKeys($keyToSend)
                $actionDescription += "Sent '$keyToSend'. "
                # Wait a very short, random time to make the action seem less instantaneous.
                Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 150)

                # Check if the sent key is one of the toggle keys.
                if ($keyToSend -in $toggleKeys) {
                    # If it's a toggle key, send it again to revert its state.
                    $Global:WShell.SendKeys($keyToSend)
                    $actionDescription += "(Toggled back). "
                }

            } catch {
                # Log a warning if sending the key fails.
                Write-Warning "Failed to send key '$keyToSend': $($_.Exception.Message)"
            }
        } else {
             # Note in the log if no keys are defined.
             $actionDescription += "No harmless keys defined to send. "
        }
    }

    # --- Action: Simulate Mouse Movement ---
    # Perform if action type is 2 (Mouse) or 3 (Both).
    if ($actionType -eq 2 -or $actionType -eq 3) {
        try {
            # Get the current mouse cursor position using System.Windows.Forms.
            $currentPosition = [System.Windows.Forms.Cursor]::Position
            # Calculate random offsets for X and Y within the specified range.
            $offsetX = Get-Random -Minimum (-$mouseJiggleRange) -Maximum ($mouseJiggleRange + 1)
            $offsetY = Get-Random -Minimum (-$mouseJiggleRange) -Maximum ($mouseJiggleRange + 1)

            # Ensure we actually move if we intended to.
            if ($offsetX -eq 0 -and $offsetY -eq 0) {
                 $offsetX = Get-Random -Minimum -1 -Maximum 2 # Result: -1, 0, or 1
                 if ($offsetX -eq 0) { $offsetY = (Get-Random -Minimum 0 -Maximum 2) * 2 - 1 } # Result: -1 or 1
            }

            # Check if the final calculated offset is still (0,0).
            if ($offsetX -eq 0 -and $offsetY -eq 0) {
                 $actionDescription += "Skipped zero mouse move. "
            } else {
                # Calculate the new X and Y coordinates.
                $newX = $currentPosition.X + $offsetX
                $newY = $currentPosition.Y + $offsetY

                # Create a new System.Drawing.Point object for the target position.
                $newPosition = [System.Drawing.Point]::new($newX, $newY)

                # Move the cursor to the new calculated position.
                [System.Windows.Forms.Cursor]::Position = $newPosition
                # Wait a short, random time before moving back.
                Start-Sleep -Milliseconds (Get-Random -Minimum 80 -Maximum 250)
                # Move the cursor back to its original position.
                [System.Windows.Forms.Cursor]::Position = $currentPosition
                $actionDescription += "Jiggled mouse ($offsetX, $offsetY)."
            }
        } catch {
            # Log a warning if moving the mouse fails.
            Write-Warning "Failed to move mouse: $($_.Exception.Message)"
        }
    }

    # Log the action(s) taken in this iteration if any specific action occurred.
    if ($actionDescription.Trim().Length -gt 0) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Action: $actionDescription" -ForegroundColor Gray
    } else {
        # Log if no specific input simulation happened this cycle.
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Action: No input simulated this cycle." -ForegroundColor DarkGray
    }


    # --- Randomized Wait with ASCII Art ---
    # Calculate the minimum sleep duration.
    $minSleep = [Math]::Max(5, $baseSleepSeconds - $sleepRandomnessSeconds)
    # Calculate the maximum sleep duration.
    $maxSleep = $baseSleepSeconds + $sleepRandomnessSeconds
    # Get a random sleep duration.
    $currentSleepSeconds = Get-Random -Minimum $minSleep -Maximum ($maxSleep + 1)

    # Call the function to draw the art slowly over the calculated sleep duration.
    Draw-AsciiArtSlowly -ArtLines $coffeeArt -DurationSeconds $currentSleepSeconds -ArtHeight $artHeight -ArtWidth $artWidth


    # Update the current time variable for the next loop's condition check.
    $now = Get-Date

    # Announce total runtime periodically based on the announcement interval.
    if ($stopwatch.IsRunning -and ($index % $announcementInterval) -eq 0) {
        # Store cursor, move up to print status, then restore cursor for art
        try {
            $statusCursorPos = $Host.UI.RawUI.CursorPosition
            # Try to print status above the art area if possible
            $statusLineY = if ($script:lastArtCursorPosition) { [Math]::Max(0, $script:lastArtCursorPosition.Y - 1) } else { $statusCursorPos.Y }
             $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $statusLineY
             # Clear the status line before writing
             Write-Host (" " * $Host.UI.RawUI.WindowSize.Width)
             $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $statusLineY
             # Write the status
             Write-Host "--- Runtime: $($stopwatch.Elapsed.ToString('dd\.hh\:mm\:ss')) --- Current Time: $($now.ToString('HH:mm:ss')) ---" -ForegroundColor Yellow
             # Restore cursor to where it was before printing status (likely after art)
             $Host.UI.RawUI.CursorPosition = $statusCursorPos
        } catch { Write-Warning "Could not print status line (console resized?)."}

        # Reset art position for next draw cycle so it doesn't drift down after status line print
        $script:lastArtCursorPosition = $null
    }
} # End of the main while loop

# --- Completion ---
Write-Host # Add a newline for clarity after the loop ends
Write-Host "--- Exiting Loop (Time window ended or script interrupted) ---" -ForegroundColor Cyan

# Reset the thread execution state to allow the system and display to sleep normally again.
# Only attempt if the type was successfully added/found earlier.
$powerUtilType = Find-Type -TypeName $powerUtilTypeName # Re-check just in case session state changed
if ($powerUtilType) {
    try {
        # Calling SetThreadExecutionState with only ES_CONTINUOUS clears previous requirements.
        # Use the type object found by Find-Type to access static members.
        $finalPowerState = $powerUtilType::SetThreadExecutionState($powerUtilType::ES_CONTINUOUS)
        # Check if the state was reset successfully.
        if ($finalPowerState -eq 0) {
             Write-Host "System and Display sleep prevention deactivated using $powerUtilTypeName." -ForegroundColor Green
        } else {
             Write-Warning "Could not definitively reset thread execution state using $powerUtilTypeName (Return code: $finalPowerState). System/Display sleep might still be prevented."
        }
    } catch {
         Write-Error "Error calling SetThreadExecutionState during cleanup: $($_.Exception.Message)"
    }

} else {
    Write-Warning "$powerUtilTypeName type was not available during cleanup. Could not reset power state."
}


# Stop the stopwatch.
$stopwatch.Stop()
# Log the completion time and the total runtime recorded by the stopwatch.
Write-Host "NoSleep routine completed at: $(Get-Date -Format "dddd MM/dd HH:mm (K)")"
Write-Host "Total active runtime: $($stopwatch.Elapsed.ToString('dd\.hh\:mm\:ss'))"

# Optional: Clean up the COM object.
# If ($Global:WShell) {
#    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Global:WShell) | Out-Null
#    Remove-Variable -Name WShell -Scope Global -ErrorAction SilentlyContinue
#    Write-Verbose "Wscript.Shell COM object released."
# }