# --- Minimal Test Script ---

# Define the C# code exactly as before
$cSharpCode = @"
using System;
using System.Runtime.InteropServices;

// Class to wrap the SetThreadExecutionState API call.
public class PowerHelper
{
    // Flags for SetThreadExecutionState:
    public const uint ES_CONTINUOUS = 0x80000000;
    public const uint ES_SYSTEM_REQUIRED = 0x00000001;
    public const uint ES_DISPLAY_REQUIRED = 0x00000002;

    // Import the SetThreadExecutionState function from kernel32.dll.
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@

# Check if type exists BEFORE trying to add
if ([System.Type]::GetType('PowerHelper')) {
    Write-Host "PowerHelper type ALREADY EXISTS before Add-Type was called." -ForegroundColor Yellow
} else {
    Write-Host "PowerHelper type does not exist yet. Attempting Add-Type..."
    try {
        # Attempt to add the type - NO SilentlyContinue, use Stop to force catch on error
        Add-Type -TypeDefinition $cSharpCode -Language CSharp -ErrorAction Stop
        Write-Host "Add-Type command completed without throwing an error." -ForegroundColor Green
        # Verify again after Add-Type
         if ([System.Type]::GetType('PowerHelper')) {
             Write-Host "PowerHelper type successfully added/found AFTER Add-Type." -ForegroundColor Green
         } else {
             Write-Warning "Add-Type completed but PowerHelper type still not found!"
         }
    } catch {
        # Catch any error during Add-Type
        Write-Error "Error occurred during Add-Type: $($_.ToString())" # Show full error
    }
}

# Final check
if ([System.Type]::GetType('PowerHelper')) {
    Write-Host "Final check: PowerHelper type exists."
} else {
     Write-Host "Final check: PowerHelper type DOES NOT exist."
}