# Enhanced Human-Like NoSleep PowerShell Script

## Overview

This PowerShell script is designed to prevent your Windows system from entering sleep mode or turning off the display during specified hours. It achieves this by simulating subtle, human-like user interactions, including randomized mouse movements and occasional key presses, alongside using Windows API calls to explicitly request the system and display remain active.

This is useful for scenarios where you need your computer to stay awake (e.g., running long tasks, monitoring, preventing status changes in communication apps like Teams/Slack) but want the activity simulation to be less robotic and predictable than simple, repetitive actions.

## Features

* **Human-Like Simulation:** Randomizes the timing and nature of actions (mouse jiggle, key press, or both) to appear less like a basic automation script.
* **Randomized Timing:** Waits a variable amount of time between actions, configured around a base interval.
* **Randomized Mouse Jiggle:** Moves the mouse cursor by a small, random offset in X and Y directions and then returns it to the original position after a brief, randomized pause.
* **Randomized Key Press:** Occasionally presses (and releases) relatively harmless keys like `Shift`, `Ctrl`, `Alt`, `ScrollLock`, `NumLock`, or `CapsLock`. Toggle keys are pressed twice to revert their state.
* **Configurable Active Hours:** Set specific start and end times during which the script should operate.
* **Adjustable Intensity:** Control the base sleep interval, the randomness range for sleep, and the maximum distance for mouse jiggles via script parameters.
* **Power Management API:** Uses `SetThreadExecutionState` to formally request the system and display stay awake, providing a more robust method than input simulation alone.
* **Status Logging:** Outputs status messages, actions performed, and periodic runtime updates to the console.
* **Graceful Exit:** Resets the power management state when the script finishes or the time window ends, allowing the system to sleep normally again.

## Prerequisites

* **Windows Operating System:** The script relies on Windows-specific components (`System.Windows.Forms`, `Wscript.Shell`, `kernel32.dll`).
* **PowerShell:** Version 5.1 or later recommended.
* **Permissions:** Ability to run PowerShell scripts. You might need to adjust your execution policy (e.g., `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`).

## Usage

1.  **Save the Script:** Save the code block provided previously as a `.ps1` file (e.g., `HumanNoSleep.ps1`).
2.  **Open PowerShell:** Launch a PowerShell terminal.
3.  **Navigate to Directory:** Use `cd` to navigate to the folder where you saved the script.
4.  **Run the Script:** Execute the script using one of the following methods:

    * **With Default Settings:**
        ```powershell
        .\HumanNoSleep.ps1
        ```
        *(This uses the default active hours 08:00-22:00, 30 +/- 10 seconds sleep, +/- 5 pixel jiggle)*

    * **With Custom Parameters:**
        ```powershell
        .\HumanNoSleep.ps1 -baseSleepSeconds 60 -sleepRandomnessSeconds 15 -mouseJiggleRange 3 -minActiveTime '09:00' -maxActiveTime '17:30'
        ```

### Parameters

* `baseSleepSeconds` (Integer, Default: `30`): The average number of seconds to wait between actions.
* `sleepRandomnessSeconds` (Integer, Default: `10`): The maximum number of seconds to randomly add or subtract from `baseSleepSeconds`. The actual sleep will be between `baseSleepSeconds - sleepRandomnessSeconds` and `baseSleepSeconds + sleepRandomnessSeconds` (minimum 5 seconds enforced).
* `mouseJiggleRange` (Integer, Default: `5`): The maximum number of pixels (positive or negative) the mouse will move randomly in the X and Y directions. Must be 1 or greater.
* `minActiveTime` (String, Default: `'08:00'`): The earliest time (HH:mm format) the script should be active.
* `maxActiveTime` (String, Default: `'22:00'`): The latest time (HH:mm format) the script should stop being active.

## How It Works

1.  **Initialization:** Sets up necessary .NET assemblies, COM objects, and defines parameters.
2.  **Power State:** Calls the Windows API function `SetThreadExecutionState` with flags `ES_DISPLAY_REQUIRED`, `ES_SYSTEM_REQUIRED`, and `ES_CONTINUOUS`. This tells Windows that the application requires the display and system to remain powered on continuously.
3.  **Main Loop:** Runs continuously as long as the current time is within the specified `minActiveTime` and `maxActiveTime`.
4.  **Random Action:** In each loop iteration, it randomly decides whether to:
    * Simulate a key press (selecting a random key from a predefined list).
    * Simulate a mouse jiggle (moving the cursor slightly and then back).
    * Do both.
    * Do nothing specific beyond waiting (implicit).
5.  **Random Wait:** Pauses execution for a randomized duration based on the `baseSleepSeconds` and `sleepRandomnessSeconds` parameters.
6.  **Time Check:** Updates the current time and checks if it's still within the active window.
7.  **Logging:** Periodically prints the elapsed runtime.
8.  **Completion:** Once the time window ends, it calls `SetThreadExecutionState` again with just `ES_CONTINUOUS` to release the requirement for the system/display to stay awake, allowing normal power management to resume.

## Important Notes & Caveats

* **Foreground Applications:** While the key presses chosen (`Shift`, `Ctrl`, `Alt`, `ScrollLock`, etc.) are generally non-disruptive when pressed alone, they *could* potentially interfere if you happen to be actively typing or interacting with an application that uses these keys at the exact moment the script sends them. The mouse jiggle is designed to return the cursor to its original position, minimizing disruption.
* **Gaming/Sensitive Apps:** Avoid running this script while gaming or using applications sensitive to unexpected input, as even minor simulation could cause issues.
* **Resource Usage:** The script itself is very lightweight, but it keeps PowerShell running.
* **Execution Policy:** You may need to adjust your PowerShell execution policy if you encounter errors running the script. `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` is often sufficient.
* **Run as Administrator:** Not typically required, but might be necessary in some locked-down environments.

## Disclaimer

This script simulates user input and modifies system power settings. Use it responsibly and at your own risk. The author is not responsible for any unintended consequences or issues arising from its use. Ensure you understand what the script does before running it, especially in corporate or sensitive environments. Always test in a controlled manner first.
