# SCOMMPBaselineAutomation

## TL;DR
1. Export your management pack to an Excel XML file and place it next to this script.
2. Run `MP_Baseline_Automation.ps1` on your SCOM management server.
3. Check the created overrides in the Operations Manager console.

## Overview
`MP_Baseline_Automation.ps1` automates baseline overrides for SCOM management packs. It scans an exported Excel XML file for monitors and rules marked as **Disabled**, **Debug** or **Activated** and creates the respective override entries. Debug overrides are assigned to a `*.DebugGrp` group that is created automatically if needed.

## Prerequisites
- Windows PowerShell with the modules `OperationsManager` and `OpsMgrExtended` installed.
- Permissions to modify the target SCOM environment.
- The script must run on a SCOM management server.

## Files
- `MP_Baseline_Automation.ps1` – main automation script.
- `LICENCE` – usage license (CC BY-SA 4.0).

## Usage
1. Export your source management pack with MPViewer as an Excel XML file and store it in the same directory as `MP_Baseline_Automation.ps1`.
2. Edit the worksheets `Monitors - Unit` and `Rules` in the XML file and set the `Enabled` column to `Debug` or `Disabled` as required.
3. Run the script. It automatically loads the XML file, creates (if necessary) an override management pack and a debug group, removes old overrides (unless `-Scope AddOnly` is used) and generates new overrides from the XML.

### Flags
- `-Scope` – `AddOnly` (only add new overrides) or `DeleteOnly` (only remove overrides previously created by this script).
- `-DebugLog` – enable verbose output.

### Example
```powershell
.\MP_Baseline_Automation.ps1 -DebugLog -Scope AddOnly
```
This runs the script with detailed debug information and adds new overrides without removing existing ones.

## License
This project is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License as noted in the `LICENCE` file.
