# Defender Toolbox 
## Objectives
Defender Toolbox uses PowerShell functions to help you with Defender Antivirus or MDE log parsing automatically. It is also a replacement of some official Python tools for Xplatform. 

## Key functions
### `Convert-MpOperationalEventLogTxt`
Format MpOperationalEvents.txt into PowerShell objects, and export the output as CSV so that you can use Excel or other tools to view and filter.

### `Convert-MacDlpPolicyBin`
Convert Endpoint DLP policy from macOS (dlp_policy.bin, dlp_sensitive_info.bin and other policy-related .bin files).

### `Convert-WdavHistory`
Convert Xplatform scan history file `wdavhistory` from JSON to a readable list.

### `Convert-MdavRtpDiag`
Convert Xplatform real-time protection diagnostic log from JSON to a readable list.

## Usage
1. Download and import the separated functions to use them.
1. Or you can use the combined PowerShell profile `Microsoft.Powershell_profile.ps1` so that the functions are automatically loaded when PowerShell is launched.


## Feature on the way
1. List help information of the tool.
1. Install and update the tool to PowerShell profile in a one-line command.