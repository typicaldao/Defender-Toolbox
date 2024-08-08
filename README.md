# Defender Toolbox 
## Objectives
Defender Toolbox uses PowerShell functions to help you with Defender Antivirus or MDE log parsing automatically. It is also a replacement of some official Python tools for Xplatform. 

## Usage
### *Recommended*: Use one-line command to download the install the module automatically via the script.
```ps1
Invoke-RestMethod "https://raw.githubusercontent.com/typicaldao/Defender-Toolbox/main/Update-DefenderToolbox.ps1" | Invoke-Expression
```

### Or its shorter version:
```ps1
irm "https://raw.githubusercontent.com/typicaldao/Defender-Toolbox/main/Update-DefenderToolbox.ps1" | iex
```

### Update the module
After you have successfully installed the module, run `Update-DefenderToolbox` to update the module, as a replacement of above commands. (Starting from version 0.2)

### *Others*:
Download and import the separated functions to use them.
Or you can use the combined PowerShell profile `Microsoft.Powershell_profile.ps1` so that the functions are automatically loaded when PowerShell is launched.

## What's new
August 8th: Update module version to 0.2. Following functions are included the the module:
1. `Convert-MpRegistrytxtToJson`
2. `Update-DefenderToolbox`

## Recommendations
1. PowerShell 7.x is recommended, but PowerShell 5.1 should work as well. However, it could encounter into some encoding issues (like `ConvertTo-Json` - resolved on August 7th). Same function could have different parameters and results, so let me know when you encounter such issue.
1. Load the module automatically in your PowerShell profile. (The update module will check and help you with it.) When you need to parse log, right click in file explorer and run the functions directly. Usually, no additional parameter is needed as the default log name is used.

## Key functions
### `Convert-MpOperationalEventLogTxt`
Format MpOperationalEvents.txt into PowerShell objects, and export the output as CSV so that you can use Excel or other tools to view and filter.

Default input: Convert *MpOperationalEvents.txt* (from *MpSupportFiles.cab*) in the current location.
Optional: Use -Path to specify your log file, and use -OutFile to save the output.

### `Convert-MacDlpPolicyBin`
Convert Endpoint DLP policy from macOS (dlp_policy.bin, dlp_sensitive_info.bin and other policy-related .bin files).

Default: Convert *dlp_policy.bin* in the current location.

### `Convert-WdavHistory`
Convert Xplatform scan history file `wdavhistory` from JSON to a readable list.

Default input: file *wdavhistory* 
Note: *wdavhistory* can be extracted from mde_diagnostic.zip <- (XMDE client analyzer result).zip, or via "mdatp diagnostic create".

### `Convert-MdavRtpDiag`
Convert Xplatform (macOS/Linux) real-time protection diagnostic log from JSON to a readable list.
Reference official documentation: https://learn.microsoft.com/en-us/defender-endpoint/linux-support-perf?view=o365-worldwide
Python version of the parser provided by Microsoft/mdatp-xplat: https://github.com/microsoft/mdatp-xplat/blob/master/linux/diagnostic/high_cpu_parser.py

### `Convert-MpRegistrytxtToJson`
Convert Defender log: MpRegistry.txt into a JSON format. Please be aware that if you are using Windows default PowerShell 5.1, the output of JSON output via function ConvertTo-Json is not that pretty. You might need to use your text editor to prettier JSON for you. PowerShell 7.x works just fine. 

## Feature on the way
1. List help information of the tool.
