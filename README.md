
# Add Windows Store into Windows LTSC

## Requirements

  
## Installation

- Right click on the Windows Start icon 🪟 then select Windows PowerShell (Admin).
- Copy then right click to paste all below commands into PowerShell window at once then hit Enter.


```ps
$url="https://raw.githubusercontent.com/bonben365/add-store-win-ltsc/main/add-store.ps1"
Set-ExecutionPolicy Bypass -Scope Process -Force
iex ((New-Object System.Net.WebClient).DownloadString($url))
```
➡️Please inspect https://github.com/bonben365/office-uninstaller/raw/main/uninstall.ps1 prior to running any of these scripts to ensure safety. We already know it's safe, but you should verify the security and contents of any script from the internet you are not familiar with.

## Screenshots


## Documentation

[🌍https://bonben365.com/](https://bonben365.com/)

