
# Uninstall All Previous Versions of Microsoft Office

## Requirements

  
## Installation

- Right click on the Windows Start icon 🪟 then select Windows PowerShell (Admin).
- Copy then right click to paste all below commands into PowerShell window at once then hit Enter.


```ps
Set-ExecutionPolicy Bypass -Scope Process -Force
$url="https://github.com/bonben365/office-uninstaller/raw/main/uninstall.ps1"
iex ((New-Object System.Net.WebClient).DownloadString($url))
```
➡️Please inspect https://github.com/bonben365/office-uninstaller/raw/main/uninstall.ps1 prior to running any of these scripts to ensure safety. We already know it's safe, but you should verify the security and contents of any script from the internet you are not familiar with.

## Screenshots


## Documentation

[🌍https://bonben365.com/](https://bonben365.com/)

