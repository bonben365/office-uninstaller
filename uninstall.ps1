$MainMenu = {
   Write-Host " ******************************************************"
   Write-Host " * Microsoft Office Installation Script               *" 
   Write-Host " * Date:    26/02/2023                                *" 
   Write-Host " * Author:  https://github.com/bonben365              *" 
   Write-Host " * Website: https://bonben365.com/                    *" 
   Write-Host " ******************************************************" 
   Write-Host 
   Write-Host " 1. Uninstall All Previous Versions of Microsoft Office"
   Write-Host " 2. Quit or Press Ctrl + C"
   Write-Host 
   Write-Host " Select an option and press Enter: "  -nonewline
   }
   cls
   
   Do { 
   cls
   Invoke-Command $MainMenu
   $Select = Read-Host
   Switch ($Select)
      {
         #1. Uninstall All Previous Versions of Microsoft Office
          1 {
               Set-Location $env:temp
               $uri = 'https://aka.ms/SaRA_CommandLineVersionFiles'
               Invoke-WebRequest -Uri $uri -OutFile 'SaRACmd.zip' -ErrorAction:SilentlyContinue | Out-Null
               Expand-Archive SaRACmd.zip -Force -ErrorAction:SilentlyContinue | Out-Null
               Set-Location "$((Get-Location).Path)\SaRACmd"
               Write-Host
               Write-Host ============================================================
               Write-Host "Uninstalling All Previous Versions of Microsoft Office...."
               Write-Host ============================================================
               Write-Host
               .\SaRAcmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion All

            cls
            }  
      }
   }
   While ($Select -ne 2)
