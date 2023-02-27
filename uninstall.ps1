$Menu = {
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
   Invoke-Command $Menu
   $Select = Read-Host

   $uninstall = {
      $null = New-Item -Path $env:temp\uninstall -ItemType Directory -Force
      Set-Location $env:temp\uninstall
      $fileName = 'configuration.xml'
      $null = New-Item $fileName -ItemType File -Force
      Add-Content $fileName -Value '<Configuration>'
      Add-Content $fileName -Value '<Remove All="True"/>'
      Add-Content $fileName -Value '</Configuration>'
      $uri = 'https://github.com/bonben365/office365-installer/raw/main/setup.exe'
      $null = Invoke-WebRequest -Uri $uri -OutFile 'setup.exe' -ErrorAction:SilentlyContinue
      .\setup.exe /configure .\configuration.xml

      Write-Host
      Write-Host ============================================================
      Write-Host "Unnstalling...."
      Write-Host ============================================================
      Write-Host

      Write-Host
      Write-Host ============================================================
      Write-Host "Done...."
      Write-Host ============================================================
      Write-Host

      Start-Sleep -Seconds 3

      # Cleanup
      Set-Location "$env:temp"
      Remove-Item $env:temp\uninstall -Recurse -Force
  }

   Switch ($Select)
      {
         #1. Uninstall All Previous Versions of Microsoft Office
          1 {Invoke-Command $uninstall}  
      }
   }
   While ($Select -ne 3)


   
