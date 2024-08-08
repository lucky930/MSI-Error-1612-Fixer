
<#
.SYNOPSIS
  PowerShell Code to Fix MSI Error 1612 or Fix "The feature you are trying to use is on network resource that is unavailable" Error"
.DESCRIPTION
  PowerShell Code to Fix MSI Error 1612 or Fix "The feature you are trying to use is on network resource that is unavailable" Error"
  This code will fetch the path where the msiexec is looking to uninstall. Usually this error comes when cache and local msi gets deleted from 
  the system and then msiexec cant find it to uninstall or repair the application
.PARAMETER <Parameter_Name>
    NA
.INPUTS
  $msipath, $msiName and $productCode
.OUTPUTS
  Log file stored in C:\Windows\Temp\ directory
.NOTES
  Version:        1.0
  Author:         Ramandeep Singh
  Creation Date:  08/08/2024
  Purpose/Change: Initial script development
  
.EXAMPLE
  NA
#>

$msiName = "IQOption.msi"            #name of the msi you will be uninstalling. Keep this file along with this script at same location.
$productCode = "{14D7E71E-ADA6-47B5-9164-36DCA8B4CEB7}"  #insert the product code of the msi here

Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /L*v C:\windows\temp\MSIUninstall.log"  #trying to uninstall msi using product code, obviously it will fail.
Start-Sleep -Seconds 5   #sleeping for 5 seconds
$MSimatch = ((Get-Content -Path "C:\temp\MSIUninstall.log" | Select-String -Pattern "C:\\.*\.msi" -AllMatches).Matches).Value  #filtering the required msi path from log file

foreach($requiredPath in $MSimatch)   #loop start for each path found in the $Msimatch variable
{
    $msiFile = Split-Path $requiredPath -leaf   #getting msi file name from full path
    $msiFolder = Split-Path $requiredPath -Parent #getting parent path from full path

    if(!(Test-path -Path "$msiFolder")) { New-Item -Path "$msiFolder" -ItemType Directory -ErrorAction SilentlyContinue}
        Copy-Item -Path "$PSScriptRoot`\$msiName" -Destination "$msiFolder" #copying the local msi file to required path that we captured in the log file
   
    if(!(Test-Path "$requiredPath")) { Rename-Item -Path "$msiFolder`\$msiName" -NewName "$requiredPath" -Force } #renaming the msi file at required path to match exactly with the one in the logs
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /L*v C:\windows\temp\$msiFile`_Uninstall.log" #uninstalling the msi using product code, this time it should success

}
