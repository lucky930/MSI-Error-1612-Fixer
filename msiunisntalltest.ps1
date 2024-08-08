

$msiPath = "C:\temp\IQOption.msi"   # change msi file path below
$msiName = "IQOption.msi"            #change msi name here
$productCode = "{14D7E71E-ADA6-47B5-9164-36DCA8B4CEB7}"  #insert the product code here

Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /L*v C:\temp\MSIUninstall.log"  #catching the log
Start-Sleep -Seconds 5
$MSimatch = ((Get-Content -Path "C:\temp\MSIUninstall.log" | Select-String -Pattern "C:\\.*\.msi" -AllMatches).Matches).Value  #capturing the required msi path from log file

foreach($requiredPath in $MSimatch)   #loop start for each path found in the log file
{
    $msiFile = Split-Path $requiredPath -leaf   #getting only msi file name from full path
    $msiFolder = Split-Path $requiredPath -Parent #getting only parent path from full path

    if(!(Test-path -Path "$msiFolder")) { New-Item -Path "$msiFolder" -ItemType Directory -ErrorAction SilentlyContinue}
        Copy-Item -Path "$PSScriptRoot`\$msiName" -Destination "$msiFolder" #copying the msi file to required path that we captured in the log file
   
    if(!(Test-Path "$requiredPath")) { Rename-Item -Path "$msiFolder`\$msiName" -NewName "$requiredPath" -Force } #renaming the msi file at required path that exactly logs were looking to uninstall
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /L*v C:\temp\$msiFile`_Uninstall.log" #uninstalling the msi using product code

}