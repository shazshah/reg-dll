#-------------------------------------------------------------------------------
#Author:            Shaz
#Purpose:           Unregister old DLL and register new DLL
#Date:              13/11/2015
#Version:           1.0.1
#Edit:              SS 27/11/2015: Changed the order of DLL Unregister to occur
#                                  before old dll rename (1.0.0 > 1.0.1).
#-------------------------------------------------------------------------------

"`n***************************************************************************"
"DLL REG"
"***************************************************************************"
"Enter the location of the new DLL and then the existing DLL."
"The existing DLL will be unregistered and renamed while the new one copied" 
"to its location and registered."
"--------------------------------------------------------------------------`n"

#Prompt for location of new DLL
Do {
$newDLL = Read-Host 'Enter the location of NEW DLL (e.g. "\\networkshare\example.dll")'} While($newDLL.Length -eq 0)

#Strip out quotes
$newDLL = $newDLL.Replace("`"", "")

#Check if New DLL exists
if (-not(Test-Path $newDLL)) {
"`nCannot find $newDLL. Exiting"

break }

#Prompt for Old DLL (existing dll)
Do {
$oldDLL = Read-Host "`nEnter the location of EXISTING DLL (e.g. `"C:\Program Files (x86)\AProgram\example.dll`")"} While($oldDLL.Length -eq 0)

#Strip out quotes
$oldDLL = $oldDLL.Replace("`"", "")

#Check if OLD DLL exists
if (-not(Test-Path $oldDLL)) {
"`nCannot find $oldDLL. Exiting"

break }

#Store the path of the old DLL - need this as the location to copy new DLL to
$copyPath = Get-ChildItem $oldDLL

#----Unregister old DLL-----

"`nUnregistering " + [System.IO.Path]::GetFileName($oldDLL)

Invoke-Expression "C:\Windows\syswow64\regsvr32.exe /u '$oldDLL'"

#----Begin renaming the old dll----

"`nRenaming EXISTING DLL with date stamp..."

[string]$directory = [System.IO.Path]::GetDirectoryName($oldDLL);
[string]$strippedFileName = [System.IO.Path]::GetFileNameWithoutExtension($oldDLL);
[string]$extension = [System.IO.Path]::GetExtension($oldDLL);
[string]$newFileName = $strippedFileName + $extension + [DateTime]::Now.ToString("yyyyMMdd-HHmmss");
[string]$renameddll = [System.IO.Path]::Combine($directory, $newFileName);

Move-Item -LiteralPath $oldDLL -Destination $renameddll;

"`n...Existing DLL has been renamed to: $renameddll"

#----Begin copying new dll----

"`nCopying new DLL to Old DLL's location..."

Copy-Item $newDLL $copyPath.DirectoryName

#Get the new path of the NEW DLL
$newDllPath = $copyPath.DirectoryName + "\" + [System.IO.Path]::GetFileName($newDLL)
"'$newdllPath'"

#-----Register New DLL-----

"`nRegistering " + [System.IO.Path]::GetFileName($newDLL)
Invoke-Expression "C:\Windows\syswow64\regsvr32.exe '$newDllPath'"

"`nDone!`n"