Install-PackageProvider NuGet -Force; 
Set-PSRepository PSGallery -InstallationPolicy Trusted 
Install-Module PSWindowsUpdate -Repository PSGallery 
Import-Module -Name PSWindowsUpdate 
ECHO "Installeren Driver updates:" 
Install-WindowsUpdate -Install -AcceptAll -UpdateType Driver -MicrosoftUpdate -ForceDownload -ForceInstall -IgnoreReboot -ErrorAction SilentlyContinue | Out-File "C:\Users\GEBRUI~1\AppData\Local\Temp\CZTEMP\Drivers_Install_1_$(get-date -f dd-MM-yyyy).log" -Force 
Install-WindowsUpdate -Install -AcceptAll -UpdateType Driver -MicrosoftUpdate -ForceDownload -ForceInstall -IgnoreReboot -ErrorAction SilentlyContinue | Out-File "C:\Users\GEBRUI~1\AppData\Local\Temp\CZTEMP\Drivers_Install_1_$(get-date -f dd-MM-yyyy).log" -Force 
ECHO "Installeren Windows updates:" 
Get-WUlist -MicrosoftUpdate 
Get-WindowsUpdate -AcceptAll -Install 
ECHO "Updaten Windows defender:" 
Update-MpSignature -UpdateSource MicrosoftUpdateServer 
