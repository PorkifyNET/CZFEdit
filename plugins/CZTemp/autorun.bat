@ECHO OFF
ECHO CZ/AutoRun Script v1.2
ECHO ---
SET CZTEMP=%TEMP%\CZTEMP
IF EXIST C:\CZTEMP\ (
  SET CZTEMP=C:\CZTEMP
)

ECHO Tijdelijke map: %CZTEMP%

IF "%~1"=="-office" GOTO OFFICE
IF "%~1"=="-hotkeys" GOTO HOTKEYS
IF "%~1"=="-update" GOTO UPDATE

ECHO ---
ECHO Verlopen van Gebruiker wachtwoord uitzetten
ECHO.
ECHO net accounts /maxpwage:unlimited >%CZTEMP%\wscript.ps1
ECHO Set-LocalUser -Name ^"Gebruiker^" -PasswordNeverExpires 1 >>%CZTEMP%\wscript.ps1
%SystemRoot%\sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File "%CZTEMP%\wscript.ps1"
DEL %CZTEMP%\wscript.ps1

ECHO ---
ECHO Fast startup uitzetten
ECHO.
Powercfg -h off
ECHO Gelukt

SET choice=
SET /p choice=Updates installeren? [y^|N]: 
IF NOT '%choice%'=='' SET choice=%choice:~0,1%
IF '%choice%'=='Y' GOTO UPDATE
IF '%choice%'=='y' GOTO UPDATE
IF '%choice%'=='N' GOTO ENDSCRIPT
IF '%choice%'=='n' GOTO ENDSCRIPT
IF '%choice%'=='' GOTO ENDSCRIPT

:UPDATE
ECHO ---
ECHO Windows Updates starten
ECHO.
ECHO Windows Update PowerShell script bouwen ...
TIMEOUT /T 2 >nul

ECHO ECHO ^"Tools downloaden^" >>%CZTEMP%\wupdate.ps1
ECHO Install-PackageProvider NuGet -Force; >%CZTEMP%\wupdate.ps1
ECHO Set-PSRepository PSGallery -InstallationPolicy Trusted >>%CZTEMP%\wupdate.ps1
ECHO Install-Module PSWindowsUpdate -Repository PSGallery >>%CZTEMP%\wupdate.ps1
ECHO Import-Module -Name PSWindowsUpdate >>%CZTEMP%\wupdate.ps1
ECHO ECHO ^"Installeren Driver updates:^" >>%CZTEMP%\wupdate.ps1
ECHO Install-WindowsUpdate -Install -AcceptAll -UpdateType Driver -MicrosoftUpdate -ForceDownload -ForceInstall -IgnoreReboot -ErrorAction SilentlyContinue ^| Out-File ^"%CZTEMP%\Drivers_Install_1_$(get-date -f dd-MM-yyyy).log^" -Force >>%CZTEMP%\wupdate.ps1
ECHO Install-WindowsUpdate -Install -AcceptAll -UpdateType Driver -MicrosoftUpdate -ForceDownload -ForceInstall -IgnoreReboot -ErrorAction SilentlyContinue ^| Out-File ^"%CZTEMP%\Drivers_Install_1_$(get-date -f dd-MM-yyyy).log^" -Force >>%CZTEMP%\wupdate.ps1
ECHO ECHO ^"Installeren Windows updates:^" >>%CZTEMP%\wupdate.ps1
ECHO Get-WUlist -MicrosoftUpdate >>%CZTEMP%\wupdate.ps1
ECHO Get-WindowsUpdate -AcceptAll -Install >>%CZTEMP%\wupdate.ps1
ECHO ECHO ^"Updaten Windows defender:^" >>%CZTEMP%\wupdate.ps1
ECHO Update-MpSignature -UpdateSource MicrosoftUpdateServer >>%CZTEMP%\wupdate.ps1

ECHO $Module = Get-Module PSWindowsUpdate >>%CZTEMP%\wupdate-clean.ps1
ECHO Remove-Module $Module.Name >>%CZTEMP%\wupdate-clean.ps1
ECHO Remove-Item $Module.ModuleBase -Recurse -Force >>%CZTEMP%\wupdate-clean.ps1
ECHO Uninstall-Module -Name PSWindowsUpdate >>%CZTEMP%\wupdate-clean.ps1
ECHO (Get-PackageProvider^|where-object{$_.name -eq ^"nuget^"}).ProviderPath^|Remove-Item -force  >>%CZTEMP%\wupdate-clean.ps1

ECHO Windows Update PowerShell uitvoeren ...
%SystemRoot%\sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File "%CZTEMP%\wupdate.ps1"
REM ECHO Windows Update cleanup ...
REM powershell.exe -ExecutionPolicy Bypass -File "%CZTEMP%\wupdate-clean.ps1"
DEL %CZTEMP%\wupdate.ps1
DEL %CZTEMP%\wupdate-clean.ps1

TIMEOUT /T 2 >nul

GOTO ENDSCRIPT

:OFFICE
ECHO ---
ECHO Microsoft Office installer
ECHO.
ECHO Microsoft Office installeren ...
%CZTEMP%\setup.exe /configure %CZTEMP%\configuration.xml
GOTO ENDSCRIPT

:HOTKEYS
ECHO ---
ECHO HP Hotkeys installer
ECHO.
ECHO HP Service stopzetten ...
sc config "HotKeyServiceUWP" start=disabled
ECHO HP Hotkeys installeren ...
"%CZTEMP%\HP hotkey support.exe"
GOTO ENDSCRIPT

:ENDSCRIPT

ECHO CZ/AutoRun Script einde
PAUSE