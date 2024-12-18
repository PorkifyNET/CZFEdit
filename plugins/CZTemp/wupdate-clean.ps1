$Module = Get-Module PSWindowsUpdate 
Remove-Module $Module.Name 
Remove-Item $Module.ModuleBase -Recurse -Force 
Uninstall-Module -Name PSWindowsUpdate 
(Get-PackageProvider|where-object{$_.name -eq "nuget"}).ProviderPath|Remove-Item -force  
