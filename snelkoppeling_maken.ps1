<# script om een snelkoppeling te maken van Beherenbestanden met de juiste opties
#>

# Declareren variabelen ----------------------------------------------------------

# startmap van het programma
#$startmap = $PSScriptRoot
$startmap="$env:USERPROFILE\beherenbestanden"

# hoofd programma die uitgevoerd wordt
[string] $hoofdprog = "beherenbestanden.ps1"

# samenvoegen installatiemap met hoofdprogramma. nodig voor aanmaken snelkoppeling
$hoofdprog = -join ("$startmap","\","$hoofdprog")

# Locatie van de executeable van Powershell in een windows 64bit systeem. 
# Let op dat bij een 32bit systeem dit anders is.
[string] $locatiePSexe = "c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe"

# Einde declareren variabelen ----------------------------------------------------

# Maken van snelkoppeling
$linkbureaublad = [Environment]::GetFolderPath("Desktop")
$SourceFilePath = $locatiePSexe
#$ShortcutPath = "$linkbureaublad\Beheren bestanden.lnk"
$ShortcutPath = "$linkbureaublad\Beheren bestanden.lnk"

$WScriptObj = New-Object -ComObject ("WScript.Shell")
$shortcut = $WscriptObj.CreateShortcut($ShortcutPath)
$shortcut.TargetPath = $SourceFilePath
$shortcut.Arguments = "-ExecutionPolicy Bypass -NoProfile -file " + '"' + "$hoofdprog" + '"' 
$shortcut.IconLocation = "$startmap" + "\" + "beheren.ico"
$shortcut.WorkingDirectory = "$startmap"
$shortcut.Save()

# Einde maken van snelkoppeling

# Info geven
write-host ""
Write-Host "Een snelkoppeling is gemaakt naar het programma Beherenbestanden met de volgende opties: "
Write-Host "Link naar programma   : $hoofdprog" 
Write-Host "Startmap              : $startmap"
write-host "Snelkoppeling gemaakt : $ShortcutPath "# Wacht een op een toets om af te sluiten.write-host ""write-host "Klaar. Druk op een toets om af te sluiten...."$x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")