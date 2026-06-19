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
write-host " ***************************************************************"
write-host "Dit script maakt een snelkoppeling naar het programma Beherenbestanden met de juiste opties."
write-host "Standaard wordt de snelkoppeling gemaakt voor PowerShell 5."
write-host "Wilt u echter kiezen voor PowerShell 7? j/n : " -NoNewline

do {
    $PressedKey = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    $keuze = $PressedKey.character
    write-host $keuze -NoNewline
} 
while ( !("j","J","n","N" -contains($keuze ) ) )

if ( $keuze -eq "j" -or $keuze -eq "J" ) {
    # Locatie van de executeable van Powershell 7 in een windows 64bit systeem. 
    try {
        $locatiePSexe = (Get-Command pwsh -ErrorAction Stop).Source
    }
    catch {
        Write-Host "PowerShell 7 niet gevonden." -ForegroundColor Yellow
        Write-Host "De snelkoppeling wordt gemaakt voor PowerShell 5.0" -ForegroundColor Yellow
    }

}

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
write-host ""
Write-Host "Een snelkoppeling is gemaakt naar het programma Beherenbestanden met de volgende opties: "
Write-Host "Link naar programma   : $hoofdprog" 
Write-Host "Startmap              : $startmap"
write-host "Snelkoppeling gemaakt : $ShortcutPath "
write-host "Pad naar PowerShell   : $locatiePSexe"

# Wacht een op een toets om af te sluiten.
write-host ""
write-host "Klaar. Druk op een toets om af te sluiten...." -NoNewline
$x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
write-host $x.character
