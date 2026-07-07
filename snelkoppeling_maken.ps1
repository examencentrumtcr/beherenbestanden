<# script om een snelkoppeling te maken van Beherenbestanden met de juiste opties

Je kan de volgende parameters meegeven aan het script:
    -auto : maakt automatisch een snelkoppeling.
            Als deze parameter niet meegegeven wordt, dan wordt de gebruiker gevraagd of hij/zij een snelkoppeling wil maken naar PowerShell 7.
    -pwsh7 : maakt automatisch een snelkoppeling naar PowerShell 7.0
             Deze paramter moet alleen meegegeven worden als de parameter -auto ook meegegeven wordt.
#>

# Declareren variabelen ----------------------------------------------------------
param (
    [switch]$auto,
    [switch]$pwsh7,
    [string]$startmap="$env:USERPROFILE\beherenbestanden"
)

# controleren van startmap van het programma
if (!(Test-Path $startmap)) {
    Write-Host "De opgegeven startmap bestaat niet: $startmap" -ForegroundColor Red
    if ($auto) {
        start-sleep -Seconds 5
    } else {
        Write-Host "Druk op een toets om af te sluiten..." -NoNewline
        $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-Host $x.character
    }
    exit 1
}

# hoofd programma die uitgevoerd wordt
[string] $hoofdprog = "beherenbestanden.ps1"

# samenvoegen installatiemap met hoofdprogramma. nodig voor aanmaken snelkoppeling
$hoofdprog = -join ("$startmap","\","$hoofdprog")

## Begin van het script -------------------------------------------------------------

if (!($auto)) {
    # als er geen parameter auto is meegegeven, dan wordt de gebruiker gevraagd of hij/zij een snelkoppeling wil maken naar PowerShell 7.
    write-host " ***************************************************************"
    write-host "Dit script maakt een snelkoppeling naar het programma Beherenbestanden met de juiste opties."
    write-host "Standaard wordt de snelkoppeling gemaakt voor PowerShell 5."
    Write-Host "Wilt u echter kiezen voor PowerShell 7? (j/n) [standaard: n]: " -NoNewline

    do {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        if ($key.VirtualKeyCode -eq 13) {
        # Enter = standaard Nee
        $keuze = 'n'
        }
        else {
        $keuze = $key.Character.ToString().ToLower()
        }
    }
    while ($keuze -notin @('j','n'))
    Write-Host $keuze

} elseif ($pwsh7) {
    # auto parameter is meegegeven en de parameter pwsh7 ook, dan wordt automatisch een snelkoppeling gemaakt naar PowerShell 7.
    $keuze = "j"
    write-host "Script Snelkoppeling_maken.ps1 wordt uitgevoerd."
    write-host "Er wordt automatisch een snelkoppeling gemaakt naar PowerShell 7."
} else {
    # auto parameter is meegegeven, maar de parameter pwsh7 niet, dan wordt automatisch een snelkoppeling gemaakt naar PowerShell 5.
    $keuze = "n"
    write-host "Script Snelkoppeling_maken.ps1 wordt uitgevoerd."
    write-host "Er wordt automatisch een snelkoppeling gemaakt naar PowerShell 5."
}

# Locatie van de executeable van Powershell 5 of 7, afhankelijk van de gemaakte keuze van de gebruiker.
if ( $keuze -eq "j" -or $keuze -eq "J" ) {
        # Locatie van de executeable van Powershell 7 in een windows 64bit systeem. 
        try {
            $locatiePSexe = (Get-Command pwsh -ErrorAction Stop).Source
        }
        catch {
            Write-Host "PowerShell 7 niet gevonden." -ForegroundColor Yellow
            Write-Host "De snelkoppeling wordt gemaakt voor PowerShell 5" -ForegroundColor Yellow
            $locatiePSexe = (get-command 'powershell.exe').source
        }

} else {
    $locatiePSexe = (get-command 'powershell.exe').source
}

# Maken van snelkoppeling
$linkbureaublad = [Environment]::GetFolderPath("Desktop")
$SourceFilePath = $locatiePSexe
$ShortcutPath = "$linkbureaublad\Beheren bestanden.lnk"

$WScriptObj = New-Object -ComObject ("WScript.Shell")
$shortcut = $WscriptObj.CreateShortcut($ShortcutPath)
$shortcut.TargetPath = $SourceFilePath
$shortcut.Arguments = "-ExecutionPolicy Bypass -NoProfile -file " + '"' + "$hoofdprog" + '"' 
$shortcut.IconLocation = "$startmap" + "\" + "script_icoon.ico"
$shortcut.WorkingDirectory = "$startmap"
$shortcut.Save()

# Einde maken van snelkoppeling

# Alleen als de parameter -auto niet meegegeven is, tekst weergeven en wachten op een toets om af te sluiten. 
if (!($auto)) {
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
} else {
    write-host "Klaar. Snelkoppeling is aangemaakt."
}