<# 
Beherenbestanden.ps1

Programma om bestanden op pc's in een netwerk te beheren, dus bestanden klaarzetten, back-uppen, verplaatsen of wissen

Zie voor versienummer hier onder de comments.

Versienummer wordt volgens Semantic Versioning uitgevoerd (zie https://semver.org/lang/nl/)

Dit programma is beschermt met auteursplicht door middel van de GNU GPL (https://www.gnu.org/licenses)

Lees ook de readme.txt op https://beherenbestanden.neveshuis.nl/readme.txt


Programma versie wordt aangegeven in de vorm : Versie.Extralabel.DATE

Versie = versie van dit programma en wordt aangegeven in de vorm major.minor.patch
         In de titelbalk staat de versienummer

Extra label =  Een extra label wordt weergegeven bij info over dit programma
   Een extra label kan een pre-release of een build zijn
   Een pre-release wordt aangegeven met alpha, beta of pre-release
   De build geeft aan hoe vaak het programma is uitgebracht en is een oplopende getal
   Een build wordt alleen meegegeven als een versie naar release gaat

DATE wordt gegeven als JJMMDD (jaar, maand en dag)

Modes:
    alpha      : Logbestanden verwijderen is uit, Automatisch updates is uit, Lokale mappen worden gebruikt en Console blijft open
    beta       : Automatisch updates is uit en Console blijft open
    prerelease : Updates worden gedownload vanuit de prerelease pagina van Github. Alle functionalliteiten kunnen getest worden.
    release    : Normale gebruik

    De modus hoeven niet persé allemaal doorlopen te worden!
    Bij het opstarten krijg je een melding als je in een testfase zit (modes alpha, beta of prelease)

#>

# Hier worden de programma variabelen gedeclareerd.

<# bepalen naam van deze script. variabele scriptnaam wordt alleen hier gebruikt (en lokaal bij bepaalde functies, nl updaten en informatie over programma)
   alleen de naam vh bestand, dus zonder bovenliggende mappen
   extensie .ps1 wordt verwijderd 
#>
$scriptnaam = $MyInvocation.InvocationName
$scriptnaam = Split-Path -leaf $scriptnaam
$scriptnaam = $scriptnaam.Replace(".ps1","")
# de naam van het programma wordt ook gebruikt in de titelbalk van het hoofdvenster.
$global:programma = @{
    versie = '4.7.1'
    extralabel = 'alpha.1.251003' # alpha, beta, update, prerelease of release
    mode = 'alpha' # alpha, beta, prerelease of release
    naam = $scriptnaam
    github = "https://api.github.com/repos/examencentrumtcr/beherenbestanden/contents/"
}

write-host ""
write-host "** Programma $scriptnaam " -f Green
write-host "** Versie is "$global:programma.versie -f Green
write-host ""
write-host "Initialiseren van het programma."

<# Manier om console af te sluiten en weer te openen.
   Het sluiten wordt uitgevoerd voor het starten van de hoofdscherm.
#>
$code = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$ShowWindowAsync = Add-Type -MemberDefinition $code -Name myAPI -PassThru
$hwnd = (Get-Process -PID $pid).MainWindowHandle

# toevoegen .NET framework klassen ---------------------------------------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# VisualStyles aan zetten ------------------------------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()

# website met updates van programma
$updatewebsite = "https://beherenbestanden.neveshuis.nl"

# startmap van het programma bepalen
$startmap=Split-Path -Parent $PSCommandPath

# map met png icoontjes bepalen
$icoontjesmap = -join ("$startmap","\","png")

# map met gif afbeeldingen bepalen
$gifjesmap = -join ("$startmap","\","png")

# map voor logbestanden bepalen
$logmap = -join ("$startmap","\","log")

# tijdelijke logbestand voor foutmeldingen
$foutmeldingsbestand = -join ("$logmap","\","Foutmeldingen.txt")

<# object met alle variabelen die in initialisatiebestand bewaard worden. Dit zijn de instellingen die gewijzigd kunnen worden.
   Zie ook functie gebruikersinstellingen onder sectie "Functions" #>
$Global:init=@{}

# de 2 bestanden met info voor venster informatieprogramma
$readmebestand="readme.md"
$changelogbestand="changelog.md"

# nodig voor werken met hashtabels. zie functies uitvoerentaken en overzichttaken
$uitvoeren = [hashtable]::Synchronized(@{})

# nodig voor informatie baloon 
$global:tooltip1 = New-Object System.Windows.Forms.ToolTip

# Afbeeldingen voor het hoofdmenu
$global:afbeeldingen = @(
    [PSCustomObject]@{
        naam = 'Automatiseren'
        bestand = 'automatiseren.gif'
    },
    [PSCustomObject]@{
        naam = 'Beoordelen'
        bestand = 'beoordelen.gif'
    },
    [PSCustomObject]@{
        naam = 'Controleren'
        bestand = 'controleren.gif'
    },
    [PSCustomObject]@{
        naam = 'Overwerken'
        bestand = 'overwerken.gif'
    },
    [PSCustomObject]@{
        naam = 'Archieveren'
        bestand = 'overzetten.gif'
    },
    [PSCustomObject]@{
        naam = 'Samenwerken'
        bestand = 'samenwerken.gif'
    },
    [PSCustomObject]@{
        naam = 'Thuiswerken'
        bestand = 'thuiswerken.gif'
    }
)

# Websites met moppen
$Global:Moppen = @(
    [PSCustomObject]@{
        naam = 'Appspot.com'
        url = 'https://official-joke-api.appspot.com/jokes/random'
        taal = 'Engels'
    },
    [PSCustomObject]@{
        naam = 'Apekool.nl'
        url = 'http://api.apekool.nl/services/jokes/getjoke.php' 
        taal = 'Nederlands'
    },
    [PSCustomObject]@{
        naam = 'Icanhazdadjoke.com'
        url = 'http://icanhazdadjoke.com/' 
        taal = 'Engels'
    }
)

# Beheerinstellingen. Dit zijn de instellingen die niet gewijzigd kunnen worden.
# Locaties worden hier gedefinieerd. Deze worden gebruikt bij het aanmaken van de lijst met rpc-nummers.
$PAL = @{
    naam = 'Prins Alexanderlaan'
    eerstenr = 1
    laatstenr= 60
}
$SHW = @{
    naam = 'Schiedamseweg'
    eerstenr = 131
    laatstenr= 180
}
$JLS = @{
    naam = 'Jan Ligthartstraat'
    eerstenr = 201
    laatstenr= 240
}
$locaties = @{
    PAL = $PAL
    SHW = $SHW
    JLS = $JLS
    }
$examenmappen = @{
    digitalebestanden = 'B:\HR_TM_T2000\examendocumenten\digitale bijlagen'
    homemapstudenten  = 'Y:\RPC'
    backupmap         = 'B:\HR_TM_T2000\backup kandidaten'
}
$opschonen=@{
    dagenbewarenbackup = 730
}
$global:beheer = @{
    examenmappen      = $examenmappen
    locaties          = $locaties
    opschonen         = $opschonen
    }

# lokale mappen worden gebruikt als programma.mode is alpha.
$lokalemappen = @{
    digitalebestanden = "C:\Users\$env:username\testwerking\digitale bijlagen"
    homemapstudenten  = "C:\Users\$env:username\testwerking\doel"
    backupmap         = "C:\Users\$env:username\testwerking\backup kandidaten"
}
if ($global:programma.mode -eq "alpha") {
        $global:beheer.examenmappen = $lokalemappen
        }


# Einde declareren variabelen


# functions ---------------------------------------------------------------------------------
# hieronder de kleine functies die door andere functies gebruikt worden ---------------------

function ShowHide-ConsoleWindow($mode) {
  
  if ($hwnd -ne [System.IntPtr]::Zero) {
    # When you got HWND of the console window:
    # (It would appear that Windows Console Host is the default terminal application)
    $null = $ShowWindowAsync::ShowWindowAsync($hwnd, $mode)
  } else {
    # When you failed to get HWND of the console window:
    # (It would appear that Windows Terminal is the default terminal application)

    # Mark the current console window with a unique string.
    $UniqueWindowTitle = New-Guid
    $Host.UI.RawUI.WindowTitle = $UniqueWindowTitle
    # $StringBuilder = New-Object System.Text.StringBuilder 1024

    # Search the process that has the window title generated above.
    $TerminalProcess = (Get-Process | Where-Object { $_.MainWindowTitle -eq $UniqueWindowTitle })
    # Get the window handle of the terminal process.
    # Note that GetConsoleWindow() in Win32 API returns the HWND of
    # powershell.exe itself rather than the terminal process.
    # When you call ShowWindowAsync(HWND, 0) with the HWND from GetConsoleWindow(),
    # the Windows Terminal window will be just minimized rather than hidden.
    $hwnd = $TerminalProcess.MainWindowHandle
    if ($hwnd -ne [System.IntPtr]::Zero) {
      # afsluiten terminal. Met $null zie je ook geen resultaat op he scherm. dus het woord "true" verschijnt niet.  
      $null = $ShowWindowAsync::ShowWindowAsync($hwnd, $mode) 
      
    } else {
      Write-Host "Niet gelukt om de status van de console window te wijzigen naar waarde $mode." -ForegroundColor Red
    }
  }
} # einde ShowHide-ConsoleWindow

Function declareren_standaardvenster ($titel, $pos_x, $pos_y)
{
$StandaardForm                            = New-Object system.Windows.Forms.Form
$StandaardForm.MaximumSize = New-Object System.Drawing.size($pos_x,$pos_y)
$StandaardForm.MinimumSize = New-Object System.Drawing.size($pos_x,$pos_y)
$StandaardForm.text                       = $titel
$StandaardForm.TopMost                    = $false
$StandaardForm.StartPosition              = 'CenterScreen'
$StandaardForm.BackColor = "white"
$StandaardForm.MaximizeBox = $False
$StandaardForm.Icon                       = [System.Drawing.Icon]::ExtractAssociatedIcon('beheren.ico')

return $StandaardForm
}

function bepaallognaam ($invoer) {
<# de lognaam wordt bepaald adhv de invoer. dit is een datum in de format yyyy-mm-dd. Zie ook function Bepaaldatum.
   functie "bepaaldatumuitlognaam" doet het omgekeerde, dus een verandering hier moet ook in deze functie worden aangepast.
#>

# haal de personeelnr uit de beheervariabele $env:username
$personeelsnr = $env:username

$jaar = $invoer.substring(6, 4)
$maand = $invoer.substring(3, 2)
$dag = $invoer.substring(0, 2)
$datumintekst = -join ($jaar,'-',$maand,'-',$dag)
$lognaam = -join ("$logmap","\log_","$datumintekst",'_',$personeelsnr,".txt")

return $lognaam
}

Function bepaaleigenlogbestanden
{
$personeelsnr = $env:username

$uitvoer = "$logmap\*$personeelsnr.txt"

return $uitvoer
}


Function Logbestandtoevoegen ($logbestand) 
{
# tijdelijke log toevoegen aan eigen logbestand als deze al bestaat. bestand hernoemen wordt altijd gedaan.

# datum bepalen
$datumvandaag = bepaaldatum;

$logvanvandaag = bepaallognaam $datumvandaag
if (test-path -path $logvanvandaag -pathtype leaf) { 
    $inhoudlog = Get-Content -path $logvanvandaag
    Add-Content -Path $logbestand -Value $inhoudlog
    Remove-Item $logvanvandaag
}
# dit wordt altijd gedaan.....    
Rename-Item -Path $logbestand -NewName $logvanvandaag

}

function Foutenloggen {
param (
    [Parameter(Mandatory = $true)] [string]$meldtekst,
    [string]$type = "MELDING!"
)
# map aanmaken voor logbestanden als deze niet bestaat
if (!(Test-Path "$logmap")) { New-Item -Path "$logmap" -ItemType Directory | Out-Null  } 

# starttijd van loggen naar variabele
$logtijd = get-date -Format "HH:mm:ss"

# loggen. $foutmeldingsbestand is in het begin gedefinieerd.
"$type" | out-file $foutmeldingsbestand -Append
"Tijd : $logtijd " | out-file $foutmeldingsbestand -Append
"$meldtekst" | out-file $foutmeldingsbestand -Append

"-------------------------------------------------------------------------
" | out-file $foutmeldingsbestand -Append

Logbestandtoevoegen $foutmeldingsbestand
}

function bepaaldatum {
# datum bepalen
$datumvandaag = get-date -Format "dd-MM-yyyy"
return $datumvandaag
}

function bepaaltijd {

# huidige tijd bepalen
$huidigetijd = get-date -Format "HH:mm:ss"
return $huidigetijd
}


function declareren_rpcnrs {

# Venster voor weergeven rpcnummer declareren. Wordt bij meerdere functies gebruikt.

$listbox_temp = New-Object System.Windows.Forms.Listbox
$listbox_temp.Location = New-Object System.Drawing.Point(10,40)
$listbox_temp.Size = New-Object System.Drawing.Size(150,20)
$listbox_temp.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$listbox_temp.SelectionMode = 'MultiExtended'
$listbox_temp.Height = 500

return $listbox_temp
}

function lijstrpcnrsaanmaken ($keuzelocatie, $listbox_temp) {
# een lijst met rpc-nummers wordt gemaakt adhv gekozen locatie

# lijst wordt leeggemaakt
$listbox_temp.Items.Clear()

# waarden uit array halen die nodig zijn om lijst te maken met rpc-nummers
[int]$keuzeeerstenr=$Global:beheer.locaties.$keuzelocatie.eerstenr
[int]$keuzelaatstenr=$Global:beheer.locaties.$keuzelocatie.laatstenr

#listbox items aanmaken
For ($i=$keuzeeerstenr; $i -le $keuzelaatstenr; $i++) {

    if ($i -lt 10) {$rpcnr="RPC-00"+"$i"}
    elseif ($i -lt 100) {$rpcnr="RPC-0"+"$i"}
    else {$rpcnr="RPC-"+"$i"}
    [void] $listbox_temp.Items.Add($rpcnr)
}

return $listbox_temp
}


function venstermetvraag {

param (
    [Parameter(Mandatory = $true)] [string]$titel,
    [Parameter(Mandatory = $true)] [string]$vraag,
    [string]$knopok = "Ok",
    [string]$knopterug = "geen",
    [string]$schuifbalk = "geen"
)


# venster declareren
if ($schuifbalk -eq "beide") { 
    $formvraag = declareren_standaardvenster $titel 600 300
} else {
    $formvraag = declareren_standaardvenster $titel 600 200
}

$objtekst2 = New-Object System.Windows.Forms.textbox
$objtekst2.Location = New-Object System.Drawing.Size(50,5) 

$objtekst2.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$objtekst2.Text = $vraag
$objtekst2.BorderStyle = 'None' # geen lijnen zichtbaar
$objtekst2.BackColor  = 'white'
$objtekst2.Forecolor  = 'blue'
$objtekst2.TabStop = $false  # tekst wordt niet geselecteerd in het begin. is dus niet blauw.
$objtekst2.ReadOnly = $true # je kan niet in het venster typen
$objtekst2.Multiline = $true # meerdere regels kan je weergeven

# Schuifbalk alleen aangeven als je true hebt toegevoegd ah eind van de functie aanroep.
if ($schuifbalk -eq "beide") { 
    $objtekst2.ScrollBars = "Both" 
    $objtekst2.Size = New-Object System.Drawing.Size(530,210)
    } else {
    $objtekst2.Size = New-Object System.Drawing.Size(530,110)
    }

$formvraag.Controls.Add($objtekst2)

# knop ok wordt altijd weergegeven. knop annuleren alleen als een 2e knop-naam gegeven is.
$vraagok = New-object System.Windows.Forms.Button 
$vraagok.text= $knopok
if ($schuifbalk -eq "beide") { 
    $vraagok.location = New-Object System.Drawing.Point(50,220)
} else {
    $vraagok.location = New-Object System.Drawing.Point(50,120)
}
$vraagok.size = "150,30"  
if ($vraagok.text -eq "Ok") { $vraagok.BackColor = 'blue' }
   else { $vraagok.BackColor = 'green' }
$vraagok.ForeColor = 'white'
$vraagok.DialogResult = [System.Windows.Forms.DialogResult]::ok
$formvraag.Controls.Add($vraagok)


if (!($knopterug -eq "geen") ) { 
$vraagescape = New-object System.Windows.Forms.Button 
$vraagescape.text= $knopterug
# $vraagescape.location = "250,220" 
if ($schuifbalk -eq "beide") { 
    $vraagescape.location = New-Object System.Drawing.Point(250,220)
} else {
    $vraagescape.location = New-Object System.Drawing.Point(250,120)
}
$vraagescape.size = "150,30"  
$vraagescape.BackColor = 'red'
$vraagescape.ForeColor = 'white'
$vraagescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel
$formvraag.Controls.Add($vraagescape)
}

# Bij Escape-toets het venster sluiten.
$formvraag.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $formvraag.Close() }
})

$formvraag.KeyPreview = $true

$result = $formvraag.ShowDialog()
    
$formvraag.close();

return $result

} # einde venstermetvraag

Function declareren_uitlegvenster ($titel, $size_x, $size_y, $icoon_x, $icoon_y, $tekst) {
<# 
Hiermee kan je mbv een vraagteken in een venster extra uitleg geven.
variabele titel wordt als titel weergegeven.
variabele size_x en size_y geven de grootte vh venster op.
variabele icoon_x en icoon_y de positie van de vraagteken.
variabele tekst is de weer te geven informatietekst.
#>

# benodigde venster declareren
$Global:Form_uitleg_taak                    = New-Object system.Windows.Forms.Form
$Global:Form_uitleg_taak.StartPosition      = 'CenterScreen'
$Global:Form_uitleg_taak.BackColor          = "white"
$Global:Form_uitleg_taak.TopMost            = $true
$Global:Form_uitleg_taak.MaximumSize        = New-Object System.Drawing.size($size_x,$size_y)
$Global:Form_uitleg_taak.MinimumSize        = New-Object System.Drawing.size($size_x,$size_y)
$Global:Form_uitleg_taak.text               = $titel
$Global:Form_uitleg_taak.ControlBox         = $False
$Global:Form_uitleg_taak.Icon               = [System.Drawing.Icon]::ExtractAssociatedIcon('beheren.ico')

[int]$tekst_x = $size_x -10
[int]$tekst_y = $size_y -85

$Global:uitlegtaaktekst                     = New-Object system.Windows.Forms.Label
$Global:uitlegtaaktekst.AutoSize            = $false
$Global:uitlegtaaktekst.location            = New-Object System.Drawing.Point(10,10)
$Global:uitlegtaaktekst.Font                = 'Microsoft Sans Serif,11'
$Global:uitlegtaaktekst.ForeColor = [System.Drawing.Color]::blue
$Global:uitlegtaaktekst.text                = $tekst
$Global:uitlegtaaktekst.width               = $tekst_x
$Global:uitlegtaaktekst.height              = $tekst_y

$Global:Form_uitleg_taak.Controls.Add($Global:uitlegtaaktekst)

[int]$knop_x = 250
[int]$knop_y = $size_y -75

$knopsluiten = New-object System.Windows.Forms.Button 
$knopsluiten.text= 'Sluiten'
$knopsluiten.location = New-Object System.Drawing.size($knop_x,$knop_y)
$knopsluiten.size = "150,30"  
$knopsluiten.BackColor = 'red'
$knopsluiten.ForeColor = 'white'
$knopsluiten.DialogResult = [System.Windows.Forms.DialogResult]::ok

$Global:Form_uitleg_taak.Controls.Add($knopsluiten)

# Bij Escape-toets het venster sluiten.
$Global:Form_uitleg_taak.KeyPreview = $true
$Global:Form_uitleg_taak.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $Global:Form_uitleg_taak.Close() }
})

# vraagteken weergeven om extra info te geven.
$Global:vraagtekenicoon                     = new-object Windows.Forms.PictureBox
$Global:vraagtekenicoon.Location            = New-Object System.Drawing.Size($icoon_x, $icoon_y)
$Global:vraagtekenicoon.Size                = New-Object System.Drawing.Size(30,60)
$Global:vraagtekenicoon.Image               = [System.Drawing.Image]::FromFile("$icoontjesmap\icoon-hulp.png")
$Global:vraagtekenicoon.add_click( { $Global:Form_uitleg_taak.showdialog() } )
$Global:vraagtekenicoon.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Klik hier voor uitleg over de taak." )
})
# $Global:vraagtekenicoon.add_MouseLeave( { $Global:Form_uitleg_taak.hide() } )

} # einde declareren_uitlegvenster

function bepaalinitnaamgebruiker {

# bepalen van de initialisatienaam van een bestand met gebruikersinstellingen

# invoer is altijd de gebruiker van de pc
# mocht dit veranderen dan hoeft alleen de volgende regel verwijderd te worden en
# ($invoer) te worden toegevoegd aan het begin van de function.
$invoer = $env:username
 
# bestandsnaam wordt bv. gebruiker0101234.ini
$uitvoer = -join ("$startmap","\gebruiker_",$invoer,".ini")

return $uitvoer
}

function gebruikersinstellingen {
# hier worden de standaardinstellingen gedeclareerd die gebruikers kunnen wijzigen

$algemeen=@{
    locatiekeuze = 'PAL'
    wissennabackup = 'Ja'
    maplegenvoorverplaatsen = 'Ja'
    gebruiker = ''
    afbeelding = 'Samenwerken'
    consolesluiten = 'Ja'
    controlevoorklaarzetten = 'Ja'
    websitemoppen = 'Apekool.nl'
}
$opschonen=@{
    dagenbewarenlogs = 365
    opschonenlogs = 'Ja'
}
$Std_inst=@{
    algemeen  = $algemeen
    opschonen = $opschonen
}

return $Std_inst
}


Function Inlezengebruikersinstellingen {

# Bepalen van de persoonlijke initialisatiebestand.
$gebruikersbestand = bepaalinitnaamgebruiker

# standaard instellingen bepalen voor object $init
$global:init = gebruikersinstellingen

# inlezen initialisatiebestand als deze bestaat en toevoegen of verwijderen waarden van $init
# dit is nodig zodat, als de initialisatiebestand wordt ingelezen, eventuele nieuwe variabelen worden behouden en verwijderde variabelen niet worden toegevoegd. 
if (test-path -path $gebruikersbestand -pathtype leaf) { 
    # inlezen van object als hashwaarden. hier staan de gewijzigde waarden in die je wil behouden.
    $myObject = Get-Content -Path $gebruikersbestand | ConvertFrom-Json

    # waarden overzetten naar $init. nieuwe waarden worden behouden. verwijderde waarden worden niet toegevoegd.
    foreach( $property in $myobject.psobject.properties.name ) {
        foreach( $subproperty in $myobject.$property.psobject.properties.name )
        {
        # alleen toevoegen als deze bij 'init' al bestaat
        if ( $global:init.$property.$subproperty -ne $null) { 
            $global:init[$property][$subproperty] = $myObject.$property.$subproperty 
            }
        } # einde foreach $subproperty - loop 
        } # einde foreach $property - loop    

    # gekozen is om dit altijd te bewaren zodat je lijst met variabelen up to date is.
    $global:init | ConvertTo-Json -depth 1 | Set-Content -Path $gebruikersbestand
    } 

}

Function Netwerkmapaanwezig ($netwerkmap, $vensterweergeven) {

# Controleren of de netwerkmap aanwezig is en indien nodig herstellen
# De Write-Host regels zie je alleen in een testfase of bij start script, als de console open is

$gevonden = $false

if (!(test-path -path "$netwerkmap")) {
    
    if ($global:programma.mode -ne "alpha") {

    # bepaal de schijf letter
    $schijf = $netwerkmap.substring(0, 2)
    # controleren of schijf aanwezig is.
    $item = Get-SmbMapping | Where-Object -property LocalPath -Value $schijf -EQ
    if ($item) {
        # controleren of schijf ook verbonden is en anders herstellen
        if ($item.Status -eq 'Unavailable') {
            Write-Host "Herstellen van netwerkschijf " $item.LocalPath " naar " $item.RemotePath
            net use $item.LocalPath $item.RemotePath
            # net use $item.LocalPath $item.RemotePath >$null 2>&1

            # controleren of map aanwezig is na herstel.
            if (test-path -path "$netwerkmap") { $gevonden = $true } 
        }
    } # einde if ($item)

    }  # einde if ($global:programma.mode -eq "alpha") 

} else {
    $gevonden = $true 
} # einde 1e controle netwerkmap

if (!($gevonden)) {
    if ($vensterweergeven) {
        $vraagdoorgaan = "Er is geen verbinding met een netwerkschijf!" + "`r`n" + "Controleer of de volgende netwerkschijf aanwezig is :" + "`r`n" + "$netwerkmap" + "`r`n" + "`r`n" + "De taak kan nu niet opgestart worden."
        $null=venstermetvraag -titel "Geen verbinding met netwerkschijf" -vraag $vraagdoorgaan 

    } else {
    Write-Host -f Red "Er is geen verbinding met de netwerkschijf " $netwerkmap
    }
} # einde if (!($gevonden))


return $gevonden
}

Function declarerenlijstlocaties ($keuzelocatienr, $wijd, $loc_x, $loc_y) {

# De lijst met locaties aanmaken met de gegeven waarden. De meeste waarden zijn bij alle functies gelijk.

$lijstlocaties                     = New-Object system.Windows.Forms.ComboBox
$lijstlocaties.text                = "Locatie"
$lijstlocaties.width               = $wijd
$lijstlocaties.autosize            = $true
$lijstlocaties.Location = New-Object System.Drawing.Size($loc_x,$loc_y) 
$lijstlocaties.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$lijstlocaties.DropDownStyle="DropDownList"
$lijstlocaties.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de locatie." )
})

# locaties toevoegen aan lijst
$teller=0

foreach( $property in $global:beheer.locaties.keys ) {

    $waarde  = $global:beheer.locaties.$property.naam
    [void] $lijstlocaties.Items.Add("$property - $waarde")

    # standaard keuzelocatie selecteren en in index van lijst zetten
    if ($property -eq "$keuzelocatienr") { $lijstlocaties.Selectedindex = $teller }
    $teller++
    }

return $lijstlocaties
}


Function Form2afsluitenbijescape {
# Bij Escape-toets het venster sluiten.
# Dit wordt bij alle vensters, behalve hoofdvenster, gebruikt.

$form2.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $form2.Close() }
})

$form2.KeyPreview = $true
}


function SortListView {
# Bestanden sorteren op de kolom waar je op klikt. Dit wordt bij functie Verkenner gebruikt.

    Param(
        [System.Windows.Forms.ListView]$sender,
        $column
    )
    $temp = $sender.Items | Foreach-Object { $_ }
    $Script:SortingDescending = !$Script:SortingDescending
    $sender.Items.Clear()
    $sender.ShowGroups = $false
    $sender.Sorting = 'none'
    $sender.Items.AddRange(($temp | Sort-Object -Descending:$script:SortingDescending -Property @{ Expression={ $_.SubItems[$column].Text } }))
}

Function Declareericoontjes {

# Op een plek de icoontjes die de bestanden weergeven declareren.
# nodig bij Functies Kopieren en Verkenner

$std_imageList = new-Object System.Windows.Forms.ImageList 
$std_imageList.ImageSize = New-Object System.Drawing.Size(30,30) 
$bitm0=[System.Drawing.Image]::FromFile("$icoontjesmap\explorer-icoon.png")
$bitm1=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon.png")
$bitm2=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-word.png")
$bitm3=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-excel.png")
$bitm4=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-pdf.png")
$bitm5=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-visio.png")
$bitm6=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-afbeelding.png")
$bitm7=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-exe.png")
$bitm8=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-link.png")
$bitm9=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-muziek.png")
$bitm10=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-pka.png")
$bitm11=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-sjabloon.png")
$bitm12=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-video.png")
$bitm13=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-zip.png")
$bitm14=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-powerpoint.png")
$bitm15=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-website.png")
$bitm16=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-programmeren.png")
$bitm17=[System.Drawing.Image]::FromFile("$icoontjesmap\file-icoon-database.png")
$std_imageList.Images.Add("explorer", $bitm0) 
$std_imageList.Images.Add("file-icoon", $bitm1) 
$std_imageList.Images.Add("word", $bitm2) 
$std_imageList.Images.Add("excel", $bitm3) 
$std_imageList.Images.Add("pdf", $bitm4) 
$std_imageList.Images.Add("visio", $bitm5) 
$std_imageList.Images.Add("afbeelding", $bitm6) 
$std_imageList.Images.Add("exe", $bitm7) 
$std_imageList.Images.Add("link", $bitm8) 
$std_imageList.Images.Add("muziek", $bitm9) 
$std_imageList.Images.Add("pka", $bitm10) 
$std_imageList.Images.Add("sjabloon", $bitm11) 
$std_imageList.Images.Add("video", $bitm12) 
$std_imageList.Images.Add("zip", $bitm13) 
$std_imageList.Images.Add("powerpoint", $bitm14) 
$std_imageList.Images.Add("website", $bitm15) 
$std_imageList.Images.Add("programmeren", $bitm16) 
$std_imageList.Images.Add("database", $bitm17) 

return $std_imageList
}


Function Bepaalicoontjenr ($controlemap, $bestand) {

# Bepaal welke icoontje moet worden weergeven afhankelijk van type, map of bestand. 
# Wordt gebruikt bij functies Bestanden kopieren en Verkenner.
# De nummering is afhankelijk van de volgorde van aanmaken bij functie Declareericoontjes. 
# een map heeft als nummer 0. Deze is als eerst gedeclareerd. zie hierboven.

if (( test-path -path "$controlemap\$bestand" -pathtype container) -eq $true)  {

    $gevondennr = 0
    } else {
    # regel hieronder werkt niet als er meerdere punten in een bestandsnaam zijn. Daarom volgt daaronder een verbetering.
    # $extensie = (Split-Path -Path $bestand -Leaf).Split(".")[1]
    $positiepunt = $bestand.LastIndexOf(".")
    $positiepunt += 1
    $extensie = $bestand.Substring($positiepunt)

    switch ($extensie) {
            "docx"  { $gevondennr = 2 }
            "xlsx"  { $gevondennr = 3 }
            "xlsm"  { $gevondennr = 3 }
            "csv"  { $gevondennr = 3 }
            "pdf"   { $gevondennr = 4 }
            "vsdx"  { $gevondennr = 5 }
            "vss"  { $gevondennr = 5 }
            "png"  { $gevondennr = 6 }
            "jpeg"  { $gevondennr = 6 }
            "jpg"   { $gevondennr = 6 }
            "bmp"   { $gevondennr = 6 }
            "gif"  { $gevondennr = 6 }
            "ico"  { $gevondennr = 6 }
            "exe"  { $gevondennr = 7 }
            "lnk"  { $gevondennr = 8 }
            "wav"   { $gevondennr = 9 }
            "wma"  { $gevondennr = 9 }
            "mp3"  { $gevondennr = 9 }
            "pka"  { $gevondennr = 10 }
            "pkt"  { $gevondennr = 10 }
            "dotx"   { $gevondennr = 11 }
            "mov"  { $gevondennr = 12 }
            "avi"  { $gevondennr = 12 }
            "mpeg"  { $gevondennr = 12 }
            "mp4"  { $gevondennr = 12 }
            "mpg"  { $gevondennr = 12 }
            "zip"  { $gevondennr = 13 }
            "7zip"  { $gevondennr = 13 }
            "gz"  { $gevondennr = 13 }
            "rar"  { $gevondennr = 13 }
            "pptx"  { $gevondennr = 14 }
            "html"  { $gevondennr = 15 }
            "php"  { $gevondennr = 15 }
            "css"  { $gevondennr = 15 }
            "asp"  { $gevondennr = 15 }
            "xps"  { $gevondennr = 15 }
            "ps1"  { $gevondennr = 16 }
            "py"  { $gevondennr = 16 }
            "js"  { $gevondennr = 16 }
            "sql"  { $gevondennr = 17 }
            default { $gevondennr = 1 }
      } # einde switch
    }

return $gevondennr
}

# hieronder de hoofdfuncties ----------------------------------------------------------------



function scriptRun {


# nodig voor afhandelen foutmeldingen
$uitvoeren.foutmelding = $false

# nodig om in logbestand te schrijven
$logbestand = $uitvoeren.logbestand

# hier wordt de taak uitgevoerd.
switch ($uitvoeren.taak) {

    "backup"   { 

    # datum en tijdstip van vandaag in variabele plaatsen. nodig om de backup-mapnaam te bepalen.
    $datumvandaag = get-date -Format "yyyy-MM-dd__HH-mm-ss"

    #kopiëren bestanden naar homemappen van de kandidaten
    foreach ($rpcitem in $uitvoeren.listbox.selecteditems) {
        # progressie laten zien op balk
        $uitvoeren.progressbar.PerformStep()
        
        # teksten met meldingen voor het logbestand
        $foutmelding_log = "[ $rpcitem - FOUT ] : "

        # bepalen bronmap. dit is de homemap van de kandidaten
        $bronmap = -join ($uitvoeren.homemap,'\',$rpcitem)

        # bepalen doelmap. dit is de map waar de backup in komt.
        $doelmap = -join ($uitvoeren.doelmap,'\',$datumvandaag,'\',$rpcitem)

        # errors leeg maken
        $error.clear()
            try {
                # extra test of de specifieke rpc-map wel bestaat. Anders wordt de catch getriggerd en heb je een foutmelding.
                if (!(Test-Path $bronmap)) { throw "Map $bronmap niet gevonden."}

                #backup maken
                # de toevoeging \\? voor de bronmap en doelmap zorgt ervoor dat lange namen - meer dan maximaal 256 tekens, geen foutmelding geven en dus het kopieren en verwijderen ook dan lukt.
                Copy-Item -path "\\?\$bronmap" -destination "\\?\$doelmap" -recurse -Force -container -ErrorAction Stop

                # en daarna wissen als dit geselecteerd is
                if ($uitvoeren.wissennabackup -eq $true) { 
                    # verwijderen van bestanden. Eerst de inhoud van mijn documenten
                    Remove-Item "\\?\$bronmap\mijn documenten\*" -Recurse -Force -ErrorAction Stop
                    # dan de root van rpc-map exclusief map mijn documenten
                    Remove-Item "\\?\$bronmap\*" -Recurse -Force -Exclude "mijn documenten" -ErrorAction Stop
                    }
                } # einde try

            catch {

                # foutmelding van PowerShell naar logbestand
                "$foutmelding_log" + $_.exception.message | out-file "$logbestand" -Append
                $uitvoeren.foutmelding = $true
                  } # einde catch

    } # einde foreach rpcitem

               } # einde taak backup

    "kopiëren" { 

    # array aanmaken en vullen met geselecteerde items. 
    # Deze items worden vervolgens gekopiëerd.
    $geselecteerdeitems = [System.Collections.ArrayList]@()

    <# Als er geen items zijn geselecteerd in de examenmap dan moet de gehele map worden gekopiëerd.
   hiervoor wordt dan de pad naar de examenmap toegevoegd aan de array $geselecteerdeitems.
   anders worden de geselecteerde items toegevoegd aan de array.
    #>
    if ($uitvoeren.listview1.selecteditems.count -eq 0) {
        $bronmap = -join ($uitvoeren.bronmap,'\*')
        $geselecteerdeitems.Add("$bronmap")
        } else {
        foreach ($item in $uitvoeren.listview1.selecteditems) {
            $bronmap = -join ($uitvoeren.bronmap,'\',$item[0].text)
            $geselecteerdeitems.Add("$bronmap")
        }
    }

    #kopiëren bestanden naar homemappen van de kandidaten

    foreach ($rpcitem in $uitvoeren.listbox.selecteditems) {
        # progressie laten zien op balk
        $uitvoeren.progressbar.PerformStep()

        # teksten met meldingen voor het logbestand
        $foutmelding_log = "[ $rpcitem - FOUT ] : "

        # doelmap bepalen. dit is een lokale variabele
        $doelmap = -join ($uitvoeren.homemap,'\',$rpcitem,'\Mijn Documenten')
    
        # kopieren geselecteerde mappen of bestanden
            
        foreach ($item in $geselecteerdeitems) {
        $error.clear()
        try {
            # extra test of de specifieke rpc-map wel bestaat. Dit is niet nodig bij kopieren blijkt na test.
            # if (!(Test-Path $doelmap)) { throw "Map $doelmap niet gevonden."}

            Copy-Item -path "$item" -destination "$doelmap" -recurse -Force -ErrorAction Stop
             
            }

        catch {
            # foutmelding van PowerShell naar logbestand
            "$foutmelding_log" + $_.exception.message | out-file "$logbestand" -Append
            $uitvoeren.foutmelding = $true
              } # einde catch
            } # einde foreach $item
        
    } # einde foreach $rpcitem

    # array legen. mss is dit niet nodig!
    $geselecteerdeitems.Clear()
    # geselecteerde rpc-nummers wissen. je moet dan opnieuw selecteren en kan niet meteen op bevestigen klikken.
    $uitvoeren.listbox.selecteditems.clear()

               } # einde taak kopiëren

    "wissen"   { 
        #wissen van bestanden in homemappen van de kandidaten
        foreach ($rpcitem in $uitvoeren.listbox.selecteditems) {
            # progressie laten zien op balk
            $uitvoeren.progressbar.PerformStep()

            # teksten met meldingen voor het logbestand
            $foutmelding_log = "[ $rpcitem - FOUT ] : "
            
            # doelmap bepalen. dit is een lokale variabele voor de deel-functie wissen hieronder.
            $doelmap = -join ($uitvoeren.homemap,'\',$rpcitem,'\Mijn Documenten')
            $doelmaproot = -join ($uitvoeren.homemap,'\',$rpcitem)

            $error.clear()
            try {
                # extra test of de specifieke rpc-map en submap wel bestaat
                if (!(Test-Path $doelmap)) { throw "Map $doelmap niet gevonden."}
                if (!(Test-Path $doelmaproot)) { throw "Map $doelmaproot niet gevonden."}

                # verwijderen van bestanden. Eerst de inhoud van mijn documenten
                Remove-Item "\\?\$doelmap\*" -Recurse -Force -ErrorAction Stop
                # dan de root van rpc-map exclusief map mijn documenten
                Remove-Item "\\?\$doelmaproot\*" -Recurse -Force -Exclude "mijn documenten" -ErrorAction Stop
                }
            catch {

                # foutmelding van PowerShell naar logbestand
                "$foutmelding_log" + $_.exception.message | out-file "$logbestand" -Append
                $uitvoeren.foutmelding = $true
                  } # einde catch
            
            } # einde foreach statement

               } # einde taak wissen

    "verplaatsen" {
        # er is maar een actie te doen dus...
        $uitvoeren.progressbar.maximum = 1

        # teksten met meldingen voor het logbestand
        $foutmelding_log = "[ FOUT ] : "

        # definiëren bron- en doelmap
        $bronmap = $uitvoeren.bronmap
        $doelmap = $uitvoeren.doelmap

        # doelmap wordt geleegd voor het verplaatsen of kopiëren, als dit gekozen is.
        if ($uitvoeren.doelmaplegen -eq $true) {
            $error.clear()
            try {
                # extra test of de specifieke rpc-map wel bestaat
                if (!(Test-Path $doelmap)) { throw "Map $doelmap niet gevonden."}
                # verwijderen van bestanden in doelmap
                Remove-Item "\\?\$doelmap\*" -Recurse -Force -ErrorAction Stop
                }
            catch {

                # foutmelding van PowerShell naar logbestand
                "$foutmelding_log" + $_.exception.message | out-file "$logbestand" -Append
                $uitvoeren.foutmelding = $true

                # Verder ook niet meer uitvoeren

                # progressie laten zien op balk
                $uitvoeren.progressbar.PerformStep()
                return;
                  } # einde catch
        }

        # kopieren of verplaatsen
        $error.clear()
        try {
            # extra test of de specifieke rpc-map wel bestaat
            if (!(Test-Path $bronmap)) { throw "Map $bronmap niet gevonden."}
            
            # eerst kopiëren
            Copy-Item -path "\\?\$bronmap" -destination "\\?\$doelmap" -recurse -ErrorAction Stop

            # verwijderen van bestanden als keuze is verplaatsen. alleen de inhoud van mijn documenten
            if ($uitvoeren.keuzeverplaatsen -eq "verplaatsen") {
                Remove-Item "\\?\$bronmap" -Recurse -Force -ErrorAction Stop
                }
            }
        catch {
              
              $uitvoeren.foutmelding = $true
              # foutmelding van PowerShell naar logbestand
              "$foutmelding_log" + $_.exception.message | out-file "$logbestand" -Append
              
              } # einde catch

        # progressie laten zien op balk
        $uitvoeren.progressbar.PerformStep()

                  } # einde taak verplaatsen

    "opschonen"   { 
        #wissen van backups van homemappen van de kandidaten
        foreach ($rpcitem in $uitvoeren.listbox.items) {
            # progressie laten zien op balk
            $uitvoeren.progressbar.PerformStep()

            # teksten met meldingen voor het logbestand
            $foutmelding_log = "[ $rpcitem - FOUT ] : "
            
            # doelmap bepalen. dit is een lokale variabele voor de deel-functie wissen hieronder.
            $doelmap = -join ($uitvoeren.doelmap,'\',$rpcitem)

            $error.clear()
            try {
                # compleet verwijderen van mappen. 
                Remove-Item "\\?\$doelmap" -Recurse -Force -ErrorAction Stop

                }
            catch {

                # foutmelding van PowerShell naar logbestand
                "$foutmelding_log" + $_.exception.message | out-file "$logbestand" -Append
                $uitvoeren.foutmelding = $true
                  } # einde catch
            
            } # einde foreach statement

               } # einde taak wissen

} # einde switch commando $uitvoeren.taak

} # einde scriptRun

function uitvoerentaken {
# hier worden de 4 belangrijkste taken uitgevoerd.

# controleren of mappen bestaan - afhankelijk van de taak
# map uitvoeren.homemap wordt in ieder geval gecontroleerd
# Eerste regel van elke foutmelding
$foutmeldingbegin = "Uitvoeren van een taak : "

if (!(Netwerkmapaanwezig $global:beheer.examenmappen.homemapstudenten $true )) { 
    $tempmap = $global:beheer.examenmappen.homemapstudenten
    Foutenloggen "$foutmeldingbegin
De volgende Netwerkmap is niet gevonden
$tempmap
    "
    return; 
} 

if ($uitvoeren.taak -eq "kopiëren") {
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.digitalebestanden $true )) { 
    $tempmap = $global:beheer.examenmappen.digitalebestanden
    Foutenloggen "$foutmeldingbegin
De volgende Netwerkmap is niet gevonden
$tempmap
    "
    return; 
    } 
} 

if ($uitvoeren.taak -eq "backup") {
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.backupmap $true )) { 
    $tempmap = $global:beheer.examenmappen.backupmap
    Foutenloggen "$foutmeldingbegin
De volgende Netwerkmap is niet gevonden
$tempmap
    "
    return; 
    } 
} 

# einde controleren of mappen bestaan

# Extra beveiliging voordat je de taak opschonen gaat uitvoeren

if ($uitvoeren.taak -eq "opschonen") {
   $vraagopschonen="LET OP : Als u nu op Opschonen klikt worden de bestanden definitief verwijderd!" + "`r`n" + "`r`n" + "Dit kan niet worden teruggedraaid." + "`r`n" + "Klik op Terug als u terug wilt."
 
   $result=venstermetvraag -titel "Starten met opschonen?" -vraag $vraagopschonen -knopok "Opschonen" -knopterug "Terug"

   if ( $result -eq "Cancel") { return }
}

# Extra beveiliging voordat je de taak wissen gaat uitvoeren

if ($uitvoeren.taak -eq "wissen") {
   $vraagwissen="LET OP : Als u nu op Wissen klikt worden de bestanden definitief verwijderd!" + "`r`n" + "`r`n" + "Dit kan niet worden teruggedraaid." + "`r`n" + "Klik op Terug als u terug wilt."

   $result=venstermetvraag -titel "Starten met wissen?" -vraag $vraagwissen -knopok "Wissen" -knopterug "Terug"

   if ( $result -eq "Cancel") { return }
}

# extra controle voordat je de taak bestanden klaarzetten gaat uitvoeren
if (($uitvoeren.taak -eq "kopiëren") -and ($uitvoeren.controlemappen)) {
    $teller = 0
    $nietlegemappen = ""

    foreach ($rpcitem in $listbox.selecteditems) {
        # doelmap bepalen. dit is een lokale variabele
        $doelmap = -join ($global:beheer.examenmappen.homemapstudenten,'\',$rpcitem,'\Mijn Documenten')
    
        # Controleren of de mappen leeg zijn. Foutmeldingen worden gelogd maar niet getoond!
        try {
            If ((Get-ChildItem -Path $doelmap -Force -ErrorAction Stop | Measure-Object).Count -gt 0) {
            # rpc-nr toevoegen aan overzicht van mappen die niet leeg zijn
            $nietlegemappen = -join ($nietlegemappen, $rpcitem, ' - ')
            $teller++
        } 
        }
        catch {
            $melding = -join ("Extra controle voor het uitvoeren van de taak Bestanden Klaarzetten : ", "`n", $_.exception.message )
            Foutenloggen $melding
        }
        
    } # einde foreach $rpcitem ...

    # als een aantal mappen niet leeg zijn
    if ($teller -gt 0 ) {
    $vraagdoorgaan="LET OP : Op een aantal RPC-nummers zijn al bestanden aanwezig!" + "`r`n" + "Als u doorgaat worden de bestanden toegevoegd." + "`r`n" +
    "Klik op Terug als u terug wilt." + "`r`n" + "`r`n" + "De volgende $teller RPC-nummers zijn niet leeg: $nietlegemappen" 

    $result=venstermetvraag -titel "Controle of homemappen van de kandidaten leeg zijn." -vraag $vraagdoorgaan -knopok "Doorgaan" -knopterug "Terug" -schuifbalk "beide"

    if ( $result -eq "Cancel") { return }
    }

} 
# Einde extra controle voordat je de taak bestanden klaarzetten gaat uitvoeren

# startknop wordt onzichtbaar en venster met functie gesloten
$StartButton.Hide()
$Btnescape.Hide()
$form2.hide();


# het proces zichtbaar maken in een balk
$uitvoeren.progressbar = New-Object System.Windows.Forms.ProgressBar
$uitvoeren.progressbar.Location = New-Object System.Drawing.Point(20, 40)
$uitvoeren.progressbar.Size = New-Object System.Drawing.Size(560, 30)
$uitvoeren.progressbar.Style = "continuous"

# de maximumwaarde van de progressbar. Deze is bij taak opschonen anders.
if ($uitvoeren.taak -eq "opschonen") {
    $uitvoeren.progressbar.maximum = $listbox.items.count
    } else {
    $uitvoeren.progressbar.maximum = $listbox.selecteditems.count
    }
$uitvoeren.progressbar.step = 1

# ProgressBar toevoegen aan form3
$uitvoeren.form3.Controls.Add($uitvoeren.progressbar);

# teksten en progressbar veranderen
$Label.hide()
$Description3.hide()
$Description2.ForeColor = 'green'
$Description2.Text = "De taak wordt uitgevoerd ..."
$uitvoeren.progressbar.visible

# tijdelijk logbestand bepalen
$tijdelijkelog = "tijdelijkelog.txt"

# logbestandsnaam definiëren en volledige pad naar bestand invoeren
$logbestand = -join ("$logmap","\",$tijdelijkelog)

# doorgeven aan hastable uitvoeren
$uitvoeren.logbestand = $logbestand

# starttijd van loggen naar variabele
$logtijd = bepaaltijd

# map aanmaken voor logbestanden als deze niet bestaat
if (!(Test-Path "$logmap")) { New-Item -Path "$logmap" -ItemType Directory | Out-Null  } 

# in logbestand info over de taak schrijven en beginnen met uitvoeren van taak ---------------------------

"De taak " + $uitvoeren.taak + " is gestart." | out-file $logbestand -Append
"Starttijd : $logtijd" | out-file $logbestand -Append

if ($uitvoeren.taak -eq "kopiëren") {
    $objtekst2.Text | out-file $logbestand -Append
    } elseif ($uitvoeren.taak -eq "backup") {
    if ($wissennabackup.checked -eq $true) { "De bestanden worden na de back-up gewist." | out-file $logbestand -Append }
    } elseif ($uitvoeren.taak -eq "verplaatsen") {
    if ($keuzeoptie1.Selectedindex -eq "1") { 
        "U gaat bestanden kopiëren van een rpc-nummer naar een andere." | out-file $logbestand -Append
        } else {
        "U gaat bestanden verplaatsen van een rpc-nummer naar een andere." | out-file $logbestand -Append
        }
     if ($doelmaplegen.checked -eq $true) { 
        "De doelmap wordt voor het verplaatsen of kopiëren eerst geleegd." | out-file $logbestand -Append
                    } else {
        "De bestanden worden toegevoegd aan de bestanden in de doelmap." | out-file $logbestand -Append
                    }
    }
$objtekst1.Text | out-file $logbestand -Append

# ------ begin runspaces ----------------------------------------------

#Configure max thread count for RunspacePool.
$maxthreads = [int]$env:NUMBER_OF_PROCESSORS
    
#Create a new session state for parsing variables ie hashtable into our runspace.
$hashVars = New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'uitvoeren',$uitvoeren,$Null
$InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
#Add the variable to the RunspacePool sessionstate
$InitialSessionState.Variables.Add($hashVars)

#Create our runspace pool. We are entering three parameters here min thread count, max thread count and host machine of where these runspaces should be made.
$script:runspace = [runspacefactory]::CreateRunspacePool(1,$maxthreads,$InitialSessionState, $Host)

#Create a PowerShell instance.
$script:powershell = [powershell]::Create()
    
#Open a RunspacePool instance.
$script:runspace.Open()
         
#Add our main code to be run via $scriptRun within our RunspacePool.
$script:powershell.AddScript(${Function:scriptRun})
$script:powershell.RunspacePool = $script:runspace
        
#Run our RunspacePool.
$script:handle = $script:powershell.BeginInvoke()

#Wait for code to complete and keep UI responsive
do {
		[System.Windows.Forms.Application]::DoEvents()
		Start-Sleep -Milliseconds 1 
} while (!$script:handle.IsCompleted)

#Cleanup our RunspacePool threads when they are complete ie. GC.
if ($script:handle.IsCompleted)
        {
            $script:powershell.EndInvoke($script:handle)
            $script:powershell.Dispose()
            $script:runspace.Dispose()
            $script:runspace.Close()
            [System.GC]::Collect()
        }

# ----- Einde runspaces -----------------------------------------------

# in logbestand eindtijd schrijven
# eerst eindtijd naar variabele
$logtijd = bepaaltijd


"" | out-file $logbestand -Append
"Eindtijd  : $logtijd" | out-file $logbestand -Append
" -------------------------------------------------------------------------" | out-file $logbestand -Append
"" | out-file $logbestand -Append

# tijdelijke log toevoegen aan eigen logbestand als deze al bestaat. 
Logbestandtoevoegen $logbestand


# aangeven dat taken zijn uitgevoerd 
if ($uitvoeren.foutmelding) { 
    $Description2.ForeColor = 'red'
    $Description2.Text = "De taak is niet correct uitgevoerd. Bekijk de foutmeldingen in het logbestand."
    } else { $Description2.Text = "De taak is uitgevoerd." }

# knoppen zichtbaar maken
$EndButton.show()
$LogButton.show()

# alleen bij taak kopieren onderstaande knop zichtbaar maken
if ($uitvoeren.taak -eq "kopiëren") {
    $Btnescape.size = New-Object System.Drawing.Size(200,40)
    # escape knop wordt nu opnieuw knop
    $Btnescape.text= "Opnieuw bestanden klaarzettenn"
    $Btnescape.BackColor = 'green'
    $Btnescape.show()
    
} else {
# knop logbestand bekijken wordt nu naar links verplaatst om aan te sluiten met knop Terug.
$LogButton.Location = New-Object System.Drawing.Size(160, 400)
}

$form2.show()

# simuleer indrukken van tab-toets om de focus terug te brengen naar huidige venster.
[System.Windows.Forms.SendKeys]::SendWait("%{TAB}")

} # einde uitvoerentaken

function overzichttaken ([string]$taak) {
# na het bevestigen van je keuze bij een van de 4 taken kom je bij deze functie.
# hier zie je een overzicht en kan je de taak starten of nog terug.

# beheer-variabelen krijgen hier een verkorte naam tbv leesbaarheid en gebruik in andere functies
$digitalebestanden = $global:beheer.examenmappen.digitalebestanden
$homemapstudenten = $global:beheer.examenmappen.homemapstudenten
$backupmap = $global:beheer.examenmappen.backupmap

# hashtable object leeg maken voor het geval het nog waarden heeft.
$uitvoeren.clear()

# definiëren venster
switch ($taak) {
    "backup"      { $titel = "Overzicht uit te voeren taak Back-up maken " }
    "kopiëren"    { $titel = "Overzicht uit te voeren taak Bestanden klaarzetten" }
    "wissen"      { $titel = "Overzicht uit te voeren taak Wissen" }
    "verplaatsen" { $titel = "Overzicht uit te voeren taak Verplaatsen of kopiëren" }
    "opschonen"   { $titel = "Overzicht uit te voeren taak Opschonen" }
}

$uitvoeren.Form3 = declareren_standaardvenster $titel 600 500

$Description2                     = New-Object system.Windows.Forms.Label
$Description2.AutoSize            = $false
$Description2.width               = 570
$Description2.height              = 20
$Description2.location            = New-Object System.Drawing.Point(20,10)
$Description2.Font                = 'Microsoft Sans Serif,11'
$Description2.ForeColor           = 'blue'

# onderstaande tekst wordt alleen zichtbaar als backuptaak is gekozen en wissen na backup is geselecteerd.
$Description3                     = New-Object system.Windows.Forms.Label
$Description3.AutoSize            = $false
$Description3.width               = 500
$Description3.height              = 20
$Description3.location            = New-Object System.Drawing.Point(20,28)
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.Text                = "De bestanden worden na de back-up gewist."
$Description3.ForeColor           = 'blue'
$Description3.hide()

$StartButton = New-Object System.Windows.Forms.Button
$StartButton.Location = New-Object System.Drawing.Size(10, 400)
$StartButton.Size = New-Object System.Drawing.Size(120, 50)
$StartButton.Text = "Start"
$StartButton.height = 40
$StartButton.BackColor = 'green'
$StartButton.ForeColor = 'white'
$StartButton.Add_click( {
    # toevoegen object en variabelen aan hastable voor uitvoeren 
    $uitvoeren.taak=$taak
    $uitvoeren.homemap=$homemapstudenten
    $uitvoeren.listbox = New-Object System.Windows.Forms.Listbox
    $uitvoeren.listbox = $listbox

    <# specifieke variabelen voor de gegeven taak.
       o.a. toevoegen tekst aan variabele uitvoeren.logbestandtekst om in logbestand te plaatsen
    #>
    switch ($taak) {

        "kopiëren" {
        $uitvoeren.ListView1 = New-Object System.Windows.Forms.ListView
        $uitvoeren.ListView1 = $global:ListView1
        $uitvoeren.controlemappen = $controlemappen.checked
        $uitvoeren.bronmap = $digitalebestanden
        foreach ($item in $global:geselecteerdebronmap) {
                    $uitvoeren.bronmap  = -join ($uitvoeren.bronmap, '\', $item)
                }
        }

        "backup" {
        $uitvoeren.doelmap="$backupmap"
        $uitvoeren.wissennabackup=$wissennabackup.checked
        }

        "verplaatsen" {
        # bepalen doel-, bronmap en doelmaplegen
        $uitvoeren.bronmap = -join ($homemapstudenten,"\",$bronselectie.selecteditem,"\Mijn Documenten\*")
        $uitvoeren.doelmap = -join ($homemapstudenten,"\",$doelselectie.selecteditem,"\Mijn Documenten")
        $uitvoeren.doelmaplegen=$doelmaplegen.checked
        # keuze tussen kopieren of verplaatsen doorgeven
        if ($keuzeoptie1.Selectedindex -eq "1") { 
                    $uitvoeren.keuzeverplaatsen = "kopiëren" 
                    } else {
                    $uitvoeren.keuzeverplaatsen = "verplaatsen" 
                    }
        }

        "opschonen" {
        $uitvoeren.doelmap="$backupmap"
        
        }
    } # einde switch $taak

    # $form2.hide()

    uitvoerentaken;
    
    });

$EndButton = New-Object System.Windows.Forms.Button
$EndButton.Location = New-Object System.Drawing.Size(10, 400)
$EndButton.Size = New-Object System.Drawing.Size(120, 50)
$EndButton.Text = "Sluiten"
$EndButton.height = 40
$EndButton.BackColor = 'red'
$EndButton.ForeColor = 'white'
$EndButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$EndButton.hide()

$Btnescape = New-object System.Windows.Forms.Button 
$Btnescape.text= "Terug"
$Btnescape.location = New-Object System.Drawing.Size(160, 400)
$Btnescape.size = New-Object System.Drawing.Size(120, 50)
$Btnescape.height = 40
$Btnescape.BackColor = 'red'
$Btnescape.ForeColor = 'white'
$Btnescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel

$LogButton = New-Object System.Windows.Forms.Button
$LogButton.Location = New-Object System.Drawing.Size(400, 400)
$LogButton.Size = New-Object System.Drawing.Size(170, 50)
$LogButton.Text = "Logbestand bekijken"
$LogButton.height = 40
$LogButton.BackColor = 'green'
$LogButton.ForeColor = 'white'
$LogButton.hide()
$LogButton.add_click({ 
    $Form2.dispose()
    $uitvoeren.Form3.dispose()

    vensterlogbestand })

$Label = New-Object System.Windows.Forms.Label
$Label.Font = 'Microsoft Sans Serif,12'
$Label.ForeColor = 'blue'
$Label.Text = "Klaar om de geselecteerde taak uit te voeren ?"
$Label.Location = New-Object System.Drawing.Point(20, 50)
$Label.Width = 480
$Label.Height = 20

switch ($taak) {
    "backup"      { 
                  $Description2.text = "U gaat een back-up uitvoeren op de volgende RPC-nummers." 
                  if ($wissennabackup.checked -eq $true) { 
                    $Description3.text = "De bestanden worden na de back-up gewist."
                    $Description3.show() 
                    }
                  }
    "kopiëren"    { $Description2.text = "U gaat bestanden of mappen klaarzetten op de volgende RPC-nummers." }
    "wissen"      { $Description2.text = "U gaat bestanden wissen van de volgende RPC-nummers." }
    "verplaatsen" { 
                  if ($keuzeoptie1.Selectedindex -eq "1") { 
                    $Description2.text = "U gaat bestanden kopiëren van een RPC-nummer naar een andere." 
                    } else {
                    $Description2.text = "U gaat bestanden verplaatsen van een RPC-nummer naar een andere." 
                    }
                  $Description3.show() 
                  if ($doelmaplegen.checked -eq $true) { 
                    $Description3.text = "De doelmap wordt voor het verplaatsen of kopiëren eerst geleegd."
                    } else {
                    $Description3.text = "De bestanden worden toegevoegd aan de bestanden in de doelmap."
                    }
                  }
    "opschonen"   { $Description2.text = "U gaat oude back-ups verwijderen." }
} # einde switch taak

$objtekst1 = New-Object System.Windows.Forms.textbox
$objtekst1.Location = New-Object System.Drawing.Size(20,80) 
$objtekst1.Size = New-Object System.Drawing.Size(200,290)
$objtekst1.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$objtekst1.Text = "RPC-nummers:" + "`r`n" + "`r`n"
$objtekst1.ReadOnly = $true
$objtekst1.Multiline = $true
$objtekst1.ScrollBars = "Both"
$objtekst1.BackColor  = 'white'

$objtekst2 = New-Object System.Windows.Forms.textbox
$objtekst2.Location = New-Object System.Drawing.Size(230,80) 
$objtekst2.Size = New-Object System.Drawing.Size(350,290)
$objtekst2.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$objtekst2.Text = ""
$objtekst2.ReadOnly = $true
$objtekst2.Multiline = $true
$objtekst2.ScrollBars = "Both"
$objtekst2.Visible = $false
$objtekst2.BackColor  = 'white'
$objtekst2.WordWrap = $false

# geselecteerde rpc-nummers links weergeven
# als taak is verplaatsen of opschonen heb je een andere inhoud dan bij overige taken
if ($taak -eq "verplaatsen") {
    $objtekst1.Text = $objtekst1.Text + " Bron is " + $bronselectie.selecteditem + "`r`n" + "`r`n"
    $objtekst1.Text = $objtekst1.Text + " Doel is " + $doelselectie.selecteditem + "`r`n"

    } elseif ($taak -eq "opschonen") {
    $objtekst1.Text = "Er zijn " + $listbox.items.count + " mappen met back-ups ouder dan " + $objTextBox1.Text + " dagen: "+ "`r`n" + "`r`n"
    foreach ($item in $listbox.items) {
        $objtekst1.Text = $objtekst1.Text + " - " + "$item" + "`r`n"
        }

    } else {
    foreach ($item in $listbox.selecteditems) {
        $objtekst1.Text = $objtekst1.Text + " - " + "$item" + "`r`n"
        }
    }

# geselecteerde mappen bij taak kopiëren rechts weergeven in overzicht. alleen als taak kopieren is geslecteerd
if ($taak -eq "kopiëren") {
    $objtekst2.Text = "Examenmap:" + "`r`n"
    $objtekst2.Visible = $true

    # bepalen van geselecteerde bronmap.
    [string] $selectie = ""
    foreach ($item in $global:geselecteerdebronmap) {
          $selectie = -join ($selectie, '\', $item)
          }
    $objtekst2.Text = $objtekst2.Text + " - " + "$selectie" + "`r`n" + "`r`n"

    # alleen als er subitems zijn geselecteerd in middelste venster, overzicht geven
    if ($listview1.selecteditems.count -gt 0) {
        
        # bepalen bronmap
        $selectie = -join ($digitalebestanden, '\', $selectie)
        # aangeven of de eerste item is gevonden. alleen bij de 1e wordt extra text weergegeven
        $eerste = $true
        # inlezen en bepalen of er mappen zijn
        foreach ($item in $ListView1.selecteditems) {
            # inlezen geselecteerde items
            $inleesitem = $item[0].text
            # bepalen van map die gecontroleerd wordt
            $controleitem = -join ($selectie, '\', $inleesitem)

            
            # alleen weergeven als dit een map is
            if (( test-path -path "$controleitem" -pathtype container) -eq $true)  {
                # alleen bij de eerste item
                if ($eerste) {
                    $objtekst2.Text = $objtekst2.Text + "Submappen :" + "`r`n"
                    $eerste = $false
                    }
                $objtekst2.Text = $objtekst2.Text + " - " + "$inleesitem" + "`r`n"
                }
            }# einde 1e for each item

        # lege regel plaatsen, alleen als er mappen zijn weergegeven
        if (!($eerste)) { $objtekst2.Text = $objtekst2.Text + "`r`n" }

        # aangeven of de eerste item is gevonden. alleen bij de 1e wordt extra text weergegeven
        $eerste = $true
        # inlezen en bepalen of er bestanden zijn
        foreach ($item in $ListView1.selecteditems) {
            # inlezen geselecteerde items
            $inleesitem = $item[0].text
            # bepalen van map die gecontroleerd wordt
            $controleitem = -join ($selectie, '\', $inleesitem)

            # alleen weergeven als dit een bestand is
            if (( test-path -path "$controleitem" -pathtype container) -eq $false)  {
                # alleen bij de eerste item
                if ($eerste) {
                    $objtekst2.Text = $objtekst2.Text + "Bestanden :" + "`r`n"
                    $eerste = $false
                    }
                $objtekst2.Text = $objtekst2.Text + " - " + "$inleesitem" + "`r`n"
                }
            } # einde 2e for each item

        } # einde ($listview1.selecteditems.count -gt 0)
    } # einde ($taak -eq "kopiëren")

$uitvoeren.Form3.Controls.AddRange(@($Description2, $Description3, $StartButton, $EndButton, $Btnescape, $LogButton, $Label, $objtekst1, $objtekst2 ))

# Toevoegen gebruik van escapetoets.
# na indrukken wordt venster gesloten en hoofdvenster geopend.
$uitvoeren.Form3.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { 

    $uitvoeren.Form3.dispose() 
    $Form2.Close()
    $form.show()
    }
})
# zorgen dat speciale toetsen worden gedetecteerd.
$uitvoeren.Form3.KeyPreview = $true

$result = $uitvoeren.form3.ShowDialog()

$uitvoeren.Form3.dispose()

# Alleen als knop Sluiten is ingedrukt.
if ($result -eq [system.windows.forms.dialogResult]::OK) { 
    $Form2.Close()
    $form.show()
    }

} # einde overzichttaken

function startknopklikbaar {
# deze functie wordt alleen gebruikt bij functie Vensterkopieren
# en bepaald of knoppen en items van vensterkopieren zichtbaar moeten zijn

# start knop
if (($global:listbox.selecteditems.count -gt 0) -and ($global:geselecteerdebronmap.count -ge 2)) {
    $Btnstart.Enabled= $true
} else {
    $Btnstart.Enabled= $false
}

} # einde startknopklikbaar

function vensterkopieren { 

function toevoegen_lijst1 ($controlemap, $toevoegitem) {

  $extensienr = Bepaalicoontjenr $controlemap $toevoegitem
  [void] $listView1.Items.Add($toevoegitem, $extensienr)
}

function inlezengekozenmap ($invoer) {
# inlezen gekozen map en in variabele plaatsen

  try { 
  $ingelezen = Get-ChildItem -Path "$invoer" -Name -ErrorAction Stop | Sort-object 
  }
  catch {
  $ingelezen = ""
  # melding loggen en weergeven
  $melding = -join ("Venster Bestanden Klaarzetten is geopend : ", "`n", $_.exception.message )
  Foutenloggen $melding
  }
  return $ingelezen
}

# function vensterkopieren begint hier

# gekozen locatie in makkelijke variabele plaatsen
$keuzelocatie=$global:init["algemeen"]["locatiekeuze"]
$digitalebestanden = $global:beheer.examenmappen.digitalebestanden

# lijst met geselecteerde examenmappen
$global:geselecteerdebronmap = [System.Collections.ArrayList]@()

if (!(Netwerkmapaanwezig $global:beheer.examenmappen.digitalebestanden $true )) { return; } 
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.homemapstudenten $true )) { return; } 


# rpcnrs declareren met een function
$global:listBox = declareren_rpcnrs;
$global:listBox = lijstrpcnrsaanmaken $keuzelocatie $global:listBox 

# De hoofdmenu onzichtbaar maken
$form.Hide()

$Form2 = declareren_standaardvenster "Bestanden klaarzetten op de homemappen van de kandidaten" 920 660;

$Description2                     = New-Object system.Windows.Forms.Label
$Description2.text                = "RPC-nummers"
$Description2.AutoSize            = $false
$Description2.width               = 150
$Description2.height              = 20
$Description2.location            = New-Object System.Drawing.Point(20,15)
$Description2.Font                = 'Microsoft Sans Serif,11'
$Description2.ForeColor = [System.Drawing.Color]::Blue

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "Locatie"
$Description3.AutoSize            = $false
$Description3.width               = 400
$Description3.height              = 42
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.location            = New-Object System.Drawing.Point(550,40)
$Description3.ForeColor = [System.Drawing.Color]::Blue

$Description4                     = New-Object system.Windows.Forms.Label
$Description4.text                = "Crebonummer"
$Description4.AutoSize            = $false
$Description4.width               = 400
$Description4.height              = 42
$Description4.Font                = 'Microsoft Sans Serif,11'
$Description4.location            = New-Object System.Drawing.Point(550,80)
$Description4.ForeColor = [System.Drawing.Color]::Blue

$Description5                     = New-Object system.Windows.Forms.Label
$Description5.text                = "Kerntaak"
$Description5.AutoSize            = $false
$Description5.width               = 400
$Description5.height              = 42
$Description5.Font                = 'Microsoft Sans Serif,11'
$Description5.location            = New-Object System.Drawing.Point(550,120)
$Description5.ForeColor = [System.Drawing.Color]::Blue

$Description8                     = New-Object system.Windows.Forms.Label
$Description8.text                = "Examen"
$Description8.AutoSize            = $false
$Description8.width               = 400
$Description8.height              = 42
$Description8.Font                = 'Microsoft Sans Serif,11'
$Description8.location            = New-Object System.Drawing.Point(550,160)
$Description8.ForeColor = [System.Drawing.Color]::Blue

$Btnstart = New-object System.Windows.Forms.Button 
$Btnstart.text= "Bevestigen"
$Btnstart.location = "50,570" 
$Btnstart.size = "150,30"  
$BtnStart.BackColor = 'green'
$BtnStart.ForeColor = 'white'
$Btnstart.Add_Click({ overzichttaken "kopiëren" }) 
$Btnstart.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bevestig je keuze en ga door naar het overzicht." )
})

$Btnescape = New-object System.Windows.Forms.Button 
$Btnescape.text= "Terug"
$Btnescape.location = "250,570" 
$Btnescape.size = "150,30"  
$Btnescape.BackColor = 'red'
$Btnescape.ForeColor = 'white'
$Btnescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel
$Btnescape.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})

$BtnOpnieuw                         = New-Object system.Windows.Forms.Button
$BtnOpnieuw.width                   = 40
$BtnOpnieuw.height                  = 40
$BtnOpnieuw.location                = New-Object System.Drawing.Point(150,200)
$BtnOpnieuw.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoontje-home.png")
$BtnOpnieuw.Add_Click({ 

      # array wordt leeggemaakt
      $geselecteerdebronmap.Clear()

      # selectie krijgt waarde van gekozen crebonummer
      $selectie = $digitalebestanden

      # inlezen gekozen map en in variabele plaatsen
      $folders = inlezengekozenmap $selectie

      # toevoegen aan venster
      $listView1.items.clear()
      foreach($folder in $folders){
           toevoegen_lijst1 $selectie $folder 
      }
      $listview1.SelectedItems.Clear()
      $listView1.AutoResizeColumns(1)

      # venster met inhoud map legen
      $listView2.items.clear()
      $objtekst2.Text = $Startexamenmap
      
      # dropdownmenu's leegmaken
      $lijstkerntaken.items.clear()
      $lijstexamens.items.clear()
      # selecties leegmaken
      $lijstcrebonrs.selectedindex = -1
      $lijstkerntaken.selectedindex = -1
      $lijstexamens.selectedindex = -1
      startknopklikbaar;
})
$BtnOpnieuw.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Alle keuzes wissen en opnieuw beginnen." )
})

$BtnTerug                         = New-Object system.Windows.Forms.Button
$BtnTerug.width                   = 40
$BtnTerug.height                  = 40
$BtnTerug.location                = New-Object System.Drawing.Point(150,245)
$BtnTerug.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoon-terug.png")
$BtnTerug.Add_Click({ 
    # selectie krijgt waarde van startmap
    $selectie = $digitalebestanden
    
    switch ($geselecteerdebronmap.count) {
           "0" { # niets doen 
               }
           "1" { 
           #verwijder alles
           $geselecteerdebronmap.Clear()
           # dropdownmenu's leegmaken
           $lijstkerntaken.items.clear()
           $lijstexamens.items.clear()
           # selecties leegmaken
           $lijstcrebonrs.selectedindex = -1
           $lijstkerntaken.selectedindex = -1
           $lijstexamens.selectedindex = -1
               }
           "2" { 
           #verwijder de laatste
           $geselecteerdebronmap.RemoveAt(1)
           # dropdownmenu's leegmaken
           $lijstexamens.items.clear()
           # selecties leegmaken
           $lijstkerntaken.selectedindex = -1
           $lijstexamens.selectedindex = -1
               }
           "3" { 
           #verwijder de laatste
           $geselecteerdebronmap.RemoveAt(2)
           # selecties leegmaken
           $lijstexamens.selectedindex = -1
               }
           default { 
           $verwijdernr=$geselecteerdebronmap.Count-1
           $geselecteerdebronmap.RemoveAt($verwijdernr)
           }
    } # einde switch

    # Standaard map
    $objtekst2.Text = $Startexamenmap
    # toevoegen aan array en geselecteerde examenmap
    foreach ($item in $geselecteerdebronmap) {
            $selectie = -join ($selectie, '\', $item)
            # Toevoegen aan geselecteerde examenmap
           $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $Item)
    } # einde foreach

    # inlezen gekozen map en in variabele plaatsen
    $folders = inlezengekozenmap $selectie

    # toevoegen aan venster 
    $listView1.items.clear()
    foreach($folder in $folders){
           toevoegen_lijst1 $selectie $folder 
           
    }
    $listview1.SelectedItems.Clear()
    $listView1.AutoResizeColumns(1)

    # venster met inhoud map legen
    $listView2.items.clear()

    startknopklikbaar;
 }) # einde BtnTerug.add_click

$BtnTerug.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "De laatste keuze wissen en één stap terug." )
})

# listbox is al in het begin gedeclareerd. onderstaande waarden gelden voor deze functie.
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(130,500)
$listBox.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$listBox.SelectionMode = 'MultiExtended'
$listBox.BackColor  = 'white'
$listBox.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de RPC-nummers waar de bestanden naar toe worden overgezet." )
})


# bij aanklikken van een rpcnr, kijken of startknop weergegeven kan worden
$listbox.add_SelectedIndexChanged(
     { startknopklikbaar;
     } )

$imageList = Declareericoontjes

$global:listView1 = New-Object System.Windows.Forms.ListView
$listView1.View = 'Details'
$listView1.Height = 278
$listView1.Width = 330
$listView1.Font = New-Object System.Drawing.Font("MS Sans Serif",12)
# zorgen dat selectie zichtbaar blijft.
$listview1.HideSelection = $false

# $listView1.AutoResizeColumns(1) 
 
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 245
 
$listView1.Location = $System_Drawing_Point
$listView1.Name = "listView1"
$listView1.Sorting = 'Ascending'
#$listView1.Columns.Add('Inhoud examenmap',600)| Out-Null
$listView1.Columns.Add('Geselecteerde mappen en bestanden',600)| Out-Null
# hieronder toch niet nodig. geeft geen effect.
# $listView1.AutoResizeColumns(1)
$listView1.SmallImageList = $imageList
$listView1.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer eventueel alleen de bestanden en mappen die je wilt overzetten. 
Klik op een map om de inhoud rechts weer te geven." )
})


# bij aanklikken van een map, inhoud weergeven in venster ernaast
$listView1.add_SelectedIndexChanged(
     { 
     # selectie krijgt waarde van volledige pad naar gekozen map
     $selectie = $digitalebestanden
     foreach ($item in $geselecteerdebronmap) {
                $selectie = -join ($selectie, '\', $item)
          }
     $selectie = -join ($selectie, '\', $listView1.SelectedItems.text)

     # alleen als 1 item is geselecteerd en de item moet een map zijn.
     if (( $listView1.selecteditems.count -eq 1) -and ((test-path -path $selectie -pathtype container) -eq $true) ) {

          # tekstbox legen
          $listView2.items.clear()

          # weergeven inhoud in rechter venster, inhoud geselecteerde map
          $folders = inlezengekozenmap $selectie

          # dan netjes in rijen plaatsen.
          foreach ($item in $folders) {
            $extensienr = Bepaalicoontjenr $selectie $item
            [void] $listView2.Items.Add($item, $extensienr)
            }
          # aanpassen vesnter aan inhoud
          $listView2.AutoResizeColumns(1) 
          startknopklikbaar;

          } else {
          # venster met inhoud map legen
          $listView2.items.clear()
          }
    }
    )


# bij dubbelklikken van een map, map selecteren en openen
# hier wordt echter alleen 1 van de dropdownmenu's geactiveerd die dit gaat uitvoeren.
$listView1.add_doubleClick(
     {
     # selectie krijgt waarde van volledige pad naar gekozen map
     $selectie = $digitalebestanden
     foreach ($item in $geselecteerdebronmap) {
                $selectie = -join ($selectie, '\', $item)
                }
     $selectie = -join ($selectie, '\', $listView1.SelectedItems.text)
      
     # Er moet een item geselecteerd zijn (je kan namelijk ook dubbelklikken op een lege plek) en de item moet een map zijn.
     # if (($listView1.SelectedIndex -ge 0) -and ((test-path -path $selectie -pathtype container) -eq $true) ) {
     if ((test-path -path $selectie -pathtype container) -eq $true) {

        # als geselecteerde aanwijzen aan een van de dropdownmenu's
        switch ($geselecteerdebronmap.count) {
           "0" { $lijstcrebonrs.selectedindex = $listView1.SelectedIndices[0] }
           "1" { $lijstkerntaken.selectedindex = $listView1.SelectedIndices[0] }
           "2" { $lijstexamens.selectedindex = $listView1.SelectedIndices[0] }
           Default {
                
                # inlezen gekozen map en in variabele plaatsen
                $folders = Get-ChildItem -Path "$selectie"  -Name | Sort-object 

                if ($folders.count -gt 0) {
                    # venster met inhoud map legen
                    $listView2.items.clear()
                    # standaard tekst in venster met geselecteerd examenmap
                    $objtekst2.Text = $Startexamenmap
                    # toevoegen aan lijst geselecteerde mappen
                    $geselecteerdebronmap.Add($listView1.SelectedItems.text)
                    # leegmaken huidige venster 
                    $listView1.items.clear()
                    # toevoegen aan geselecteerde examenmap
                    foreach ($item in $geselecteerdebronmap) {
                        $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $Item)
                        }

                    # toevoegen aan venster
                    foreach($folder in $folders){ toevoegen_lijst1 $selectie $folder  }
                    # breedte listview1 aanpassen aan de inhoud
                    $listView1.AutoResizeColumns(1) 
                    startknopklikbaar;
                } # if ($folders.count -gt 0)
           } # einde default switch keuze
        }  #einde switch

      }   # einde if (test-path -path $selectie -pathtype container)
     } )

$listView2 = New-Object System.Windows.Forms.ListView
$listView2.View = 'Details'
$listView2.Height = 278
$listView2.Width = 350
$listView2.Font = New-Object System.Drawing.Font("MS Sans Serif",12)
# zorgen dat selectie zichtbaar blijft.
$listview2.MultiSelect = $false
 
$System_Drawing_Point2 = New-Object System.Drawing.Point
$System_Drawing_Point2.X = 530
$System_Drawing_Point2.Y = 245
 
$listView2.Location = $System_Drawing_Point2
$listView2.Name = "listView2"
$listView2.Sorting = 'Ascending'
$listView2.Columns.Add('Inhoud geselecteerde examenmap',600)| Out-Null
# hieronder toch niet nodig. geeft geen effect.
# $listView1.AutoResizeColumns(1)
$listView2.SmallImageList = $imageList
$listView2.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Dit is de inhoud van de geselecteerde map." )
})

$listView2.add_SelectedIndexChanged( { 
  $listview2.SelectedItems.Clear()
  } )

# standaard tekst in venster met geselecteerd examenmap
$Startexamenmap = "Examenmap"
# Tekst tussen twee geselecteerde mappen, om deze uit elkaar te houden
$Scheidingstekst = " » "

$objtekst2 = New-Object System.Windows.Forms.textbox
$objtekst2.Location = New-Object System.Drawing.Size(200,200) 
$objtekst2.Size = New-Object System.Drawing.Size(680,47)
$objtekst2.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$objtekst2.Text = $Startexamenmap
$objtekst2.ReadOnly = $true
$objtekst2.Multiline = $true
$objtekst2.BackColor  = 'white'
$objtekst2.Forecolor  = 'blue'
$objtekst2.WordWrap = $true
$objtekst2.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Hier zie je de huidige geselecteerde examenmap." )
})

$Controlemappen = New-Object System.Windows.Forms.Checkbox 
$Controlemappen.Location = New-Object System.Drawing.Point(50,530)
$Controlemappen.Size = New-Object System.Drawing.Size(700,30)
$Controlemappen.Text = "Controleer of de homemappen van de kandidaten leeg zijn voordat bestanden worden overgezet."
$Controlemappen.Font = 'Microsoft Sans Serif,11'
$Controlemappen.ForeColor = [System.Drawing.Color]::green
$Controlemappen.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Geef aan of gecontroleerd moet worden of de homemappen van de kandidaten leeg zijn." )
})

if (($global:init["algemeen"]["controlevoorklaarzetten"]) -eq "Ja") {
    $Controlemappen.checked = $true
    } else {
    $Controlemappen.checked = $false
    }

# Standaard lijst met alle locaties maken, met standaardwaarden
$lijstlocaties = declarerenlijstlocaties $keuzelocatie 330 200 40


# bij wijzigen van selectie lijstlocaties
$lijstlocaties.add_SelectedIndexChanged(
     { 
     [string]$waarde = $lijstlocaties.selecteditem
     $keuzelocatie = $waarde.Substring(0,3)

     # nieuwe rpcnrs declareren
     $global:listBox = lijstrpcnrsaanmaken $keuzelocatie $global:listBox 
     # startknop niet klikbaar maken
     startknopklikbaar;
     } ) 

$lijstcrebonrs                     = New-Object system.Windows.Forms.ComboBox
$lijstcrebonrs.text                = "Crebonummer"
$lijstcrebonrs.width               = 330
$lijstcrebonrs.autosize            = $true
$lijstcrebonrs.Location = New-Object System.Drawing.Size(200,80) 
$lijstcrebonrs.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$lijstcrebonrs.DropDownStyle="DropDownList"
$lijstcrebonrs.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de crebonummer." )
})


# crebonummers in de dropdownlijst zetten. Deze lijst verandert niet.
# en weergeven in venster listview1. Deze lijst verandert wel.
Get-ChildItem -Path $digitalebestanden -Name | Sort-object | ForEach-Object {
    [void] $lijstcrebonrs.Items.Add($_)
    toevoegen_lijst1 $digitalebestanden $_ 
    }
# breedte listview1 aanpassen aan de inhoud
$listView1.AutoResizeColumns(1) 

# uitvoeren als selectie van crebonummer wijzigt
$lijstcrebonrs.add_SelectedIndexChanged(
     { 
      
      # alleen als er geselecteerd is 
      if ($lijstcrebonrs.SelectedIndex -ge 0) {
        # venster met inhoud map legen
        $listView2.items.clear()

        <# toevoegen aan lijst geselecteerde mappen
           array wordt leeggemaakt omdat je inhoud geeft aan de 1e item aan de array "geselecteerdebronmap"
        #>
        $geselecteerdebronmap.Clear()

        # leegmaken huidige venster en dropdownmenu's
        $listView1.items.clear()
        $lijstkerntaken.items.clear()
        $lijstexamens.items.clear()
        $objtekst2.Text = $Startexamenmap

        # selectie krijgt waarde van gekozen crebonummer
        $selectie = -join ($digitalebestanden, '\', $lijstcrebonrs.SelectedItem)

        # als gekozen item een map is dan toevoegen aan lijst en weergeven in venster en in dropdownmenu
        # anders selecteren van hoofdmap en deze weergeven in venster
        if ((test-path -path $selectie -pathtype container) -eq $true) {
            # toevoegen aan lijst geselecteerde mappen 
            $geselecteerdebronmap.Add($lijstcrebonrs.SelectedItem)

            # Toevoegen aan geselecteerde examenmap
            $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $lijstcrebonrs.SelectedItem)

            # inlezen gekozen map en in variabele plaatsen
            $folders = inlezengekozenmap $selectie

            # toevoegen aan betreffende dropdown lijst en aan venster in het midden.
            foreach($folder in $folders) {
                $lijstkerntaken.Items.Add($folder)
                toevoegen_lijst1 $selectie $folder
                # $listView1.AutoResizeColumns(1)
                }
            } else {
            # selectie leeg maken
            $lijstcrebonrs.selectedindex = -1
            # hoofdmap selecteren om weer te geven
            $selectie = $digitalebestanden

            # inlezen gekozen map en in variabele plaatsen
            $folders = inlezengekozenmap $selectie

            # toevoegen aan venster in het midden.
            foreach($folder in $folders) {
                toevoegen_lijst1 $selectie $folder
                }
            }
     # breedte listview1 aanpassen aan de inhoud
     $listView1.AutoResizeColumns(1)

     startknopklikbaar;
     } else { 
       $lijstcrebonrs.selectedindex = -1
     } 
     # einde ($lijstkerntaken.SelectedIndex -ge 0) .. else
     } )


$lijstkerntaken                     = New-Object system.Windows.Forms.ComboBox
$lijstkerntaken.text                = "Kerntaak"
$lijstkerntaken.width               = 330
$lijstkerntaken.autosize            = $true
$lijstkerntaken.Location = New-Object System.Drawing.Size(200,120) 
$lijstkerntaken.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$lijstkerntaken.DropDownStyle="DropDownList"
$lijstkerntaken.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de kerntaak." )
})


# uitvoeren als selectie van kerntaken wijzigt
 $lijstkerntaken.add_SelectedIndexChanged(
    {
      

      if ($lijstkerntaken.SelectedIndex -ge 0) {

        # venster met inhoud map legen
        $listView2.items.clear()

        # leegmaken huidige venster en dropdownmenu's
        $listView1.items.clear()
        $lijstexamens.items.clear()
        $objtekst2.Text = $Startexamenmap

        <# controleren of een aantal geselecteerde mappen verwijderd moeten worden. Dit komt voor als je al mappen hebt geselecteerd en terug gaat.
           verwvanaf =1 omdat hiermee de 2e item in de array wordt aangesproken
        #>
        $verwvanaf = 1
        [int] $aantal = $geselecteerdebronmap.count -$verwvanaf
        if ($geselecteerdebronmap.count -gt $verwvanaf) { $geselecteerdebronmap.RemoveRange($verwvanaf,$aantal) }

        # selectie krijgt waarde van gekozen kerntaak
        $selectie = -join ($digitalebestanden, '\', $geselecteerdebronmap[0], '\', $lijstkerntaken.SelectedItem)

        # Toevoegen aan geselecteerde examenmap
        $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $geselecteerdebronmap[0])

        # inlezen gekozen map en in variabele plaatsen
        $folders = inlezengekozenmap $selectie

        # controle of geselecteerde een map is of een bestand
        if ((test-path -path $selectie -pathtype container) -eq $true) {
            # en dan nu pas toevoegen
            $geselecteerdebronmap.Add($lijstkerntaken.SelectedItem)
        
            # Toevoegen aan geselecteerde examenmap
            $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $lijstkerntaken.SelectedItem)

            # toevoegen aan betreffende dropdown lijst
            foreach($folder in $folders){
                $lijstexamens.Items.Add($folder)
                }
        } else {
            # selectie leeg maken
            $lijstkerntaken.selectedindex = -1

            # selectie krijgt waarde van eerder gekozen crebonr
            $selectie = -join ($digitalebestanden, '\', $geselecteerdebronmap[0])
            # opnieuw inlezen gekozen map en in variabele plaatsen
            $folders = inlezengekozenmap $selectie
        }
        
        # toevoegen aan venster in het midden. 
        foreach($folder in $folders){
           toevoegen_lijst1 $selectie $folder
           }
        # breedte listview1 aanpassen aan de inhoud
        $listView1.AutoResizeColumns(1)
        startknopklikbaar;
    } # einde ($lijstkerntaken.SelectedIndex -ge 0)
    } )

$lijstexamens                     = New-Object system.Windows.Forms.ComboBox
$lijstexamens.text                = "Examen"
$lijstexamens.width               = 330
$lijstexamens.autosize            = $true
$lijstexamens.Location = New-Object System.Drawing.Size(200,160) 
$lijstexamens.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$lijstexamens.DropDownStyle="DropDownList"
$lijstexamens.Name =  "Kies een examen"
$lijstexamens.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer het examen." )
})


# uitvoeren als selectie van examens wijzigt
$lijstexamens.add_SelectedIndexChanged(
    {
      if ($lijstexamens.SelectedIndex -ge 0) {
      
        # venster met inhoud map legen
        $listView2.items.clear()

        # leegmaken huidige vensters
        $listView1.items.clear()
        $objtekst2.Text = $Startexamenmap

        <# controleren of een aantal geselecteerde mappen verwijderd moeten worden. Dit komt voor als je al mappen hebt geselecteerd en terug gaat.
           verwvanaf = 2 omdat hiermee de 3e item in de array wordt aangesproken.
        #>
        $verwvanaf = 2
        [int] $aantal = $geselecteerdebronmap.count -$verwvanaf
        if ($geselecteerdebronmap.count -gt $verwvanaf) { $geselecteerdebronmap.RemoveRange($verwvanaf,$aantal) }

        # selectie krijgt waarde van gekozen examen
        $selectie = -join ($digitalebestanden, '\', $geselecteerdebronmap[0], '\', $geselecteerdebronmap[1], '\', $lijstexamens.SelectedItem)

        # Toevoegen aan geselecteerde examenmap
        $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $geselecteerdebronmap[0], $Scheidingstekst, $geselecteerdebronmap[1])

        # inlezen gekozen map en in variabele plaatsen
        $folders = inlezengekozenmap $selectie

        # controle of geselecteerde een map is of een bestand
        if ((test-path -path $selectie -pathtype container) -eq $true) {
            # en dan nu pas toevoegen
            $geselecteerdebronmap.Add($lijstexamens.SelectedItem)

            # Toevoegen aan geselecteerde examenmap
            $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $lijstexamens.SelectedItem)
        } else {
            # selectie leeg maken
            $lijstexamens.selectedindex = -1

            # selectie krijgt waarde van eerder gekozen kerntaak
            $selectie = -join ($digitalebestanden, '\', $geselecteerdebronmap[0], '\', $geselecteerdebronmap[1])

            # inlezen gekozen map en in variabele plaatsen
            $folders = inlezengekozenmap $selectie
        }

        # toevoegen aan venster
        foreach($folder in $folders){
           toevoegen_lijst1 $selectie $folder
           }
        # breedte listview1 aanpassen aan de inhoud
        $listView1.AutoResizeColumns(1)
        startknopklikbaar;
    }  # einde if ($lijstexamens.SelectedIndex -ge 0)
    } )

# venster met uitleg over deze taak wordt gedeclareerd. hieronder worden enkele variabelen aangepast aan deze taakvenster
declareren_uitlegvenster "Uitleg over de taak Bestanden klaarzetten." 690 340 520 570 "Hier kunt u bestanden klaarzetten op de homemappen, of RPC-nummers, van de kandidaten.
In het linkerkolom selecteert u de RPC-nummers waar de bestanden naar toe worden overgezet.
In het midden selecteert u de locatie, crebonummer, kerntaak en examen.

U kunt ook 'door de mappen bladeren' door dubbel te klikken op een map in het middelste kolom.
Met het terug-icoontje gaat u een map terug.
Met het home-icoontje gaat u terug naar het begin.

Eventueel kunt u ervoor kiezen om alleen enkele mappen over te zetten.
Hiertoe moet u in het vakje onderaan deze bestanden of mappen selecteren, 
met een enkele muisklik of met de CTRL-toets in combinatie met een muisklik.

Als u op een map klikt in het vakje onderaan wordt de inhoud hiervan rechts weergegeven.

Om naar het overzicht te gaan waar u het overzetten kan starten moet u op Bevestigen klikken.
"

startknopklikbaar;

$Form2.controls.AddRange(@($listBox, $lijstlocaties, $lijstcrebonrs, $lijstkerntaken, $lijstexamens, $objtekst2, $BtnOpnieuw, $BtnTerug, 
$listview1, $listView2, $Btnstart, $Btnescape, $Controlemappen, $Description2,
$Description3, $Description4, $Description5, $Description8, $Global:vraagtekenicoon ))

Form2afsluitenbijescape;

# venster tonen
$null = $form2.ShowDialog()
    
$form2.close();
# $form2.hide();

# De hoofdmenu zichtbaar maken
$form.show()

} # einde vensterkopieren

function vensterbackup {
# taak backuppen begint hier

# variabelen
$keuzelocatie=$global:init["algemeen"]["locatiekeuze"]

# controleren of de mappen beschikbaar zijn
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.homemapstudenten $true )) { return; } 

if (!(Netwerkmapaanwezig $global:beheer.examenmappen.backupmap $true )) { return; } 

# De hoofdmenu onzichtbaar maken
$form.Hide()

# rpcnrs declareren met een function
$global:listBox = declareren_rpcnrs;
$global:listBox = lijstrpcnrsaanmaken $keuzelocatie $global:listBox 

# venster declareren
$Form2 = declareren_standaardvenster "Back-up maken van de homemappen" 480 680

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "Locatie"
$Description3.AutoSize            = $false
$Description3.width               = 800
$Description3.height              = 40
$Description3.location            = New-Object System.Drawing.Point(290,15)
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.ForeColor = [System.Drawing.Color]::Blue

$Description4                     = New-Object system.Windows.Forms.Label
$Description4.text                = "RPC-nummers"
$Description4.AutoSize            = $false
$Description4.width               = 800
$Description4.height              = 40
$Description4.location            = New-Object System.Drawing.Point(290,55)
$Description4.Font                = 'Microsoft Sans Serif,11'
$Description4.ForeColor = [System.Drawing.Color]::Blue

$Btnstart = New-object System.Windows.Forms.Button 
$Btnstart.text= "Bevestigen"
$Btnstart.location = "50,570" 
$Btnstart.size = "150,30"
$BtnStart.BackColor = 'green'
$BtnStart.ForeColor = 'white'
$Btnstart.Add_Click({ overzichttaken "backup" }) 
$Btnstart.Enabled= $false
$Btnstart.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bevestig je keuze en ga door naar het overzicht." )
})

$Btnescape = New-object System.Windows.Forms.Button 
$Btnescape.text= "Terug"
$Btnescape.location = "250,570" 
$Btnescape.size = "150,30"  
$Btnescape.BackColor = 'red'
$Btnescape.ForeColor = 'white'
$Btnescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel
$Btnescape.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})

# listbox is in het begin al gedeclareerd.
$listBox.Location = New-Object System.Drawing.Point(10,55)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$listBox.SelectionMode = 'MultiExtended'
$listBox.Height = 475
$listBox.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de RPC-nummers waarvan de back-up wordt gemaakt." )
})

# start knop zichtbaar of niet
$listbox.add_SelectedIndexChanged(
     { 
     if ($listbox.selecteditems.count -gt 0) {
        $Btnstart.Enabled= $true
        } else {
        $Btnstart.Enabled= $false
        }
    } )

$wissennabackup = New-Object System.Windows.Forms.Checkbox 
$wissennabackup.Location = New-Object System.Drawing.Point(50,530)
$wissennabackup.Size = New-Object System.Drawing.Size(300,30)
$wissennabackup.Text = "Bestanden na de back-up ook wissen"
$wissennabackup.Font = 'Microsoft Sans Serif,11'
$wissennabackup.ForeColor = [System.Drawing.Color]::green
$wissennabackup.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Geef aan of de homemappen van de kandidaten na de bac-kup worden gewist." )
})

if (($global:init["algemeen"]["wissennabackup"]) -eq "Ja") {
    $wissennabackup.checked = $true
    } else {
    $wissennabackup.checked = $false
    }

# Standaard lijst met alle locaties maken, met standaardwaarden
$lijstlocaties = declarerenlijstlocaties $keuzelocatie 260 10 15


# bij wijzigen van selectie lijstlocaties
$lijstlocaties.add_SelectedIndexChanged(
     { 
     [string]$waarde = $lijstlocaties.selecteditem
     $keuzelocatie = $waarde.Substring(0,3)

     # nieuwe rpcnrs declareren
     $global:listBox = lijstrpcnrsaanmaken $keuzelocatie $global:listBox 
     # startknop niet klikbaar maken
     $Btnstart.Enabled= $false
     } ) 

# venster met uitleg over deze taak wordt gedeclareerd. hieronder worden enkele variabelen aangepast aan deze taakvenster
declareren_uitlegvenster "Uitleg over de taak Back-up maken." 730 230 420 580 "Hier kunt u een back-up maken van de geselecteerde homemappen van de kandidaten.
Bovenaan kunt u de locatie en daaronder de homemappen, of RPC-nummers, selecteren.

Standaard wordt na de back-up de homemappen van de kandidaten gewist.
Als u dit niet wilt moet u het vinkje bij de bijbehorende tekst onderaan weghalen.

Om naar het overzicht te gaan waar u de back-up kan starten moet u op Bevestigen klikken.
"

$form2.controls.AddRange(@($lijstlocaties, $listBox, $Btnstart, $Btnescape, $wissennabackup, $Description3, $Description4, $Global:vraagtekenicoon ))

Form2afsluitenbijescape;

$null = $form2.ShowDialog()
    
$form2.close();

# De hoofdmvenster zichtbaar maken
$form.show()
} # einde vensterbackup

function vensterwissen {
# taak wissen begint hier

# variabelen
$keuzelocatie=$global:init["algemeen"]["locatiekeuze"]


# controleren of de mappen beschikbaar zijn
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.homemapstudenten $true )) { return; } 

# De hoofdmenu onzichtbaar maken
$form.Hide()

# rpcnrs declareren met een function
$global:listBox = declareren_rpcnrs;
$global:listBox = lijstrpcnrsaanmaken $keuzelocatie $global:listBox 

# venster declareren
$Form2 = declareren_standaardvenster "Wissen van homemappen van de kandidaten" 480 640;

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "Locatie"
$Description3.AutoSize            = $false
$Description3.width               = 800
$Description3.height              = 40
$Description3.location            = New-Object System.Drawing.Point(290,15)
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.ForeColor = [System.Drawing.Color]::Blue

$Description4                     = New-Object system.Windows.Forms.Label
$Description4.text                = "RPC-nummers"
$Description4.AutoSize            = $false
$Description4.width               = 800
$Description4.height              = 40
$Description4.location            = New-Object System.Drawing.Point(290,55)
$Description4.Font                = 'Microsoft Sans Serif,11'
$Description4.ForeColor = [System.Drawing.Color]::Blue


$Btnstart = New-object System.Windows.Forms.Button 
$Btnstart.text= "Bevestigen"
$Btnstart.location = "50,540" 
$Btnstart.size = "150,30"  
$BtnStart.BackColor = 'green'
$BtnStart.ForeColor = 'white'
$Btnstart.Add_Click({ overzichttaken "wissen" }) 
$Btnstart.Enabled= $false
$Btnstart.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bevestig je keuze en ga door naar het overzicht." )
})

$Btnescape = New-object System.Windows.Forms.Button 
$Btnescape.text= "Terug"
$Btnescape.location = "250,540" 
$Btnescape.size = "150,30"  
$Btnescape.BackColor = 'red'
$Btnescape.ForeColor = 'white'
$Btnescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel
$form2.cancelbutton = $Btnescape
$Btnescape.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})


# listbox is in het begin al gedeclareerd.
$listBox.Location = New-Object System.Drawing.Point(10,55)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$listBox.SelectionMode = 'MultiExtended'
$listBox.Height = 475
$listBox.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de RPC-nummers waarvan de back-up wordt gemaakt." )
})

# start knop zichtbaar of niet
$listbox.add_SelectedIndexChanged(
     { 
     if ($listbox.selecteditems.count -gt 0) {
        $Btnstart.Enabled= $true
        } else {
        $Btnstart.Enabled= $false
        }
    } )

# Standaard lijst met alle locaties maken, met standaardwaarden
$lijstlocaties = declarerenlijstlocaties $keuzelocatie 260 10 15

# bij wijzigen van selectie lijstlocaties
$lijstlocaties.add_SelectedIndexChanged(
     { 
     [string]$waarde = $lijstlocaties.selecteditem
     $keuzelocatie = $waarde.Substring(0,3)

     # nieuwe rpcnrs declareren
     $global:listBox = lijstrpcnrsaanmaken $keuzelocatie $global:listBox 
     # startknop niet klikbaar maken
     $Btnstart.Enabled= $false
     } ) 

# venster met uitleg over deze taak wordt gedeclareerd. hieronder worden enkele variabelen aangepast aan deze taakvenster
declareren_uitlegvenster "Uitleg over de taak Wissen." 690 180 430 550 "Hier kunt u bestanden van de geselecteerde RPC-nummers verwijderen.
Bovenaan kunt u de locatie en daaronder de RPC-nummers die verwijderd moet worden, selecteren.

Om naar het overzicht te gaan waar u het verwijderen kan starten moet u op Bevestigen klikken.
"

$form2.controls.AddRange(@($lijstlocaties, $listBox, $Btnstart, $Btnescape, $Description3, $Description4, $Global:vraagtekenicoon))

Form2afsluitenbijescape;

$null = $form2.ShowDialog()
    
$form2.close();

# De hoofdvenster zichtbaar maken
$form.show()
} # einde vensterwissen

function vensterverplaatsen {
# taak verplaatsen van bestanden begint hier

# variabelen
$keuzelocatie=$global:init["algemeen"]["locatiekeuze"]
$keuzedoelmaplegen=$global:init["algemeen"]["maplegenvoorverplaatsen"]


# controleren of de mappen beschikbaar zijn
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.homemapstudenten $true )) { return; } 

# De hoofdmenu onzichtbaar maken
$form.Hide()

# venster declareren
$Form2 = declareren_standaardvenster "Verplaatsen of kopiëren van bestanden" 550 380;

$Description2                     = New-Object system.Windows.Forms.Label
$Description2.text                = "Bron"
$Description2.AutoSize            = $false
$Description2.width               = 150
$Description2.height              = 50
$Description2.location            = New-Object System.Drawing.Point(300,120)
$Description2.Font                = 'Microsoft Sans Serif,11'
$Description2.ForeColor = [System.Drawing.Color]::Blue

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "Doel"
$Description3.AutoSize            = $false
$Description3.width               = 150
$Description3.height              = 50
$Description3.location            = New-Object System.Drawing.Point(300,170)
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.ForeColor = [System.Drawing.Color]::Blue

$Description5                     = New-Object system.Windows.Forms.Label
$Description5.text                = "Locatie"
$Description5.AutoSize            = $false
$Description5.width               = 300
$Description5.height              = 42
$Description5.Font                = 'Microsoft Sans Serif,11'
$Description5.location            = New-Object System.Drawing.Point(300,20)
$Description5.ForeColor = [System.Drawing.Color]::Blue

$Description6                     = New-Object system.Windows.Forms.Label
$Description6.text                = "Taak"
$Description6.AutoSize            = $false
$Description6.width               = 300
$Description6.height              = 42
$Description6.Font                = 'Microsoft Sans Serif,11'
$Description6.location            = New-Object System.Drawing.Point(300,70)
$Description6.ForeColor = [System.Drawing.Color]::Blue

$Btnstart = New-object System.Windows.Forms.Button 
$Btnstart.text= "Bevestigen"
$Btnstart.location = "50,270" 
$Btnstart.size = "150,30"  
$BtnStart.BackColor = 'green'
$BtnStart.ForeColor = 'white'
$Btnstart.Add_Click({ overzichttaken "verplaatsen" }) 
$Btnstart.Enabled= $false
$Btnstart.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bevestig je keuze en ga door naar het overzicht." )
})

$Btnescape = New-object System.Windows.Forms.Button 
$Btnescape.text= "Terug"
$Btnescape.location = "250,270" 
$Btnescape.size = "150,30"  
$Btnescape.BackColor = 'red'
$Btnescape.ForeColor = 'white'
$Btnescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel
$Btnescape.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})
$form2.cancelbutton = $Btnescape

# de selectie vd de bron
$bronselectie                     = New-Object system.Windows.Forms.ComboBox
$bronselectie.width               = 260
$bronselectie.autosize            = $true
$bronselectie.DropDownStyle       = "DropDownList"
$bronselectie.Font                = 'Microsoft Sans Serif,12'
$bronselectie.location = "20,120" 
$bronselectie.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer het rpc-nummer met de bronbestanden." )
})
$bronselectie = lijstrpcnrsaanmaken $keuzelocatie $bronselectie

#  de selectie vd de doel rpc-nummer 
$doelselectie                     = New-Object system.Windows.Forms.ComboBox
$doelselectie.width               = 260
$doelselectie.autosize            = $true
$doelselectie.DropDownStyle       = "DropDownList"
$doelselectie.Font                = 'Microsoft Sans Serif,12'
$doelselectie.location = "20,170" 
$doelselectie.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer het rpc-nummer waar de bestanden naar toe moeten." )
})
$doelselectie = lijstrpcnrsaanmaken $keuzelocatie $doelselectie

#  de keuze vd taak
$keuzeoptie1                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie1.width               = 260
$keuzeoptie1.autosize            = $true
$keuzeoptie1.DropDownStyle       = "DropDownList"
$keuzeoptie1.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie1.location = "20,70" 
[void] $keuzeoptie1.Items.Add("Verplaatsen")
[void] $keuzeoptie1.Items.Add("Kopiëren")
$keuzeoptie1.Selectedindex = 0
$keuzeoptie1.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de uit te voeren taak." )
})


$doelselectie.add_SelectedIndexChanged(
     { 
    # kijken of dezelfde selectie is gemaakt en zo ja, opheffen
    if ($doelselectie.selecteditem -eq $bronselectie.selecteditem) { 
        $doelselectie.selecteditem = $null
        }
    # kijken of startknop zichtbaar mag zijn
    if (($doelselectie.selecteditem -ne $null) -and ($bronselectie.selecteditem -ne $null)) {
        $Btnstart.Enabled= $true
        } else {
        $Btnstart.Enabled= $false
        }  
    } )

$bronselectie.add_SelectedIndexChanged(
     { 
    # kijken of dezelfde selectie is gemaakt en zo ja, opheffen
    if ($doelselectie.selecteditem -eq $bronselectie.selecteditem) { 
        $bronselectie.selecteditem = $null
        }
    # kijken of startknop zichtbaar mag zijn
     if (($doelselectie.selecteditem -ne $null) -and ($bronselectie.selecteditem -ne $null)) {
        $Btnstart.Enabled= $true
        } else {
        $Btnstart.Enabled= $false
        }  
    } )

# Standaard lijst met alle locaties maken, met standaardwaarden
$lijstlocaties = declarerenlijstlocaties $keuzelocatie 260 20 20


# bij wijzigen van selectie lijstlocaties
$lijstlocaties.add_SelectedIndexChanged(
     { 
     [string]$waarde = $lijstlocaties.selecteditem
     $keuzelocatie = $waarde.Substring(0,3)

     # nieuwe rpcnrs declareren
     $bronselectie = lijstrpcnrsaanmaken $keuzelocatie $bronselectie
     $doelselectie = lijstrpcnrsaanmaken $keuzelocatie $doelselectie

     # selectie van bron en doel ophefen
     $doelselectie.selecteditem = $null
     $bronselectie.selecteditem = $null

     # startknop niet klikbaar maken
     $Btnstart.Enabled= $false
     } ) 

$doelmaplegen = New-Object System.Windows.Forms.Checkbox 
$doelmaplegen.Location = New-Object System.Drawing.Point(20,220)
$doelmaplegen.Size = New-Object System.Drawing.Size(500,30)
$doelmaplegen.Text = "Doelmap wissen alvorens het verplaatsen of kopiëren."
$doelmaplegen.Font = 'Microsoft Sans Serif,12'
$doelmaplegen.ForeColor = [System.Drawing.Color]::Green
if ($keuzedoelmaplegen -eq "Ja") {
    $doelmaplegen.Checked = $true
    } else {
    $doelmaplegen.Checked = $false
    }
$doelmaplegen.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Geef aan of de doelmap geleegd moet worden voor het uitvoeren van de taak." )
})

# venster met uitleg over deze taak wordt gedeclareerd. hieronder worden enkele variabelen aangepast aan deze taakvenster
declareren_uitlegvenster "Uitleg over de taak Verplaatsen of kopiëren." 680 280 460 290 "Hier kunt u bestanden van een RPC-nummer verplaatsten of kopiëren naar een andere RPC-nummer.

Geef aan in welke RPC-nummer de bestanden staan, dit is de bron. 
Dan naar welke RPC-nummer de bestanden worden gekopiëerd, dat is het doel. 
Eventueel kan de locatie of taak nog worden aangepast.

Als Doelmap legen voor het verplaatsen of kopiëren is aangevinkt zal de inhoud van de doelmap 
gewist worden voor het verplaatsen of kopiëren.

Om naar het overzicht te gaan waar u het verplaatsen kan starten moet u op Bevestigen klikken."

$form2.controls.AddRange(@($lijstlocaties, $keuzeoptie1, $bronselectie, $doelselectie, $doelmaplegen, $Btnstart, $Btnescape, $Description2, 
$Description3, $Description4, $Description5, $Description6, $Global:vraagtekenicoon ))

Form2afsluitenbijescape;

$null = $form2.ShowDialog()
    
$form2.close();

# De hoofdmvenster zichtbaar maken
$form.show()
} # einde vensterverplaatsen


function vensterlogbestand {

function bepaaldatumuitlognaam ($invoer)
{
<# Uit de invoer wordt de datum bepaald. De format hiervoor staat in functie "bepaallognaam".
  Een wijziging in deze functie moet hier ook worden toegepast.
#>
$jaar = $invoer.substring(4, 4)
$maand = $invoer.substring(9, 2)
$dag = $invoer.substring(12, 2)
$datumintekst = -join ($dag,'-',$maand,'-',$jaar)
return $datumintekst
} # einde bepaaldatumuitlognaam

function Inlezenlogs { 

# De logbestanden worden ingelezen en met een try-catch methode de fouten opgevangen
try {
$uitvoer = Get-ChildItem -Path $weertegevenlogs -Name -ErrorAction Stop
return $uitvoer
}
catch {
      $melding = -join ("Taak logbestanden bekijken is gestart : ", "`n", $_.exception.message )
      Foutenloggen $melding
}
} # einde function Inlezenlogs

# function vensterlogbestand bekijken begint hier

# Deze variabele bepaalt de "weer te geven logbestanden". Dit is is een function geplaatst om op 1 plek veranderingen door te voeren.
$weertegevenlogs = bepaaleigenlogbestanden

# controle of er logbestanden zijn
if (!(Test-Path -Path $weertegevenlogs)) {
    $null = venstermetvraag -titel "Geen logbestanden" -vraag "`r`nEr zijn geen logbestanden om te tonen."
    return;
    }

# De hoofdmenu onzichtbaar maken
$form.Hide()

# venster declareren
$Form2 = declareren_standaardvenster "Logbestanden bekijken" 1060 670

$Description2                     = New-Object system.Windows.Forms.Label
$Description2.text                = "Filter op jaar"
$Description2.AutoSize            = $false
$Description2.width               = 100
$Description2.height              = 42
$Description2.location            = New-Object System.Drawing.Point(20,18)
$Description2.Font                = 'Microsoft Sans Serif,11'
$Description2.ForeColor = [System.Drawing.Color]::Blue

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "en op maand"
$Description3.AutoSize            = $false
$Description3.width               = 100
$Description3.height              = 42
$Description3.location            = New-Object System.Drawing.Point(210,18)
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.ForeColor = [System.Drawing.Color]::Blue

$Description4                     = New-Object system.Windows.Forms.Label
$Description4.text                = "Logbestanden"
$Description4.AutoSize            = $false
$Description4.width               = 120
$Description4.height              = 20
$Description4.location            = New-Object System.Drawing.Point(20,65)
$Description4.Font                = 'Microsoft Sans Serif,11'
$Description4.ForeColor = [System.Drawing.Color]::Blue

$Description5                     = New-Object system.Windows.Forms.Label
$Description5.text                = "Inhoud van geselecteerde logbestand"
$Description5.AutoSize            = $false
$Description5.width               = 280
$Description5.height              = 20
$Description5.location            = New-Object System.Drawing.Point(160,65)
$Description5.Font                = 'Microsoft Sans Serif,11'
$Description5.ForeColor = [System.Drawing.Color]::Blue

$filterjaar                     = New-Object system.Windows.Forms.ComboBox
$filterjaar.width               = 70
$filterjaar.autosize            = $true
$filterjaar.DropDownStyle       = "DropDownList"
$filterjaar.Font                = 'Microsoft Sans Serif,12'
$filterjaar.location = "120,15" 
# vullen met jaren
for ($i=2022; $i -le 2040; $i++) {
    [void] $filterjaar.Items.Add($i)
}
$filterjaar.add_SelectedIndexChanged({
    $filtermaand.Enabled = $true
    $filtermaand.selecteditem = $null
    $listbox2.Items.clear()

    # Inlezen logs en in array plaatsen
    Inlezenlogs | Sort-Object -Descending | ForEach-Object {
        $datumlog = bepaaldatumuitlognaam "$_"

        if ($datumlog.Contains($filterjaar.selecteditem)) { 
            [void] $listbox2.Items.Add($datumlog)
        }
    }
    if ($listbox2.Items.count -gt 0) { $listBox2.SelectedIndex = 0 }
        else { $objtekst1.Text = "" }
})
$filterjaar.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Toon alleen de logbestanden van de geselecteerde jaar." )
})

$filtermaand                     = New-Object system.Windows.Forms.ComboBox
$filtermaand.width               = 50
$filtermaand.autosize            = $true
$filtermaand.DropDownStyle       = "DropDownList"
$filtermaand.Font                = 'Microsoft Sans Serif,12'
$filtermaand.location = "320,15" 
$filtermaand.Enabled = $false
# vullen met maanden
for ($i=01; $i -le 12; $i++) {
    if ($i -lt 10) {
        [void] $filtermaand.Items.Add("0$i")
        } else {
        [void] $filtermaand.Items.Add("$i")
        }
}
$filtermaand.add_SelectedIndexChanged({
    $listbox2.Items.clear()

    Inlezenlogs | Sort-Object -Descending | ForEach-Object {
        $datumlog = bepaaldatumuitlognaam "$_"
        $filter = -join ($filtermaand.selecteditem,"-",$filterjaar.selecteditem)
        if ($datumlog.Contains($filter)) { 
            [void] $listbox2.Items.Add($datumlog)
        }
    }
    if ($listbox2.Items.count -gt 0) { $listBox2.SelectedIndex = 0 }
        else { $objtekst1.Text = "" }
})
$filtermaand.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Toon alleen de logbestanden van de geselecteerde maand." )
})

$filterwissen = New-object System.Windows.Forms.Button 
$filterwissen.text= "Filters wissen"
$filterwissen.location = "400,15" 
$filterwissen.size = "150,30"  
$filterwissen.BackColor = 'blue'
$filterwissen.ForeColor = 'white'
$filterwissen.add_click({
    
    $filtermaand.selecteditem = $null
    $filterjaar.selecteditem = $null
    $filtermaand.Enabled = $false

    $listbox2.Items.clear()
    Inlezenlogs | Sort-Object -Descending | ForEach-Object {
    $datumlog = bepaaldatumuitlognaam "$_"
    [void] $listbox2.Items.Add($datumlog)        
    }
    # eerste logbestand is geselecteerd.
    $listBox2.SelectedIndex = 0
})
$filterwissen.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Wis alle filters en laat alle logbestanden weer zien." )
})

$Buttenok = New-object System.Windows.Forms.Button 
$Buttenok.text= "Terug"
$Buttenok.location = "50,590" 
$Buttenok.size = "150,30"  
$Buttenok.BackColor = 'red'
$Buttenok.ForeColor = 'white'
$Buttenok.DialogResult = [System.Windows.Forms.DialogResult]::ok
$Buttenok.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})

$listBox2 = New-Object System.Windows.Forms.Listbox
$listBox2.Location = New-Object System.Drawing.Point(10,90)
$listBox2.Size = New-Object System.Drawing.Size(120,20)
$listBox2.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$listBox2.Height = 495
$listBox2.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de datum van het logbestand die u wilt zien." )
})

# bij aanklikken van een datum, inhoud weergeven in venster ernaast
$listbox2.add_SelectedIndexChanged( {
    $objtekst1.Text = ""
    # datum omzetten naar bestandsnaam
    $gekozendatum = bepaallognaam $ListBox2.SelectedItem
    # inhoud bestand inlezen
    $volledigetext = Get-Content -Path "$gekozendatum"
    # dan netjes in rijen plaatsen.
    foreach ($item in $volledigetext) {
               $objtekst1.Text = $objtekst1.Text + "$item" + "`r`n"
               }
    } )

# inhoud van logmap weergeven met listbox2
Inlezenlogs | Sort-Object -Descending | ForEach-Object {
    $datumlog = bepaaldatumuitlognaam "$_"
    [void] $listbox2.Items.Add($datumlog)
    }

$objtekst1 = New-Object System.Windows.Forms.textbox
$objtekst1.Location = New-Object System.Drawing.Size(150,90) 
$objtekst1.Size = New-Object System.Drawing.Size(880,484)
$objtekst1.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$objtekst1.Text = ""
$objtekst1.ReadOnly = $true
$objtekst1.Multiline = $true
$objtekst1.ScrollBars = "Both"
$objtekst1.BackColor  = 'white'
$objtekst1.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Inhoud van het geselecteerde logbestand." )
})

# venster met uitleg over deze taak wordt gedeclareerd. hieronder worden enkele variabelen aangepast aan deze taakvenster
declareren_uitlegvenster "Uitleg over het venster Logbestanden bekijken." 680 180 500 590 "Als u links de datum van het logbestand dat je wilt bekijken selecteert, dan wordt rechts 
de inhoud weergegeven.

U kunt om sneller een logbestand te vinden, bovenaan filteren op jaar en maand.
Door om de knop Filters wissen te klikken ziet u weer alle logbestanden." 

$Form2.Controls.AddRange(@($Description2, $Description3, $Description4, $Description5, $filterjaar, $filtermaand, $filterwissen, $listbox2, $objtekst1, $Buttenok, $Global:vraagtekenicoon ))

# eerste logbestand is geselecteerd.
$listBox2.SelectedIndex = 0

Form2afsluitenbijescape;

$null = $form2.ShowDialog()

# De hoofdmenu zichtbaar maken
$form.show()
} # einde vensterlogbestand


function updateuitvoeren {

    Function vergelijk_versies ($huidigeversie, $updateversie) {
    # Haal eventuele +rc.x of -rc.x suffixen eruit en bewaar ze
    $huidigeMain = $huidigeversie -replace '\.rc\.\d+$', ''
    $updateMain = $updateversie -replace '\.rc\.\d+$', ''

    # Haal rc-nummer indien aanwezig
    $huidigeRc = $null
    if ($huidigeversie -match '\.rc\.(\d+)$') { $huidigeRc = [int]$Matches[1] }
    $updateRc = $null
    if ($updateversie -match '\.rc\.(\d+)$') { $updateRc = [int]$Matches[1] }

    # Split de hoofdversies in onderdelen
    $huidigeDeel = $huidigeMain -split '\.'
    $updateDeel = $updateMain -split '\.'

    for ($i = 0; $i -lt [Math]::Max($huidigeDeel.Length, $updateDeel.Length); $i++) {
        $huidigeNummer = if ($i -lt $huidigeDeel.Length) { [int]$huidigeDeel[$i] } else { 0 }
        $updateNummer = if ($i -lt $updateDeel.Length) { [int]$updateDeel[$i] } else { 0 }

        if ($huidigeNummer -lt $updateNummer) {
            return -1  # Huidige versie is lager
        } elseif ($huidigeNummer -gt $updateNummer) {
            return 1   # Huidige versie is hoger
        }
    }

    # Vergelijk RC's indien hoofdversies gelijk zijn
    if ($null -ne $huidigeRc -and $null -ne $updateRc) {
        if ($huidigeRc -lt $updateRc) { return -1 }
        elseif ($huidigeRc -gt $updateRc) { return 1 }
        else { return 0 }
    } elseif ($null -ne $huidigeRc -and $null -eq $updateRc) {
        # Release Candidate is altijd lager dan een definitieve release
        return -1
    } elseif ($null -eq $huidigeRc -and $null -ne $updateRc) {
        return 1
    }

    return 0  # Versies zijn gelijk
} # einde vergelijk_versies

# de function updateuitvoeren begint hier ****************************************

# alleen starten als programma.mode niet de status alpha of beta heeft.
if ("alpha","beta" -contains($global:programma.mode)) {
    return
}

# bepalen programma naam tbv controle hieronder en doorzoeken website naar laatste versie
$programmanaam = $global:programma.naam

# Eerste regel van elke foutmelding
$foutmeldingbegin = "Uitvoeren van een update :
"
# huidige versie van het programma
$huidigeversie = $global:programma.versie

# Standaard waarde. Als er geen update is gevonden, dan blijft deze waarde 0.0.0. 
# Let op, dit betekent dat er een probleem is met de update.
$updateto = "0.0.0" 

# zip-bestand met update op lokale pc
$zip_download = -join ("$startmap","\","updatebestand.zip")

if (Test-Path -path $zip_download -pathtype leaf) {
        write-host "Er is net een update uitgevoerd. Het gedownloade bestand wordt verwijderd."
        Remove-Item $zip_download -Recurse -Force
        return
    }

# controleren of het script al is opgestart. Als dit zo is kan het mis gaan bij het updaten.
try {
    $gevonden = Get-CimInstance Win32_Process -Filter "Name='powershell.exe' AND CommandLine like '%$programmanaam%'" -ErrorAction Stop
    }
catch {
        # melding loggen
        $melding = -join ($foutmeldingbegin, $_.exception.message )
        Foutenloggen $melding
        $melding2 = -join ($foutmeldingbegin, "Fout tijdens de controle of het script al is gestart. Zie logbestand voor details." )
        Write-Host $melding2 -ForegroundColor Yellow
        return
}
 
 if ($gevonden) { 
     # er moet altijd een proces gevonden worden omdat dit script in ieder geval draait. Probleem is er als er meerdere processen draaien.
    if ($gevonden.processid.count -gt 1) { 
        $melding = -join ($foutmeldingbegin, "Het updateproces wordt niet uitgevoerd omdat een andere proces van het script al is opgestart." )
        Write-Host $melding -ForegroundColor Yellow
        # Melding ook loggen
        Foutenloggen -meldtekst $melding -type "OPGELET:"
        Start-Sleep -Seconds 3
        return
    } 
 }

# Info geven
write-host "Controleren op een update."

# dit is de url van de github repository. Hier wordt bepaald of de release of prerelease wordt gedownload.
$url = $global:programma.github 
if ($global:programma.mode -eq "release") {
        $url = -join ($url,"release")
        } else {
        $url = -join ($url,"prerelease")
        }

# inhoud van de map in github ophalen
try {
        $response = Invoke-RestMethod -Uri $url -Headers @{ "User-Agent" = "PowerShell" }
    } 
catch {
        $melding = -join ($foutmeldingbegin, "Het is niet gelukt om verbinding te maken met de website!
Neem contact op met de eigenaar van $url"  ) 
        # melding loggen
        Foutenloggen $melding
        write-host $melding -f Red
        Start-Sleep -Seconds 8
        return
    }
       
# Loop door de items in de response en controleer of er een nieuw update is
foreach ($item in $response) {
        # alleen bestanden controleren
        if ($item.type -eq "file") {
            # bestnaam is de naam van het te controleren bestand
            $bestnaam = $($item.name) 

            # controleer of het bestand de juiste format heeft.
            if ($bestnaam -match "$programmanaam") {
                # conroleer of het bestand de juiste versieformat heeft.
                if (( $bestnaam -match "_\d+\.\d+\.\d+\.zip$") -or ( $bestnaam -match "_\d+\.\d+\.\d+\.rc\.\d+\.zip$") ){

                    # bepaal laatste versie van het script
                    # versiemetzip is de versie met .zip in de naam.
                    $versiemetzip = $bestnaam.split('_')[1]
                    # de positie van de laatste punt in de versie bepalen
                    $positiepunt = $versiemetzip.LastIndexOf(".")
                    # updateto is alleen de versie zonder .zip. Dit is de versie die we willen vergelijken met de huidige versie.
                    $updateto = $versiemetzip.Substring(0, $positiepunt) 
                    # tedownloadebestand is de naam van het bestand dat gedownload moet worden
                    $tedownloadenbestand = $item.download_url

                    break # we zoeken maar 1 bestand, dus na de eerste gevonden versie stoppen met zoeken
                } # einde controle bestnaam met versieformat
            } # einde controle bestnaam met programmanaam                     
        } # einde if item.type = file
    } # einde foreach loop

# Als er geen update is gevonden, melding geven en loggen
if ($updateto -eq "0.0.0") {
        Write-Host "" -f Red
        $melding = -join ($foutmeldingbegin, "Er is geen update gevonden in de GitHub repository: $url
Neem contact op met de eigenaar van $url"  ) 
        # melding loggen
        Foutenloggen $melding
        write-host $melding -f Red
        Start-Sleep -Seconds 8
    } 

# vergelijken van de huidige versie met de update versie
$resultaat = vergelijk_versies $huidigeversie $updateto
if ($resultaat -eq 0) {
        # Huidige versie is gelijk aan de update versie
        write-host "Huidige versie is gelijk aan de update versie."
        return
    } elseif ($resultaat -gt 0) {
        # Huidige versie is hoger dan de update versie
        write-host "Huidige versie is hoger dan de update versie." 
        return
     }


# Hier aangekomen dan is er een update beschikbaar.
write-host "Programma wordt geupdatet naar versie $updateto "

# downloaden zip_download van Github en foutmeldingen opvangen
$error.clear()
try {
    Invoke-WebRequest -Uri $tedownloadenbestand -OutFile $zip_download -ErrorAction Stop
    }

catch {
    # foutmelding weergeven, loggen en stoppen
    $melding = -join ($foutmeldingbegin, "Het updaten is niet gelukt omdat het updatebestand niet is gevonden op de website. 
Neem contact op met de eigenaar van $url" )
    Foutenloggen $melding
    write-host $melding -f Red
    Start-Sleep -Seconds 8
    return
    }

# backup maken van hele map voor het geval het mis gaat bij het uitpakken
# eerst de backupbestand een naam geven
$backupzip = "$startmap\backup.zip"
# als deze al bestaat, verwijderen...
if (test-path -path "$backupzip" -PathType Leaf) { Remove-Item "$backupzip" }
# Dan backup maken....
try {
    Write-Host "Er wordt een veiligheidsbackup gemaakt."
    Compress-Archive -Path "$startmap\*" -DestinationPath $backupzip -ErrorAction Stop
}
catch {
    # foutmelding weergeven en stoppen
    $melding = -join ($foutmeldingbegin, "Het updaten is niet gelukt omdat er geen veiligheidsback-up gemaakt kon worden. 
Neem contact op met de eigenaar van $url")
    Foutenloggen $melding
    write-host $melding -f Red
    Start-Sleep -Seconds 8
    return
    }

# uitpakken en installeren van programma
$error.clear()
try {
    Expand-Archive -Path "$zip_download" -DestinationPath "$startmap" -Force -ErrorAction Stop
}
catch {
    # foutmelding weergeven en stoppen
    
    $melding = -join ($foutmeldingbegin, "Het updaten is niet gelukt omdat er iets fout ging bij het uitpakken van de nieuwe bestanden. 
Neem contact op met de eigenaar van $url" )
    Foutenloggen $melding
    write-host $melding -f Red
    Start-Sleep -Seconds 8
    
    # Terugzetten van backup 
    Expand-Archive -Path "$backupzip" -DestinationPath "$startmap" -Force

    # zipbestand en backup na het uitpakken verwijderen
    Remove-Item "$zip_download"
    if (test-path -path "$backupzip") { Remove-Item "$backupzip" }

    return
    }

# de backup na het uitpakken verwijderen. de zipbestand niet omdat bij het starten dan gezien kan worden dat al een update is uitgevoerd.
# Remove-Item "$zip_download"
if (test-path -path "$backupzip") { Remove-Item "$backupzip" }

Write-Host -f Yellow "Het programma heeft een update uitgevoerd en heeft nu de versie $updateto.
Het programma wordt opnieuw opgestart  ..."

Foutenloggen -meldtekst "Het programma heeft een update uitgevoerd en heeft nu de versie $updateto." -type "INFO"

Start-Sleep -Seconds 5

# opnieuw opstarten script met een andere procesnummer, zie de regel met start-proces hieronder, wordt niet meer gebruikt sinds 4.5.3.
# Dit was nodig toen Sharepoint werd gebruikt. Nu wordt weer de oude methode gebruikt.
# Nu toch weer met start-proces omdat je dan een nieuw proces krijgt en de oude kan worden afgesloten.
start-process PowerShell.exe -argumentlist '-file',".\$programmanaam.ps1"
# powershell -file "$PSScriptRoot\$scriptnaam.ps1"

# beëindigen van huidige proces. Script gaat verder in het nieuwe proces dat met start-process is gestart.
exit;


} # einde updateuitvoeren

function vensterinstellingen {

# De hoofdmenu onzichtbaar maken
$form.Hide()

# variabelen
$keuzelocatie=$global:init["algemeen"]["locatiekeuze"]

# venster declareren
$Form2 = declareren_standaardvenster "Instellingen wijzigen" 1000 500

$keuzeoptie1                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie1.width               = 80
$keuzeoptie1.autosize            = $true
$keuzeoptie1.DropDownStyle       = "DropDownList"
$keuzeoptie1.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie1.location = "450,55" 

# locaties toevoegen aan lijst
$teller=0
foreach( $property in $global:beheer.locaties.keys ) {
    # locatiecode toevoegen
    [void] $keuzeoptie1.Items.Add("$property")

    # standaard keuzelocatie selecteren en in index van lijst zetten
    if ($property -eq "$keuzelocatie") { $keuzeoptie1.Selectedindex = $teller }
    $teller++
    } 
  

$keuzeoptie2                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie2.width               = 80
$keuzeoptie2.autosize            = $true
$keuzeoptie2.DropDownStyle       = "DropDownList"
$keuzeoptie2.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie2.location = "450,95" 
[void] $keuzeoptie2.Items.Add("Ja")
[void] $keuzeoptie2.Items.Add("Nee")
if ($global:init["algemeen"]["wissennabackup"] -eq "Ja") {
    $keuzeoptie2.Selectedindex = 0
    } else {
    $keuzeoptie2.Selectedindex = 1
    }

$keuzeoptie3                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie3.width               = 80
$keuzeoptie3.autosize            = $true
$keuzeoptie3.DropDownStyle       = "DropDownList"
$keuzeoptie3.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie3.location = "450,135" 
[void] $keuzeoptie3.Items.Add("Ja")
[void] $keuzeoptie3.Items.Add("Nee")
if ($global:init["algemeen"]["maplegenvoorverplaatsen"] -eq "Ja") {
    $keuzeoptie3.Selectedindex = 0
    } else {
    $keuzeoptie3.Selectedindex = 1
    }

$keuzeoptie4                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie4.width               = 80
$keuzeoptie4.autosize            = $true
$keuzeoptie4.DropDownStyle       = "DropDownList"
$keuzeoptie4.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie4.location = "450,175" 
[void] $keuzeoptie4.Items.Add("Ja")
[void] $keuzeoptie4.Items.Add("Nee")
if ($global:init["opschonen"]["opschonenlogs"] -eq "Ja") {
    $keuzeoptie4.Selectedindex = 0
    } else {
    $keuzeoptie4.Selectedindex = 1
    }

$keuzeoptie5 = New-Object System.Windows.Forms.TextBox 
$keuzeoptie5.Location = New-Object System.Drawing.Size(450,215) 
$keuzeoptie5.Size = New-Object System.Drawing.Size(80,60)
$keuzeoptie5.MaxLength = 4
$keuzeoptie5.Font = 'Microsoft Sans Serif,11'
$keuzeoptie5.Text=$Global:init["opschonen"]["dagenbewarenlogs"]
$keuzeoptie5.Add_TextChanged({
    $this.Text = $this.Text -replace '\D'
})

$keuzeoptie6                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie6.width               = 160
$keuzeoptie6.autosize            = $true
$keuzeoptie6.DropDownStyle       = "DropDownList"
$keuzeoptie6.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie6.location = "780,15" 

[int]$teller = 0
$global:afbeeldingen.ForEach( {
    [void] $keuzeoptie6.Items.Add($_.naam)
    if ($_.naam -eq $global:init["algemeen"]["afbeelding"]) {$keuzeoptie6.Selectedindex = $teller}
    $teller+=1

})

$keuzeoptie6.add_SelectedIndexChanged({
# een match zoeken met de gekozen afbeelding in je persoonlijke instellingen
$global:afbeeldingen.ForEach( {
    if ($_.naam -eq $keuzeoptie6.Selecteditem) {
        $keuzeafbeelding = $_.bestand
        $gifbox2.Image    = [System.Drawing.Image]::FromFile("$gifjesmap\$keuzeafbeelding")
        }
})

})

$keuzeoptie7                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie7.width               = 230
$keuzeoptie7.autosize            = $true
$keuzeoptie7.DropDownStyle       = "DropDownList"
$keuzeoptie7.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie7.location = "300,335" 

[int]$teller = 0
$global:moppen.ForEach( {
    [void] $keuzeoptie7.Items.Add($_.naam)
    if ($_.naam -eq $global:init["algemeen"]["websitemoppen"]) {$keuzeoptie7.Selectedindex = $teller}
    $teller+=1
})

$keuzeoptie8 = New-Object System.Windows.Forms.TextBox 
$keuzeoptie8.Location = New-Object System.Drawing.Size(300,15) 
$keuzeoptie8.Size = New-Object System.Drawing.Size(230,60)
$keuzeoptie8.MaxLength = 40
$keuzeoptie8.Font = 'Microsoft Sans Serif,11'
$keuzeoptie8.Text=$Global:init["algemeen"]["gebruiker"]

$keuzeoptie9                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie9.width               = 80
$keuzeoptie9.autosize            = $true
$keuzeoptie9.DropDownStyle       = "DropDownList"
$keuzeoptie9.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie9.location = "450,255" 
[void] $keuzeoptie9.Items.Add("Ja")
[void] $keuzeoptie9.Items.Add("Nee")
if ($global:init["algemeen"]["consolesluiten"] -eq "Ja") {
    $keuzeoptie9.Selectedindex = 0
    } else {
    $keuzeoptie9.Selectedindex = 1
    }

$keuzeoptie10                     = New-Object system.Windows.Forms.ComboBox
$keuzeoptie10.width               = 80
$keuzeoptie10.autosize            = $true
$keuzeoptie10.DropDownStyle       = "DropDownList"
$keuzeoptie10.Font                = 'Microsoft Sans Serif,12'
$keuzeoptie10.location = "450,295" 
[void] $keuzeoptie10.Items.Add("Ja")
[void] $keuzeoptie10.Items.Add("Nee")
if ($global:init["algemeen"]["controlevoorklaarzetten"] -eq "Ja") {
    $keuzeoptie10.Selectedindex = 0
    } else {
    $keuzeoptie10.Selectedindex = 1
    }

$Description1                     = New-Object system.Windows.Forms.Label
$Description1.text                = "Standaard locatie"
$Description1.AutoSize            = $false
$Description1.width               = 400
$Description1.height              = 30
$Description1.location            = New-Object System.Drawing.Point(40,60)
$Description1.Font                = 'Microsoft Sans Serif,11'
$Description1.ForeColor = [System.Drawing.Color]::Blue
$Description1.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Met de standaard locatie wordt bepaald van welke locatie de rpc-nummers geselecteerd kunnen worden." )
})

$Description2                     = New-Object system.Windows.Forms.Label
$Description2.text                = "Homemap kandidaten wissen na uitvoeren van een back-up"
$Description2.AutoSize            = $false
$Description2.width               = 400
$Description2.height              = 40
$Description2.location            = New-Object System.Drawing.Point(40,100)
$Description2.Font                = 'Microsoft Sans Serif,11'
$Description2.ForeColor = [System.Drawing.Color]::Blue
$Description2.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bepaal of de optie om homemappen te wissen standaard is geselecteerd bij het uitvoeren van een back-up." )
})

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "Doelmap wissen alvorens het verplaatsen van bestanden"
$Description3.AutoSize            = $false
$Description3.width               = 400
$Description3.height              = 30
$Description3.location            = New-Object System.Drawing.Point(40,140)
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.ForeColor = [System.Drawing.Color]::Blue
$Description3.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bepaal of de optie om de doelmap te wissen standaard is geselecteerd bij het verplaatsen van bestanden." )
})

$Description4                     = New-Object system.Windows.Forms.Label
$Description4.text                = "Automatisch verwijderen oude logbestanden bij de start van het programma"
$Description4.AutoSize            = $false
$Description4.width               = 400
$Description4.height              = 40
$Description4.location            = New-Object System.Drawing.Point(40,180)
$Description4.Font                = 'Microsoft Sans Serif,11'
$Description4.ForeColor = [System.Drawing.Color]::Blue
$Description4.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bepaal of bij het opstarten logbestanden verwijderd worden. Hierbij wordt gebruikt gemaakt van de instelling hieronder." )
})

$Description5                     = New-Object system.Windows.Forms.Label
$Description5.text                = "Aantal dagen dat de logbestanden van het programma bewaard blijven"
$Description5.AutoSize            = $false
$Description5.width               = 400
$Description5.height              = 40
$Description5.location            = New-Object System.Drawing.Point(40,220)
$Description5.Font                = 'Microsoft Sans Serif,11'
$Description5.ForeColor = [System.Drawing.Color]::Blue
$Description5.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bepaal tot hoeveel dagen de logbestanden bewaard blijven bij het verwijderen van oude logbestanden." )
})

$Description6                     = New-Object system.Windows.Forms.Label
$Description6.text                = "Afbeelding in het hoofdvenster"
$Description6.AutoSize            = $false
$Description6.width               = 400
$Description6.height              = 40
$Description6.location            = New-Object System.Drawing.Point(560,20)
$Description6.Font                = 'Microsoft Sans Serif,11'
$Description6.ForeColor = [System.Drawing.Color]::Blue
$Description6.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Wijzigen van albeelding in hoofdmenu" )
})


$Description7                     = New-Object system.Windows.Forms.Label
$Description7.text                = "Keuze voor website met moppen"
$Description7.AutoSize            = $false
$Description7.width               = 400
$Description7.height              = 40
$Description7.location            = New-Object System.Drawing.Point(40,340)
$Description7.Font                = 'Microsoft Sans Serif,11'
$Description7.ForeColor = [System.Drawing.Color]::Blue
$Description7.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bepaal welke website standaard de moppen genereerd." )
})


$Description8                     = New-Object system.Windows.Forms.Label
$Description8.text                = "Jouw naam"
$Description8.AutoSize            = $false
$Description8.width               = 400
$Description8.height              = 40
$Description8.location            = New-Object System.Drawing.Point(40,20)
$Description8.Font                = 'Microsoft Sans Serif,11'
$Description8.ForeColor = [System.Drawing.Color]::Blue
$Description8.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Je naam zoals die wordt weergegeven in het hoofdvenster." )
})

$Description9                     = New-Object system.Windows.Forms.Label
$Description9.text                = "Console, venster met informatie over het script, verbergen als het hoofdvenster wordt getoond."
$Description9.AutoSize            = $false
$Description9.width               = 400
$Description9.height              = 40
$Description9.location            = New-Object System.Drawing.Point(40,260)
$Description9.Font                = 'Microsoft Sans Serif,11'
$Description9.ForeColor = [System.Drawing.Color]::Blue
$Description9.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bepaal of je de console, na het opstartproces, wil sluiten." )
})

$Description10                     = New-Object system.Windows.Forms.Label
$Description10.text                = "Controleer of de homemappen van de kandidaten leeg zijn voordat bestanden worden klaargezet."
$Description10.AutoSize            = $false
$Description10.width               = 400
$Description10.height              = 40
$Description10.location            = New-Object System.Drawing.Point(40,300)
$Description10.Font                = 'Microsoft Sans Serif,11'
$Description10.ForeColor = [System.Drawing.Color]::Blue
$Description10.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bepaal of je de homemappen van de kandidaten wil controleren voordat je bestanden klaarzet." )
})

$Btnstandaard = New-object System.Windows.Forms.Button 
$Btnstandaard.text= "Herstel de standaardinstellingen"
$Btnstandaard.location = "450,390" 
$Btnstandaard.size = "250,30"  
$Btnstandaard.BackColor = 'blue'
$Btnstandaard.ForeColor = 'white'
$Btnstandaard.add_click({
    # inlezen ini-bestand naar tijdelijke object
    $temp_init = gebruikersinstellingen

    # alle opties krijgen hun standaard waarde. Er wordt nog niets bewaard!
    $keuzeoptie1.SelectedItem=$temp_init["algemeen"]["locatiekeuze"]
    $keuzeoptie2.SelectedItem=$temp_init["algemeen"]["wissennabackup"]
    $keuzeoptie3.SelectedItem=$temp_init["algemeen"]["maplegenvoorverplaatsen"]
    $keuzeoptie4.SelectedItem=$temp_init["opschonen"]["opschonenlogs"]
    $keuzeoptie5.Text=$temp_init["opschonen"]["dagenbewarenlogs"]
    $keuzeoptie6.SelectedItem=$temp_init["algemeen"]["afbeelding"]
    $keuzeoptie7.SelectedItem=$temp_init["algemeen"]["websitemoppen"]
    $keuzeoptie9.SelectedItem=$temp_init["algemeen"]["consolesluiten"]
    $keuzeoptie10.SelectedItem=$temp_init["algemeen"]["controlevoorklaarzetten"]
}) # einde Btnstandaard.add_click
$Btnstandaard.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Zet alle instellingen terug naar de standaard waarden." )
})

$Btnaccept = New-object System.Windows.Forms.Button 
$Btnaccept.text= "Bewaren"
$Btnaccept.location = "50,390" 
$Btnaccept.size = "150,30"  
$Btnaccept.BackColor = 'green'
$Btnaccept.ForeColor = 'white'
$Btnaccept.DialogResult = [System.Windows.Forms.DialogResult]::yes
$Btnaccept.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Sla de wijzingen op en ga terug naar het hoofdvenster." )
})

$Btnescape = New-object System.Windows.Forms.Button 
$Btnescape.text= "Terug"
$Btnescape.location = "250,390" 
$Btnescape.size = "150,30"  
$Btnescape.BackColor = 'red'
$Btnescape.ForeColor = 'white'
$Btnescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel
$Btnescape.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster zonder de wijzigingen op te slaan." )
})

#Gekozen afbeelding tonen naast de opties
$gifbox2 = New-Object Windows.Forms.picturebox
$gifbox2.AutoSize = $false

# standaard waarde geven voor het geval geen match gevonden is hieronder
$keuzeafbeelding = 'Samenwerken'

# een match zoeken met de gekozen afbeelding in je persoonlijke instellingen
$global:afbeeldingen.ForEach( {
    if ($_.naam -eq $keuzeoptie6.Selecteditem) {
        $keuzeafbeelding = $_.bestand
        }
})
$gifbox2.Image    = [System.Drawing.Image]::FromFile("$gifjesmap\$keuzeafbeelding")
$gifbox2.location = New-Object System.Drawing.Point(570,40)
$gifbox2.Size     = "370,360" 
$gifbox2.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$gifbox2.add_click({
    if ($keuzeoptie6.Selectedindex -eq 6) {
        $keuzeoptie6.Selectedindex = 0
        } else {
        $keuzeoptie6.Selectedindex++
        }
})
$gifbox2.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Klik hier voor de volgende afbeelding." )
})


# venster met uitleg over deze taak wordt gedeclareerd. hieronder worden enkele variabelen aangepast aan deze taakvenster
declareren_uitlegvenster "Uitleg over het venster Instellingen." 680 300 920 400 "Wijzig hier de standaard instellingen van het programma. 
Met deze instellingen wijzig je de voorkeuren die standaard zijn ingesteld bij een taak
maar vaak kan je de voorkeuren bij een taak nog voor het uitvoeren aanpassen.
De instellingen voor je naam en de afbeelding zie je alleen terug in het hoofdscherm.

U krijgt extra informatie over een instelling als de muiscursor op een tekst staat.
U kunt alle instellingen herstellen naar de standaardwaarde door op de knop 
Herstel de standaardinstellingen te klikken.

Als u de instellingen wilt bewaren klikt u op Bewaren.
Als u terug wilt zonder de instellingen te bewaren klikt u op Terug.
" 

$Form2.Controls.AddRange(@($keuzeoptie8, $keuzeoptie6, $keuzeoptie1, $keuzeoptie2, $keuzeoptie3, $keuzeoptie4, $keuzeoptie5, $keuzeoptie9, $keuzeoptie10, $keuzeoptie7,
$description1, $description2, $description3, $description4, $description5, $description6, $description8, $description9, $description10, $description7,
$Btnaccept, $Btnescape, $Btnstandaard, $Global:vraagtekenicoon, $gifbox2 ))

Form2afsluitenbijescape;

# openen venster
$result = $form2.ShowDialog()

# bewaren van instellingen
if ($result -eq [system.windows.forms.dialogResult]::yes) { 

    # keuzes worden ingesteld
    $global:init["algemeen"]["locatiekeuze"]=$keuzeoptie1.Selecteditem
    $global:init["algemeen"]["wissennabackup"]=$keuzeoptie2.Selecteditem
    $global:init["algemeen"]["maplegenvoorverplaatsen"]=$keuzeoptie3.Selecteditem
    $global:init["opschonen"]["opschonenlogs"]=$keuzeoptie4.Selecteditem
    $global:init["algemeen"]["websitemoppen"]=$keuzeoptie7.Selecteditem
    $global:init["algemeen"]["consolesluiten"]=$keuzeoptie9.Selecteditem
    $global:init["algemeen"]["controlevoorklaarzetten"]=$keuzeoptie10.Selecteditem
    # alle spaties aan begin en eind weghalen
    # $keuzeoptie8.Text.Trim()
    $global:init["algemeen"]["gebruiker"]=$keuzeoptie8.Text.Trim()

    # Ook gelijk weergeven in hoofdvenster
    $Description.text = "Welkom " + $global:init.algemeen.gebruiker

    if (!($keuzeoptie5.Text -eq "" )) {
        # eventueel de nullen ervoor weghalen
        [int32]$getal1=$keuzeoptie5.Text
        # alleen als getal1 groter of gelijk is aan 0 is wordt de wijging doorgevoerd
        if ($getal1 -ge 0) { $global:init["opschonen"]["dagenbewarenlogs"]=$getal1 }
    }
    $global:init["algemeen"]["afbeelding"]=$keuzeoptie6.Selecteditem
    # afbeelding in hoofdmenu wordt aangepast
    # een match zoeken met de gekozen afbeelding in je persoonlijke instellingen
    $global:afbeeldingen.ForEach( {
    if ($_.naam -eq $global:init["algemeen"]["afbeelding"]) {
        $keuzeafbeelding = $_.bestand
        $gifBox.Image    = [System.Drawing.Image]::FromFile("$gifjesmap\$keuzeafbeelding")
        }
        })

    # bestand met nieuwe variabele bewaren.
    # Eerst wordt de persoonlijke initialisatiebestand bepaald.
    $gebruikersbestand = bepaalinitnaamgebruiker
    $global:init | ConvertTo-Json -depth 1 | Set-Content -Path $gebruikersbestand

    # Console open houden of sluiten
    if (($global:programma.mode -eq 'release') -or ($global:programma.mode -eq 'prerelease')) {
        if ($keuzeoptie9.Selectedindex -eq 0) {
        ShowHide-ConsoleWindow 0;
        } else {
        ShowHide-ConsoleWindow 5;
        }
    }

} # einde bewaren van instellingen

# De hoofdmenu zichtbaar maken
$form.show()

} # einde vensterinstellingen

function vensteropschonen {

# inlezen van backupmappen, bepalen aantal te verwijderen mappen en vullen van venster
function inlezenmappen {

if ($objTextBox1.Text -eq "") {
    $null = venstermetvraag -titel "Geen getal ingevuld" -vraag "`r`nVul een geldige getal in voor het aantal dagen waarna een back-up wordt verwijderd."
    # $objTextBox1.Text = "0"
    $Btnaccept.Enabled=$false 
    $listBox.items.clear()
    return;
}

# ingevoerde getallen uit de invulbox halen en eventueel de nullen ervoor eruit halen. dit laatste is een controle.
# getal1 is voor de backups
[int32]$getal1=$objTextBox1.Text
$objTextBox1.Text=$getal1.Tostring()

# aantal backups die worden verwijderd tellen
$tellerbackups = 0
# hier wordt er ingelezen, aantal te verwijderen mappen bepaald en venster gevuld.
$listBox.items.clear()

# beheer-variabele krijgt hier een verkorte naam tbv leesbaarheid
$backupmap = $global:beheer.examenmappen.backupmap
# inlezen backupmap en fouten opvangen
try {
   Get-ChildItem -Path "$backupmap" -ErrorAction Stop | Where-Object {($_.psiscontainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-$getal1))} | foreach-object {
        $tellerbackups++
        $listbox.Items.Add($_) 
        }
   
}
catch {
        $melding = -join ("Taak verwijderen van backups van de homemappen is gestart : ", "`n", $_.exception.message )
        Foutenloggen $melding
       }
# Knop Bevestigen laten zien als er mappen zijn om te verwijderen
if ($tellerbackups -eq 0 ) { $Btnaccept.Enabled=$false }
      else { $Btnaccept.Enabled=$true }
}
# einde function inlezenmappen


# ------ begin van de function vensteropschonen ---------------------

# controleren of de mappen beschikbaar zijn
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.backupmap $true )) { return; } 

# De hoofdmenu onzichtbaar maken
$form.Hide()

# venster declareren
$Form2 = declareren_standaardvenster "Verwijder gemaakte back-ups van de homemappen van de kandidaten" 650 600

$Description1                     = New-Object system.Windows.Forms.Label
$Description1.text                = "Back-ups van de homemappen die ouder zijn dan"
$Description1.AutoSize            = $false
$Description1.width               = 360
$Description1.height              = 30
$Description1.location            = New-Object System.Drawing.Point(20,15)
$Description1.Font                = 'Microsoft Sans Serif,11'
$Description1.ForeColor = [System.Drawing.Color]::blue


$Description2                     = New-Object system.Windows.Forms.Label
$Description2.text                = "dagen worden verwijderd."
$Description2.AutoSize            = $false
$Description2.width               = 200
$Description2.height              = 30
$Description2.location            = New-Object System.Drawing.Point(440,15)
$Description2.Font                = 'Microsoft Sans Serif,11'
$Description2.ForeColor = [System.Drawing.Color]::blue

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "Onderstaande mappen worden verwijderd."
$Description3.AutoSize            = $false
$Description3.width               = 800
$Description3.height              = 30
$Description3.location            = New-Object System.Drawing.Point(20,85)
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.ForeColor = [System.Drawing.Color]::Blue

$Description4                     = New-Object system.Windows.Forms.Label
$Description4.text                = "Wijzig eventueel het aantal dagen hierboven en klik op Opnieuw inlezen."
$Description4.AutoSize            = $false
$Description4.width               = 800
$Description4.height              = 30
$Description4.location            = New-Object System.Drawing.Point(20,50)
$Description4.Font                = 'Microsoft Sans Serif,11'
$Description4.ForeColor = [System.Drawing.Color]::Blue

$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Location = New-Object System.Drawing.Size(380,13) 
$objTextBox1.Size = New-Object System.Drawing.Size(55,50)
$objTextBox1.MaxLength = 4
$objTextBox1.Font = 'Microsoft Sans Serif,11'
$objTextBox1.Text=$Global:beheer.opschonen.dagenbewarenbackup
$objTextBox1.Add_TextChanged({
    $this.Text = $this.Text -replace '\D'
})
$objTextBox1.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Wijzigen hier eventueel tot hoeveel dagen oud de back-ups bewaard blijven." )
})

$global:listbox = New-Object System.Windows.Forms.Listbox
$global:listbox.Location = New-Object System.Drawing.Point(20,115)
$global:listbox.Size = New-Object System.Drawing.Size(250,20)
$global:listbox.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$global:listbox.Height = 400
$global:listbox.SelectionMode = 0
$global:listbox.HorizontalScrollbar = $true

$Btnaccept = New-object System.Windows.Forms.Button 
$Btnaccept.text= "Bevestigen"
$Btnaccept.location = "50,510" 
$Btnaccept.size = "150,30"  
$Btnaccept.BackColor = 'green'
$Btnaccept.ForeColor = 'white'
$Btnaccept.add_click({ 
    inlezenmappen
    if ($Btnaccept.Enabled) { overzichttaken "opschonen"} 
    })

# alleen klikbaar als aantal gevonden bestanden hoger is dan 0
if ($tellerbackups -eq 0 ) { $Btnaccept.Enabled=$false }
    
$Btnaccept.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bevestig je keuze en ga door naar het overzicht." )
})

$Btnescape = New-object System.Windows.Forms.Button 
$Btnescape.text= "Terug"
$Btnescape.location = "250,510" 
$Btnescape.size = "150,30"
$Btnescape.BackColor = 'red'
$Btnescape.ForeColor = 'white'
$Btnescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel
$Btnescape.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})

$BtInlezen = New-object System.Windows.Forms.Button 
$BtInlezen.text= "Opnieuw inlezen"
$BtInlezen.location = "300,115" 
$BtInlezen.size = "150,30"  
$BtInlezen.BackColor = 'blue'
$BtInlezen.ForeColor = 'white'
$BtInlezen.add_click({ 
    inlezenmappen
    })
# $BtInlezen.Enabled=$false
$BtInlezen.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Opnieuw inlezen van de mappen als het aantal dagen is gewijzigd." )
})

# venster met uitleg over deze taak wordt gedeclareerd.
declareren_uitlegvenster "Uitleg over het venster Verwijder oude back-ups en oude logbestanden." 700 200 450 510 "Hier kunnen oude back-ups van de homemappen van de kandidaten verwijderd worden.

In het venster zie je een overzicht van de mappen die vewijderd worden met de huidige instellingen.
Wijzigen eventueel tot hoeveel dagen oud de bestanden bewaard blijven en klik op Opnieuw inlezen.

Om naar het overzicht te gaan waar u het verwijderen kan starten moet u op Bevestigen klikken." 


$Form2.Controls.AddRange(@( $objTextBox1, $global:listbox, $description1, $description2, $description3, $description4,  $BtInlezen, $Btnaccept, $Btnescape, $Global:vraagtekenicoon ))


# eerst inlezen mappen
inlezenmappen;

Form2afsluitenbijescape;

# openen venster
$null = $form2.ShowDialog()

# De hoofdmenu zichtbaar maken
$form.show()

} # einde vensteropschonen


function informatieprogramma {

function info_venster_vullen ($keuze) {

# infovenster en tijdelijke object legen. wordt gebruikt bij function informatieprogramma
$objtekst1.Text = "Het bestand wordt geladen ..."
$objtekst_temp.text = ""

# tijdelijke variabelen benoemen
$tempreadmebestand="$startmap\$readmebestand"
$tempchangelogbestand="$startmap\$changelogbestand"

# inhoud bestand inlezen
if ($keuze -eq "readme") { 
    if ((test-path -path $tempreadmebestand -pathtype leaf)) { 
        $volledigetext = Get-Content -Path "$tempreadmebestand"
        } else {
        $volledigetext = "Het Readme-document is niet gedownload en kan dus niet getoont worden."
        }
    } else {
    if ((test-path -path $tempchangelogbestand -pathtype leaf)) { 
        $volledigetext = Get-Content -Path "$tempchangelogbestand"
        } else {
        $volledigetext = "Het Changelog-document is niet gedownload en kan dus niet getoont worden."
        }
    }

# dan netjes in rijen plaatsen.
foreach ($item in $volledigetext) {
               $objtekst_temp.Text = $objtekst_temp.Text + "$item" + "`r`n"
               }

$objtekst1.Text = $objtekst_temp.text 
} # einde info_venster_vullen

# Begin van function info_venster_vullen

# De hoofdmenu onzichtbaar maken
$form.Hide()

# aangeven wat de inhoud is van de infovenster, readme of changelog
$global:infovenster = "changelog"

# Powershell versie
$psmajor = $PSVersionTable.PSVersion.Major
$psminor = $PSVersionTable.PSVersion.Minor
$psbuild = $PSVersionTable.PSVersion.Build
$psrevision = $PSVersionTable.PSVersion.Revision

$psversie = "$PSMajor.$PSMinor.$PSbuild.$psrevision"
 
# venster declareren
$Form2 = declareren_standaardvenster "Informatie over het programma" 960 690

$Description2                     = New-Object system.Windows.Forms.Label
$Description2.text                = "Naam van het programma :
Versie : 
Build :
PowerShell versie :
Installatiemap :
Programmeur :
Afbeeldingen :"

$Description2.AutoSize            = $false
$Description2.width               = 220
$Description2.height              = 120
$Description2.location            = New-Object System.Drawing.Point(20,10)
$Description2.Font                = 'Microsoft Sans Serif,11'
$Description2.ForeColor = [System.Drawing.Color]::Blue

$versie = $global:programma.versie
$extralabel = $global:programma.extralabel
$scriptnaam = $global:programma.naam

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "$scriptnaam
$versie
$extralabel
$psversie
$startmap
Benvindo Neves
Marco Laluan"

$Description3.AutoSize            = $false
$Description3.width               = 1200
$Description3.height              = 115
$Description3.location            = New-Object System.Drawing.Point(250,10)
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.ForeColor = [System.Drawing.Color]::Blue


$Btninfovenster = New-object System.Windows.Forms.Button 
#$Btninfovenster.text= "Changelog bekijken"
$Btninfovenster.location = "230,600" 
$Btninfovenster.size = "180,30"  
$Btninfovenster.BackColor = 'blue'
$Btninfovenster.ForeColor = 'white'
if ($global:infovenster -eq "readme") { 
        $Btninfovenster.text= "Changelog bekijken"
        } else { 
        $Btninfovenster.text= "Readme bekijken"
        } 
$Btninfovenster.add_click({ 
    if ($global:infovenster -eq "readme") { 
        $global:infovenster="changelog" 
        $Btninfovenster.text= "Readme bekijken"
        } else { 
        $global:infovenster="readme" 
        $Btninfovenster.text= "Changelog bekijken"
        } 
    # infovenster vullen met infobestand
    info_venster_vullen $global:infovenster
})
$Btninfovenster.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Verander de inhoud van het informatievakje tussen Readme en Changelog." )
})


$Buttenok = New-object System.Windows.Forms.Button 
$Buttenok.text= "Terug"
$Buttenok.location = "50,600" 
$Buttenok.size = "150,30"  
$Buttenok.BackColor = 'red'
$Buttenok.ForeColor = 'white'
$Buttenok.DialogResult = [System.Windows.Forms.DialogResult]::ok
$Buttenok.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})

$objtekst1 = New-Object System.Windows.Forms.textbox
$objtekst1.Location = New-Object System.Drawing.Size(25,135) 
$objtekst1.Size = New-Object System.Drawing.Size(920,445)
$objtekst1.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$objtekst1.Text = ""
$objtekst1.ReadOnly = $true
$objtekst1.Multiline = $true
$objtekst1.ScrollBars = "Both"
$objtekst1.BackColor  = 'white'

# dit object maakt dat het laden van het bestand en laten zien in het venster sneller gaat
# zie ook function info_venster_vullen
$objtekst_temp = New-Object System.Windows.Forms.textbox

# infovenster vullen met infobestand
info_venster_vullen $global:infovenster

# venster met uitleg over deze taak wordt gedeclareerd. 
declareren_uitlegvenster "Uitleg over het venster Informatie over het programma." 680 250 500 600 "Bovenaan ziet u enkele gegevens over dit programma.

In het grote vakje kunt u eventueel de readme- of de changelog-bestand bekijken.
De README-bestand is een bestand die eerst gelezen moet worden, voorafgaand aan compilatie, installatie of eerste gebruik.
De CHANGELOG-bestand is een bestand met de wijzigingen per versie.

Door op de blauwe knop onderaan te klikken wijzigt u de inhoud." 

$Form2.Controls.AddRange(@($Description2, $Description3, $Buttenok, $Btninfovenster, $objtekst1, $Global:vraagtekenicoon))

Form2afsluitenbijescape;

# venster starten
$null = $form2.ShowDialog()


# De hoofdmenu zichtbaar maken
$form.show()
} # einde informatieprogramma

function opschonenlogsbijstart {


# alleen starten als dit is ingesteld in opschonenlogs en programma.mode niet de status alpha heeft.
if (($global:init.opschonen.opschonenlogs -eq "Nee") -or ("alpha" -contains($global:programma.mode)) ) {
    return;
}

# alleen starten als map met logbestanden bestaat
if (!(test-path -path "$logmap")) {
return;
}

# variabelen definiëren -----------------

# tijdelijk logbestand bepalen
$tijdelijkelog = "tijdelijkelog.txt"

# logbestandsnaam definiëren en volledige pad naar bestand invoeren
$logbestand = -join ("$logmap","\",$tijdelijkelog)

# starttijd van loggen naar variabele
$logtijd = bepaaltijd

#foutmelding voor logbestand
$foutmelding_log="[ FOUT ] "

# Eerste regel voor logbestand over het starten van opschonen
$melding_opschonen_begint = "Het opschonen van logbestanden is gestart. "

# Deze variabele bepaalt de "weer te geven logbestanden". Dit is in een function geplaatst om op 1 plek veranderingen door te voeren.
$weertegevenlogs = bepaaleigenlogbestanden

# aantal dagen dat bestanden bewaard blijven uit object halen
[int32]$getal2 = $global:init.opschonen.dagenbewarenlogs

# einde variabelen definiëren ----------------------

<# Controleren of er logbestanden zijn om te verwijderen. 
   Hiervoor worden de eigen logbestanden die kunnen worden verwijderd geteld.
   Met try-catch fouten worden fouten opvangen
#>

$tellerlogs=0
try {
    Get-ChildItem -Path $weertegevenlogs -ErrorAction Stop | Where-Object {(!($_.psiscontainer) -and $_.LastWriteTime -lt (Get-Date).AddDays(-$getal2))} | foreach-object {
        $tellerlogs++
    }
} 

catch {
        # melding loggen en weergeven
        $melding = -join ($melding_opschonen_begint, "`n", $_.exception.message )
        Foutenloggen $melding
        Write-Host $melding -ForegroundColor Red
        return;
}

# Stoppen als er niets is om te verwijderen
if ( $tellerlogs -eq 0) {
    return;
}

# start proces verwijderen

Write-Host "Opschonen van $tellerlogs logbestanden."

# in logbestand schrijven
$melding_opschonen_begint | out-file $logbestand -Append
"Starttijd : $logtijd" | out-file $logbestand -Append

# in logbestand info schrijven
"De volgende $tellerlogs logbestanden ouder dan $getal2 dagen verwijderen : " | out-file $logbestand -Append

# bepaal de lijst met mappen die verwijderd worden en opvangen foutmelding met try- catch methode
try {
    Get-ChildItem -Path $weertegevenlogs -ErrorAction Stop | Where-Object {(!($_.psiscontainer) -and $_.LastWriteTime -lt (Get-Date).AddDays(-$getal2))} | foreach-object {
        # map om te verwijderen
        $todelete = $_.Name
        # in logbestand schrijven
        " - $todelete " | out-file $logbestand -Append

        #verwijderen map 
        $error.clear()
        Remove-Item "$logmap\$todelete" -force -ErrorAction Stop 
   
    } # einde get-childitem        
} # einde try

catch {
      # foutmelding van PowerShell naar logbestand
      "$foutmelding_log" + $_ | out-file $logbestand -Append
      Write-Host $_ -ForegroundColor Red
} # einde try catch

# in logbestand eindtijd schrijven
# eerst eindtijd naar variabele
# alleen uitvoeren als aantal dagen verwijderen logs > 0. Zo wordt alles verwijderd als je dit als 0 instelt. 
# Nu blijft de laatste log met het opschonen bestaan.
if ( $getal2 -gt 0) {
    $logtijd = bepaaltijd

    "" | out-file $logbestand -Append
    "Eindtijd  : $logtijd" | out-file $logbestand -Append
    " -------------------------------------------------------------------------" | out-file $logbestand -Append
    "" | out-file $logbestand -Append

    # tijdelijke log toevoegen aan eigen logbestand als deze al bestaat, anders bestand hernoemen.
    Logbestandtoevoegen $logbestand

} else {
    # tijdelijke logbestand verwijderen
    Remove-Item -Path $logbestand -Force
}

} # einde opschonenlogsbijstart

function Venstermetgrap {
# Een venster verschijnt met een grap of een leuke weetje

function laadgrap {
# een random grap ophalen en in tekst-variabele plaatsen
# eerst de reload knop zichtbaar. voor als het goed gaat.
$reload.Enabled = $true
try {
    switch ($keuzewebsite.Selecteditem) {
        'Apekool.nl' {
        # toevoegen van -ContentType "application/json; charset=utf-8bom" heeft geen zin.
        $quote = Invoke-RestMethod -Method Get -Uri 'http://api.apekool.nl/services/jokes/getjoke.php' -erroraction Stop
        $tekst = $quote.joke 
          }
        'Appspot.com' {
        $quote = Invoke-RestMethod -Method Get -Uri 'https://official-joke-api.appspot.com/jokes/random' -erroraction Stop
        $tekst = $quote.setup + "`r`n" + $quote.punchline
          }
        'Icanhazdadjoke.com' {
        $Header = @{'Accept' = 'application/json' }
        $quote = (Invoke-RestMethod -Method Get -Uri 'https://icanhazdadjoke.com/' -Headers $Header -UseBasicParsing).joke
        $tekst = $quote
          }
        Default { 
            $tekst = 'Kies een website met moppen bij de keuzelijst hieronder.'
            $reload.Enabled = $false 
        }
    }
    
}
catch {
$tekst = "Het is niet gelukt een mop te krijgen van de API van " + $keuzewebsite.selecteditem + "`r`n" + "De website geeft de volgende melding: " + "`r`n" + $_.exception.message
Foutenloggen $tekst
$reload.Enabled = $false  
}

return $tekst
}

#venster aanmaken
$Form2 = declareren_standaardvenster "Moppen" 900 300;

$objtekst1 = New-Object System.Windows.Forms.textbox
$objtekst1.Location = New-Object System.Drawing.Size(20,20) 
$objtekst1.Size = New-Object System.Drawing.Size(850,180)
$objtekst1.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$objtekst1.ReadOnly = $true
$objtekst1.Multiline = $true
$objtekst1.TabStop = $false
#$objtekst1.ScrollBars = "Both"
$objtekst1.ScrollBars = 'vertical'
$objtekst1.BackColor  = 'white'
$Form2.Controls.Add($objtekst1);

$Buttenok = New-object System.Windows.Forms.Button 
$Buttenok.text= "Terug"
$Buttenok.location = "50,210" 
$Buttenok.size = "150,30"  
$Buttenok.BackColor = 'red'
$Buttenok.ForeColor = 'white'
$Buttenok.DialogResult = [System.Windows.Forms.DialogResult]::ok
$Buttenok.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})
$Form2.Controls.Add($Buttenok);

$reload = New-object System.Windows.Forms.Button 
$reload.text= "Nog een grap"
$reload.location = "250,210" 
$reload.size = "150,30"  
$reload.BackColor = 'green'
$reload.ForeColor = 'white'
$reload.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Nog een grap genereren." )
})
$reload.add_click({
    $objtekst1.Text = laadgrap
})
# hier naar toe verplaatst omdat $reload eerst gedefinieerd moet worden.
$objtekst1.Text = laadgrap

$Form2.Controls.Add($reload);

$keuzewebsite                     = New-Object system.Windows.Forms.ComboBox
$keuzewebsite.width               = 360
$keuzewebsite.autosize            = $true
$keuzewebsite.DropDownStyle       = "DropDownList"
$keuzewebsite.Font                = 'Microsoft Sans Serif,12'
$keuzewebsite.location = "500,210" 

[int]$teller = 0
$global:moppen.ForEach( {
    [void] $keuzewebsite.Items.Add($_.naam)
    # adhv ingestelde keuze bij instellingen bepaal je de geselecteerde item
    if ($_.naam -eq $global:init["algemeen"]["websitemoppen"]) {$keuzewebsite.Selectedindex = $teller}
    $teller+=1

})

$keuzewebsite.add_SelectedIndexChanged({
    $objtekst1.Text = laadgrap
})

# $reload en $keuzewebsite moeten eerst gedefinieerd moeten worden.
# dit is nodig zodat als er geen mop geladen kan worden, de knop reload uitgeschakeld is.
$objtekst1.Text = laadgrap

$Form2.Controls.Add($keuzewebsite);

# venster met uitleg over deze taak wordt gedeclareerd.
declareren_uitlegvenster "Uitleg over het venster Moppen." 700 200 450 210 "Hier wordt een mop gegenereerd en weergegeven.
Rechtsonder staat de website waaruit de mop is gehaald.
Klik hierop om een andere website te kiezen.
Met de knop Nog een grap kan je een nieuwe mop genereren.
De knop Terug brengt je terug naar het hoofdvenster." 

$Form2.controls.add($Global:vraagtekenicoon)
Form2afsluitenbijescape;

# venster starten
$null = $form2.ShowDialog()

} # einde functie Venstermetgrap

function vensterverkenner {

function toevoegen_lijst1 ($controlemap, $toevoegitem, $toevoegdatum, $toevoeggrootte) {
# toevoegen items en icoontjes, dit is een map of bestand, aan lijst

    if (( test-path -path "$controlemap\$toevoegitem" -pathtype container) -eq $true)  {
        [void] $listView1.Items.Add($toevoegitem, 0).SubItems.Add($toevoegdatum.ToString())
        } else {

        # Grootte van bestand weergeven in kilobytes
        if ($toevoeggrootte -eq 0) { $lengte = '0 kB'}
            elseif ($toevoeggrootte -lt 1024) { $lengte = '1 kB'}
            else { 
            $waarde =  [math]::round($toevoeggrootte / 1024)
            $lengte = -join ($waarde," kB")
            }
        $extensienr = Bepaalicoontjenr $controlemap $toevoegitem
        [void] $listView1.Items.Add($toevoegitem, $extensienr).SubItems.Addrange( @($toevoegdatum.ToString(),$lengte ) )
        }
}

Function Selecteermap {
# Selecteert de huidige gekozen map of submap.

    $selectie = $global:beheer.examenmappen.homemapstudenten
    $selectie = -join ($selectie,'\',$listBox.SelectedItem)

# eventuele geselecteerde submappen toevoegen
    foreach ($item in $geselecteerdebronmap) {
                $selectie = -join ($selectie, '\', $item)
    }
    return $selectie
}

Function geselecteerdebronmaplegen {

$mijndocumentenmap = 'Mijn Documenten'

$geselecteerdebronmap.Clear()
$geselecteerdebronmap.Add($mijndocumentenmap)
}

Function geselecteerdemap_vullen ($selectie) {


  try { 
  # weergeven inhoud in rechter venster (inhoud geselecteerde map). En fouten opvangen.
  $inhoud = Get-ChildItem -Path "$selectie" -ErrorAction Stop | Sort-object 

  # dan netjes in rijen plaatsen.
  foreach ($item in $inhoud) {
        toevoegen_lijst1 $selectie $item.Name $item.LastWriteTime $item.Length
        $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $Item.Name)
       }
  }
  catch {
  # melding loggen en weergeven
  $melding = -join ("Venster Verkenner is geopend : ", "`n", $_.exception.message )
  Foutenloggen $melding
  }

  # aanpassen venster aan inhoud
  $listView1.AutoResizeColumns(1) 

  # Weergeven submappen van homemap kandidaat in bovenste venster (submappenvenster)
  if ( $listBox.selecteditems.count -eq 1) {
        $objtekst2.Text = $listBox.SelectedItem
     } else {
        $objtekst2.Text = ''
     }

  foreach ($item in $geselecteerdebronmap) {
          $objtekst2.Text = -join ($objtekst2.Text, $Scheidingstekst, $Item)
  }
} # einde geselecteerdemap_vullen


# ----------- Start function VensterVerkenner -------------------------------

# variabelen krijgen hier een verkorte naam tbv leesbaarheid en gebruik in andere functies
$keuzelocatie     = $global:init["algemeen"]["locatiekeuze"]
# lijst met geselecteerde examenmappen
$global:geselecteerdebronmap = [System.Collections.ArrayList]@()

# controleren of de mappen beschikbaar zijn
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.homemapstudenten $true )) { return; } 

# rpcnrs declareren met een function
$global:listBox = declareren_rpcnrs;
$global:listBox = lijstrpcnrsaanmaken $keuzelocatie $global:listBox 

# De hoofdmenu onzichtbaar maken
$form.Hide()

$Form2 = declareren_standaardvenster "Bestanden van homemappen van de kandidaten bekijken" 900 650;

$Description2                     = New-Object system.Windows.Forms.Label
$Description2.text                = "RPC-nummers"
$Description2.AutoSize            = $false
$Description2.width               = 150
$Description2.height              = 20
$Description2.location            = New-Object System.Drawing.Point(20,15)
$Description2.Font                = 'Microsoft Sans Serif,11'
$Description2.ForeColor = [System.Drawing.Color]::Blue
$Form2.controls.add($Description2)

$Description3                     = New-Object system.Windows.Forms.Label
$Description3.text                = "Locatie"
$Description3.AutoSize            = $false
$Description3.width               = 400
$Description3.height              = 20
$Description3.Font                = 'Microsoft Sans Serif,11'
$Description3.location            = New-Object System.Drawing.Point(210,15)
$Description3.ForeColor = [System.Drawing.Color]::Blue
$Form2.controls.add($Description3)

# listbox is al in het begin gedeclareerd. onderstaande waarden gelden voor deze functie.
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(130,500)
$listBox.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$listBox.SelectionMode = 'One'
$listBox.BackColor  = 'white'
$listBox.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Selecteer de RPC-nummer waarvan je de bestanden wilt zien." )
})
$Form2.controls.add($listBox)

$listBox.add_SelectedIndexChanged( { 

    # Venster inhoud kandidatenmap legen
    $listView1.items.clear()
    # geselecteerde submappen van inhoud kandidatenmap legen
    geselecteerdebronmaplegen
    # selecteert de gekozen map of submap
    $selectie = Selecteermap

    # alleen als 1 item is geselecteerd 
    if ( $listBox.selecteditems.count -eq 1) {
          # weergeven inhoud in rechter venster, inhoud geselecteerde map
          geselecteerdemap_vullen $selectie

          }  # einde $listView1.selecteditems.count -eq 1 
     } ) # einde $listBox.add_SelectedIndexChanged

# Standaard lijst met alle locaties maken, met standaardwaarden
$lijstlocaties = declarerenlijstlocaties $keuzelocatie 330 200 40

$Form2.controls.add($lijstlocaties)

# bij wijzigen van selectie lijstlocaties, andere rpc-nummers weergeven
$lijstlocaties.add_SelectedIndexChanged(
     { 
     [string]$waarde = $lijstlocaties.selecteditem
     $keuzelocatie = $waarde.Substring(0,3)
     # nieuwe rpcnrs declareren
     $global:listBox = lijstrpcnrsaanmaken $keuzelocatie $global:listBox 

     # Venster inhoud kandidatenmap legen
     $listView1.items.clear()
     $objtekst2.Clear()
     # geselecteerde submappen van inhoud kandidatenmap legen
     $geselecteerdebronmap.Clear()

     } ) 



$BtnOpnieuw                         = New-Object system.Windows.Forms.Button
$BtnOpnieuw.width                   = 40
$BtnOpnieuw.height                  = 40
$BtnOpnieuw.location                = New-Object System.Drawing.Point(150,80)
$BtnOpnieuw.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoontje-home.png")
$BtnOpnieuw.Add_Click({ 
      # alleen als 1 item is geselecteerd 
      if ( $listBox.selecteditems.count -eq 1) {
      # array wordt leeggemaakt en krijgt eventueel de Mijn documentenmap als 1e item
      geselecteerdebronmaplegen

      # selectie krijgt waarde van home-map nu array terug is naar standaard.
      $selectie = Selecteermap

      # vullen van rechter venster
      $listView1.items.clear()
      geselecteerdemap_vullen ($selectie)
      }
})
$BtnOpnieuw.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Terug naar de home-map van de geselecteerde RPC-nummer." )
})
$Form2.controls.add($BtnOpnieuw)

$BtnTerug                         = New-Object system.Windows.Forms.Button
$BtnTerug.width                   = 40
$BtnTerug.height                  = 40
$BtnTerug.location                = New-Object System.Drawing.Point(150,125)
$BtnTerug.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoon-terug.png")
$BtnTerug.Add_Click({ 
      # alleen als 1 item is geselecteerd EN er items zijn in array
      if (( $listBox.selecteditems.count -eq 1) -and ($geselecteerdebronmap.Count -gt 0)) {
        # laatste item in array verwijderen
        $verwijdernr=$geselecteerdebronmap.Count-1
        $geselecteerdebronmap.RemoveAt($verwijdernr)
        
        # selectie krijgt waarde van home-map nu array terug is naar standaard.
        $selectie = Selecteermap

        # vullen van rechter venster
        $listView1.items.clear()
        geselecteerdemap_vullen ($selectie)
      }
 }) # einde BtnTerug.add_click

$BtnTerug.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "één map terug gaan." )
})
$Form2.controls.add($BtnTerug)

# standaard tekst in venster met geselecteerd examenmap. Nodig voor Objtekst, weergave van geselecteerde homemap.
$Startexamenmap = "Homemap kandidaat"
# Tekst tussen twee geselecteerde mappen, om deze uit elkaar te houden
$Scheidingstekst = " » "

$objtekst2 = New-Object System.Windows.Forms.textbox
$objtekst2.Location = New-Object System.Drawing.Size(200,80)
$objtekst2.Size = New-Object System.Drawing.Size(660,47)
$objtekst2.font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$objtekst2.Text = ''
$objtekst2.ReadOnly = $true
$objtekst2.Multiline = $true
$objtekst2.BackColor  = 'white'
$objtekst2.Forecolor  = 'blue'
$objtekst2.WordWrap = $true
$objtekst2.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Hier zie je de huidige geselecteerde map." )
})
$Form2.controls.add($objtekst2)


# afbeeldingen declareren voor venster met inhoud van kandidaat-map
$imageList = Declareericoontjes

# Venster met inhoud van kandidaat-map
$global:listView1 = New-Object System.Windows.Forms.ListView
$listView1.View = 'Details'
$listView1.Height = 400
$listView1.Width = 660
$listView1.Font = New-Object System.Drawing.Font("MS Sans Serif",12)
# zorgen dat selectie zichtbaar blijft.
$listview1.HideSelection = $false
# Alleen 1 item kan geselecteert worden
$listview1.MultiSelect = $false
$listView1.FullRowSelect = $true
# locatie
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 200
$System_Drawing_Point.Y = 125
$listView1.Location = $System_Drawing_Point
# Naam
$listView1.Name = "listView1"
$listView1.Sorting = 'Ascending'
$listView1.Columns.Add('Inhoud geselecteerde RPC-nummer',300)| Out-Null
$listView1.Columns.Add('Gewijzigd op',200)| Out-Null
# $listView1.Columns.Add('Type',400)| Out-Null
$listView1.Columns.Add('Grootte',70)| Out-Null
# kolom 1 en 2 rechts centreren (kolom 1 mss toch niet doen?????)
$listView1.Columns[2].textalign=1
$listView1.Columns[1].textalign=1

$listView1.SmallImageList = $imageList
$listView1.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Hier zie je de inhoud van de geselecteerde RPC-nummer." )
})
$Form2.controls.add($listView1)

# bij dubbelklikken van een map, map selecteren en openen
$listView1.add_doubleClick( {
     # selecteert de gekozen map of submap
     $selectie = Selecteermap
     # toevoegen gekozen item aan selectie
     $selectie = -join ($selectie, '\', $listView1.SelectedItems.text)
      
     # Er moet een item geselecteerd zijn (je kan namelijk ook dubbelklikken op een lege plek) en de item moet een map zijn.
     if (( $listView1.selecteditems.count -eq 1) -and ((test-path -path $selectie -pathtype container) -eq $true) ) {
          # toevoegen aan array
          $geselecteerdebronmap.Add($listView1.SelectedItems.text)
          # Vullen van rechter venster
          $listView1.items.clear()
          geselecteerdemap_vullen $selectie
     }
} ) # einde $listView1.add_doubleClick

# Sorteren op een kolom die je aanklikt
$listView1.add_ColumnClick({SortListView $this $_.Column})

$Btnescape = New-object System.Windows.Forms.Button 
$Btnescape.text= "Terug"
$Btnescape.location = "50,550" 
$Btnescape.size = "150,30"  
$Btnescape.BackColor = 'red'
$Btnescape.ForeColor = 'white'
$Btnescape.DialogResult = [System.Windows.Forms.DialogResult]::cancel
$Btnescape.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Ga terug naar het hoofdvenster." )
})
$Form2.controls.add($Btnescape)

# venster met uitleg over deze taak wordt gedeclareerd. 
declareren_uitlegvenster "Uitleg over het venster Verkenner." 680 320 480 560 "Hier kunt u de inhoud van de homemappen van de kandidaten bekijken.
Dit kunt u gebruiken als controle na het uitvoeren van een taak.

Eventueel kunt u eerst kiezen om de standaardlocatie te wijzigen.
Links kiest u vervolgens het RPC-nummer van de betreffende homemap.
Rechts ziet u dan de mappen en bestanden van de homemap.

U kunt op een map dubbelklikken om de inhoud weer te geven.
Met het terug-icoontje gaat u een map terug.
Met het home-icoontje gaat u terug naar het begin.

Door op de blauwe knop onderaan te klikken gaat u terug naar het hoofdscherm." 
$Form2.controls.add($Global:vraagtekenicoon)

# venster tonen
# $form2.Topmost = $false

Form2afsluitenbijescape;

$null = $form2.ShowDialog()

# venster sluiten    
$form2.close();

# De hoofdmenu zichtbaar maken
$form.show()

} # einde functie vensterverkenner


<# ---------------------      Start script ---------------------------------------------------------------


#>


# controleren op een update. 
updateuitvoeren;

# Bepalen van de persoonlijke initialisatiebestand.
$gebruikersbestand = bepaalinitnaamgebruiker

write-host "Gebruikersinstellingen inlezen."

# Inlezen oude algemene initialisatiebestand en omzetten naar gebruikersbestand. Dit kan op een gegeven moment eruit.
# Oude initialisatiebestand werd gebruikt tot versie 4.5.0
$global:initbestand = -join ("$startmap","\","initialisatie.ini")
if (test-path -path $global:initbestand -pathtype leaf) { 
    # als nieuwe bestand al bestaat, verwijderen
    if (test-path -path $gebruikersbestand -pathtype leaf) {
         Remove-Item $gebruikersbestand
         }
    Rename-Item -Path $global:initbestand -NewName $gebruikersbestand
}

# Vanaf versie 4.7.0 naam gebruikersbestand wijzigen. Dit kan op een gegeven moment weg.
$oudenaam = -join ("$startmap","\gebruiker",$env:username,".ini")
$nieuwenaam = bepaalinitnaamgebruiker
if (test-path -path $oudenaam -pathtype leaf) { 
    Rename-Item -Path $oudenaam -NewName $nieuwenaam
}

# inlezen van gebruikers instellingen
Inlezengebruikersinstellingen;

#Overige controles en tijdelijke taken uitvoeren
write-host "Oude bestanden opschonen of herstellen."

# verwijderen updater.ps1. vanaf versie 4.5.0
if (test-path -path "$startmap\updater.ps1") { Remove-Item "$startmap\updater.ps1" } 
# verwijderen 2 bestanden vanaf versie 4.5.1. snelkoppeling_maken.ps1 wordt weer gebruikt vanaf 4.6.2
# if (test-path -path "$startmap\snelkoppeling_maken.ps1") { Remove-Item "$startmap\snelkoppeling_maken.ps1" } 
if (test-path -path "$startmap\beheren.ini") { Remove-Item "$startmap\beheren.ini" } 
if (test-path -path "$startmap\updateinfo.ini") { Remove-Item "$startmap\updateinfo.ini" }
# verwijderen beheer.ini vanaf versie 4.6.0
if (test-path -path "$startmap\beheer.ini") { Remove-Item "$startmap\beheer.ini" }  
# verwijderen filetransfer.gif vanaf versie 4.6.0
if (test-path -path "$startmap\png\filetransfer.gif") { Remove-Item "$startmap\png\filetransfer.gif" }  
# verwijderen netwerkschijvencontroleren.exe vanaf versie 4.6.1
if (test-path -path "$startmap\netwerkschijvencontroleren.exe") { Remove-Item "$startmap\netwerkschijvencontroleren.exe" }  
# verwijderen 1 map en 2 bestanden vanaf versie 4.6.2
if (test-path -path "$startmap\png\controleren-2.gif") { Remove-Item "$startmap\png\controleren-2.gif" }  
if (test-path -path "$startmap\png\nieuwe map") { Remove-Item "$startmap\png\nieuwe map" -Force -Recurse }
if (test-path -path "$startmap\snelkoppeling_maken.exe") { Remove-Item "$startmap\snelkoppeling_maken.exe" }   

# Dit staat hier voor versies lager dan 4.5.0 om de logbestanden te hernoemen. Kan op een gegeven moment verwijderd worden.
if (test-path -path "$logmap") {
$persnr = $env:username
Get-ChildItem -Path "$logmap\log_??????????.txt" -Name | ForEach-Object {
    $file = $_
    $nieuw = $file.Insert(14,'_'+$persnr)
    $oudbestand = -join ($logmap,'\',$file)
    Rename-Item -Path $oudbestand -NewName $nieuw
}
}

# Dit staat hier voor versies vanaf 4.5.3 om de logbestanden te hernoemen die fout zijn gegaan bij versie 4.5.2. Kan op een gegeven moment verwijderd worden.
if (test-path -path "$logmap") {
$persnr = $env:username
Get-ChildItem -Path "log\log_??????????_.txt" -Name | ForEach-Object {
    $file = $_
    $nieuw = $file.Insert(15,$persnr)
    $oudbestand = -join ($logmap,'\',$file)
    Rename-Item -Path $oudbestand -NewName $nieuw
}
}


# Controleren of netwerkschijven aanwezig zijn en eventueel herstellen.
# Eerste regel van elke foutmelding tijdens deze controle.
$foutmeldingbegin = "Initialiseren van het programma : "
# Als er een fout is dan hiermee onthouden.
$netwerkmapfout = $false

if (!(Netwerkmapaanwezig $global:beheer.examenmappen.digitalebestanden $false )) { 
    $tempmap = $global:beheer.examenmappen.digitalebestanden
    Foutenloggen "$foutmeldingbegin
De volgende Netwerkmap is niet gevonden
$tempmap
    "
    $netwerkmapfout = $true 
} 
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.homemapstudenten $false )) { 
    $tempmap = $global:beheer.examenmappen.homemapstudenten
    Foutenloggen "$foutmeldingbegin
De volgende Netwerkmap is niet gevonden
$tempmap
    "
    $netwerkmapfout = $true 
    } 
if (!(Netwerkmapaanwezig $global:beheer.examenmappen.backupmap $false )) { 
    $tempmap = $global:beheer.examenmappen.backupmap
    Foutenloggen "$foutmeldingbegin
De volgende Netwerkmap is niet gevonden
$tempmap
    "
    $netwerkmapfout = $true 
    } 

if ($netwerkmapfout) { Start-Sleep -Seconds 5}

# Einde controles en tijdelijke taken uitvoeren


# opschonen logbestanden als dit is ingesteld.
opschonenlogsbijstart;

write-host "Opstarten hoofdvenster."

# Hide Console. Alleen bij mode = release of prerelease
if (( ($global:programma.mode -eq 'release') -or ($global:programma.mode -eq 'prerelease')) -and ($global:init.algemeen.consolesluiten -eq 'Ja')) { 
    write-host "De console wordt afgesloten..."
    ShowHide-ConsoleWindow 0;
    }

# Hoofdvenster declareren --------------------------------------------------------------------

if ($global:programma.mode -eq 'release') {
    $koptekst = "Beheren bestanden versie " + $global:programma.versie
    } else {
    $koptekst = "Beheren bestanden versie " + $global:programma.versie + ". LET OP : Programma heeft de status " + $global:programma.mode
    }

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(700,480) #575
$Form.text                       = $koptekst
$Form.TopMost                    = $false
$Form.StartPosition              = 'CenterScreen'
$form.BackColor                  = "white"
$form.MaximizeBox                = $False
$form.Icon                       = [System.Drawing.Icon]::ExtractAssociatedIcon('beheren.ico')

$Description                     = New-Object system.Windows.Forms.Label
$Description.text                = "Welkom " + $global:init.algemeen.gebruiker
$Description.AutoSize            = $false
$Description.width               = 500
$Description.height              = 35
$Description.location            = New-Object System.Drawing.Point(164,15)
$Description.Font                = 'Microsoft Sans Serif,12'
$Form.controls.add($Description)


$uitleg1                     = New-Object system.Windows.Forms.Label
$uitleg1.text                = "Bestanden klaarzetten"
$uitleg1.AutoSize            = $false
$uitleg1.width               = 160
$uitleg1.height              = 20
$uitleg1.location            = New-Object System.Drawing.Point(168, 215)
$uitleg1.Font                = 'Microsoft Sans Serif,11'
$uitleg1.ForeColor = [System.Drawing.Color]::Blue
$Form.controls.add($uitleg1)


$uitleg2                     = New-Object system.Windows.Forms.Label
$uitleg2.text                = "Back-up maken"
$uitleg2.AutoSize            = $false
$uitleg2.width               = 160
$uitleg2.height              = 20
$uitleg2.location            = New-Object System.Drawing.Point(402, 215)
$uitleg2.Font                = 'Microsoft Sans Serif,11'
$uitleg2.ForeColor = [System.Drawing.Color]::blue
$Form.controls.add($uitleg2)

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.width                   = 160
$Button1.height                  = 160
$Button1.location                = New-Object System.Drawing.Point(164,50)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$Button1.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icon bestanden klaarzetten.png")
$Button1.BackColor = [System.Drawing.Color]::green
$Button1.Add_Click({ vensterkopieren })
$Button1.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bestanden klaarzetten voor de geselecteerde homemappen van de kandidaten" )
})
$Form.controls.add($Button1)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.width                   = 160
$Button2.height                  = 160
$Button2.location                = New-Object System.Drawing.Point(376,50)
$Button2.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$Button2.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icon backup maken.png")
$Button2.BackColor = [System.Drawing.Color]::green
$Button2.Add_Click({ vensterbackup })
$Button2.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Back-up maken van de geselecteerde homemappen van de kandidaten." )
})
$Form.controls.add($Button2)

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.width                   = 60
$Button4.height                  = 60
$Button4.location                = New-Object System.Drawing.Point(588,50)
$Button4.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$Button4.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icon verplaatsen.png")
# $Button4.BackColor = [System.Drawing.Color]::green
$Button4.Add_Click({ vensterverplaatsen })
$Button4.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bestanden van een pc verplaatsen of kopieëren naar een andere pc." )
})
$Form.controls.add($Button4)

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.width                   = 60
$Button3.height                  = 60
$Button3.location                = New-Object System.Drawing.Point(588,150)
$Button3.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$Button3.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icon wissen.png")
# $Button3.BackColor = [System.Drawing.Color]::green
$Button3.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bestanden verwijderen uit de geselecteerde pcs in het netwerk." )
})
$Button3.Add_Click({ vensterwissen })
$Form.controls.add($Button3)

$Button9                         = New-Object system.Windows.Forms.Button
$Button9.width                   = 60
$Button9.height                  = 60
$Button9.location                = New-Object System.Drawing.Point(588,250)
$Button9.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$Button9.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoon-settings.png")
$Button9.Add_Click({ vensterinstellingen })
$Button9.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "De persoonlijke keuzes voor het programma wijzigen." )
})
$Form.controls.add($Button9)

$Button10                         = New-Object system.Windows.Forms.Button
$Button10.width                   = 60
$Button10.height                  = 60
$Button10.location                = New-Object System.Drawing.Point(588,350)
$Button10.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoon-clean.png")
$Button10.Add_Click({ vensteropschonen })
$Button10.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Oude back-ups verwijderen." )
})
$Form.controls.add($Button10)

$Button7                         = New-Object system.Windows.Forms.Button
$Button7.width                   = 60
$Button7.height                  = 60
$Button7.location                = New-Object System.Drawing.Point(52,350)
$Button7.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$Button7.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoon-stoppen.png")
$Button7.BackColor = [System.Drawing.Color]::red
$Button7.Add_Click({ 
    # $result = programmaafsluiten
    $result = venstermetvraag -titel "Stoppen?" -vraag "`r`nWilt u het programma stoppen?" -knopok "Stoppen" -knopterug "Terug"

    if ($result -eq 'OK') {
        # programma sluiten
        $form.DialogResult = [System.Windows.Forms.DialogResult]::cancel
        return
        } 
    })
$Button7.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Het programma afsluiten." )
})
$Form.controls.add($Button7)

$Button6                         = New-Object system.Windows.Forms.Button
$Button6.width                   = 60
$Button6.height                  = 60
$Button6.location                = New-Object System.Drawing.Point(52,250)
$Button6.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoon-info.png")
$Button6.Add_Click({ informatieprogramma })
$Button6.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Algemene informatie over het programma lezen." )
})
$Form.controls.add($Button6)

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.width                   = 60
$Button5.height                  = 60
$Button5.location                = New-Object System.Drawing.Point(52,150)
$Button5.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$Button5.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoon-log.png")
$Button5.Add_Click({ vensterlogbestand })
$Button5.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "De logbestanden van de uitgevoerde taken bekijken." )
})
$Form.controls.add($Button5)

$Button11                         = New-Object system.Windows.Forms.Button
$Button11.width                   = 60
$Button11.height                  = 60
$Button11.location                = New-Object System.Drawing.Point(52,50)
$Button11.Image=[System.Drawing.Image]::FromFile("$icoontjesmap\icoon-verkenner.png")
$Button11.Add_Click({ vensterverkenner })
$Button11.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Bestanden verkennen." )
})
$Form.controls.add($Button11)


<# handmatig updaten is niet meer mogelijk!
   Hiervoor was $Button8 gedefinieerd
#>


# declareren venster met uitleg over programma. verschijnt als de muis over de vraagteken gaat.

$Form_uitlegprog = declareren_standaardvenster "Uitleg over het programma" 690 480
$Form_uitlegprog.ControlBox = $False

# Bij Escape-toets het venster sluiten.
$Form_uitlegprog.KeyPreview = $true
$Form_uitlegprog.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $Form_uitlegprog.Close() }
})


$uitlegprogtekst                     = New-Object system.Windows.Forms.Label
$uitlegprogtekst.text                = "Met dit programma kunnen bestanden op meerdere pc's in een netwerk worden beheerd.
De twee belangrijkste taken van dit programma worden met een groene icoontje weergegeven.

Met het groene icoontje 
    - 'Bestanden klaarzetten' worden bestanden en volledige mappen overgezet naar 
       de geselecteerde homemappen van de kandidaten,
    - 'Back-up maken' wordt een back-up gemaakt van de geselecteerde homemappen 
       van de kandidaten.

Met de vier icoontjes aan de linkerkant kunnen:
    - de bestanden in de homemappen van de kandidaten bekijken,
    - de logbestanden van de uitgevoerde taken bekeken worden,
    - algemene informatie over het programma worden gelezen,
    - het programma worden afgesloten.

Met de vier icoontjes aan de rechterkant kunnen:
    - de bestanden van een homemap verplaatst of gekopieërd worden naar een andere homemap,
    - de bestanden van de geselecteerde homemappen worden gewist,
    - de persoonlijke keuzes voor het programma worden gewijzigd,    
    - oude back-ups worden verwijderd.

Informatie over een taak kan verkregen worden door met de muiscursor over een object te gaan.
"
$uitlegprogtekst.AutoSize            = $false
$uitlegprogtekst.width               = 680
$uitlegprogtekst.height              = 390
$uitlegprogtekst.location            = New-Object System.Drawing.Point(10,10)
$uitlegprogtekst.Font                = 'Microsoft Sans Serif,11'
$uitlegprogtekst.ForeColor = [System.Drawing.Color]::blue
$Form_uitlegprog.Controls.Add($uitlegprogtekst)

$knopsluiten = New-object System.Windows.Forms.Button 
$knopsluiten.text= 'Sluiten'
$knopsluiten.location = "250,400" 
$knopsluiten.size = "150,30"  
$knopsluiten.BackColor = 'red'
$knopsluiten.ForeColor = 'white'
$knopsluiten.DialogResult = [System.Windows.Forms.DialogResult]::ok
$Form_uitlegprog.Controls.Add($knopsluiten)

# einde declareren venster met uitleg over programma

# declareren vraag-icoontje om extra informatiete geven over het programma
$picturequestion = new-object Windows.Forms.PictureBox
$picturequestion.Location = New-Object System.Drawing.Size(610,430) 
$picturequestion.Size = New-Object System.Drawing.Size(30,60)
$picturequestion.Image = [System.Drawing.Image]::FromFile("$icoontjesmap\icoon-hulp.png")
$picturequestion.add_click( { $Form_uitlegprog.showdialog() } )
$picturequestion.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Klik hier voor uitleg over het programma." )
})
$Form.controls.add($picturequestion)

#GIF afbeelding toevoegen
$gifBox = New-Object Windows.Forms.picturebox

# standaard waarde geven voor het geval geen match gevonden is hieronder
$keuzeafbeelding = 'samenwerken.gif'
$keuze_xas = 130
$keuze_yas = 230

# een match zoeken met de gekozen afbeelding in je persoonlijke instellingen
$global:afbeeldingen.ForEach( {
    if ($_.naam -eq $global:init["algemeen"]["afbeelding"]) {
        $keuzeafbeelding = $_.bestand
        }
})
$gifBox.Image    = [System.Drawing.Image]::FromFile("$gifjesmap\$keuzeafbeelding")
$gifBox.location = New-Object System.Drawing.Point($keuze_xas,$keuze_yas)
$gifBox.AutoSize = $false
$gifbox.Size     = "440,250" 
$gifbox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
# $gifbox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Centerimage

$gifBox.add_click( { Venstermetgrap } )
$gifBox.add_MouseHover({
    $global:tooltip1.SetToolTip($this, "Klik hier voor een leuke grap." )
})

$Form.Controls.Add($gifbox)

# Einde hoofdvenster declareren

# Als het programma in testfase is, dan venster tonen met melding
$form.add_Shown({ 
    $mode=$global:programma.mode
    if ($global:programma.mode -ne 'release') {
        $null = venstermetvraag -titel "Programma is nog in testfase" -vraag "`r`nLET OP : Programma is nog in testfase $mode.`r`nGebruik het programma nu alleen voor testdoeleinden."
    }
 } )

# Form wordt getoond. Dit is het hoofdvenster. Het programma is opgestart.
$null = $Form.ShowDialog()

# hoofdmenu sluiten
$Form.dispose()


# einde script.