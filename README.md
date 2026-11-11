# Readme

## 1. Introductie

Dit project is gestart om te voorzien in de behoefte aan een hulpmiddel dat eenvoudig bestanden kan kopiëren naar meerdere computers.  
Het kreeg de naam **"Beherenbestanden"** omdat het programma niet alleen bestanden kopieert, maar ook kan verwijderen en back-uppen.
Later zijn nog extra functionaliteiten toegevoegd aan het programma.
Het project begon in 2013, oorspronkelijk geschreven in **DOS**. 
Sindsdien heeft het programma veel wijzigingen ondergaan.  
De belangrijkste wijziging was de herschrijving in **PowerShell**.

### Hoofdfunctionaliteiten
- Bestanden kopiëren naar mappen  
- Een back-up maken van mappen

### Extra functionaliteiten
- Bestanden in mappen bekijken, openen of verwijderen
- Bestanden verplaatsen naar een andere map
- Standaard instellingen van het programma aanpassen  
- Oude back-ups verwijderen
- Gebeurtenissen (logs) bekijken
- Programmainfo, readme, changelog en licentie, tonen 
- Een grapje ophalen van een website

### Functionaliteiten die automatisch worden uitgevoerd
- Gebeurtenissen (logs) bijhouden en oude gebeurtenissen verwijderen.
- Het programma updaten. Dit is ook op commando uit te voeren.

---

## 2. Installatie

Dit programma werkt op **Windows 10 of nieuwer** en vereist **PowerShell 7 of hoger**.
Het script installeert automatisch de laatste versie van **PowerShell 7** en de benodigde modulen.

### Methode 1: Installer (aanbevolen)
- Download: [setup-beherenbestanden.zip](https://github.com/examencentrumtcr/beherenbestanden/tree/main/setup)
- Uitpakken
- Dubbelklik op setup-beheren.exe
> ⚠️ Microsoft Defender SmartScreen kan het bestand blokkeren, maar deze waarschuwing kan veilig genegeerd worden.

Tijdens de installatie kun je kiezen om:
- de installatiemap te wijzigen,
- een bureaubladsnelkoppeling te maken,
- een schone installatie uit te voeren.

De standaardopties zijn aanbevolen. Klik op **"Installeren"** om verder te gaan.

### Methode 2: ZIP-bestand
> ⚠️ Bij deze methode moet je handmatig een snelkoppeling maken en rekening houden met beperkingen van PowerShell-scripts.  
- Download: [beherenbestanden_versie.zip](https://github.com/examencentrumtcr/beherenbestanden/tree/main/release)
- Uitpakken: Pak het bestand uit in een map waar de gebruiker schrijfrechten heeft.  
- Voer het script `Wijzig_Executionpolicy_bypass.bat` uit om het uitvoeren van PowerShell-scripts mogelijk te maken.
- Voer eventueel het script `snelkoppeling_maken.ps1` uit → klik met de rechtermuisknop op het programma en kies  
  `Run with PowerShell` of `Openen met > PowerShell 7`.  

---

## 3. Configuratie

Configuratie is alleen nodig als:
- Je geen snelkoppeling hebt gemaakt, of  
- PowerShell beperkt is in het uitvoeren van scripts.  

### Snelkoppeling maken
Voer `snelkoppeling_maken.ps1` uit → klik met de rechtermuisknop op het programma en kies  
  `Run with PowerShell` of `Openen met > PowerShell 7`.  
Dit maakt een snelkoppeling en zet PowerShell in een onbeperkte modus.  

### PowerShell-uitvoering toestaan
Voer `Wijzig_Executionpolicy_bypass.bat` uit.  
Dit zet PowerShell in *Bypass*-modus.  
> ⚠️ Zonder dit kan PowerShell een **Execution Policy-waarschuwing** tonen.  

---

## 4. Gebruik

Er is een account met schrijfrechten nog voor de SharePointsite van het RPC.
Het programma gebruikt de SharePointsite om bestanden te downloaden en te backuppen.

Het programma starten :
- Als je een snelkoppeling hebt gemaakt op het bureaublad → **dubbelklik erop**.  
  Anders → klik in de installatiemap met de rechtermuisknop op `beherenbestanden.ps1` en kies  
  `Run with PowerShell` of `Openen met > PowerShell 7`.  
- Er wordt gevraagd in te loggen bij de Sharepoint site van het RPC. 
  Bij gebruik van een vaste pc in het netwerk van de organisatie gaat het aanmelden automatisch.
> ⚠️ Zonder een account met de juiste rechten zijn een aantal taken niet uit te voeren.

---

## 5. Bestandenoverzicht

- `Beherenbestanden.ps1` – Hoofdprogramma 
- `Snelkoppeling_maken.ps1` – Hulpmiddel om een snelkoppeling te maken  
- `Wijzig_Executionpolicy_bypass.bat` – Hulpmiddel om PowerShell Execution Policy te wijzigen  
- `Changelog.md` – Overzicht van wijzigingen per versie  
- `Readme.md` – Dit bestand  
- `License` - Licentieovereenkomst
- `Png` – Map met iconen en afbeeldingen  

---

## 6. Licentie

Dit project valt onder de **GNU General Public License v3.0**.  
Beherenbestanden is vrije software: je mag het verspreiden en/of aanpassen onder de voorwaarden van de GPL, zoals gepubliceerd door de Free Software Foundation.  
Het wordt verspreid **zonder enige garantie**; zelfs zonder de impliciete garantie van **Verkoopbaarheid** of **Geschiktheid voor een bepaald doel**.  
Zie <https://www.gnu.org/licenses/> voor meer informatie.  

---

## 7. Bekende bugs

- Het programma kan crashen tijdens het uitvoeren van taken.  
  De achtergrondtaak gaat echter door, en na enige tijd verschijnt een melding dat de taak is voltooid.  

Als dit gebeurt:  
➡️ Wacht tot de taak is afgerond en controleer of het programma nog werkt.  

✅ Dit probleem is opgelost vanaf **versie 4.2.2** (Zie changelog).

---

## 8. Auteurs en dankwoord

**Hoofdauteur:**  
- [Benvindo Neves](https://neveshuis.nl/over-mij)  

**Lay-out en afbeeldingen:**  
- Marco Laluan  

**Met dank aan:**  
- Rene de Bruin  
- Edgar Seedorf  
- Robby Tosasi  
- Rob Schortemeijer  

---

## 9. Changelog

Zie: <https://github.com/examencentrumtcr/beherenbestanden/tree/main/changelog.md>
