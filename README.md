# Beherenbestanden.ps1

## 1. Introductie

Dit project is gestart om te voorzien in de behoefte aan een hulpmiddel dat eenvoudig bestanden kan kopiëren naar meerdere computers.  
Het kreeg de naam **"Beherenbestanden"** omdat het script niet alleen bestanden kopieert, maar ook kan verwijderen en back-uppen.

Later is de mogelijkheid toegevoegd om bestanden tussen mappen te verplaatsen.  
Het project begon in 2013, oorspronkelijk geschreven in **DOS**. Sindsdien heeft het script veel wijzigingen ondergaan.  
De belangrijkste update was de herschrijving in **PowerShell**.

### Hoofdfunctionaliteiten
- Bestanden kopiëren naar mappen  
- Een back-up maken van mappen

### Extra functionaliteiten
- Bestanden in mappen verwijderen  
- Bestanden tussen mappen verplaatsen of kopiëren  
- Inhoud van mappen verkennen  
- Standaard persoonlijke instellingen aanpassen  
- Oude back-ups opruimen  
- Gebeurtenissen (logs) bekijken
- Gebeurtenissen (logs) verwijderen bij de start van het programma    
- Programmainfo, README en changelog tonen  
- Een grapje ophalen van een website (Er zijn 3 websites waaruit te kiezen is) 

### Functionaliteiten die automatisch worden uitgevoerd
- Script updaten bij de start van het programma 
- Gebeurtenissen (logs) in een bestand bijhouden

---

## 2. Installatie

Dit script werkt op **Windows 10 of nieuwer** en vereist **PowerShell 5.1 of hoger**.

### Methode 1: Installer (aanbevolen)
- Download: [setup-beherenbestanden.zip](https://github.com/examencentrumtcr/beherenbestanden/tree/main/setup)
- Uitpakken
- Dubbelklik op setup-beheren.exe

> ⚠️ Microsoft Defender SmartScreen kan het bestand blokkeren, maar deze waarschuwing kan veilig genegeerd worden.

Tijdens de installatie kun je kiezen voor:
- Wijzig de installatiemap  
- Maak een bureaubladsnelkoppeling  
- Voer een schone installatie uit  

De standaardopties zijn aanbevolen. Klik op **"Installeren"** om verder te gaan.

### Methode 2: ZIP-bestand
- Download: [beherenbestanden_versie.zip](https://github.com/examencentrumtcr/beherenbestanden/tree/main/release)
- Uitpakken

Pak het bestand uit in een map waar de gebruiker schrijfrechten heeft.  
> ⚠️ Bij deze methode moet je handmatig een snelkoppeling maken en rekening houden met beperkingen van PowerShell-scripts.  

---

## 3. Configuratie

Configuratie is alleen nodig als:
- Je geen snelkoppeling hebt gemaakt, of  
- PowerShell beperkt is in het uitvoeren van scripts.  

### Snelkoppeling maken
Voer `snelkoppeling_maken.ps1` uit → klik met de rechtermuisknop op het script en kies  
  `Openen met > Windows PowerShell` of `Uitvoeren met PowerShell`.  
Dit maakt een snelkoppeling en zet PowerShell in een onbeperkte modus.  

### PowerShell-uitvoering toestaan
Voer `Wijzig_Executionpolicy_bypass.bat` uit.  
Dit zet PowerShell in *Bypass*-modus.  

Zonder dit kan PowerShell een **Execution Policy-waarschuwing** tonen.  

---

## 4. Gebruik

- Als je een snelkoppeling hebt gemaakt → **dubbelklik erop**.  
- Anders → klik met de rechtermuisknop op het `beherenbestanden.ps1` en kies  
  `Openen met > Windows PowerShell` of `Uitvoeren met PowerShell`.  

---

## 5. Bestandenoverzicht

- `Beherenbestanden.ps1` – Hoofdscript  
- `Beheren.ico` – Programma-icoon  
- `Snelkoppeling_maken.exe` – Hulpmiddel om een snelkoppeling te maken  
- `Wijzig_Executionpolicy_bypass.bat` – Hulpmiddel om PowerShell Execution Policy te wijzigen  
- `Changelog.txt` – Overzicht van wijzigingen per versie  
- `Readme.txt` – Dit bestand  
- `PNG/` – Map met iconen en afbeeldingen  

---

## 6. Licentie

Dit project valt onder de **GNU General Public License v3.0**.  

Beherenbestanden is vrije software: je mag het verspreiden en/of aanpassen onder de voorwaarden van de GPL, zoals gepubliceerd door de Free Software Foundation.  

Het wordt verspreid **zonder enige garantie**; zelfs zonder de impliciete garantie van **Verkoopbaarheid** of **Geschiktheid voor een bepaald doel**.  

Zie <https://www.gnu.org/licenses/> voor meer informatie.  

---

## 7. Bekende bugs

- Het script kan crashen tijdens het uitvoeren van taken.  
  De achtergrondtaak gaat echter door, en na enige tijd verschijnt een melding dat de taak is voltooid.  

Als dit gebeurt:  
➡️ Wacht tot de taak is afgerond en controleer of het script nog werkt.  

✅ Dit probleem is opgelost vanaf **versie 4.2.2**.  
Zie changelog: <https://github.com/examencentrumtcr/beherenbestanden/tree/main>  

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

Zie: <https://github.com/examencentrumtcr/beherenbestanden/tree/main>
