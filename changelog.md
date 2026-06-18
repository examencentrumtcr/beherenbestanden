# Changelog

Alle belangrijke wijzigingen in **Beherenbestanden.ps1** worden hieronder bijgehouden.  

---

## [4.8.0] - 2026-06-08
### Nieuwe functies
- Met de functie **Verkenner** kan je nu ook bestanden, in de studentmappen, openen.
- Nieuwe indeling en nieuwe knop in het hoofdvenster;
  Hoofdvenster heeft 2 knoppen voor de standaardtaken, een menuknop en een afsluitknop.
- Na elke update verschijnt er eenmalig een venster met informatie over de nieuwe versie.
- Programma werkt nu ook met PowerShell vanaf versie 7.
  Je kan bij **Instellingen** dit aanzetten. PowerShell 7 wordt indien nodig geïnstalleerd.
- Updateproces aangepast. Bij **Instellingen** kan je updaten uitzetten of eenmalig controleren.

### Bugfixes
- Bij **Verkenner** werkt het 'op volgorde zetten' door te klikken op de kolommen 'Gewijzigd' 
  en 'Grootte' niet zoals het de bedoeling is. Dit is gewijzigd.
- Bepaalde tekens in de tekst bij de moppenbot werden niet goed weergeven.
  Opgelost door UTF-8 te gebruiken.

### Overige wijzigingen
- README is aangepast. Een aantal verwijzingen naar bestanden stonden verkeerd, nieuwe functies
  zijn toegevoegd en wijzigingen zijn doorgevoerd.
- De naam van het programma-icoontje is gewijzigd.
- Het venster **Informatie** heeft nu 3 tabbladen: Changelog, Readme en License;
  De tabblad License met licentie-informatie is nieuw.
- Logbestanden ouder dan 2025 worden niet meer getoond.
- De manier waarop de escapetoets wordt gebruikt bij uitvoeren van taken gewijzigd.
- Code gestructureerd en opgeschoond.

---

## [4.7.1] - 2025-10-08
### Bugfixes
- Tekstuele fouten bij functies **Moppenbot** en **Instellingen** opgelost.
- Een niet gebruikte waarde van de gebruikersinstellingen verwijderd.

### Overige wijzigingen
- In de programma code wijzigingen aangebracht in hoe met bepaalde fases wordt omgegaan.
-  Van 2 functies, console openen en -sluiten, 1 gemaakt.
- Bij de **Moppenbot** heb je de keuze uit moppen van 3 websites.
- Het is mogelijk om het aantal dagen dat de logbestanden bewaard blijven op 0 te zetten.
- Het programma wordt nu geüpdatet vanuit een repository in Github.
- Code optimalisatie. Updateproces gewijzigd; 
  Na een update start het programma in een andere sessie en controleert niet opnieuw.
  Updateproces kan nu omgaan met getallen > 9 in de versienummer.
  Het updateproces wordt nu ook gelogd.
- De manier waarop meldingen worden gelogd aangepast. Nu 3 typen meldingen.

---

## [4.7.0] - 2025-06-17
### Nieuwe functies
- Console kan opengehouden worden bij het starten van het hoofdvenster.  
  Handig voor foutopsporing en testen. Optie beschikbaar in **Instellingen**.  

### Bugfixes
- Controle toegevoegd of individuele homemappen bestaan (niet alleen de hoofdmap).  
- Console werd niet afgesloten bij instellingen → opgelost door oude methode terug te plaatsen.  

### Overige wijzigingen
- Persoonlijke instellingenbestand krijgt nieuwe naam; code migreert automatisch oude naam.  
- Functie **Bestanden klaarzetten**: controle of studentmappen leeg zijn (instelbaar).  
- Functie **Verkenner**: sorteren op grootte of type toegevoegd.  
- Verschillende iconen voor bestandstypen toegevoegd in **Verkenner** en **Bestanden klaarzetten**.  
- Systeemfouten worden nu in logbestand opgeslagen.  
- Vensters met informatie/vragen verbeterd (scrollbaar en schaalbaar).  
- Optie *Bestanden na de back-up wissen* verplaatst.  
- Teksten gestandaardiseerd (kandidaat, homemappen, RPC-nummers).  
- Grapjes komen nu van [apekool.nl](https://apekool.nl) i.p.v. moppenbot.nl.  
- Tool toegevoegd voor wijzigen PowerShell ExecutionPolicy.  

---

## [4.6.2] - 2025-02-18
### Bugfixes
- Netwerkmappen werden niet automatisch hersteld.  
- Updateproces crashte bij meerdere openstaande processen → nu back-up en herstel ingebouwd.  
- Focusprobleem na taakuitvoering opgelost.  
- Infovenster past nu bij lange installatienamen.  

### Overige wijzigingen
- Console kan opengehouden worden via variabele in gebruikersbestand.  
- Afbeeldingen in instellingen tonen nu extra tekst.  
- Navigatie verbeterd (Esc sluit vensters).  
- Ongebruikte bestanden en mappen worden bij start verwijderd.  
- Tool **snelkoppeling_maken** weer omgezet naar PowerShell-script (.ps1).  

---

## [4.6.1] - 2025-01-21
### Bugfixes
- Tabvolgorde van icoontjes hersteld.  
- Console-venster error opgelost door overbodige code te verwijderen.  

### Overige wijzigingen
- Hoofdscherm aangepast: alleen belangrijkste functies hebben grote iconen.  
- Consistente kleuren en teksten voor knoppen.  
- Tool **netwerkschijvencontroleren.exe** verwijderd bij start.  
- Venster **Verkenner** toont meer bestandsinfo (datum, grootte).  
- Mop-venster: optie voor schuine moppen toegevoegd.  
- PowerShell-versie zichtbaar in infovenster.  

---

## [4.6.0] - 2024-11-06
### Nieuwe functies
- Functie **Verkenner** toegevoegd om studentmappen te bekijken.  
- Afbeelding in hoofdscherm kan gewijzigd worden (8 keuzes).  

### Overige wijzigingen
- Mop-functie uitgebreid met schuine moppen (uit).  
- Welkomstnaam instelbaar via instellingen.  
- Beheer.ini-bestand wordt niet meer bewaard.  
- Infoteksten verbeterd.  
- Readme en Changelog lokaal i.p.v. van website.  
- Netwerkmappen automatisch hersteld bij problemen.  
- Console zichtbaar bij testmodi, verborgen bij productie.  
- Updateproces plaatst bestanden direct in juiste mappen.  

---

## [4.5.3] - 2024-09-21
### Bugfixes
- Logbestanden niet zichtbaar door naamgevingsfout → opgelost.  

### Overige wijzigingen
- Updates weer via website.  
- Alle SharePoint-functies verwijderd.  
- Welkomsttekst standaard "Welkom gebruiker".  
- Beheerinstellingen lokaal bewaard.  
- Readme en Changelog via website geladen.  
- Controle op modules verwijderd.  
- Installatiemap zichtbaar in infovenster.  

---

## [4.5.2] - 2024-09-11
### Bugfixes
- MFA-inloggen uitgeschakeld → code in commentaar.  

### Overige wijzigingen
- Tekstuele correcties.  
- Naamgeving aangepast (bv. “Inhoud examenmap” → “Geselecteerde mappen en bestanden”).  
- Afmeldoptie verwijderd.  
- Automatisch updaten uitgeschakeld (SharePoint).  
- Readme en Changelog blijven lokaal bewaard.  

---

## [4.5.1] - 2024-05-23
### Bugfixes
- Console sluit nu correct.  

### Overige wijzigingen
- Updates via SharePoint in plaats van website.  
- Extra bevestiging bij wissen van bestanden.  
- Kleine tekstuele correcties.  
- Taken verplaatsen/wissen van plek gewisseld.  
- Changelog/Readme via SharePoint opgehaald.  
- Logvolgorde aangepast.  
- Extra tools toegevoegd en bijgewerkt.  

---

## [4.5.0] - 2023-12-12
### Nieuwe functies
- Inloggen verplicht via Microsoft MFA.  
- Persoonlijke instellingen en logbestanden per gebruiker.  
- Afmeldoptie toegevoegd bij afsluiten.  
- Beheerinstellingen deels in SharePoint.  

### Bugfixes
- Tekstcorrecties bij icoontjes.  
- Back-up controle verbeterd.  
- Netwerkschijven automatisch hersteld.  

### Overige wijzigingen
- Mop/weetje-venster uitgebreid en verbeterd.  
- Instellingen beperkt (updates en back-up dagen vast).  
- Updater.ps1 verwijderd.  
- Logbestanden tonen nieuwste bovenaan.  
- Testfase-indicator toegevoegd.  
- Script voor initialisatiebestand toegevoegd.  
- Startconsole zichtbaar met statusmeldingen.  

---

## [4.4.1] - 2023-09-22
### Bugfixes
- Opschonen langzamer → opgelost met runspaces.  

### Overige wijzigingen
- Venster **Bestanden overzetten** verbeterd met iconen.  
- Venster **Opschonen** aangepast (alleen back-ups wissen).  
- Meldingen en logstructuur verbeterd.  
- Update-notificaties verwijderd.  
- Website aangepast naar [beherenbestanden.neveshuis.nl](https://beherenbestanden.neveshuis.nl).  

---

## [4.4.0] - 2023-06-01
### Nieuwe functies
- **Bestanden overzetten**: mappenstructuur doorbladeren met navigatie-iconen.  
- Klikken op afbeelding toont mop/weetje.  

### Bugfixes
- Downloadproblemen opgelost.  
- Teksten onder knoppen niet meer klikbaar.  
- Tabvolgorde hersteld.  

### Overige wijzigingen
- Handmatige update-optie verwijderd.  
- Taaknaam zichtbaar in titelbalk.  

---

## [4.3.2] - 2023-02-10
### Overige wijzigingen
- Back-up mapnaam bevat nu datum en tijd (HH-MM-SS).  

---

## [4.3.1] - 2023-01-10
### Bugfixes
- Lange bestandsnamen geen probleem meer.  

### Overige wijzigingen
- Bij kopiëren worden standaard nieuwe bestanden toegevoegd.  

---

## [4.3.0] - 2022-11-04
### Nieuwe functies
- Uitleg toegevoegd via vraagtekens en tooltips.  
- Instelbare bewaartermijn log- en back-upbestanden.  
- Logbestanden automatisch opruimen.  

### Bugfixes
- Kopieerproces logischer gemaakt.  
- Vensters correct verborgen bij taken.  

### Overige wijzigingen
- Layout aangepast (grotere iconen voor hoofdtaken, witte achtergrond).  
- RPC-nummers voor PAL gewijzigd naar 1–60.  

---

## [4.2.2] - 2022-00-29
### Bugfixes
- Scriptproblemen opgelost met runspaces.  
- Betere methode om variabelen te beheren via `initialisatie.ini`.  

### Overige wijzigingen
- Meldingsvensters met kleurgecodeerde knoppen.  
- Tekstcorrecties en verduidelijkingen.  

---

## [4.2.1] - 2022-08-25
### Overige wijzigingen
- Init-bestand hernoemd naar `gebruiker.ini`.  
- Infovenster aangepast.  
- Changelog bevat alleen functionele informatie.  

---

## [4.2.0] - 2022-08-15
### Nieuwe functies
- Infovenster met Readme en Changelog.  
- Opschoonfunctie voor back-ups en logs.  
- Filteroptie bij logbestanden.  
- Instellingenvenster toegevoegd.  
- Post-update taken via `updateinfo`.  

### Bugfixes
- Back-up weergave verbeterd.  
- Snellere laadtijd infovenster.  
- Foutmeldingen bij afbeeldingen opgelost.  

### Overige wijzigingen
- Layout en kleuren aangepast.  
- Dropdownmenu’s toegevoegd.  

---

## [4.1.0] - 2022-07-04
### Nieuwe functies
- Controle op inhoud doelmappen bij kopiëren, back-up en verplaatsen.  
- Optie om doelmap te legen bij verplaatsen.  

### Overige wijzigingen
- Logstructuur verbeterd met `[INFO]` en `[FOUT]`.  

---

## [4.0.0] - 2022-06-23
### Eerste officiële versie
- Herschreven in PowerShell (niet updatebaar vanaf 3.7.18).  
- **Hoofdtaken**: bestanden klaarzetten, back-up, wissen, verplaatsen/kopiëren.  
- **Overige taken**: log bekijken, programma updaten.  
