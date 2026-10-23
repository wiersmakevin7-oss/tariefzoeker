# Tariefzoeker

Streamlit app om exporttarieven per vervoerder te vergelijken op basis van leverland, postcode, afmetingen, gewicht, diesel, road tax en marge.

## Starten

```powershell
streamlit run app.py
```

## Tariefbestanden

De standaardstaffels staan in `data/` en worden automatisch geladen.

Je kunt optioneel extra `.xlsx` tariefbestanden uploaden via de sidebar. De app leest per workbook de tariefstaffels, zoekt de juiste postcodezone per vervoerder, en kiest de eerste passende staffel op basis van gewicht, laadmeters en pallet-aantallen.

## Excel-zendingen

De tab Excel-zendingen verwerkt een .xlsx met kolomkoppen op rij 1.
Kies het werkblad, leverland (standaard FR) en de kolommen voor postcode,
totaalgewicht in kg, totale laadmeters en aantal pallets. Iedere rij is een
zelfstandige zending; regels worden niet op datum of postcode samengevoegd.
De upload blijft in het geheugen en wordt niet op GitHub gepubliceerd.

Het resultaat behoudt bestaande cellen en werkbladen en voegt per vervoerder
alleen totale inkoop inclusief fuel en roadtax toe. Optioneel wordt een
verkoopkolom inclusief de ingestelde marge toegevoegd; marge is opslag over
inkoop inclusief toeslagen. Vervoerders zonder staffel voor het gekozen land
worden weggelaten, ook op het palletstaffelblad.
Geen passende staffel geeft Geen tarief in de prijscel, niet een nulprijs.
Ongeldige invoer wordt per regel vermeld zonder de overige regels te stoppen.

Bij onbekende palletsoort worden pallet-aantalgebonden staffels uitgesloten;
selecteer Europallet of Blokpallet alleen als dit voor alle regels klopt.
De ingevoerde laadmeters en het gewicht worden altijd meegenomen.
De berekening gebruikt actuele app-instellingen en geladen staffels, geen
historische selectie op losdatum. Controleer geldigheid voordat je de
uitkomst als historische kostprijs gebruikt. Limiet: 5.000 regels / 30 MB.

Tests: `python -m unittest discover -s tests`.

Optie Staffel voor 1-4 europallets toevoegen maakt een apart werkblad met
vier scenario's per unieke postcode. Voer verplicht een gewicht per europallet
in; het tariefgewicht is aantal maal dit gewicht, de laadmeters zijn aantal
maal 0,4. Dit zijn scenario's, geen afleiding van historische palletgewichten.
Ook dit blad bevat alle vervoerders en optioneel de ingestelde marge.

## Prijsformule

```text
inkoop = basistarief + dieselbedrag + road tax bedrag
dieselbedrag = basistarief * diesel%
road tax bedrag = basistarief * road tax%
verkoop = inkoop * (1 + marge%)
```

## Streamlit Community Cloud

Gebruik bij deployment:

```text
Repository: wiersmakevin7-oss/tariefzoeker
Branch: main
Main file path: app.py
```

### Permanente toeslag-instellingen

De app kan diesel, road tax en marge permanent opslaan in deze GitHub-repository als `cloud_settings.json`.

Maak hiervoor een GitHub fine-grained personal access token aan met toegang tot deze repository en permissie:

```text
Contents: Read and write
```

Zet daarna in Streamlit Community Cloud bij **App settings > Secrets**:

```toml
GITHUB_TOKEN = "jouw_github_token"
APP_PASSWORD = "kies_een_veilig_wachtwoord"
```

Zonder deze secret werkt de app nog steeds, maar opgeslagen toeslagen zijn dan alleen lokaal/per sessie betrouwbaar.

`APP_PASSWORD` is optioneel. Als je deze secret invult, toont de app eerst een login-scherm. Laat je hem leeg, dan is de app zonder login bereikbaar.
