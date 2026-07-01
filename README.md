# Tariefzoeker

Streamlit app om exporttarieven per vervoerder te vergelijken op basis van leverland, postcode, afmetingen, gewicht, diesel, road tax en marge.

## Starten

```powershell
streamlit run app.py
```

## Tariefbestanden

De standaardstaffels staan in `data/` en worden automatisch geladen.

Je kunt optioneel extra `.xlsx` tariefbestanden uploaden via de sidebar. De app leest per workbook de tariefstaffels, zoekt de juiste postcodezone per vervoerder, en kiest de eerste passende staffel op basis van gewicht, laadmeters en pallet-aantallen.

## Berekening

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
```

Zonder deze secret werkt de app nog steeds, maar opgeslagen toeslagen zijn dan alleen lokaal/per sessie betrouwbaar.
