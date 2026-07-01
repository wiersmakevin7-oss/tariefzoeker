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
