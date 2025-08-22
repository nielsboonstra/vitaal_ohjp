# VITAAL | Ultimo naar OHJP Conversie Tool

Deze webapplicatie helpt je om een OMS-export uit Ultimo om te zetten naar een OHJP-worksheet voor het opzetten van een onderhoudsjaarplan.

## Functionaliteit
- Upload een Excel-bestand (OMS-export) via de zijbalk.
- Bekijk de ingelezen data in een overzichtelijke tabel.
- Kies het startjaar, startweek en een naam voor het exportbestand.
- Start de conversie: de tool verwerkt de data, normaliseert frequenties, en maakt per complex een heatmap-worksheet aan.
- Download het resultaat als een Excel-bestand met meerdere sheets (één per complex).

## Installatie
1. Installeer de benodigde Python packages:
	```powershell
	pip install -r requirements.txt
	```
2. Start de webapp lokaal:
	```powershell
	streamlit run webapp.py
	```

## Benodigdheden
- Python 3.8+
- OMS-exportbestand in Excel-formaat (.xlsx)

## Gebruik
1. Open de webapp in je browser.
2. Upload je OMS-exportbestand.
3. Bekijk en controleer de data.
4. Vul het gewenste startjaar, startweek en exportnaam in.
5. Klik op "Start conversie".
6. Download het gegenereerde OHJP Excel-bestand.

## Opmerkingen
- Alleen frequenties "WK" (week) en "JR" (jaar) worden genormaliseerd naar "MD" (maand).
- Verkeerscentrales worden automatisch uitgesloten van de planning
- Het resultaat bevat per complex een eigen worksheet.