# Geld-planner

Generator voor **Geldplanner.xlsx** — een Excel-sjabloon voor persoonlijke financiën,
gericht op Nederlandse particulieren, stellen en zzp'ers.

## Gebruik

```bash
pip install openpyxl
python generate_geldplanner.py            # schrijft Geldplanner.xlsx
python generate_geldplanner.py Mijn.xlsx  # of een eigen bestandsnaam
```

## Tabbladen

| Tabblad | Inhoud |
| --- | --- |
| **Instructies** | Uitleg per tabblad, kleurcodering en de disclaimer. |
| **Cashflow Overzicht** | Maandelijkse inkomsten, vaste lasten en variabele kosten; totalen, overschot/tekort, spaarquote en buffer, plus een staafdiagram inkomsten vs. uitgaven. |
| **ZZP Belastingtool** | Omzet, zakelijke kosten en urencriterium in; belastbare winst, inkomstenbelasting via de schijven, zelfstandigen- en startersaftrek, MKB-winstvrijstelling, heffingskortingen, Zvw-bijdrage en netto maandinkomen uit. |
| **Pensioengat Calculator** | Leeftijd, pensioenleeftijd, gewenst inkomen, opgebouwd vermogen en AOW in; benodigd kapitaal, het pensioengat en de benodigde maandelijkse inleg uit. |
| **Vermogensgroei** | Startkapitaal, maandinleg, rendement en looptijd in; projectie per jaar tot 40 jaar met lijndiagram. |

## Opmaak

* **Gele cellen** zijn invoer en als enige ontgrendeld.
* **Grijze en blauwe cellen** bevatten formules en zijn vergrendeld; elk blad is beveiligd
  zonder wachtwoord, dus *Controleren → Blad-beveiliging opheffen* volstaat om formules aan te passen.
* Bedragen gebruiken de opmaakcode `€ #,##0.00`, die op een Nederlandse Excel-installatie
  als `€ 1.234,56` wordt weergegeven.
* Alle labels en toelichtingen zijn Nederlands.

## Formules aanpassen

Alle bedragen worden als Excel-formule weggeschreven, niet als voorberekende waarde. In het
script staan de rijnummers in variabelen (`omzet_rij`, `belastbaar_rij`, …), zodat je regels kunt
toevoegen of verplaatsen zonder celverwijzingen met de hand bij te werken.

De belastingtarieven, schijfgrenzen, aftrekposten, heffingskortingen en de arbeidskortingstabel
staan als **gele invoercellen** op het tabblad *ZZP Belastingtool*. Ze kunnen dus in Excel zelf
worden bijgewerkt zonder het script te draaien.

## Disclaimer

Deze tool is uitsluitend bedoeld voor informatieve doeleinden en biedt geen financieel advies.
De opgenomen cijfers voor 2026 zijn indicatief; controleer de actuele bedragen op
[belastingdienst.nl](https://www.belastingdienst.nl). Rendementen uit het verleden bieden geen
garantie voor de toekomst.
