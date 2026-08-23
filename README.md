# Geld-planner

Generator voor **Geldplanner.xlsx** — een Excel-sjabloon voor persoonlijke financiën,
gericht op Nederlandse particulieren en stellen.

## Gebruik

```bash
pip install openpyxl
python generate_geldplanner.py            # schrijft Geldplanner.xlsx
python generate_geldplanner.py Mijn.xlsx  # of een eigen bestandsnaam
```

## Tabbladen

| Tabblad | Inhoud |
| --- | --- |
| **Titelblad** | Voorblad met een plek voor een eigen logo, de productnaam en een korte introductie. |
| **Dashboard** | Vier KPI-kaarten die automatisch bijwerken vanuit de andere tabbladen: maandelijks overschot/tekort, spaarquote, pensioengat-status en verwacht vermogen. Groen/oranje/rood signaleert of iets gezond is of aandacht nodig heeft. |
| **Instructies** | Uitleg per tabblad, kleurcodering en de disclaimer. |
| **Cashflow Overzicht** | Maandelijkse inkomsten, vaste lasten en variabele kosten; totalen, overschot/tekort (uitgelicht als hoofduitkomst), spaarquote en buffer, plus een staafdiagram inkomsten vs. uitgaven. |
| **Pensioengat Calculator** | Leeftijd, pensioenleeftijd, gewenst inkomen, opgebouwd vermogen en AOW in; benodigd kapitaal, het pensioengat en de benodigde maandelijkse inleg uit. |
| **Vermogensgroei** | Startkapitaal, rendement en looptijd in; projectie per jaar tot 40 jaar met lijndiagram. |

### Koppeling tussen de tabbladen

Het maandelijkse overschot uit **Cashflow Overzicht** loopt automatisch door als voorstel voor de
maandelijkse inleg op **Pensioengat Calculator** en **Vermogensgroei**. Elk van die tabbladen heeft
een "Koppeling met Cashflow Overzicht"-sectie met:

1. het voorstel (een berekende cel die naar Cashflow Overzicht verwijst),
2. een Ja/Nee-keuze "Gebruik dit voorstel?" (standaard "Ja"),
3. een eigen invoercel die alleen wordt gebruikt bij "Nee".

Zo werkt de doorstroom automatisch, maar kan de gebruiker hem per tabblad overschrijven.

## Opmaak

* **Gele cellen** zijn invoer en als enige ontgrendeld.
* **Grijze cellen** bevatten formules; **blauwe cellen** zijn een sectieresultaat; **goudkleurige
  cellen** (groot, 20pt) zijn de hoofduitkomst van het tabblad (bijv. het pensioengat). Alle drie
  zijn vergrendeld; elk blad is beveiligd zonder wachtwoord, dus *Controleren → Blad-beveiliging
  opheffen* volstaat om formules aan te passen.
* Elke sectie staat in een lichte kaart met subtiele rand, met een gekleurde kopbalk erboven —
  donkerblauw is de hoofdkleur, goud de accentkleur (ook gebruikt voor de "koppeling"-secties).
* Bedragen gebruiken de opmaakcode `€ #,##0.00`, die op een Nederlandse Excel-installatie
  als `€ 1.234,56` wordt weergegeven.
* Alle labels en toelichtingen zijn Nederlands.

## Formules aanpassen

Alle bedragen worden als Excel-formule weggeschreven, niet als voorberekende waarde. In het
script staan de rijnummers in variabelen (`overschot_rij`, `gat_rij`, `eindwaarde_rij`, …), zodat
je regels kunt toevoegen of verplaatsen zonder celverwijzingen met de hand bij te werken. Elk
tabblad-functie (`bouw_cashflow`, `bouw_pensioen`, `bouw_vermogen`) geeft een `dict` met de
rijnummers terug die het Dashboard en de koppeling-secties nodig hebben.

## Disclaimer

Deze tool is uitsluitend bedoeld voor informatieve doeleinden en biedt geen financieel advies.
Rendementen uit het verleden bieden geen garantie voor de toekomst; beleggen brengt risico's met
zich mee.
