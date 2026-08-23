# Geld-planner

Generator voor **Geldplanner.xlsx** — een Excel-cashflow- en budgetruimte-dashboard voor
Nederlandse particulieren en stellen, in de stijl van een professioneel financieel dashboard
(KPI-cards, donut- en staafdiagrammen, groen/rood-signalering).

## Gebruik

```bash
pip install openpyxl
python generate_geldplanner.py            # schrijft Geldplanner.xlsx
python generate_geldplanner.py Mijn.xlsx  # of een eigen bestandsnaam
```

## Tabbladen

| Tabblad | Inhoud |
| --- | --- |
| **Overzicht Cashflow** | Invoerpaneel voor maandelijkse inkomsten, vaste lasten en variabele kosten, met 4 KPI-cards bovenaan: totale inkomsten, totale uitgaven, overschot/tekort (groen/rood) en spaarquote. |
| **Statistieken Cashflow** | Geen eigen invoer — leest uit Overzicht Cashflow. Donut van uitgaven per categorie (7 categorieën), staafdiagram inkomsten (groen) vs. uitgaven (rood) per maand, en een KPI-card die het overschot vergelijkt met vorige maand (bedrag + %, groen/rood). |
| **Overzicht Budgetruimte** | Invoerpaneel voor gewenst spaar- en investeerbedrag, met 4 KPI-cards: beschikbare budgetruimte (= overschot uit Overzicht Cashflow), gewenst spaarbedrag, gewenst investeerbedrag en vrij besteedbaar. Toont een rode waarschuwing als het gewenste bedrag hoger is dan het beschikbare overschot. |
| **Statistieken Budgetruimte** | Geen eigen invoer — leest uit Overzicht Budgetruimte. Donut van de verdeling sparen/investeren/vrij besteedbaar, gestapeld staafdiagram van de budgetruimte per maand, en een KPI-card met het totaal gespaard + geïnvesteerd over 12 maanden. |

### Koppeling tussen de tabbladen

Het maandelijkse overschot uit **Overzicht Cashflow** loopt automatisch door als beschikbare
budgetruimte op **Overzicht Budgetruimte**. Beide "Statistieken"-tabbladen hebben geen eigen
invoer: alle cijfers, grafieken en KPI's daar zijn live formules die naar het bijbehorende
"Overzicht"-tabblad verwijzen. De 12-maands staafdiagrammen zijn een prognose: ze trekken het
huidige maandbudget vlak door over het jaar (er wordt geen aparte historie per maand bijgehouden).

## Opmaak

* **Gele cellen** zijn invoer en als enige ontgrendeld; alles daarbuiten is vergrendeld achter
  werkbladbeveiliging zonder wachtwoord (*Controleren → Blad-beveiliging opheffen* om formules
  aan te passen).
* **KPI-cards**: gekleurde kopbalk (donkerblauw, wit) met het bedrag in groot vet lettertype
  (22pt) eronder in een lichte kaart. De overschot/tekort- en vrij-besteedbaar-cards kleuren
  automatisch groen of rood.
* Elke sectie staat in een lichte kaart met subtiele rand, gescheiden door witruimte — geen
  kale, ongeformatteerde rijen.
* Grafieken gebruiken één consistente kleurstijl door de hele werkmap: donkerblauw/goud-palet
  voor categorieën, groen voor inkomsten/positief, rood voor uitgaven/aandachtspunten.
* Bedragen gebruiken de opmaakcode `€ #,##0.00`, die op een Nederlandse Excel-installatie als
  `€ 1.234,56` wordt weergegeven.

## Formules aanpassen

Alle bedragen worden als Excel-formule weggeschreven, niet als voorberekende waarde. In het
script staan de rijnummers in variabelen (`overschot_rij`, `beschikbaar_rij`, …), zodat je regels
kunt toevoegen of verplaatsen zonder celverwijzingen met de hand bij te werken. `bouw_cashflow` en
`bouw_budgetruimte` geven een `dict` met rijnummers terug die de bijbehorende Statistieken-tabbladen
en de onderlinge koppeling nodig hebben.

## Disclaimer

Deze tool is uitsluitend bedoeld voor informatieve doeleinden en biedt geen financieel advies.
