#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Geldplanner NL — generator voor een Excel-sjabloon (.xlsx) voor persoonlijke
financien, gericht op Nederlandse particulieren, stellen en zzp'ers.

Gebruik:
    pip install openpyxl
    python generate_geldplanner.py [uitvoerbestand.xlsx]

Het script bouwt een werkmap met vijf tabbladen:
    1. Instructies
    2. Cashflow Overzicht
    3. ZZP Belastingtool
    4. Pensioengat Calculator
    5. Vermogensgroei

Alle bedragen worden als Excel-formules weggeschreven (geen voorberekende
waarden), zodat de formules in Excel zelf zichtbaar en aanpasbaar zijn.

DISCLAIMER: uitsluitend informatief, geen financieel advies.
"""

import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Huisstijl / opmaak
# ---------------------------------------------------------------------------

FONT_NAME = "Calibri"

CLR_DONKERBLAUW = "1F3864"
CLR_BLAUW = "2E5C8A"
CLR_LICHTBLAUW = "D9E2F3"
CLR_GEEL = "FFF2CC"
CLR_GEEL_RAND = "BF8F00"
CLR_GRIJS = "EDEDED"
CLR_GRIJS_RAND = "BFBFBF"
CLR_GROEN = "E2EFDA"
CLR_ROOD = "FCE4E4"
CLR_WIT = "FFFFFF"

F_TITEL = Font(name=FONT_NAME, size=18, bold=True, color=CLR_WIT)
F_SUBTITEL = Font(name=FONT_NAME, size=10, italic=True, color=CLR_WIT)
F_SECTIE = Font(name=FONT_NAME, size=11, bold=True, color=CLR_WIT)
F_KOP = Font(name=FONT_NAME, size=11, bold=True, color=CLR_DONKERBLAUW)
F_LABEL = Font(name=FONT_NAME, size=11)
F_LABEL_VET = Font(name=FONT_NAME, size=11, bold=True)
F_NOTITIE = Font(name=FONT_NAME, size=9, italic=True, color="7F7F7F")
F_RESULTAAT = Font(name=FONT_NAME, size=12, bold=True, color=CLR_DONKERBLAUW)
F_TEKST = Font(name=FONT_NAME, size=11)
F_TEKST_VET = Font(name=FONT_NAME, size=11, bold=True, color=CLR_DONKERBLAUW)

FILL_TITEL = PatternFill("solid", fgColor=CLR_DONKERBLAUW)
FILL_SECTIE = PatternFill("solid", fgColor=CLR_BLAUW)
FILL_INVOER = PatternFill("solid", fgColor=CLR_GEEL)
FILL_BEREKEND = PatternFill("solid", fgColor=CLR_GRIJS)
FILL_RESULTAAT = PatternFill("solid", fgColor=CLR_LICHTBLAUW)
FILL_GROEN = PatternFill("solid", fgColor=CLR_GROEN)
FILL_ROOD = PatternFill("solid", fgColor=CLR_ROOD)

_dun_geel = Side(style="thin", color=CLR_GEEL_RAND)
_dun_grijs = Side(style="thin", color=CLR_GRIJS_RAND)
_dun_blauw = Side(style="thin", color=CLR_BLAUW)

RAND_INVOER = Border(left=_dun_geel, right=_dun_geel, top=_dun_geel, bottom=_dun_geel)
RAND_BEREKEND = Border(left=_dun_grijs, right=_dun_grijs, top=_dun_grijs, bottom=_dun_grijs)
RAND_RESULTAAT = Border(left=_dun_blauw, right=_dun_blauw, top=_dun_blauw, bottom=_dun_blauw)

# Getalnotaties. openpyxl schrijft opmaakcodes in en-US-conventie weg; Excel
# toont ze in de landinstelling van de gebruiker, dus op een Nederlandse
# installatie verschijnt "€ 1.234,56" met de komma als decimaalteken.
FMT_EURO = '€ #,##0.00;[Red]-€ #,##0.00'
FMT_EURO_ROND = '€ #,##0;[Red]-€ #,##0'
FMT_PCT = '0.0%'
FMT_PCT2 = '0.00%'
FMT_GEHEEL = '0'
FMT_JAAR = '0 "jaar"'

# Cellen die de gebruiker mag invullen zijn ontgrendeld; al het andere blijft
# vergrendeld zodat werkbladbeveiliging de berekeningen beschermt.
BESCHERMD = Protection(locked=True)
ONBESCHERMD = Protection(locked=False)


# ---------------------------------------------------------------------------
# Kleine hulpfuncties voor het opbouwen van de bladen
# ---------------------------------------------------------------------------

def zet_kolombreedtes(ws, breedtes):
    for kolom, breedte in breedtes.items():
        ws.column_dimensions[kolom].width = breedte


def titelblok(ws, titel, ondertitel, laatste_kolom="F"):
    """Zet de donkerblauwe titelbalk bovenaan een blad. Retourneert de volgende rij."""
    ws.merge_cells(f"B2:{laatste_kolom}2")
    ws.merge_cells(f"B3:{laatste_kolom}3")
    cel = ws["B2"]
    cel.value = titel
    cel.font = F_TITEL
    cel.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    sub = ws["B3"]
    sub.value = ondertitel
    sub.font = F_SUBTITEL
    sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 32
    ws.row_dimensions[3].height = 18
    for rij in (2, 3):
        for kol in range(2, ws[f"{laatste_kolom}1"].column + 1):
            ws.cell(row=rij, column=kol).fill = FILL_TITEL
    return 5


def sectiebalk(ws, rij, tekst="", laatste_kolom="D"):
    ws.merge_cells(f"B{rij}:{laatste_kolom}{rij}")
    cel = ws.cell(row=rij, column=2)
    cel.value = tekst
    cel.font = F_SECTIE
    cel.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for kol in range(2, ws[f"{laatste_kolom}1"].column + 1):
        ws.cell(row=rij, column=kol).fill = FILL_SECTIE
    ws.row_dimensions[rij].height = 20
    return rij + 1


def invoerregel(ws, rij, label, standaardwaarde, notatie=FMT_EURO, notitie="",
                label_kolom=2, waarde_kolom=3, notitie_kolom=4):
    """Een gele, bewerkbare invoercel met label en optionele toelichting."""
    lbl = ws.cell(row=rij, column=label_kolom, value=label)
    lbl.font = F_LABEL
    lbl.alignment = Alignment(vertical="center", indent=1)

    cel = ws.cell(row=rij, column=waarde_kolom, value=standaardwaarde)
    cel.fill = FILL_INVOER
    cel.border = RAND_INVOER
    cel.number_format = notatie
    cel.font = F_LABEL
    cel.protection = ONBESCHERMD
    cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)

    if notitie:
        nt = ws.cell(row=rij, column=notitie_kolom, value=notitie)
        nt.font = F_NOTITIE
        nt.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[rij].height = 17
    return rij + 1


def berekendregel(ws, rij, label, formule, notatie=FMT_EURO, notitie="",
                  accent=False, label_kolom=2, waarde_kolom=3, notitie_kolom=4):
    """Een grijze, vergrendelde uitvoercel met formule."""
    lbl = ws.cell(row=rij, column=label_kolom, value=label)
    lbl.font = F_RESULTAAT if accent else F_LABEL_VET
    lbl.alignment = Alignment(vertical="center", indent=1)

    cel = ws.cell(row=rij, column=waarde_kolom, value=formule)
    cel.fill = FILL_RESULTAAT if accent else FILL_BEREKEND
    cel.border = RAND_RESULTAAT if accent else RAND_BEREKEND
    cel.number_format = notatie
    cel.font = F_RESULTAAT if accent else F_LABEL_VET
    cel.protection = BESCHERMD
    cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)

    if notitie:
        nt = ws.cell(row=rij, column=notitie_kolom, value=notitie)
        nt.font = F_NOTITIE
        nt.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[rij].height = 20 if accent else 17
    return rij + 1


def notitieregel(ws, rij, tekst="", laatste_kolom="D", vet=False):
    ws.merge_cells(f"B{rij}:{laatste_kolom}{rij}")
    cel = ws.cell(row=rij, column=2, value=tekst)
    cel.font = F_TEKST_VET if vet else F_NOTITIE
    cel.alignment = Alignment(vertical="center", indent=1, wrap_text=False)
    return rij + 1


def beveilig(ws):
    """Werkbladbeveiliging aan (zonder wachtwoord, zodat de eigenaar hem kan opheffen).

    Alle cellen zijn standaard vergrendeld; alleen de gele invoercellen zijn via
    ``Protection(locked=False)`` ontgrendeld en blijven dus bewerkbaar.
    """
    ws.protection.sheet = True
    ws.protection.enable()


def blad_basis(ws, tabkleur=CLR_DONKERBLAUW):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tabkleur
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


DISCLAIMER = ("Deze tool is uitsluitend bedoeld voor informatieve doeleinden "
              "en biedt geen financieel advies")


# ---------------------------------------------------------------------------
# TABBLAD 1 — Instructies
# ---------------------------------------------------------------------------

def bouw_instructies(wb):
    ws = wb.create_sheet("Instructies")
    blad_basis(ws, CLR_DONKERBLAUW)
    zet_kolombreedtes(ws, {"A": 2, "B": 26, "C": 78, "D": 2})

    r = titelblok(ws, "Geldplanner", "Persoonlijk financieel overzicht voor Nederlandse huishoudens en zzp'ers", "C")

    r = notitieregel(ws, r, "Welkom", "C", vet=True)
    for regel in [
        "Deze werkmap helpt je in kaart te brengen wat er maandelijks binnenkomt en weggaat, wat je",
        "als zzp'er netto overhoudt, hoe groot je pensioengat is en hoe je vermogen groeit.",
        "Vul alleen de gele cellen in. De grijze en blauwe cellen bevatten formules en zijn vergrendeld.",
    ]:
        r = notitieregel(ws, r, regel, "C")
    r += 1

    r = sectiebalk(ws, r, "DE TABBLADEN", "C")
    kop_b = ws.cell(row=r, column=2, value="Tabblad")
    kop_c = ws.cell(row=r, column=3, value="Wat doe je hier?")
    for cel in (kop_b, kop_c):
        cel.font = F_KOP
        cel.fill = FILL_RESULTAAT
        cel.border = RAND_RESULTAAT
        cel.alignment = Alignment(vertical="center", indent=1)
    r += 1

    uitleg = [
        ("1. Instructies",
         "Deze pagina: uitleg, kleurcodering en de disclaimer."),
        ("2. Cashflow Overzicht",
         "Vul je maandelijkse inkomsten, vaste lasten en variabele kosten in. Je ziet direct je "
         "totale inkomsten, totale uitgaven, het maandelijkse overschot of tekort en je spaarquote."),
        ("3. ZZP Belastingtool",
         "Vul je jaaromzet, zakelijke kosten en of je aan het urencriterium voldoet in. Je krijgt een "
         "schatting van je belastbare winst, inkomstenbelasting, heffingskortingen, "
         "Zvw-bijdrage en wat je netto per maand overhoudt."),
        ("4. Pensioengat Calculator",
         "Vul je leeftijd, gewenste pensioenleeftijd, gewenst pensioeninkomen, opgebouwd "
         "pensioenvermogen en verwachte AOW in. Je ziet het gat en het bedrag dat je maandelijks "
         "moet sparen of beleggen om het te dichten."),
        ("5. Vermogensgroei",
         "Vul je startkapitaal, maandelijkse inleg, verwacht rendement en looptijd in. Je ziet de "
         "geprojecteerde waarde per jaar plus een grafiek van de groei."),
    ]
    for naam, tekst in uitleg:
        b = ws.cell(row=r, column=2, value=naam)
        b.font = F_LABEL_VET
        b.alignment = Alignment(vertical="top", indent=1)
        c = ws.cell(row=r, column=3, value=tekst)
        c.font = F_TEKST
        c.alignment = Alignment(vertical="top", indent=1, wrap_text=True)
        ws.row_dimensions[r].height = 44 if len(tekst) > 95 else 30
        for cel in (b, c):
            cel.border = RAND_BEREKEND
        r += 1
    r += 1

    r = sectiebalk(ws, r, "KLEURCODERING", "C")
    legenda = [
        (FILL_INVOER, RAND_INVOER, "Invoercel (geel)",
         "Hier vul je zelf iets in. Alleen deze cellen zijn bewerkbaar."),
        (FILL_BEREKEND, RAND_BEREKEND, "Berekende cel (grijs)",
         "Bevat een formule. Vergrendeld zodat je niet per ongeluk een berekening overschrijft."),
        (FILL_RESULTAAT, RAND_RESULTAAT, "Eindresultaat (blauw)",
         "De belangrijkste uitkomst van een sectie. Ook vergrendeld."),
    ]
    for vulling, rand, naam, tekst in legenda:
        b = ws.cell(row=r, column=2, value=naam)
        b.font = F_LABEL_VET
        b.fill = vulling
        b.border = rand
        b.alignment = Alignment(vertical="center", indent=1)
        c = ws.cell(row=r, column=3, value=tekst)
        c.font = F_TEKST
        c.alignment = Alignment(vertical="center", indent=1)
        r += 1
    r += 1

    r = sectiebalk(ws, r, "ZO GEBRUIK JE DE WERKMAP", "C")
    stappen = [
        "1. Begin bij 'Cashflow Overzicht' en vul je maandbudget in. Alles wat niet van toepassing is laat je op 0 staan.",
        "2. Ben je zzp'er? Ga daarna naar 'ZZP Belastingtool' en schat je netto maandinkomen.",
        "3. Gebruik dat netto bedrag als inkomen in het cashflow-overzicht.",
        "4. Stop je maandelijkse overschot in 'Vermogensgroei' om te zien wat dat op termijn oplevert.",
        "5. Controleer met 'Pensioengat Calculator' of dat genoeg is voor je gewenste pensioen.",
        "6. De bladen zijn beveiligd zonder wachtwoord. Wil je formules aanpassen? Controleren > Blad-beveiliging opheffen.",
        "7. Bedragen staan in euro's. Op een Nederlandse Excel-installatie verschijnen ze als € 1.234,56.",
    ]
    for stap in stappen:
        cel = ws.cell(row=r, column=2, value=stap)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        cel.font = F_TEKST
        cel.alignment = Alignment(vertical="center", indent=1)
        r += 1
    r += 1

    r = sectiebalk(ws, r, "BELANGRIJK — DISCLAIMER", "C")
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 3, end_column=3)
    cel = ws.cell(row=r, column=2)
    cel.value = (
        DISCLAIMER + ".\n\n"
        "De belastingtarieven, heffingskortingen en aftrekposten in dit bestand zijn indicatieve "
        "cijfers voor 2026 en kunnen door de wetgever worden gewijzigd. Controleer de actuele "
        "bedragen altijd op belastingdienst.nl. Rendementen uit het verleden bieden geen garantie "
        "voor de toekomst; beleggen brengt risico's met zich mee en je kunt je inleg verliezen.\n"
        "Raadpleeg voor persoonlijk advies een financieel adviseur, boekhouder of belastingadviseur."
    )
    cel.font = Font(name=FONT_NAME, size=11, bold=True, color="9C0006")
    cel.fill = FILL_ROOD
    cel.border = Border(left=Side(style="medium", color="C00000"),
                        right=Side(style="medium", color="C00000"),
                        top=Side(style="medium", color="C00000"),
                        bottom=Side(style="medium", color="C00000"))
    cel.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
    for rij in range(r, r + 4):
        ws.row_dimensions[rij].height = 26

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# TABBLAD 2 — Cashflow Overzicht
# ---------------------------------------------------------------------------

def bouw_cashflow(wb):
    ws = wb.create_sheet("Cashflow Overzicht")
    blad_basis(ws, "2E75B6")
    zet_kolombreedtes(ws, {"A": 2, "B": 38, "C": 16, "D": 44, "E": 2,
                           "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12})

    r = titelblok(ws, "Cashflow Overzicht", "Wat komt er maandelijks binnen en wat gaat eruit?", "D")

    # --- Inkomsten ---
    r = sectiebalk(ws, r, "INKOMSTEN PER MAAND")
    inkomsten_start = r
    r = invoerregel(ws, r, "Nettosalaris persoon 1", 3600, notitie="Bedrag dat op je rekening staat")
    r = invoerregel(ws, r, "Nettosalaris persoon 2", 0, notitie="Laat op 0 staan als je alleenstaand bent")
    r = invoerregel(ws, r, "Freelance-inkomsten (netto)", 0, notitie="Zie tabblad 'ZZP Belastingtool'")
    r = invoerregel(ws, r, "Toeslagen (huur-, zorg-, kinderopvang)", 0)
    r = invoerregel(ws, r, "Huurinkomsten / alimentatie", 0)
    r = invoerregel(ws, r, "Overige inkomsten", 0)
    inkomsten_eind = r - 1
    tot_inkomsten_rij = r
    r = berekendregel(ws, r, "Totale inkomsten",
                      f"=SUM(C{inkomsten_start}:C{inkomsten_eind})")
    r += 1

    # --- Vaste lasten ---
    r = sectiebalk(ws, r, "VASTE LASTEN PER MAAND")
    vast_start = r
    r = invoerregel(ws, r, "Huur of hypotheek", 1250)
    r = invoerregel(ws, r, "Energie, water en warmte", 180)
    r = invoerregel(ws, r, "Zorgverzekering", 290, notitie="Beide partners samen")
    r = invoerregel(ws, r, "Overige verzekeringen", 65, notitie="Inboedel, aansprakelijkheid, opstal, ORV")
    r = invoerregel(ws, r, "Gemeentelijke belastingen en waterschap", 60, notitie="Jaarbedrag gedeeld door 12")
    r = invoerregel(ws, r, "Internet, tv en telefoon", 70)
    r = invoerregel(ws, r, "Abonnementen en lidmaatschappen", 45, notitie="Streaming, sportschool, kranten")
    r = invoerregel(ws, r, "Aflossing leningen en studieschuld", 0)
    r = invoerregel(ws, r, "Kinderopvang / school", 0)
    r = invoerregel(ws, r, "Overige vaste lasten", 0)
    vast_eind = r - 1
    tot_vast_rij = r
    r = berekendregel(ws, r, "Totaal vaste lasten", f"=SUM(C{vast_start}:C{vast_eind})")
    r += 1

    # --- Variabele kosten ---
    r = sectiebalk(ws, r, "VARIABELE KOSTEN PER MAAND")
    var_start = r
    r = invoerregel(ws, r, "Boodschappen", 550)
    r = invoerregel(ws, r, "Vervoer (brandstof, OV, onderhoud)", 200)
    r = invoerregel(ws, r, "Vrije tijd, uitgaan en horeca", 200)
    r = invoerregel(ws, r, "Kleding en persoonlijke verzorging", 100)
    r = invoerregel(ws, r, "Zorgkosten (eigen risico, tandarts)", 40)
    r = invoerregel(ws, r, "Vakantie (gereserveerd per maand)", 150)
    r = invoerregel(ws, r, "Cadeaus, goede doelen en overig", 75)
    var_eind = r - 1
    tot_var_rij = r
    r = berekendregel(ws, r, "Totaal variabele kosten", f"=SUM(C{var_start}:C{var_eind})")
    r += 1

    # --- Resultaat ---
    r = sectiebalk(ws, r, "RESULTAAT")
    res_inkomsten = r
    r = berekendregel(ws, r, "Totale inkomsten", f"=C{tot_inkomsten_rij}")
    res_uitgaven = r
    r = berekendregel(ws, r, "Totale uitgaven", f"=C{tot_vast_rij}+C{tot_var_rij}")
    overschot_rij = r
    r = berekendregel(ws, r, "Maandelijks overschot / tekort",
                      f"=C{res_inkomsten}-C{res_uitgaven}", accent=True,
                      notitie="Positief = je houdt over. Negatief = je komt tekort.")
    r = berekendregel(ws, r, "Spaarquote",
                      f'=IF(C{res_inkomsten}=0,"",C{overschot_rij}/C{res_inkomsten})',
                      notatie=FMT_PCT, notitie="Deel van je inkomen dat je overhoudt")
    r = berekendregel(ws, r, "Overschot / tekort per jaar",
                      f"=C{overschot_rij}*12")
    r = berekendregel(ws, r, "Vaste lasten als deel van inkomen",
                      f'=IF(C{res_inkomsten}=0,"",C{tot_vast_rij}/C{res_inkomsten})',
                      notatie=FMT_PCT, notitie="Vuistregel: houd dit onder 50%")
    buffer_bereken_rij = r
    buffer_invoer_rij = r + 2   # de gele cel twee regels lager
    r = berekendregel(
        ws, buffer_bereken_rij, "Buffer: hoeveel maanden uitgaven heb je?",
        f'=IF(C{tot_vast_rij}+C{tot_var_rij}=0,"",'
        f'C{buffer_invoer_rij}/(C{tot_vast_rij}+C{tot_var_rij}))',
        notatie='0.0 "maanden"', notitie="Vuistregel: drie tot zes maanden")
    r += 1
    r = invoerregel(ws, buffer_invoer_rij, "Huidig spaargeld (buffer)", 5000,
                    notitie="Wordt gebruikt voor de bufferberekening hierboven")
    r += 1

    # Voorwaardelijke opmaak: groen bij overschot, rood bij tekort.
    ws.conditional_formatting.add(
        f"C{overschot_rij}",
        CellIsRule(operator="lessThan", formula=["0"], fill=FILL_ROOD,
                   font=Font(name=FONT_NAME, size=12, bold=True, color="9C0006")))
    ws.conditional_formatting.add(
        f"C{overschot_rij}",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=FILL_GROEN,
                   font=Font(name=FONT_NAME, size=12, bold=True, color="375623")))

    # --- Grafiekgegevens ---
    r += 1
    grafiek_kop = r
    kop = ws.cell(row=r, column=2, value="Grafiekgegevens")
    kop.font = F_KOP
    ws.cell(row=r, column=3, value="Bedrag per maand").font = F_KOP
    r += 1
    grafiek_start = r
    grafiek_rijen = [
        ("Inkomsten", f"=C{res_inkomsten}"),
        ("Vaste lasten", f"=C{tot_vast_rij}"),
        ("Variabele kosten", f"=C{tot_var_rij}"),
        ("Totale uitgaven", f"=C{res_uitgaven}"),
    ]
    for label, formule in grafiek_rijen:
        lbl = ws.cell(row=r, column=2, value=label)
        lbl.font = F_LABEL
        lbl.alignment = Alignment(vertical="center", indent=1)
        cel = ws.cell(row=r, column=3, value=formule)
        cel.fill = FILL_BEREKEND
        cel.border = RAND_BEREKEND
        cel.number_format = FMT_EURO
        cel.font = F_LABEL
        cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        r += 1
    grafiek_eind = r - 1

    # --- Staafdiagram ---
    grafiek = BarChart()
    grafiek.type = "col"
    grafiek.style = 10
    grafiek.title = "Inkomsten vs. uitgaven per maand"
    grafiek.y_axis.title = "Bedrag per maand (€)"
    grafiek.x_axis.title = None
    grafiek.height = 9
    grafiek.width = 17
    gegevens = Reference(ws, min_col=3, min_row=grafiek_kop, max_row=grafiek_eind)
    categorieen = Reference(ws, min_col=2, min_row=grafiek_start, max_row=grafiek_eind)
    grafiek.add_data(gegevens, titles_from_data=True)
    grafiek.set_categories(categorieen)
    grafiek.legend = None
    grafiek.dLbls = DataLabelList()
    grafiek.dLbls.showVal = True
    grafiek.gapWidth = 60
    ws.add_chart(grafiek, "F5")

    notitieregel(ws, grafiek_eind + 2, DISCLAIMER + ".", "D")

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# TABBLAD 3 — ZZP Belastingtool
# ---------------------------------------------------------------------------

def bouw_zzp(wb):
    ws = wb.create_sheet("ZZP Belastingtool")
    blad_basis(ws, "548235")
    zet_kolombreedtes(ws, {"A": 2, "B": 42, "C": 16, "D": 50, "E": 2,
                           "F": 14, "G": 14, "H": 14, "I": 14})

    r = titelblok(ws, "ZZP Belastingtool", "Schatting van je inkomstenbelasting en netto maandinkomen (box 1)", "E")

    # --- Waarschuwing ---
    ws.merge_cells(f"B{r}:E{r + 1}")
    waarschuwing = ws.cell(row=r, column=2)
    waarschuwing.value = ("LET OP: dit is een schatting, geen garantie. De uitkomst is geen aangifte en geen "
                          "financieel of fiscaal advies. Tarieven en kortingen voor 2026 zijn indicatief — "
                          "controleer ze op belastingdienst.nl en pas ze zo nodig hieronder aan.")
    waarschuwing.font = Font(name=FONT_NAME, size=10, bold=True, color="9C6500")
    waarschuwing.fill = PatternFill("solid", fgColor="FFF2CC")
    waarschuwing.border = Border(left=Side(style="medium", color=CLR_GEEL_RAND),
                                 right=Side(style="medium", color=CLR_GEEL_RAND),
                                 top=Side(style="medium", color=CLR_GEEL_RAND),
                                 bottom=Side(style="medium", color=CLR_GEEL_RAND))
    waarschuwing.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 20
    ws.row_dimensions[r + 1].height = 20
    r += 3

    # --- Invoer ---
    r = sectiebalk(ws, r, laatste_kolom="E", tekst="JOUW GEGEVENS")
    omzet_rij = r
    r = invoerregel(ws, r, "Jaaromzet als freelancer (excl. btw)", 75000,
                    notitie="Alle facturen samen, exclusief btw")
    kosten_rij = r
    r = invoerregel(ws, r, "Zakelijke kosten per jaar", 8000,
                    notitie="Laptop, verzekeringen, kantoor, reiskosten, boekhouder")
    uren_rij = r
    r = invoerregel(ws, r, "Voldoe je aan het urencriterium?", "Ja", notatie="General",
                    notitie="1.225 uur per jaar → recht op zelfstandigenaftrek")
    starter_rij = r
    r = invoerregel(ws, r, "Recht op startersaftrek dit jaar?", "Nee", notatie="General",
                    notitie="Max. 3x in de eerste 5 jaar als ondernemer")
    for_rij = r
    r = invoerregel(ws, r, "Overige aftrekposten", 0,
                    notitie="Bijv. meewerkaftrek of stakingsaftrek")
    aow_leeftijd_rij = r
    r = invoerregel(ws, r, "Heb je de AOW-leeftijd bereikt?", "Nee", notatie="General",
                    notitie="Vanaf de AOW-leeftijd gelden lagere tarieven (hier niet doorgerekend)")
    r += 1

    ja_nee = DataValidation(type="list", formula1='"Ja,Nee"', allow_blank=False,
                            showDropDown=False)
    ja_nee.error = "Kies Ja of Nee."
    ja_nee.errorTitle = "Ongeldige invoer"
    ws.add_data_validation(ja_nee)
    for rij in (uren_rij, starter_rij, aow_leeftijd_rij):
        ja_nee.add(ws.cell(row=rij, column=3))

    # --- Parameters ---
    r = sectiebalk(ws, r, laatste_kolom="E", tekst="BELASTINGPARAMETERS 2026 (INDICATIEF — AANPASBAAR)")
    p_grens1 = r
    r = invoerregel(ws, r, "Schijf 1 — grens", 38883, notatie=FMT_EURO_ROND,
                    notitie="Inkomen tot en met dit bedrag valt in schijf 1")
    p_tarief1 = r
    r = invoerregel(ws, r, "Schijf 1 — tarief", 0.3570, notatie=FMT_PCT2,
                    notitie="Inclusief premies volksverzekeringen")
    p_grens2 = r
    r = invoerregel(ws, r, "Schijf 2 — grens", 79137, notatie=FMT_EURO_ROND)
    p_tarief2 = r
    r = invoerregel(ws, r, "Schijf 2 — tarief", 0.3756, notatie=FMT_PCT2)
    p_tarief3 = r
    r = invoerregel(ws, r, "Schijf 3 — tarief (boven schijf 2)", 0.4950, notatie=FMT_PCT2)
    p_zelfst = r
    r = invoerregel(ws, r, "Zelfstandigenaftrek", 1200, notatie=FMT_EURO_ROND,
                    notitie="Wordt jaarlijks afgebouwd")
    p_starters = r
    r = invoerregel(ws, r, "Startersaftrek", 2123, notatie=FMT_EURO_ROND)
    p_mkb = r
    r = invoerregel(ws, r, "MKB-winstvrijstelling", 0.127, notatie=FMT_PCT2,
                    notitie="Over de winst na ondernemersaftrek")
    p_ahk_max = r
    r = invoerregel(ws, r, "Algemene heffingskorting — maximum", 3115, notatie=FMT_EURO_ROND)
    p_ahk_start = r
    r = invoerregel(ws, r, "Algemene heffingskorting — afbouw vanaf", 29529, notatie=FMT_EURO_ROND)
    p_ahk_pct = r
    r = invoerregel(ws, r, "Algemene heffingskorting — afbouwpercentage", 0.06337, notatie=FMT_PCT2)
    p_zvw_pct = r
    r = invoerregel(ws, r, "Zvw-bijdrage (inkomensafhankelijk)", 0.0526, notatie=FMT_PCT2,
                    notitie="Betaal je als ondernemer zelf via de aanslag")
    p_zvw_max = r
    r = invoerregel(ws, r, "Zvw — maximumgrondslag", 78727, notatie=FMT_EURO_ROND)
    r += 1

    # --- Arbeidskorting-tabel ---
    r = sectiebalk(ws, r, laatste_kolom="E", tekst="ARBEIDSKORTING — OPBOUW EN AFBOUW (INDICATIEF 2026)")
    ak_kop = r
    koppen = ["Schijf", "Vanaf", "Basisbedrag", "Percentage"]
    for i, tekst in enumerate(koppen):
        cel = ws.cell(row=r, column=2 + i, value=tekst)
        cel.font = F_KOP
        cel.fill = FILL_RESULTAAT
        cel.border = RAND_RESULTAAT
        cel.alignment = Alignment(horizontal="center", vertical="center")
    r += 1
    ak_start = r
    ak_rijen = [
        ("1 — opbouw", 0, 0, 0.08053),
        ("2 — opbouw", 12400, 999, 0.30030),
        ("3 — opbouw", 26800, 5322, 0.02258),
        ("4 — afbouw", 43900, 5712, 0.06510),
    ]
    for naam, vanaf, basis, pct in ak_rijen:
        lbl = ws.cell(row=r, column=2, value=naam)
        lbl.font = F_LABEL
        lbl.alignment = Alignment(vertical="center", indent=1)
        for kol, waarde, notatie in ((3, vanaf, FMT_EURO_ROND),
                                     (4, basis, FMT_EURO_ROND),
                                     (5, pct, FMT_PCT2)):
            cel = ws.cell(row=r, column=kol, value=waarde)
            cel.fill = FILL_INVOER
            cel.border = RAND_INVOER
            cel.number_format = notatie
            cel.font = F_LABEL
            cel.protection = ONBESCHERMD
            cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        r += 1
    ak_1, ak_2, ak_3, ak_4 = ak_start, ak_start + 1, ak_start + 2, ak_start + 3
    ws.column_dimensions["E"].width = 14
    r = notitieregel(ws, r, "In schijf 4 wordt de arbeidskorting afgebouwd; het percentage geldt daar als aftrek.", "E")
    r += 1

    # --- Berekening ---
    r = sectiebalk(ws, r, laatste_kolom="E", tekst="BEREKENING")
    winst_rij = r
    r = berekendregel(ws, r, "Winst uit onderneming", f"=C{omzet_rij}-C{kosten_rij}",
                      notitie="Omzet minus zakelijke kosten")
    zelfst_rij = r
    r = berekendregel(ws, r, "Zelfstandigenaftrek",
                      f'=IF(C{uren_rij}="Ja",MIN(C{p_zelfst},MAX(0,C{winst_rij})),0)',
                      notitie="Niet hoger dan de winst")
    starters_rij = r
    r = berekendregel(ws, r, "Startersaftrek",
                      f'=IF(AND(C{uren_rij}="Ja",C{starter_rij}="Ja"),C{p_starters},0)',
                      notitie="Mag de winst wel overschrijden")
    ondernemersaftrek_rij = r
    r = berekendregel(ws, r, "Totale ondernemersaftrek",
                      f"=C{zelfst_rij}+C{starters_rij}+C{for_rij}")
    winst_na_aftrek_rij = r
    r = berekendregel(ws, r, "Winst na ondernemersaftrek",
                      f"=MAX(0,C{winst_rij}-C{ondernemersaftrek_rij})")
    mkb_rij = r
    r = berekendregel(ws, r, "MKB-winstvrijstelling",
                      f"=C{winst_na_aftrek_rij}*C{p_mkb}")
    belastbaar_rij = r
    r = berekendregel(ws, r, "Belastbare winst (box 1)",
                      f"=MAX(0,C{winst_na_aftrek_rij}-C{mkb_rij})", accent=True)
    ib_bruto_rij = r
    r = berekendregel(
        ws, r, "Inkomstenbelasting vóór heffingskortingen",
        f"=MIN(C{belastbaar_rij},C{p_grens1})*C{p_tarief1}"
        f"+MAX(0,MIN(C{belastbaar_rij},C{p_grens2})-C{p_grens1})*C{p_tarief2}"
        f"+MAX(0,C{belastbaar_rij}-C{p_grens2})*C{p_tarief3}",
        notitie="Schijventarief toegepast op de belastbare winst")
    ahk_rij = r
    r = berekendregel(
        ws, r, "Algemene heffingskorting",
        f"=MAX(0,C{p_ahk_max}-MAX(0,C{belastbaar_rij}-C{p_ahk_start})*C{p_ahk_pct})",
        notitie="Bouwt af naarmate je inkomen stijgt")
    ak_rij = r
    r = berekendregel(
        ws, r, "Arbeidskorting",
        f"=IF(C{belastbaar_rij}<=C{ak_2},D{ak_1}+C{belastbaar_rij}*E{ak_1},"
        f"IF(C{belastbaar_rij}<=C{ak_3},D{ak_2}+(C{belastbaar_rij}-C{ak_2})*E{ak_2},"
        f"IF(C{belastbaar_rij}<=C{ak_4},D{ak_3}+(C{belastbaar_rij}-C{ak_3})*E{ak_3},"
        f"MAX(0,D{ak_4}-(C{belastbaar_rij}-C{ak_4})*E{ak_4}))))",
        notitie="Op basis van de tabel hierboven")
    kortingen_rij = r
    r = berekendregel(ws, r, "Totaal heffingskortingen", f"=C{ahk_rij}+C{ak_rij}")
    ib_netto_rij = r
    r = berekendregel(ws, r, "Te betalen inkomstenbelasting",
                      f"=MAX(0,C{ib_bruto_rij}-C{kortingen_rij})",
                      notitie="Heffingskortingen worden niet uitbetaald onder nul")
    zvw_rij = r
    r = berekendregel(ws, r, "Zvw-bijdrage",
                      f"=MIN(C{belastbaar_rij},C{p_zvw_max})*C{p_zvw_pct}")
    totale_heffing_rij = r
    r = berekendregel(ws, r, "Totaal te betalen (IB + Zvw)",
                      f"=C{ib_netto_rij}+C{zvw_rij}", accent=True)
    r += 1

    # --- Uitkomst ---
    r = sectiebalk(ws, r, laatste_kolom="E", tekst="WAT HOUD JE OVER?")
    netto_jaar_rij = r
    r = berekendregel(ws, r, "Netto besteedbaar per jaar",
                      f"=C{winst_rij}-C{totale_heffing_rij}", accent=True,
                      notitie="Winst minus inkomstenbelasting en Zvw-bijdrage")
    netto_maand_rij = r
    r = berekendregel(ws, r, "Netto besteedbaar per maand",
                      f"=C{netto_jaar_rij}/12", accent=True,
                      notitie="Neem dit bedrag over in 'Cashflow Overzicht'")
    r = berekendregel(ws, r, "Gemiddelde belastingdruk",
                      f'=IF(C{winst_rij}<=0,"",C{totale_heffing_rij}/C{winst_rij})',
                      notatie=FMT_PCT, notitie="Als percentage van je winst")
    reservering_rij = r
    r = berekendregel(ws, r, "Reserveer per maand voor de aanslag",
                      f"=C{totale_heffing_rij}/12",
                      notitie="Zet dit apart op een aparte rekening")
    r = berekendregel(ws, r, "Marginaal tarief op de volgende euro winst",
                      f"=IF(C{belastbaar_rij}<C{p_grens1},C{p_tarief1},"
                      f"IF(C{belastbaar_rij}<C{p_grens2},C{p_tarief2},C{p_tarief3}))*(1-C{p_mkb})",
                      notatie=FMT_PCT,
                      notitie="Bij benadering, inclusief MKB-winstvrijstelling")
    r += 1

    r = notitieregel(ws, r, "Niet meegenomen in deze schatting:", "E", vet=True)
    for tekst in [
        "• Btw-aangifte (btw loopt buiten je winst om), inkomen van een fiscaal partner en box 3-vermogen.",
        "• Fiscale oudedagsreserve, investeringsaftrek (KIA), afschrijvingen en middeling.",
        "• Loon uit dienstbetrekking naast je onderneming; dat verhoogt je schijf en verlaagt je kortingen.",
        "• Inkomensafhankelijke toeslagen en de tarieven vanaf de AOW-leeftijd.",
    ]:
        r = notitieregel(ws, r, tekst, "E")
    r += 1
    r = notitieregel(ws, r, DISCLAIMER + ". Dit is een schatting, geen garantie.", "E")

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# TABBLAD 4 — Pensioengat Calculator
# ---------------------------------------------------------------------------

def bouw_pensioen(wb):
    ws = wb.create_sheet("Pensioengat Calculator")
    blad_basis(ws, "7030A0")
    zet_kolombreedtes(ws, {"A": 2, "B": 44, "C": 16, "D": 52, "E": 2})

    r = titelblok(ws, "Pensioengat Calculator", "Hoeveel moet je maandelijks opzijzetten voor het pensioen dat je wilt?", "D")

    r = sectiebalk(ws, r, "JOUW SITUATIE")
    leeftijd_rij = r
    r = invoerregel(ws, r, "Huidige leeftijd", 38, notatie=FMT_JAAR)
    pensioenleeftijd_rij = r
    r = invoerregel(ws, r, "Gewenste pensioenleeftijd", 67, notatie=FMT_JAAR,
                    notitie="De AOW-leeftijd is nu 67 jaar en beweegt mee met de levensverwachting")
    gewenst_rij = r
    r = invoerregel(ws, r, "Gewenst pensioeninkomen per maand", 3500,
                    notitie="In euro's van vandaag, netto besteedbaar")
    aow_rij = r
    r = invoerregel(ws, r, "Verwachte AOW per maand", 1500,
                    notitie="Ca. € 1.050 alleenstaand, ca. € 750 p.p. samenwonend (netto, 2025)")
    pensioen_rij = r
    r = invoerregel(ws, r, "Verwacht werkgeverspensioen per maand", 700,
                    notitie="Zie mijnpensioenoverzicht.nl")
    vermogen_rij = r
    r = invoerregel(ws, r, "Huidig opgebouwd pensioenvermogen", 30000,
                    notitie="Lijfrente, banksparen, beleggingen bestemd voor pensioen")
    inleg_rij = r
    r = invoerregel(ws, r, "Wat leg je nu al maandelijks in?", 150,
                    notitie="Wordt van de benodigde inleg afgetrokken")
    r += 1

    r = sectiebalk(ws, r, "AANNAMES")
    rendement_rij = r
    r = invoerregel(ws, r, "Verwacht rendement vóór pensioen (per jaar)", 0.05, notatie=FMT_PCT,
                    notitie="Historisch wereldwijd aandelenrendement ligt rond 6-7% nominaal")
    rendement_na_rij = r
    r = invoerregel(ws, r, "Verwacht rendement tijdens pensioen (per jaar)", 0.03, notatie=FMT_PCT,
                    notitie="Meestal lager: je bouwt risico af")
    inflatie_rij = r
    r = invoerregel(ws, r, "Verwachte inflatie (per jaar)", 0.02, notatie=FMT_PCT,
                    notitie="Doelstelling van de ECB is 2%")
    duur_rij = r
    r = invoerregel(ws, r, "Aantal jaren dat je pensioen moet duren", 25, notatie=FMT_JAAR,
                    notitie="Levensverwachting bij 67 is ruim 18 jaar; 25 is voorzichtig")
    r += 1

    r = sectiebalk(ws, r, "HET PENSIOENGAT")
    jaren_rij = r
    r = berekendregel(ws, r, "Jaren tot je pensioen",
                      f"=MAX(0,C{pensioenleeftijd_rij}-C{leeftijd_rij})", notatie=FMT_JAAR)
    tekort_nu_rij = r
    r = berekendregel(ws, r, "Maandelijks tekort in euro's van vandaag",
                      f"=MAX(0,C{gewenst_rij}-C{aow_rij}-C{pensioen_rij})",
                      notitie="Gewenst inkomen minus AOW en werkgeverspensioen")
    tekort_toekomst_rij = r
    r = berekendregel(ws, r, "Maandelijks tekort bij pensionering",
                      f"=C{tekort_nu_rij}*(1+C{inflatie_rij})^C{jaren_rij}",
                      notitie="Hetzelfde tekort, gecorrigeerd voor inflatie")
    reeel_rij = r
    r = berekendregel(ws, r, "Reëel rendement tijdens pensioen",
                      f"=(1+C{rendement_na_rij})/(1+C{inflatie_rij})-1", notatie=FMT_PCT2,
                      notitie="Rendement minus inflatie; houdt je uitkering koopkrachtvast")
    kapitaal_rij = r
    r = berekendregel(
        ws, r, "Benodigd pensioenkapitaal bij pensionering",
        f"=IF(ABS(C{reeel_rij})<0.0001,C{tekort_toekomst_rij}*12*C{duur_rij},"
        f"C{tekort_toekomst_rij}*12*(1-(1+C{reeel_rij})^(-C{duur_rij}))/C{reeel_rij})",
        accent=True,
        notitie="Contante waarde van het tekort over de hele pensioenperiode")
    groei_rij = r
    r = berekendregel(ws, r, "Waarde van je huidige vermogen bij pensionering",
                      f"=C{vermogen_rij}*(1+C{rendement_rij})^C{jaren_rij}",
                      notitie="Je opgebouwde vermogen groeit door tot je pensioendatum")
    groei_inleg_rij = r
    r = berekendregel(
        ws, r, "Waarde van je huidige maandinleg bij pensionering",
        f"=IF(OR(C{jaren_rij}=0,C{rendement_rij}=0),C{inleg_rij}*12*C{jaren_rij},"
        f"C{inleg_rij}*(((1+(1+C{rendement_rij})^(1/12)-1)^(C{jaren_rij}*12)-1)"
        f"/((1+C{rendement_rij})^(1/12)-1)))",
        notitie="Wat je huidige inleg oplevert als je die volhoudt")
    gat_rij = r
    r = berekendregel(ws, r, "PENSIOENGAT (tekort aan kapitaal)",
                      f"=MAX(0,C{kapitaal_rij}-C{groei_rij}-C{groei_inleg_rij})", accent=True,
                      notitie="Nul betekent: je ligt op koers")
    r += 1

    r = sectiebalk(ws, r, "WAT MOET JE DOEN?")
    maandrente_rij = r
    r = berekendregel(ws, r, "Maandelijks rendement (samengesteld)",
                      f"=(1+C{rendement_rij})^(1/12)-1", notatie='0.000%',
                      notitie="Jaarrendement omgerekend naar maandbasis")
    extra_rij = r
    r = berekendregel(
        ws, r, "Extra maandelijks sparen of beleggen",
        f'=IF(C{jaren_rij}<=0,"",IF(C{gat_rij}<=0,0,'
        f"IF(C{maandrente_rij}=0,C{gat_rij}/(C{jaren_rij}*12),"
        f"C{gat_rij}*C{maandrente_rij}/((1+C{maandrente_rij})^(C{jaren_rij}*12)-1))))",
        accent=True, notitie="Bovenop wat je nu al inlegt")
    totaal_rij = r
    r = berekendregel(ws, r, "Totale maandelijkse inleg vanaf nu",
                      f'=IF(C{extra_rij}="","",C{inleg_rij}+C{extra_rij})',
                      notitie="Je huidige inleg plus het extra bedrag")
    r = berekendregel(ws, r, "Totaal dat je zelf inlegt tot pensioendatum",
                      f'=IF(C{totaal_rij}="","",C{totaal_rij}*12*C{jaren_rij})',
                      notitie="Zonder rendement, alleen je eigen geld")
    r = berekendregel(ws, r, "Dekkingsgraad zonder extra inleg",
                      f'=IF(C{kapitaal_rij}<=0,"",(C{groei_rij}+C{groei_inleg_rij})/C{kapitaal_rij})',
                      notatie=FMT_PCT, notitie="100% = je huidige plan is toereikend")
    r += 1

    ws.conditional_formatting.add(
        f"C{gat_rij}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_ROOD,
                   font=Font(name=FONT_NAME, size=12, bold=True, color="9C0006")))
    ws.conditional_formatting.add(
        f"C{gat_rij}",
        CellIsRule(operator="lessThanOrEqual", formula=["0"], fill=FILL_GROEN,
                   font=Font(name=FONT_NAME, size=12, bold=True, color="375623")))

    r = notitieregel(ws, r, "Hoe deze berekening werkt", "D", vet=True)
    for tekst in [
        "• Je gewenste inkomen wordt met de inflatie opgehoogd naar je pensioendatum.",
        "• Het benodigde kapitaal is de contante waarde van je maandelijkse tekort over de hele pensioenperiode,",
        "   verdisconteerd tegen het reële rendement, zodat je uitkering koopkrachtvast blijft.",
        "• De benodigde inleg volgt uit de standaardformule voor samengestelde rente met periodieke stortingen.",
        "• Belastingvoordeel op lijfrente-inleg (jaarruimte) is niet meegerekend; in de praktijk kan de netto inleg lager uitvallen.",
        "",
        DISCLAIMER + ". Uitkomsten zijn schattingen op basis van je eigen aannames.",
    ]:
        r = notitieregel(ws, r, tekst, "D")

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# TABBLAD 5 — Vermogensgroei
# ---------------------------------------------------------------------------

MAX_JAREN = 40


def bouw_vermogen(wb):
    ws = wb.create_sheet("Vermogensgroei")
    blad_basis(ws, "BF8F00")
    zet_kolombreedtes(ws, {"A": 2, "B": 40, "C": 16, "D": 46, "E": 2,
                           "F": 8, "G": 15, "H": 15, "I": 15, "J": 15, "K": 15, "L": 15})

    r = titelblok(ws, "Vermogensgroei", "Projectie van je vermogen met samengestelde rente", "D")

    r = sectiebalk(ws, r, "JOUW INVOER")
    start_rij = r
    r = invoerregel(ws, r, "Huidig spaargeld en beleggingen", 25000,
                    notitie="Je startkapitaal")
    maandinleg_rij = r
    r = invoerregel(ws, r, "Maandelijkse inleg", 500,
                    notitie="Bijvoorbeeld je maandelijkse overschot uit tabblad 2")
    rendement_rij = r
    r = invoerregel(ws, r, "Verwacht rendement per jaar", 0.06, notatie=FMT_PCT,
                    notitie="Spaarrekening ca. 1,5-2%, wereldwijd aandelenfonds historisch ca. 7%")
    jaren_rij = r
    r = invoerregel(ws, r, "Aantal jaren", 25, notatie=FMT_JAAR,
                    notitie=f"Maximaal {MAX_JAREN} jaar")
    inflatie_rij = r
    r = invoerregel(ws, r, "Verwachte inflatie per jaar", 0.02, notatie=FMT_PCT,
                    notitie="Voor de koopkracht-kolom in de tabel")
    kosten_rij = r
    r = invoerregel(ws, r, "Beleggingskosten per jaar", 0.0035, notatie=FMT_PCT2,
                    notitie="Fondskosten plus servicekosten van je broker")
    r += 1

    jaren_val = DataValidation(type="whole", operator="between",
                               formula1="0", formula2=str(MAX_JAREN), allow_blank=False)
    jaren_val.error = f"Vul een aantal jaren tussen 0 en {MAX_JAREN} in."
    jaren_val.errorTitle = "Ongeldige looptijd"
    ws.add_data_validation(jaren_val)
    jaren_val.add(ws.cell(row=jaren_rij, column=3))

    r = sectiebalk(ws, r, "AFGELEIDE WAARDEN")
    netto_rendement_rij = r
    r = berekendregel(ws, r, "Netto rendement na kosten",
                      f"=C{rendement_rij}-C{kosten_rij}", notatie=FMT_PCT2,
                      notitie="Wordt in de projectie gebruikt")
    maandrente_rij = r
    r = berekendregel(ws, r, "Maandelijks rendement (samengesteld)",
                      f"=(1+C{netto_rendement_rij})^(1/12)-1", notatie='0.000%')
    r += 1
    eind_rij_verwijzing = r  # samenvattingsblok, wordt na de tabel gevuld
    r += 7

    # --- Projectietabel ---
    tabel_kop = r
    koppen = ["Jaar", "Beginwaarde", "Inleg dit jaar", "Rendement",
              "Eindwaarde", "Totaal ingelegd", "Waarde in euro's van nu"]
    for i, tekst in enumerate(koppen):
        cel = ws.cell(row=tabel_kop, column=6 + i, value=tekst)
        cel.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_WIT)
        cel.fill = FILL_SECTIE
        cel.border = RAND_RESULTAAT
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[tabel_kop].height = 30

    tabel_start = tabel_kop + 1
    for jaar in range(0, MAX_JAREN + 1):
        rij = tabel_start + jaar
        toon = f"C${jaren_rij}>={jaar}"

        ws.cell(row=rij, column=6, value=f'=IF({toon},{jaar},NA())').number_format = FMT_GEHEEL
        if jaar == 0:
            begin = f'=IF({toon},C${start_rij},NA())'
            inleg = f'=IF({toon},0,NA())'
            rendement = f'=IF({toon},0,NA())'
            eind = f'=IF({toon},C${start_rij},NA())'
            ingelegd = f'=IF({toon},C${start_rij},NA())'
        else:
            vorige = rij - 1
            begin = f'=IF({toon},J{vorige},NA())'
            inleg = f'=IF({toon},C${maandinleg_rij}*12,NA())'
            # Eindwaarde: beginwaarde groeit een jaar door, inleg maandelijks bijgestort.
            eind = (f'=IF({toon},G{rij}*(1+C${netto_rendement_rij})'
                    f'+IF(C${maandrente_rij}=0,C${maandinleg_rij}*12,'
                    f'C${maandinleg_rij}*((1+C${maandrente_rij})^12-1)/C${maandrente_rij}),NA())')
            rendement = f'=IF({toon},J{rij}-G{rij}-H{rij},NA())'
            ingelegd = f'=IF({toon},K{vorige}+H{rij},NA())'
        reeel = f'=IF({toon},J{rij}/(1+C${inflatie_rij})^{jaar},NA())'

        waarden = [(7, begin), (8, inleg), (9, rendement), (10, eind), (11, ingelegd), (12, reeel)]
        for kol, formule in waarden:
            cel = ws.cell(row=rij, column=kol, value=formule)
            cel.number_format = FMT_EURO_ROND
            cel.fill = FILL_BEREKEND
            cel.border = RAND_BEREKEND
            cel.font = Font(name=FONT_NAME, size=10)
            cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        jaarcel = ws.cell(row=rij, column=6)
        jaarcel.fill = FILL_BEREKEND
        jaarcel.border = RAND_BEREKEND
        jaarcel.font = Font(name=FONT_NAME, size=10, bold=True)
        jaarcel.alignment = Alignment(horizontal="center", vertical="center")

    tabel_eind = tabel_start + MAX_JAREN

    # Rijen buiten de gekozen looptijd tonen #N/A; die verbergen we met witte tekst
    # (de grafiek laat ze dan als onderbreking weg in plaats van als nul).
    ws.conditional_formatting.add(
        f"F{tabel_start}:L{tabel_eind}",
        FormulaRule(formula=[f"ISNA(F{tabel_start})"],
                    font=Font(name=FONT_NAME, size=10, color=CLR_WIT),
                    fill=PatternFill("solid", fgColor=CLR_WIT)))

    # --- Samenvatting (verwijst naar de laatste zichtbare rij) ---
    rr = eind_rij_verwijzing
    rr = sectiebalk(ws, rr, "RESULTAAT NA DE GEKOZEN LOOPTIJD")
    eindwaarde_rij = rr
    rr = berekendregel(ws, rr, "Geprojecteerde eindwaarde",
                       f"=INDEX(J{tabel_start}:J{tabel_eind},C{jaren_rij}+1)", accent=True)
    ingelegd_totaal_rij = rr
    rr = berekendregel(ws, rr, "Totaal zelf ingelegd",
                       f"=INDEX(K{tabel_start}:K{tabel_eind},C{jaren_rij}+1)")
    rr = berekendregel(ws, rr, "Waarvan rendement",
                       f"=C{eindwaarde_rij}-C{ingelegd_totaal_rij}",
                       notitie="Het deel dat je geld voor je heeft verdiend")
    rr = berekendregel(ws, rr, "Eindwaarde in euro's van nu",
                       f"=INDEX(L{tabel_start}:L{tabel_eind},C{jaren_rij}+1)",
                       notitie="Gecorrigeerd voor inflatie: dit is je koopkracht")

    # --- Lijndiagram ---
    grafiek = LineChart()
    grafiek.title = "Groei van je vermogen"
    grafiek.style = 12
    grafiek.y_axis.title = "Waarde (€)"
    grafiek.x_axis.title = "Jaren"
    grafiek.height = 11
    grafiek.width = 24
    gegevens = Reference(ws, min_col=10, max_col=12, min_row=tabel_kop, max_row=tabel_eind)
    categorieen = Reference(ws, min_col=6, min_row=tabel_start, max_row=tabel_eind)
    grafiek.add_data(gegevens, titles_from_data=True)
    grafiek.set_categories(categorieen)
    grafiek.dispBlanksAs = "gap"
    for serie in grafiek.series:
        serie.smooth = False
        serie.marker.symbol = "none"
    ws.add_chart(grafiek, f"N{tabel_kop}")

    laatste = tabel_eind + 2
    notitieregel(ws, laatste, "De projectie rekent met maandelijkse bijstortingen en samengestelde groei "
                              "op basis van het rendement ná kosten.", "D")
    notitieregel(ws, laatste + 1, DISCLAIMER + ". Rendementen uit het verleden bieden geen garantie voor de toekomst.", "D")

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# Hoofdprogramma
# ---------------------------------------------------------------------------

def bouw_werkmap():
    wb = Workbook()
    wb.remove(wb.active)

    bouw_instructies(wb)
    bouw_cashflow(wb)
    bouw_zzp(wb)
    bouw_pensioen(wb)
    bouw_vermogen(wb)

    wb.properties.title = "Geldplanner — persoonlijke financien"
    wb.properties.subject = "Cashflow, ZZP-belasting, pensioengat en vermogensgroei"
    wb.properties.creator = "Geldplanner"
    wb.properties.description = DISCLAIMER + "."
    wb.properties.language = "nl-NL"
    wb.active = 0
    return wb


def main():
    uitvoer = sys.argv[1] if len(sys.argv) > 1 else "Geldplanner.xlsx"
    wb = bouw_werkmap()
    wb.save(uitvoer)
    print(f"Bestand aangemaakt: {uitvoer}")
    print("Tabbladen: " + ", ".join(wb.sheetnames))


if __name__ == "__main__":
    main()
