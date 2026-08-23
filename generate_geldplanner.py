#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Geldplanner NL — generator voor een Excel-sjabloon (.xlsx) voor persoonlijke
financien, gericht op Nederlandse particulieren en stellen.

Gebruik:
    pip install openpyxl
    python generate_geldplanner.py [uitvoerbestand.xlsx]

Het script bouwt een werkmap met vier tabbladen, in professionele
dashboard-stijl (KPI-cards, donut- en staafdiagrammen):
    1. Overzicht Cashflow       — invoer + KPI-cards
    2. Statistieken Cashflow    — donut + staafdiagram + vergelijkings-KPI
    3. Overzicht Budgetruimte   — invoer + KPI-cards
    4. Statistieken Budgetruimte — donut + staafdiagram + KPI

Alle bedragen worden als Excel-formules weggeschreven (geen voorberekende
waarden). Tabblad 2 en 4 hebben geen eigen invoer: ze lezen uitsluitend uit
tabblad 1 resp. 3. Tabblad 3 (Budgetruimte) leest op zijn beurt het
maandelijkse overschot uit tabblad 1.

DISCLAIMER: uitsluitend informatief, geen financieel advies.
"""

import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# ---------------------------------------------------------------------------
# Huisstijl / opmaak
# ---------------------------------------------------------------------------

FONT_NAME = "Calibri"

CLR_HOOFD = "1F3864"          # donkerblauw — hoofdkleur voor koppen en KPI-balken
CLR_HOOFD_LICHT = "2E5C8A"    # secundaire tint voor sectiekoppen binnen een tabblad
CLR_KAART_BG = "F4F6FA"
CLR_KAART_RAND = "C9D2E3"
CLR_GEEL = "FFF2CC"
CLR_GEEL_RAND = "BF8F00"
CLR_GRIJS = "EDEDED"
CLR_GRIJS_RAND = "BFBFBF"
CLR_GROEN = "2E7D46"
CLR_GROEN_BG = "E5F3E8"
CLR_ROOD = "C0392B"
CLR_ROOD_BG = "FBE2E1"
CLR_ORANJE = "B15C00"
CLR_ORANJE_BG = "FDECD2"
CLR_WIT = "FFFFFF"

# Categorische paletkleuren voor donut-/staafdiagram-categorieën. Bewust geen
# puur rood of groen: die kleuren zijn gereserveerd voor goed/fout-signalen.
PALET = ["1F3864", "2E5C8A", "4F81BD", "8DB4E2", "BF8F00", "D9A441", "6B7A99"]

F_TITEL = Font(name=FONT_NAME, size=18, bold=True, color=CLR_WIT)
F_SUBTITEL = Font(name=FONT_NAME, size=10, italic=True, color=CLR_WIT)
F_SECTIE = Font(name=FONT_NAME, size=11, bold=True, color=CLR_WIT)
F_KOP = Font(name=FONT_NAME, size=11, bold=True, color=CLR_HOOFD)
F_LABEL = Font(name=FONT_NAME, size=11)
F_LABEL_VET = Font(name=FONT_NAME, size=11, bold=True)
F_NOTITIE = Font(name=FONT_NAME, size=9, italic=True, color="7F7F7F")
F_RESULTAAT = Font(name=FONT_NAME, size=12, bold=True, color=CLR_HOOFD)
F_KPI_GROOT = Font(name=FONT_NAME, size=22, bold=True, color=CLR_HOOFD)
F_TEKST = Font(name=FONT_NAME, size=11)
F_TEKST_VET = Font(name=FONT_NAME, size=11, bold=True, color=CLR_HOOFD)

FILL_TITEL = PatternFill("solid", fgColor=CLR_HOOFD)
FILL_SECTIE = PatternFill("solid", fgColor=CLR_HOOFD_LICHT)
FILL_INVOER = PatternFill("solid", fgColor=CLR_GEEL)
FILL_BEREKEND = PatternFill("solid", fgColor=CLR_GRIJS)
FILL_RESULTAAT = PatternFill("solid", fgColor=CLR_KAART_BG)
FILL_KAART = PatternFill("solid", fgColor=CLR_KAART_BG)
FILL_GROEN = PatternFill("solid", fgColor=CLR_GROEN_BG)
FILL_ORANJE = PatternFill("solid", fgColor=CLR_ORANJE_BG)
FILL_ROOD = PatternFill("solid", fgColor=CLR_ROOD_BG)

_dun_geel = Side(style="thin", color=CLR_GEEL_RAND)
_dun_grijs = Side(style="thin", color=CLR_GRIJS_RAND)
_dun_kaart = Side(style="thin", color=CLR_KAART_RAND)

RAND_INVOER = Border(left=_dun_geel, right=_dun_geel, top=_dun_geel, bottom=_dun_geel)
RAND_BEREKEND = Border(left=_dun_grijs, right=_dun_grijs, top=_dun_grijs, bottom=_dun_grijs)
RAND_RESULTAAT = Border(left=_dun_kaart, right=_dun_kaart, top=_dun_kaart, bottom=_dun_kaart)

# Getalnotaties. openpyxl schrijft opmaakcodes in en-US-conventie weg; Excel
# toont ze in de landinstelling van de gebruiker, dus op een Nederlandse
# installatie verschijnt "€ 1.234,56" met de komma als decimaalteken.
FMT_EURO = '€ #,##0.00;[Red]-€ #,##0.00'
FMT_EURO_ROND = '€ #,##0;[Red]-€ #,##0'
FMT_PCT = '0.0%'

BESCHERMD = Protection(locked=True)
ONBESCHERMD = Protection(locked=False)

DISCLAIMER = ("Deze tool is uitsluitend bedoeld voor informatieve doeleinden "
              "en biedt geen financieel advies")

MAANDEN = ["jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"]


# ---------------------------------------------------------------------------
# Kleine hulpfuncties voor het opbouwen van de bladen
# ---------------------------------------------------------------------------

def zet_kolombreedtes(ws, breedtes):
    for kolom, breedte in breedtes.items():
        ws.column_dimensions[kolom].width = breedte


def titelblok(ws, titel, ondertitel, laatste_kolom="L"):
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
    cel = ws.cell(row=rij, column=2, value=f"■  {tekst}")
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
        nt.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[rij].height = 17
    return rij + 1


def berekendregel(ws, rij, label, formule, notatie=FMT_EURO, notitie="",
                  accent=False, groot=False, label_kolom=2, waarde_kolom=3, notitie_kolom=4):
    """Een grijze, vergrendelde uitvoercel met formule. groot=True markeert de
    hoofduitkomst van het tabblad: groot, vet, lichte kaart eromheen."""
    lbl = ws.cell(row=rij, column=label_kolom, value=label)
    lbl.font = F_RESULTAAT if (accent or groot) else F_LABEL_VET
    lbl.alignment = Alignment(vertical="center", indent=1, wrap_text=True)

    cel = ws.cell(row=rij, column=waarde_kolom, value=formule)
    cel.fill = FILL_RESULTAAT if (accent or groot) else FILL_BEREKEND
    cel.border = RAND_RESULTAAT if (accent or groot) else RAND_BEREKEND
    cel.font = F_KPI_GROOT if groot else (F_RESULTAAT if accent else F_LABEL_VET)
    cel.number_format = notatie
    cel.protection = BESCHERMD
    cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)

    if notitie:
        nt = ws.cell(row=rij, column=notitie_kolom, value=notitie)
        nt.font = F_NOTITIE
        nt.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[rij].height = 30 if groot else (20 if accent else 17)
    return rij + 1


def notitieregel(ws, rij, tekst="", laatste_kolom="D", vet=False):
    ws.merge_cells(f"B{rij}:{laatste_kolom}{rij}")
    cel = ws.cell(row=rij, column=2, value=tekst)
    cel.font = F_TEKST_VET if vet else F_NOTITIE
    cel.alignment = Alignment(vertical="center", indent=1, wrap_text=False)
    return rij + 1


def kader_vulling(ws, top, bottom, kolommen, fill):
    for rij in range(top, bottom + 1):
        for kol in kolommen:
            ws[f"{kol}{rij}"].fill = fill


def kader_rand(ws, top, bottom, kol_links, kol_rechts, kleur, dikte="thin"):
    """Tekent een dunne rand rond een rechthoekig blok cellen — een 'kaart' of
    'paneel' om een samenhangende sectie heen, zonder bestaande randen te wissen."""
    zijde = Side(style=dikte, color=kleur)
    links_idx = column_index_from_string(kol_links)
    rechts_idx = column_index_from_string(kol_rechts)
    for rij in range(top, bottom + 1):
        for kol in range(links_idx, rechts_idx + 1):
            boven = zijde if rij == top else None
            onder = zijde if rij == bottom else None
            links = zijde if kol == links_idx else None
            rechts = zijde if kol == rechts_idx else None
            if not (boven or onder or links or rechts):
                continue
            cel = ws.cell(row=rij, column=kol)
            bestaand = cel.border
            cel.border = Border(
                top=boven or bestaand.top,
                bottom=onder or bestaand.bottom,
                left=links or bestaand.left,
                right=rechts or bestaand.right,
            )


def sectie_kaart(ws, top, bottom, laatste_kolom="D"):
    """Lichte achtergrond + subtiele rand om een sectie heen (paneel-effect),
    zonder de vulling van invoer-/uitvoercellen in kolom C te overschrijven."""
    kolommen = [get_column_letter(k) for k in range(2, column_index_from_string(laatste_kolom) + 1)]
    randkolommen = [k for k in kolommen if k != "C"]
    kader_vulling(ws, top, bottom, randkolommen, FILL_KAART)
    kader_rand(ws, top, bottom, "B", laatste_kolom, CLR_KAART_RAND)


def kpi_kaart(ws, top, kol_links, kol_rechts, titel, waarde_formule, notatie,
              subtekst_formule=None, kleur_balk=None):
    """Eén KPI-card: gekleurde kopbalk (witte tekst) + groot bedrag eronder in
    een lichte kaart, met een optionele subtekstregel (bijv. een delta).
    Retourneert (waarde_celadres, subtekst_celadres_of_None, volgende_rij)."""
    kleur_balk = kleur_balk or CLR_HOOFD
    links_idx = column_index_from_string(kol_links)
    rechts_idx = column_index_from_string(kol_rechts)
    kolommen = [get_column_letter(k) for k in range(links_idx, rechts_idx + 1)]

    kop_rij = top
    ws.merge_cells(start_row=kop_rij, start_column=links_idx, end_row=kop_rij, end_column=rechts_idx)
    kop = ws.cell(row=kop_rij, column=links_idx, value=titel.upper())
    kop.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_WIT)
    kop.fill = PatternFill("solid", fgColor=kleur_balk)
    kop.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[kop_rij].height = 22

    waarde_rij = kop_rij + 1
    ws.merge_cells(start_row=waarde_rij, start_column=links_idx, end_row=waarde_rij + 1, end_column=rechts_idx)
    waarde_cel = ws.cell(row=waarde_rij, column=links_idx, value=waarde_formule)
    waarde_cel.number_format = notatie
    waarde_cel.font = F_KPI_GROOT
    waarde_cel.protection = BESCHERMD
    waarde_cel.alignment = Alignment(horizontal="center", vertical="center")
    for rij in range(waarde_rij, waarde_rij + 2):
        ws.row_dimensions[rij].height = 24

    onderkant = waarde_rij + 1
    subtekst_adres = None
    if subtekst_formule is not None:
        sub_rij = onderkant + 1
        ws.merge_cells(start_row=sub_rij, start_column=links_idx, end_row=sub_rij, end_column=rechts_idx)
        sub_cel = ws.cell(row=sub_rij, column=links_idx, value=subtekst_formule)
        sub_cel.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_HOOFD)
        sub_cel.protection = BESCHERMD
        sub_cel.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[sub_rij].height = 18
        onderkant = sub_rij
        subtekst_adres = sub_cel.coordinate

    kader_vulling(ws, waarde_rij, onderkant, kolommen, FILL_KAART)
    kader_rand(ws, kop_rij, onderkant, kol_links, kol_rechts, CLR_KAART_RAND, dikte="medium")
    return waarde_cel.coordinate, subtekst_adres, onderkant + 1


def tabel_kader(ws, top, bottom, kol_links, kol_rechts):
    """Rand + lichte kaart om een gegevenstabel (bijv. grafiekbron) heen, zodat
    er geen kale/ongeformatteerde rijen tussen de secties staan."""
    kader_rand(ws, top, bottom, kol_links, kol_rechts, CLR_KAART_RAND, dikte="medium")


def beveilig(ws):
    """Werkbladbeveiliging aan (zonder wachtwoord). Alleen gele invoercellen
    (Protection(locked=False)) blijven bewerkbaar."""
    ws.protection.sheet = True
    ws.protection.enable()


def blad_basis(ws, tabkleur=CLR_HOOFD):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tabkleur
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def kleur_donut_slices(serie, kleuren):
    serie.data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=k)) for i, k in enumerate(kleuren)
    ]


# ---------------------------------------------------------------------------
# TABBLAD 1 — Overzicht Cashflow
# ---------------------------------------------------------------------------

CF_KOLOMBREEDTES = {"A": 2, "B": 34, "C": 16, "D": 38, "E": 15, "F": 15,
                    "G": 4, "H": 15, "I": 15, "J": 4, "K": 15, "L": 15, "M": 2}


def bouw_cashflow(wb):
    ws = wb.create_sheet("Overzicht Cashflow")
    blad_basis(ws, CLR_HOOFD)
    zet_kolombreedtes(ws, CF_KOLOMBREEDTES)

    titelblok(ws, "Overzicht Cashflow", "Jouw maandelijkse inkomsten, lasten en overschot", "L")
    kaart_top = 5  # de KPI-cards worden pas na de invoersectie geschreven (zie onderaan)

    # --- Invoerpaneel: begint onder de KPI-cards ---
    r = 9

    r = sectiebalk(ws, r, "INKOMSTEN PER MAAND")
    inkomsten_start = r
    r = invoerregel(ws, r, "Nettosalaris persoon 1", 3600, notitie="Bedrag dat op je rekening staat")
    r = invoerregel(ws, r, "Nettosalaris persoon 2", 0, notitie="Laat op 0 staan als je alleenstaand bent")
    r = invoerregel(ws, r, "Freelance-inkomsten (netto)", 0, notitie="Netto bedrag, na belasting")
    r = invoerregel(ws, r, "Toeslagen (huur-, zorg-, kinderopvang)", 0)
    r = invoerregel(ws, r, "Huurinkomsten / alimentatie", 0)
    r = invoerregel(ws, r, "Overige inkomsten", 0)
    inkomsten_eind = r - 1
    tot_inkomsten_rij = r
    r = berekendregel(ws, r, "Totale inkomsten", f"=SUM(C{inkomsten_start}:C{inkomsten_eind})")
    sectie_kaart(ws, inkomsten_start, tot_inkomsten_rij, "D")
    r += 1

    r = sectiebalk(ws, r, "VASTE LASTEN PER MAAND")
    vast_start = r
    huur_rij = r
    r = invoerregel(ws, r, "Huur of hypotheek", 1250)
    energie_rij = r
    r = invoerregel(ws, r, "Energie, water en warmte", 180)
    zorgverz_rij = r
    r = invoerregel(ws, r, "Zorgverzekering", 290, notitie="Beide partners samen")
    overige_verz_rij = r
    r = invoerregel(ws, r, "Overige verzekeringen", 65, notitie="Inboedel, aansprakelijkheid, opstal, ORV")
    gemeente_rij = r
    r = invoerregel(ws, r, "Gemeentelijke belastingen en waterschap", 60, notitie="Jaarbedrag gedeeld door 12")
    internet_rij = r
    r = invoerregel(ws, r, "Internet, tv en telefoon", 70)
    abonnementen_rij = r
    r = invoerregel(ws, r, "Abonnementen en lidmaatschappen", 45, notitie="Streaming, sportschool, kranten")
    aflossing_rij = r
    r = invoerregel(ws, r, "Aflossing leningen en studieschuld", 0)
    kinderopvang_rij = r
    r = invoerregel(ws, r, "Kinderopvang / school", 0)
    overige_vast_rij = r
    r = invoerregel(ws, r, "Overige vaste lasten", 0)
    vast_eind = r - 1
    tot_vast_rij = r
    r = berekendregel(ws, r, "Totaal vaste lasten", f"=SUM(C{vast_start}:C{vast_eind})")
    sectie_kaart(ws, vast_start, tot_vast_rij, "D")
    r += 1

    r = sectiebalk(ws, r, "VARIABELE KOSTEN PER MAAND")
    var_start = r
    boodschappen_rij = r
    r = invoerregel(ws, r, "Boodschappen", 550)
    vervoer_rij = r
    r = invoerregel(ws, r, "Vervoer (brandstof, OV, onderhoud)", 200)
    vrije_tijd_rij = r
    r = invoerregel(ws, r, "Vrije tijd, uitgaan en horeca", 200)
    kleding_rij = r
    r = invoerregel(ws, r, "Kleding en persoonlijke verzorging", 100)
    zorgkosten_rij = r
    r = invoerregel(ws, r, "Zorgkosten (eigen risico, tandarts)", 40)
    vakantie_rij = r
    r = invoerregel(ws, r, "Vakantie (gereserveerd per maand)", 150)
    cadeaus_rij = r
    r = invoerregel(ws, r, "Cadeaus, goede doelen en overig", 75)
    var_eind = r - 1
    tot_var_rij = r
    r = berekendregel(ws, r, "Totaal variabele kosten", f"=SUM(C{var_start}:C{var_eind})")
    sectie_kaart(ws, var_start, tot_var_rij, "D")
    r += 1

    r = sectiebalk(ws, r, "RESULTAAT")
    resultaat_kop = r - 1
    res_inkomsten = r
    r = berekendregel(ws, r, "Totale inkomsten", f"=C{tot_inkomsten_rij}")
    res_uitgaven = r
    r = berekendregel(ws, r, "Totale uitgaven", f"=C{tot_vast_rij}+C{tot_var_rij}")
    overschot_rij = r
    r = berekendregel(ws, r, "Maandelijks overschot / tekort",
                      f"=C{res_inkomsten}-C{res_uitgaven}", groot=True,
                      notitie="Positief = je houdt over. Negatief = je komt tekort.")
    spaarquote_rij = r
    r = berekendregel(ws, r, "Spaarquote",
                      f'=IF(C{res_inkomsten}=0,"",C{overschot_rij}/C{res_inkomsten})',
                      notatie=FMT_PCT, notitie="Deel van je inkomen dat je overhoudt")
    buffer_bereken_rij = r
    buffer_invoer_rij = r + 2
    r = berekendregel(
        ws, buffer_bereken_rij, "Buffer: hoeveel maanden uitgaven heb je?",
        f'=IF(C{tot_vast_rij}+C{tot_var_rij}=0,"",'
        f'C{buffer_invoer_rij}/(C{tot_vast_rij}+C{tot_var_rij}))',
        notatie='0.0 "maanden"', notitie="Vuistregel: drie tot zes maanden")
    r += 1
    r = invoerregel(ws, buffer_invoer_rij, "Huidig spaargeld (buffer)", 5000,
                    notitie="Wordt gebruikt voor de bufferberekening hierboven")
    resultaat_eind = r
    sectie_kaart(ws, resultaat_kop + 1, resultaat_eind, "D")
    r += 1

    ws.conditional_formatting.add(
        f"C{overschot_rij}",
        FormulaRule(formula=[f"C{overschot_rij}<0"], fill=FILL_ROOD,
                   font=Font(name=FONT_NAME, size=22, bold=True, color=CLR_ROOD)))
    ws.conditional_formatting.add(
        f"C{overschot_rij}",
        FormulaRule(formula=[f"C{overschot_rij}>=0"], fill=FILL_GROEN,
                   font=Font(name=FONT_NAME, size=22, bold=True, color=CLR_GROEN)))

    r = sectiebalk(ws, r, "VERGELIJKING MET VORIGE MAAND")
    vorige_start = r
    vorige_overschot_rij = r
    r = invoerregel(ws, r, "Overschot / tekort vorige maand", 260,
                    notitie="Gebruikt op tabblad 'Statistieken Cashflow' voor de maand-op-maandvergelijking")
    sectie_kaart(ws, vorige_start, r - 1, "D")
    r += 1

    notitieregel(ws, r, DISCLAIMER + ".", "D")

    # --- KPI-cards bovenaan (na het paneel geschreven, want ze verwijzen naar
    # rijen die pas hierboven zijn vastgesteld) ---
    kpi_kaart(ws, kaart_top, "B", "C", "Totale inkomsten", f"=C{res_inkomsten}", FMT_EURO_ROND)
    kpi_kaart(ws, kaart_top, "E", "F", "Totale uitgaven", f"=C{res_uitgaven}", FMT_EURO_ROND)
    w_overschot, _, _ = kpi_kaart(ws, kaart_top, "H", "I", "Overschot / tekort",
                                  f"=C{overschot_rij}", FMT_EURO_ROND)
    kpi_kaart(ws, kaart_top, "K", "L", "Spaarquote", f"=C{spaarquote_rij}", FMT_PCT)

    ws.conditional_formatting.add(
        f"H{kaart_top}:I{kaart_top}",
        FormulaRule(formula=[f"C{overschot_rij}<0"], fill=PatternFill("solid", fgColor=CLR_ROOD)))
    ws.conditional_formatting.add(
        f"H{kaart_top}:I{kaart_top}",
        FormulaRule(formula=[f"C{overschot_rij}>=0"], fill=PatternFill("solid", fgColor=CLR_GROEN)))
    ws.conditional_formatting.add(
        w_overschot,
        FormulaRule(formula=[f"C{overschot_rij}<0"], font=Font(name=FONT_NAME, size=22, bold=True, color=CLR_ROOD)))

    ws.freeze_panes = "A9"
    beveilig(ws)
    return ws, {
        "blad": "Overzicht Cashflow",
        "inkomsten_rij": res_inkomsten,
        "uitgaven_rij": res_uitgaven,
        "overschot_rij": overschot_rij,
        "spaarquote_rij": spaarquote_rij,
        "vorige_overschot_rij": vorige_overschot_rij,
        "categorieen": [
            ("Wonen", [huur_rij, energie_rij, gemeente_rij]),
            ("Verzekeringen", [zorgverz_rij, overige_verz_rij]),
            ("Abonnementen", [internet_rij, abonnementen_rij]),
            ("Schulden & zorgplicht", [aflossing_rij, kinderopvang_rij, overige_vast_rij]),
            ("Boodschappen", [boodschappen_rij]),
            ("Vervoer", [vervoer_rij]),
            ("Vrije tijd & overig", [vrije_tijd_rij, kleding_rij, zorgkosten_rij, vakantie_rij, cadeaus_rij]),
        ],
    }


# ---------------------------------------------------------------------------
# TABBLAD 2 — Statistieken Cashflow
# ---------------------------------------------------------------------------

def bouw_statistieken_cashflow(wb, cashflow):
    ws = wb.create_sheet("Statistieken Cashflow")
    blad_basis(ws, CLR_HOOFD_LICHT)
    zet_kolombreedtes(ws, {"A": 2, "B": 24, "C": 14, "D": 3,
                           "E": 9, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9,
                           "K": 9, "L": 9, "M": 9, "N": 9, "O": 2})
    cf = f"'{cashflow['blad']}'"

    titelblok(ws, "Statistieken Cashflow", "Deze pagina leest automatisch uit 'Overzicht Cashflow' — geen eigen invoer", "N")

    # --- KPI-card: vergelijking met vorige maand ---
    delta = f"({cf}!C{cashflow['overschot_rij']}-{cf}!C{cashflow['vorige_overschot_rij']})"
    vorige_abs = f"ABS({cf}!C{cashflow['vorige_overschot_rij']})"
    subtekst = (
        f'=IF({cf}!C{cashflow["vorige_overschot_rij"]}=0,"Geen vergelijking beschikbaar",'
        f'CONCATENATE(IF({delta}>=0,"▲ +","▼ "),TEXT(ABS({delta}),"€ #,##0"),'
        f'"  (",TEXT({delta}/{vorige_abs},"0.0%"),")"))'
    )
    waarde_adres, sub_adres, r = kpi_kaart(
        ws, 5, "B", "E", "Overschot t.o.v. vorige maand",
        f"={cf}!C{cashflow['overschot_rij']}", FMT_EURO_ROND, subtekst_formule=subtekst)
    for adres in (waarde_adres, sub_adres):
        ws.conditional_formatting.add(
            adres, FormulaRule(formula=[f"{delta}>=0"], fill=FILL_GROEN,
                               font=Font(name=FONT_NAME, size=10, bold=True, color=CLR_GROEN)))
        ws.conditional_formatting.add(
            adres, FormulaRule(formula=[f"{delta}<0"], fill=FILL_ROOD,
                               font=Font(name=FONT_NAME, size=10, bold=True, color=CLR_ROOD)))
    ws.conditional_formatting.add(
        waarde_adres, FormulaRule(formula=[f"{delta}>=0"], font=Font(name=FONT_NAME, size=22, bold=True, color=CLR_GROEN)))
    ws.conditional_formatting.add(
        waarde_adres, FormulaRule(formula=[f"{delta}<0"], font=Font(name=FONT_NAME, size=22, bold=True, color=CLR_ROOD)))
    r += 1

    # --- Donut: uitgaven per categorie ---
    r = sectiebalk(ws, r, "UITGAVEN PER CATEGORIE", "C")
    cat_kop = r
    hk1 = ws.cell(row=r, column=2, value="Categorie")
    hk2 = ws.cell(row=r, column=3, value="Bedrag / mnd")
    for cel in (hk1, hk2):
        cel.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_WIT)
        cel.fill = FILL_SECTIE
        cel.alignment = Alignment(horizontal="center", vertical="center")
    r += 1
    cat_start = r
    for naam, rijen in cashflow["categorieen"]:
        som = "+".join(f"{cf}!C{rij}" for rij in rijen)
        lbl = ws.cell(row=r, column=2, value=naam)
        lbl.font = F_LABEL
        lbl.alignment = Alignment(vertical="center", indent=1)
        cel = ws.cell(row=r, column=3, value=f"={som}")
        cel.number_format = FMT_EURO_ROND
        cel.font = F_LABEL
        cel.fill = FILL_BEREKEND
        cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        r += 1
    cat_eind = r - 1
    kader_vulling(ws, cat_kop, cat_eind, ["B", "C"], FILL_KAART)
    tabel_kader(ws, cat_kop, cat_eind, "B", "C")

    donut1 = DoughnutChart()
    donut1.title = "Uitgaven per categorie"
    donut1.height = 9
    donut1.width = 14
    donut1.holeSize = 55
    gegevens = Reference(ws, min_col=3, min_row=cat_kop, max_row=cat_eind)
    categorieen = Reference(ws, min_col=2, min_row=cat_start, max_row=cat_eind)
    donut1.add_data(gegevens, titles_from_data=True)
    donut1.set_categories(categorieen)
    donut1.dLbls = DataLabelList()
    donut1.dLbls.showPercent = True
    kleur_donut_slices(donut1.series[0], PALET[: cat_eind - cat_start + 1])
    ws.add_chart(donut1, f"E{cat_kop}")

    r = cat_eind + 2

    # --- Staafdiagram: inkomsten vs. uitgaven per maand (prognose) ---
    r = sectiebalk(ws, r, "INKOMSTEN VS. UITGAVEN — PROGNOSE PER MAAND", "D")
    r = notitieregel(ws, r, "Gebaseerd op je huidige maandbudget uit 'Overzicht Cashflow', vlak doorgetrokken over het jaar.", "D")
    maand_kop = r
    koppen = ["Maand", "Inkomsten", "Uitgaven"]
    for i, tekst in enumerate(koppen):
        cel = ws.cell(row=r, column=2 + i, value=tekst)
        cel.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_WIT)
        cel.fill = FILL_SECTIE
        cel.alignment = Alignment(horizontal="center", vertical="center")
    r += 1
    maand_start = r
    for maand in MAANDEN:
        ws.cell(row=r, column=2, value=maand.capitalize()).font = F_LABEL
        ws.cell(row=r, column=2).alignment = Alignment(vertical="center", indent=1)
        c_ink = ws.cell(row=r, column=3, value=f"={cf}!C{cashflow['inkomsten_rij']}")
        c_uit = ws.cell(row=r, column=4, value=f"={cf}!C{cashflow['uitgaven_rij']}")
        for cel in (c_ink, c_uit):
            cel.number_format = FMT_EURO_ROND
            cel.font = F_LABEL
            cel.fill = FILL_BEREKEND
            cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        r += 1
    maand_eind = r - 1
    kader_vulling(ws, maand_kop, maand_eind, ["B", "C", "D"], FILL_KAART)
    tabel_kader(ws, maand_kop, maand_eind, "B", "D")
    r = maand_eind + 1

    bar1 = BarChart()
    bar1.type = "col"
    bar1.grouping = "clustered"
    bar1.style = 10
    bar1.title = "Inkomsten vs. uitgaven per maand"
    bar1.y_axis.title = "Bedrag per maand (€)"
    bar1.height = 9
    bar1.width = 24
    gegevens = Reference(ws, min_col=3, max_col=4, min_row=maand_kop, max_row=maand_eind)
    categorieen = Reference(ws, min_col=2, min_row=maand_start, max_row=maand_eind)
    bar1.add_data(gegevens, titles_from_data=True)
    bar1.set_categories(categorieen)
    bar1.series[0].graphicalProperties.solidFill = CLR_GROEN
    bar1.series[1].graphicalProperties.solidFill = CLR_ROOD
    bar1.gapWidth = 40
    # De donut hierboven is 9cm hoog (~17 rijen); anker deze grafiek daaronder
    # zodat ze elkaar niet overlappen, ook als de brontabel korter is.
    bar1_anker_rij = max(maand_kop, cat_kop + 19)
    ws.add_chart(bar1, f"E{bar1_anker_rij}")

    # Plaats de disclaimer onder de onderkant van de staafdiagram (~17 rijen
    # hoog), zodat de zwevende grafiek de tekst niet overlapt.
    r = max(r, bar1_anker_rij + 19)
    notitieregel(ws, r, DISCLAIMER + ".", "D")

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# TABBLAD 3 — Overzicht Budgetruimte
# ---------------------------------------------------------------------------

def bouw_budgetruimte(wb, cashflow):
    ws = wb.create_sheet("Overzicht Budgetruimte")
    blad_basis(ws, CLR_HOOFD)
    zet_kolombreedtes(ws, CF_KOLOMBREEDTES)
    cf = f"'{cashflow['blad']}'"

    titelblok(ws, "Overzicht Budgetruimte", "Hoeveel van je overschot wil je sparen, investeren of vrij besteden?", "L")
    kaart_top = 5

    r = 9
    r = sectiebalk(ws, r, "BESCHIKBARE BUDGETRUIMTE")
    beschikbaar_start = r
    beschikbaar_rij = r
    r = berekendregel(ws, r, "Beschikbaar overschot uit Overzicht Cashflow",
                      f"={cf}!C{cashflow['overschot_rij']}", accent=True,
                      notitie="Wordt automatisch bijgewerkt zodra je dat tabblad invult")
    sectie_kaart(ws, beschikbaar_start, r - 1, "D")
    r += 1

    r = sectiebalk(ws, r, "JOUW SPAAR- EN INVESTERINGSDOEL")
    doel_start = r
    spaar_rij = r
    r = invoerregel(ws, r, "Gewenst spaarbedrag per maand", 150,
                    notitie="Bijvoorbeeld op een spaarrekening of buffer")
    invest_rij = r
    r = invoerregel(ws, r, "Gewenst investeerbedrag per maand", 100,
                    notitie="Bijvoorbeeld beleggen of extra pensioeninleg")
    sectie_kaart(ws, doel_start, r - 1, "D")
    r += 1

    r = sectiebalk(ws, r, "RESULTAAT")
    resultaat_start = r
    gewenst_rij = r
    r = berekendregel(ws, r, "Totaal gewenst (sparen + investeren)",
                      f"=C{spaar_rij}+C{invest_rij}")
    vrij_rij = r
    r = berekendregel(ws, r, "Vrij besteedbaar na sparen/investeren",
                      f"=C{beschikbaar_rij}-C{gewenst_rij}", groot=True,
                      notitie="Wat overblijft nadat je spaar- en investeringsdoel is afgetrokken")
    waarschuwing_rij = r
    waarschuwing_formule = (
        f'=IF(C{gewenst_rij}>C{beschikbaar_rij},'
        f'"⚠ Let op: je gewenste bedrag is hoger dan je beschikbare overschot deze maand.","")'
    )
    ws.merge_cells(f"B{r}:D{r}")
    w_cel = ws.cell(row=r, column=2, value=waarschuwing_formule)
    w_cel.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_HOOFD)
    w_cel.fill = FILL_KAART
    w_cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    w_cel.protection = BESCHERMD
    ws.row_dimensions[r].height = 20
    ws.conditional_formatting.add(
        f"B{r}", FormulaRule(formula=[f"C{gewenst_rij}>C{beschikbaar_rij}"], fill=FILL_ROOD,
                             font=Font(name=FONT_NAME, size=10, bold=True, color=CLR_ROOD)))
    r += 1
    resultaat_eind = r - 1
    sectie_kaart(ws, resultaat_start, resultaat_eind, "D")
    r += 1

    ws.conditional_formatting.add(
        f"C{vrij_rij}", FormulaRule(formula=[f"C{vrij_rij}<0"], fill=FILL_ROOD,
                                    font=Font(name=FONT_NAME, size=22, bold=True, color=CLR_ROOD)))
    ws.conditional_formatting.add(
        f"C{vrij_rij}", FormulaRule(formula=[f"C{vrij_rij}>=0"], fill=FILL_GROEN,
                                    font=Font(name=FONT_NAME, size=22, bold=True, color=CLR_GROEN)))

    notitieregel(ws, r, DISCLAIMER + ".", "D")

    kpi_kaart(ws, kaart_top, "B", "C", "Beschikbare budgetruimte", f"=C{beschikbaar_rij}", FMT_EURO_ROND)
    kpi_kaart(ws, kaart_top, "E", "F", "Gewenst spaarbedrag", f"=C{spaar_rij}", FMT_EURO_ROND)
    kpi_kaart(ws, kaart_top, "H", "I", "Gewenst investeerbedrag", f"=C{invest_rij}", FMT_EURO_ROND)
    w_vrij, _, _ = kpi_kaart(ws, kaart_top, "K", "L", "Vrij besteedbaar", f"=C{vrij_rij}", FMT_EURO_ROND)
    ws.conditional_formatting.add(
        f"K{kaart_top}:L{kaart_top}",
        FormulaRule(formula=[f"C{vrij_rij}<0"], fill=PatternFill("solid", fgColor=CLR_ROOD)))
    ws.conditional_formatting.add(
        f"K{kaart_top}:L{kaart_top}",
        FormulaRule(formula=[f"C{vrij_rij}>=0"], fill=PatternFill("solid", fgColor=CLR_GROEN)))
    ws.conditional_formatting.add(
        w_vrij, FormulaRule(formula=[f"C{vrij_rij}<0"], font=Font(name=FONT_NAME, size=22, bold=True, color=CLR_ROOD)))

    ws.freeze_panes = "A9"
    beveilig(ws)
    return ws, {
        "blad": "Overzicht Budgetruimte",
        "beschikbaar_rij": beschikbaar_rij,
        "spaar_rij": spaar_rij,
        "invest_rij": invest_rij,
        "vrij_rij": vrij_rij,
    }


# ---------------------------------------------------------------------------
# TABBLAD 4 — Statistieken Budgetruimte
# ---------------------------------------------------------------------------

def bouw_statistieken_budgetruimte(wb, budget):
    ws = wb.create_sheet("Statistieken Budgetruimte")
    blad_basis(ws, CLR_HOOFD_LICHT)
    zet_kolombreedtes(ws, {"A": 2, "B": 24, "C": 14, "D": 3,
                           "E": 9, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9,
                           "K": 9, "L": 9, "M": 9, "N": 9, "O": 2})
    bg = f"'{budget['blad']}'"
    kleuren3 = PALET[:3]

    titelblok(ws, "Statistieken Budgetruimte", "Deze pagina leest automatisch uit 'Overzicht Budgetruimte' — geen eigen invoer", "N")

    waarde_adres, _, r = kpi_kaart(
        ws, 5, "B", "E", "Totaal gespaard + geïnvesteerd (12 mnd)",
        f"=({bg}!C{budget['spaar_rij']}+{bg}!C{budget['invest_rij']})*12", FMT_EURO_ROND)
    r += 1

    r = sectiebalk(ws, r, "VERDELING VAN JE BUDGETRUIMTE", "C")
    cat_kop = r
    hk1 = ws.cell(row=r, column=2, value="Bestemming")
    hk2 = ws.cell(row=r, column=3, value="Bedrag / mnd")
    for cel in (hk1, hk2):
        cel.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_WIT)
        cel.fill = FILL_SECTIE
        cel.alignment = Alignment(horizontal="center", vertical="center")
    r += 1
    cat_start = r
    for naam, formule in [
        ("Sparen", f"={bg}!C{budget['spaar_rij']}"),
        ("Investeren", f"={bg}!C{budget['invest_rij']}"),
        ("Vrij besteedbaar", f"=MAX(0,{bg}!C{budget['vrij_rij']})"),
    ]:
        lbl = ws.cell(row=r, column=2, value=naam)
        lbl.font = F_LABEL
        lbl.alignment = Alignment(vertical="center", indent=1)
        cel = ws.cell(row=r, column=3, value=formule)
        cel.number_format = FMT_EURO_ROND
        cel.font = F_LABEL
        cel.fill = FILL_BEREKEND
        cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        r += 1
    cat_eind = r - 1
    kader_vulling(ws, cat_kop, cat_eind, ["B", "C"], FILL_KAART)
    tabel_kader(ws, cat_kop, cat_eind, "B", "C")

    donut2 = DoughnutChart()
    donut2.title = "Sparen / investeren / vrij besteedbaar"
    donut2.height = 9
    donut2.width = 14
    donut2.holeSize = 55
    gegevens = Reference(ws, min_col=3, min_row=cat_kop, max_row=cat_eind)
    categorieen = Reference(ws, min_col=2, min_row=cat_start, max_row=cat_eind)
    donut2.add_data(gegevens, titles_from_data=True)
    donut2.set_categories(categorieen)
    donut2.dLbls = DataLabelList()
    donut2.dLbls.showPercent = True
    kleur_donut_slices(donut2.series[0], kleuren3)
    ws.add_chart(donut2, f"E{cat_kop}")

    r = cat_eind + 2

    r = sectiebalk(ws, r, "BUDGETRUIMTE PER MAAND — PROGNOSE", "E")
    r = notitieregel(ws, r, "Gebaseerd op je huidige instellingen uit 'Overzicht Budgetruimte', vlak doorgetrokken over het jaar.", "E")
    maand_kop = r
    koppen = ["Maand", "Sparen", "Investeren", "Vrij besteedbaar"]
    for i, tekst in enumerate(koppen):
        cel = ws.cell(row=r, column=2 + i, value=tekst)
        cel.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_WIT)
        cel.fill = FILL_SECTIE
        cel.alignment = Alignment(horizontal="center", vertical="center")
    r += 1
    maand_start = r
    for maand in MAANDEN:
        ws.cell(row=r, column=2, value=maand.capitalize()).font = F_LABEL
        ws.cell(row=r, column=2).alignment = Alignment(vertical="center", indent=1)
        c_sp = ws.cell(row=r, column=3, value=f"={bg}!C{budget['spaar_rij']}")
        c_iv = ws.cell(row=r, column=4, value=f"={bg}!C{budget['invest_rij']}")
        c_vr = ws.cell(row=r, column=5, value=f"=MAX(0,{bg}!C{budget['vrij_rij']})")
        for cel in (c_sp, c_iv, c_vr):
            cel.number_format = FMT_EURO_ROND
            cel.font = F_LABEL
            cel.fill = FILL_BEREKEND
            cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        r += 1
    maand_eind = r - 1
    kader_vulling(ws, maand_kop, maand_eind, ["B", "C", "D", "E"], FILL_KAART)
    tabel_kader(ws, maand_kop, maand_eind, "B", "E")
    r = maand_eind + 1

    bar2 = BarChart()
    bar2.type = "col"
    bar2.grouping = "stacked"
    bar2.overlap = 100
    bar2.style = 10
    bar2.title = "Budgetruimte per maand"
    bar2.y_axis.title = "Bedrag per maand (€)"
    bar2.height = 9
    bar2.width = 24
    gegevens = Reference(ws, min_col=3, max_col=5, min_row=maand_kop, max_row=maand_eind)
    categorieen = Reference(ws, min_col=2, min_row=maand_start, max_row=maand_eind)
    bar2.add_data(gegevens, titles_from_data=True)
    bar2.set_categories(categorieen)
    for i, serie in enumerate(bar2.series):
        serie.graphicalProperties.solidFill = kleuren3[i]
    bar2.gapWidth = 40
    # De donut hierboven is 9cm hoog (~17 rijen); anker deze grafiek daaronder
    # zodat ze elkaar niet overlappen, ook als de brontabel korter is.
    bar2_anker_rij = max(maand_kop, cat_kop + 19)
    ws.add_chart(bar2, f"F{bar2_anker_rij}")

    # Plaats de disclaimer onder de onderkant van de staafdiagram (~17 rijen
    # hoog), zodat de zwevende grafiek de tekst niet overlapt.
    r = max(r, bar2_anker_rij + 19)
    notitieregel(ws, r, DISCLAIMER + ".", "E")

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# Hoofdprogramma
# ---------------------------------------------------------------------------

def bouw_werkmap():
    wb = Workbook()
    wb.remove(wb.active)

    _, cashflow_ref = bouw_cashflow(wb)
    bouw_statistieken_cashflow(wb, cashflow_ref)
    _, budget_ref = bouw_budgetruimte(wb, cashflow_ref)
    bouw_statistieken_budgetruimte(wb, budget_ref)

    wb.properties.title = "Geldplanner — cashflow en budgetruimte"
    wb.properties.subject = "Cashflow-dashboard en budgetruimte-planner"
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
