#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Geldplanner NL — generator voor een Excel-sjabloon (.xlsx) voor persoonlijke
financien, gericht op Nederlandse particulieren en stellen.

Gebruik:
    pip install openpyxl
    python generate_geldplanner.py [uitvoerbestand.xlsx]

Het script bouwt een werkmap met zes tabbladen:
    1. Titelblad
    2. Dashboard
    3. Instructies
    4. Cashflow Overzicht
    5. Pensioengat Calculator
    6. Vermogensgroei

Alle bedragen worden als Excel-formules weggeschreven (geen voorberekende
waarden), zodat de formules in Excel zelf zichtbaar en aanpasbaar zijn. Het
Dashboard en de koppeling tussen Cashflow, Pensioengat en Vermogensgroei
verwijzen naar cellen op andere tabbladen, dus die werken automatisch bij
zodra de gebruiker de brontabbladen invult.

DISCLAIMER: uitsluitend informatief, geen financieel advies.
"""

import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

# ---------------------------------------------------------------------------
# Huisstijl / opmaak
# ---------------------------------------------------------------------------

FONT_NAME = "Calibri"

# Hoofdkleur (donkerblauw) + accentkleur (goud), consistent op elk tabblad.
CLR_DONKERBLAUW = "132A46"
CLR_BLAUW = "2E5C8A"
CLR_LICHTBLAUW = "D9E2F3"
CLR_ACCENT = "C9A227"
CLR_ACCENT_DONKER = "8C6D0F"
CLR_ACCENT_LICHT = "F7EFD9"
CLR_KAART_BG = "F3F5FA"
CLR_KAART_RAND = "C9D2E3"
CLR_GEEL = "FFF2CC"
CLR_GEEL_RAND = "BF8F00"
CLR_GRIJS = "EDEDED"
CLR_GRIJS_RAND = "BFBFBF"
CLR_GROEN_BG = "E5F3E8"
CLR_GROEN_TEKST = "1E7145"
CLR_GROEN_RAND = "7FC49A"
CLR_ORANJE_BG = "FDECD2"
CLR_ORANJE_TEKST = "B15C00"
CLR_ORANJE_RAND = "F0B462"
CLR_ROOD_BG = "FCE4E4"
CLR_ROOD_TEKST = "9C0006"
CLR_ROOD_RAND = "E39494"
CLR_WIT = "FFFFFF"

F_TITEL = Font(name=FONT_NAME, size=18, bold=True, color=CLR_WIT)
F_SUBTITEL = Font(name=FONT_NAME, size=10, italic=True, color=CLR_WIT)
F_SECTIE = Font(name=FONT_NAME, size=11, bold=True, color=CLR_WIT)
F_SECTIE_ACCENT = Font(name=FONT_NAME, size=11, bold=True, color=CLR_DONKERBLAUW)
F_KOP = Font(name=FONT_NAME, size=11, bold=True, color=CLR_DONKERBLAUW)
F_LABEL = Font(name=FONT_NAME, size=11)
F_LABEL_VET = Font(name=FONT_NAME, size=11, bold=True)
F_NOTITIE = Font(name=FONT_NAME, size=9, italic=True, color="7F7F7F")
F_RESULTAAT = Font(name=FONT_NAME, size=12, bold=True, color=CLR_DONKERBLAUW)
F_RESULTAAT_LABEL = Font(name=FONT_NAME, size=13, bold=True, color=CLR_DONKERBLAUW)
F_KPI_GROOT = Font(name=FONT_NAME, size=20, bold=True, color=CLR_DONKERBLAUW)
F_KPI_MEGA = Font(name=FONT_NAME, size=26, bold=True, color=CLR_DONKERBLAUW)
F_BADGE = Font(name=FONT_NAME, size=11, bold=True, color=CLR_DONKERBLAUW)
F_TEKST = Font(name=FONT_NAME, size=11)
F_TEKST_VET = Font(name=FONT_NAME, size=11, bold=True, color=CLR_DONKERBLAUW)

FILL_TITEL = PatternFill("solid", fgColor=CLR_DONKERBLAUW)
FILL_SECTIE = PatternFill("solid", fgColor=CLR_BLAUW)
FILL_SECTIE_ACCENT = PatternFill("solid", fgColor=CLR_ACCENT)
FILL_INVOER = PatternFill("solid", fgColor=CLR_GEEL)
FILL_BEREKEND = PatternFill("solid", fgColor=CLR_GRIJS)
FILL_RESULTAAT = PatternFill("solid", fgColor=CLR_LICHTBLAUW)
FILL_KPI = PatternFill("solid", fgColor=CLR_ACCENT_LICHT)
FILL_KAART = PatternFill("solid", fgColor=CLR_KAART_BG)
FILL_GROEN = PatternFill("solid", fgColor=CLR_GROEN_BG)
FILL_ORANJE = PatternFill("solid", fgColor=CLR_ORANJE_BG)
FILL_ROOD = PatternFill("solid", fgColor=CLR_ROOD_BG)

_dun_geel = Side(style="thin", color=CLR_GEEL_RAND)
_dun_grijs = Side(style="thin", color=CLR_GRIJS_RAND)
_dun_blauw = Side(style="thin", color=CLR_BLAUW)
_dun_accent = Side(style="thin", color=CLR_ACCENT_DONKER)

RAND_INVOER = Border(left=_dun_geel, right=_dun_geel, top=_dun_geel, bottom=_dun_geel)
RAND_BEREKEND = Border(left=_dun_grijs, right=_dun_grijs, top=_dun_grijs, bottom=_dun_grijs)
RAND_RESULTAAT = Border(left=_dun_blauw, right=_dun_blauw, top=_dun_blauw, bottom=_dun_blauw)
RAND_ACCENT = Border(left=_dun_accent, right=_dun_accent, top=_dun_accent, bottom=_dun_accent)

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

DISCLAIMER = ("Deze tool is uitsluitend bedoeld voor informatieve doeleinden "
              "en biedt geen financieel advies")


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


def sectiebalk(ws, rij, tekst="", laatste_kolom="D", accent=False):
    """Gekleurde sectiekop. accent=True gebruikt de goudkleur (koppel-secties)."""
    ws.merge_cells(f"B{rij}:{laatste_kolom}{rij}")
    cel = ws.cell(row=rij, column=2)
    cel.value = f"◆  {tekst}"
    cel.font = F_SECTIE_ACCENT if accent else F_SECTIE
    cel.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    vulling = FILL_SECTIE_ACCENT if accent else FILL_SECTIE
    for kol in range(2, ws[f"{laatste_kolom}1"].column + 1):
        ws.cell(row=rij, column=kol).fill = vulling
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
    """Een grijze, vergrendelde uitvoercel met formule.

    accent=True markeert een sectie-eindresultaat (lichtblauw).
    groot=True markeert de belangrijkste uitkomst van het hele tabblad: groot,
    vet en met een goudkleurige kaart eromheen.
    """
    lbl = ws.cell(row=rij, column=label_kolom, value=label)
    lbl.font = F_RESULTAAT_LABEL if groot else (F_RESULTAAT if accent else F_LABEL_VET)
    lbl.alignment = Alignment(vertical="center", indent=1, wrap_text=True)

    cel = ws.cell(row=rij, column=waarde_kolom, value=formule)
    if groot:
        cel.fill = FILL_KPI
        cel.border = RAND_ACCENT
        cel.font = F_KPI_GROOT
    else:
        cel.fill = FILL_RESULTAAT if accent else FILL_BEREKEND
        cel.border = RAND_RESULTAAT if accent else RAND_BEREKEND
        cel.font = F_RESULTAAT if accent else F_LABEL_VET
    cel.number_format = notatie
    cel.protection = BESCHERMD
    cel.alignment = Alignment(horizontal="right", vertical="center", indent=1)

    if notitie:
        nt = ws.cell(row=rij, column=notitie_kolom, value=notitie)
        nt.font = F_NOTITIE
        nt.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[rij].height = 34 if groot else (20 if accent else 17)
    return rij + 1


def notitieregel(ws, rij, tekst="", laatste_kolom="D", vet=False):
    ws.merge_cells(f"B{rij}:{laatste_kolom}{rij}")
    cel = ws.cell(row=rij, column=2, value=tekst)
    cel.font = F_TEKST_VET if vet else F_NOTITIE
    cel.alignment = Alignment(vertical="center", indent=1, wrap_text=False)
    return rij + 1


def kader_vulling(ws, top, bottom, kolommen, fill):
    """Zet een lichte achtergrondkleur op specifieke kolommen (bijv. label en
    toelichting), zonder de vulling van invoer-/uitvoercellen te overschrijven."""
    for rij in range(top, bottom + 1):
        for kol in kolommen:
            ws[f"{kol}{rij}"].fill = fill


def kader_rand(ws, top, bottom, kol_links, kol_rechts, kleur, dikte="thin"):
    """Tekent een dunne rand rond een rechthoekig blok cellen — een 'kaart'
    om een samenhangende sectie heen, zonder bestaande randen te wissen."""
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
    """Combineert kader_vulling + kader_rand tot de standaard 'kaart'-stijl die
    onder elke sectiebalk wordt gebruikt: lichte achtergrond + subtiele rand."""
    kolommen = [get_column_letter(k) for k in range(2, column_index_from_string(laatste_kolom) + 1)]
    randkolommen = [k for k in kolommen if k not in ("C",)]
    kader_vulling(ws, top, bottom, randkolommen, FILL_KAART)
    kader_rand(ws, top, bottom, "B", laatste_kolom, CLR_KAART_RAND)


def ja_nee_validatie(ws, rijen, kolom=3):
    dv = DataValidation(type="list", formula1='"Ja,Nee"', allow_blank=False, showDropDown=False)
    dv.error = "Kies Ja of Nee."
    dv.errorTitle = "Ongeldige invoer"
    ws.add_data_validation(dv)
    for rij in rijen:
        dv.add(ws.cell(row=rij, column=kolom))
    return dv


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


# ---------------------------------------------------------------------------
# TABBLAD 1 — Titelblad
# ---------------------------------------------------------------------------

def bouw_titelblad(wb):
    ws = wb.create_sheet("Titelblad")
    blad_basis(ws, CLR_DONKERBLAUW)
    zet_kolombreedtes(ws, {"A": 3, "B": 8, "C": 16, "D": 16, "E": 16, "F": 16, "G": 8, "H": 3})

    for rij in range(1, 3):
        for kol in range(1, 9):
            ws.cell(row=rij, column=kol).fill = FILL_TITEL
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 8

    # --- Logo-plek ---
    ws.merge_cells("C4:F11")
    logo = ws["C4"]
    logo.value = "[ Plaats hier je logo ]"
    logo.font = Font(name=FONT_NAME, size=11, italic=True, color="9AA5B1")
    logo.fill = PatternFill("solid", fgColor="F5F6F8")
    logo.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    stip = Side(style="dashed", color="B7C0CC")
    for rij in range(4, 12):
        for kol in range(3, 7):
            cel = ws.cell(row=rij, column=kol)
            cel.fill = PatternFill("solid", fgColor="F5F6F8")
            cel.border = Border(
                top=stip if rij == 4 else None,
                bottom=stip if rij == 11 else None,
                left=stip if kol == 3 else None,
                right=stip if kol == 6 else None,
            )
        ws.row_dimensions[rij].height = 18

    r = 13
    ws.merge_cells(f"B{r}:G{r}")
    titel = ws.cell(row=r, column=2, value="GELDPLANNER")
    titel.font = Font(name=FONT_NAME, size=40, bold=True, color=CLR_DONKERBLAUW)
    titel.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 56
    r += 1

    ws.merge_cells(f"B{r}:G{r}")
    tagline = ws.cell(row=r, column=2,
                      value="Grip op je geld — cashflow, pensioen en vermogensopbouw in één overzicht")
    tagline.font = Font(name=FONT_NAME, size=13, italic=True, color=CLR_ACCENT_DONKER)
    tagline.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 22
    r += 2

    ws.merge_cells(f"C{r}:F{r}")
    for kol in range(3, 7):
        ws.cell(row=r, column=kol).fill = PatternFill("solid", fgColor=CLR_ACCENT)
    ws.row_dimensions[r].height = 4
    r += 2

    ws.merge_cells(f"B{r}:G{r + 3}")
    intro = ws.cell(row=r, column=2)
    intro.value = (
        "Welkom bij Geldplanner. Deze werkmap geeft je in een paar minuten overzicht over je "
        "maandelijkse cashflow, je pensioengat en de groei van je vermogen — met heldere, "
        "aanpasbare formules in plaats van een kant-en-klaar zwart doosje.\n\n"
        "Vul de gele cellen op de tabbladen in; het Dashboard werkt automatisch bij."
    )
    intro.font = Font(name=FONT_NAME, size=11, color="404040")
    intro.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    for rij in range(r, r + 4):
        ws.row_dimensions[rij].height = 20
    r += 5

    ws.merge_cells(f"B{r}:G{r}")
    kop = ws.cell(row=r, column=2, value="WAT ZIT ERIN?")
    kop.font = Font(name=FONT_NAME, size=11, bold=True, color=CLR_ACCENT_DONKER)
    kop.alignment = Alignment(horizontal="center", vertical="center")
    r += 1

    onderdelen = [
        "◆  Dashboard — je financiële situatie in één oogopslag",
        "◆  Cashflow Overzicht — inkomsten, lasten en je maandelijkse overschot",
        "◆  Pensioengat Calculator — wat je AOW en pensioen missen, en wat je daarvoor opzij zet",
        "◆  Vermogensgroei — de projectie van je spaargeld en beleggingen",
    ]
    for tekst in onderdelen:
        ws.merge_cells(f"C{r}:F{r}")
        cel = ws.cell(row=r, column=3, value=tekst)
        cel.font = Font(name=FONT_NAME, size=11, color="404040")
        cel.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r].height = 19
        r += 1
    r += 1

    ws.merge_cells(f"B{r}:G{r}")
    cta = ws.cell(row=r, column=2, value="→  Ga naar het tabblad 'Dashboard' om te beginnen")
    cta.font = Font(name=FONT_NAME, size=12, bold=True, color=CLR_DONKERBLAUW)
    cta.fill = PatternFill("solid", fgColor=CLR_ACCENT_LICHT)
    cta.alignment = Alignment(horizontal="center", vertical="center")
    kader_rand(ws, r, r, "B", "G", CLR_ACCENT, dikte="medium")
    ws.row_dimensions[r].height = 26
    r += 2

    ws.merge_cells(f"B{r}:G{r}")
    footer = ws.cell(row=r, column=2,
                     value=f"Geldplanner · Versie 1.0  ·  {DISCLAIMER}. Zie het tabblad 'Instructies' voor de volledige toelichting.")
    footer.font = F_NOTITIE
    footer.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 26

    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# TABBLAD 2 — Dashboard
# ---------------------------------------------------------------------------

def dashboard_kaart(ws, top, kol_links, kol_rechts, icoon_label_formule):
    """Bouwt de opmaak van één KPI-kaart (kop, grote waarde, badge, rand) en
    geeft de celadressen terug die de aanroeper met formules moet vullen."""
    links_idx = column_index_from_string(kol_links)
    rechts_idx = column_index_from_string(kol_rechts)
    kolommen = [get_column_letter(k) for k in range(links_idx, rechts_idx + 1)]

    kop_rij = top
    ws.merge_cells(start_row=kop_rij, start_column=links_idx, end_row=kop_rij, end_column=rechts_idx)
    kop = ws.cell(row=kop_rij, column=links_idx, value=icoon_label_formule)
    kop.font = Font(name=FONT_NAME, size=10, bold=True, color=CLR_WIT)
    kop.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[kop_rij].height = 22

    waarde_rij = kop_rij + 1
    ws.merge_cells(start_row=waarde_rij, start_column=links_idx, end_row=waarde_rij + 2, end_column=rechts_idx)
    waarde_cel = ws.cell(row=waarde_rij, column=links_idx)
    waarde_cel.font = F_KPI_MEGA
    waarde_cel.alignment = Alignment(horizontal="center", vertical="center")
    for rij in range(waarde_rij, waarde_rij + 3):
        ws.row_dimensions[rij].height = 26

    badge_rij = waarde_rij + 3
    ws.merge_cells(start_row=badge_rij, start_column=links_idx, end_row=badge_rij, end_column=rechts_idx)
    badge_cel = ws.cell(row=badge_rij, column=links_idx)
    badge_cel.font = F_BADGE
    badge_cel.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[badge_rij].height = 22

    kader_vulling(ws, waarde_rij, badge_rij, kolommen, FILL_KAART)
    kader_rand(ws, kop_rij, badge_rij, kol_links, kol_rechts, CLR_ACCENT, dikte="medium")
    for kol in range(links_idx, rechts_idx + 1):
        ws.cell(row=kop_rij, column=kol).fill = FILL_TITEL

    return waarde_cel.coordinate, badge_cel.coordinate, badge_rij + 3


def bouw_dashboard(wb, cashflow, pensioen, vermogen):
    ws = wb.create_sheet("Dashboard")
    blad_basis(ws, CLR_ACCENT)
    zet_kolombreedtes(ws, {"A": 2, "B": 24, "C": 24, "D": 3, "E": 24, "F": 24, "G": 2})

    cf = f"'{cashflow['blad']}'"
    pen = f"'{pensioen['blad']}'"
    verm = f"'{vermogen['blad']}'"

    r = titelblok(
        ws, "Dashboard",
        "Jouw financiële situatie in één oogopslag — werkt automatisch bij zodra je de andere tabbladen invult",
        "F")
    r += 1

    # --- Kaart 1: Maandelijks overschot / tekort ---
    top1 = r
    waarde, badge, _ = dashboard_kaart(ws, top1, "B", "C", "€   MAANDELIJKS OVERSCHOT / TEKORT")
    ws[waarde] = f"={cf}!C{cashflow['overschot_rij']}"
    ws[waarde].number_format = FMT_EURO_ROND
    ws[badge] = (f'=IF({cf}!C{cashflow["overschot_rij"]}>=0,'
                f'"✓ Je houdt geld over","⚠ Je geeft meer uit dan je binnenkrijgt")')
    ws.conditional_formatting.add(
        waarde, FormulaRule(formula=[f"{cf}!C{cashflow['overschot_rij']}<0"],
                            font=Font(name=FONT_NAME, size=26, bold=True, color=CLR_ROOD_TEKST)))
    ws.conditional_formatting.add(
        badge, FormulaRule(formula=[f"{cf}!C{cashflow['overschot_rij']}>=0"],
                           fill=FILL_GROEN, font=Font(name=FONT_NAME, size=11, bold=True, color=CLR_GROEN_TEKST)))
    ws.conditional_formatting.add(
        badge, FormulaRule(formula=[f"{cf}!C{cashflow['overschot_rij']}<0"],
                           fill=FILL_ROOD, font=Font(name=FONT_NAME, size=11, bold=True, color=CLR_ROOD_TEKST)))

    # --- Kaart 2: Spaarquote ---
    sq = f"{cf}!C{cashflow['spaarquote_rij']}"
    waarde, badge, volgende1 = dashboard_kaart(ws, top1, "E", "F", "%   SPAARQUOTE")
    ws[waarde] = f'=IF({sq}="","–",{sq})'
    ws[waarde].number_format = FMT_PCT
    ws[badge] = (f'=IF({sq}="","Nog geen data",IF({sq}>=0.1,"✓ Gezonde spaarquote",'
                f'IF({sq}>=0,"⚠ Kan beter","✕ Tekort")))')
    ws.conditional_formatting.add(
        badge, FormulaRule(formula=[f'AND({sq}<>"",{sq}>=0.1)'],
                           fill=FILL_GROEN, font=Font(name=FONT_NAME, size=11, bold=True, color=CLR_GROEN_TEKST)))
    ws.conditional_formatting.add(
        badge, FormulaRule(formula=[f'AND({sq}<>"",{sq}>=0,{sq}<0.1)'],
                           fill=FILL_ORANJE, font=Font(name=FONT_NAME, size=11, bold=True, color=CLR_ORANJE_TEKST)))
    ws.conditional_formatting.add(
        badge, FormulaRule(formula=[f'AND({sq}<>"",{sq}<0)'],
                           fill=FILL_ROOD, font=Font(name=FONT_NAME, size=11, bold=True, color=CLR_ROOD_TEKST)))

    # --- Kaart 3: Pensioengat status ---
    top2 = volgende1
    gat = f"{pen}!C{pensioen['gat_rij']}"
    waarde, badge, _ = dashboard_kaart(ws, top2, "B", "C", "◆   PENSIOENGAT")
    ws[waarde] = f"={gat}"
    ws[waarde].number_format = FMT_EURO_ROND
    ws[badge] = f'=IF({gat}<=0,"✓ Op koers","⚠ Actie nodig")'
    ws.conditional_formatting.add(
        badge, FormulaRule(formula=[f"{gat}<=0"],
                           fill=FILL_GROEN, font=Font(name=FONT_NAME, size=11, bold=True, color=CLR_GROEN_TEKST)))
    ws.conditional_formatting.add(
        badge, FormulaRule(formula=[f"{gat}>0"],
                           fill=FILL_ORANJE, font=Font(name=FONT_NAME, size=11, bold=True, color=CLR_ORANJE_TEKST)))

    # --- Kaart 4: Verwacht vermogen ---
    waarde, badge, volgende2 = dashboard_kaart(
        ws, top2, "E", "F", f'=CONCATENATE("↗   VERWACHT VERMOGEN OVER ",{verm}!C{vermogen["jaren_invoer_rij"]}," JAAR")')
    ws[waarde] = f"={verm}!C{vermogen['eindwaarde_rij']}"
    ws[waarde].number_format = FMT_EURO_ROND
    ws[badge] = f'=CONCATENATE("Koopkracht van nu: € ",TEXT({verm}!C{vermogen["reeel_eind_rij"]},"#,##0"))'
    # Geen CF nodig: deze badge is informatief, geen goed/fout-oordeel — houdt de neutrale kaartkleur.

    r = volgende2 + 1
    r = sectiebalk(ws, r, "AAN DE SLAG", "F")
    for stap in [
        "1  Vul je cijfers in op 'Cashflow Overzicht', 'Pensioengat Calculator' en 'Vermogensgroei'.",
        "2  Kom terug naar dit Dashboard — de kaarten hierboven werken automatisch bij.",
        "3  Je maandelijkse overschot loopt automatisch door als voorstel voor je inleg bij pensioen en vermogen; "
        "dit kun je op elk tabblad overschrijven.",
        "4  Twijfel je over de kleurcodering of de formules? Kijk op het tabblad 'Instructies'.",
    ]:
        r = notitieregel(ws, r, stap, "F")
    r += 1
    r = notitieregel(ws, r, DISCLAIMER + ".", "F")

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws


# ---------------------------------------------------------------------------
# TABBLAD 3 — Instructies
# ---------------------------------------------------------------------------

def bouw_instructies(wb):
    ws = wb.create_sheet("Instructies")
    blad_basis(ws, CLR_DONKERBLAUW)
    zet_kolombreedtes(ws, {"A": 2, "B": 26, "C": 80, "D": 2})

    r = titelblok(ws, "Geldplanner", "Persoonlijk financieel overzicht voor Nederlandse huishoudens", "C")

    r = notitieregel(ws, r, "Welkom", "C", vet=True)
    for regel in [
        "Deze werkmap helpt je in kaart te brengen wat er maandelijks binnenkomt en weggaat, hoe groot je",
        "pensioengat is en hoe je vermogen groeit. Het Dashboard geeft een direct overzicht van je situatie.",
        "Vul alleen de gele cellen in. De grijze, blauwe en goudkleurige cellen bevatten formules en zijn vergrendeld.",
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
        ("1. Titelblad",
         "Voorblad met een plek voor je eigen logo, de productnaam en een korte introductie."),
        ("2. Dashboard",
         "Vier KPI-kaarten die automatisch bijwerken vanuit de andere tabbladen: je maandelijkse "
         "overschot, je spaarquote, de status van je pensioengat en je verwachte vermogen."),
        ("3. Instructies",
         "Deze pagina: uitleg, kleurcodering en de disclaimer."),
        ("4. Cashflow Overzicht",
         "Vul je maandelijkse inkomsten, vaste lasten en variabele kosten in. Je ziet direct je "
         "totale inkomsten, totale uitgaven, het maandelijkse overschot of tekort en je spaarquote."),
        ("5. Pensioengat Calculator",
         "Vul je leeftijd, gewenste pensioenleeftijd, gewenst pensioeninkomen, opgebouwd "
         "pensioenvermogen en verwachte AOW in. Je maandelijkse overschot uit 'Cashflow Overzicht' "
         "loopt automatisch door als voorstel voor je pensioeninleg; je kunt dit overschrijven."),
        ("6. Vermogensgroei",
         "Vul je startkapitaal, verwacht rendement en looptijd in. Ook hier loopt je overschot uit "
         "'Cashflow Overzicht' automatisch door als voorstel voor je maandelijkse inleg. Je ziet de "
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
        (FILL_RESULTAAT, RAND_RESULTAAT, "Sectieresultaat (blauw)",
         "Het eindresultaat van een sectie. Ook vergrendeld."),
        (FILL_KPI, RAND_ACCENT, "Hoofduitkomst (goud, groot)",
         "De belangrijkste cijfer van het hele tabblad, bijvoorbeeld je pensioengat."),
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
    r = notitieregel(ws, r,
                     "Op het Dashboard betekent een groene badge dat je situatie gezond is, "
                     "oranje dat er een aandachtspunt is en rood dat directe actie nodig is.", "C")
    r += 1

    r = sectiebalk(ws, r, "ZO GEBRUIK JE DE WERKMAP", "C")
    stappen = [
        "1. Begin bij 'Cashflow Overzicht' en vul je maandbudget in. Alles wat niet van toepassing is laat je op 0 staan.",
        "2. Bekijk het 'Dashboard' voor een overzicht van je situatie.",
        "3. Ga naar 'Pensioengat Calculator' en 'Vermogensgroei'. Je maandelijkse overschot uit stap 1 loopt daar "
        "automatisch door als voorstel voor je inleg — kies 'Nee' bij de koppeling om je eigen bedrag te gebruiken.",
        "4. De bladen zijn beveiligd zonder wachtwoord. Wil je formules aanpassen? Controleren > Blad-beveiliging opheffen.",
        "5. Bedragen staan in euro's. Op een Nederlandse Excel-installatie verschijnen ze als € 1.234,56.",
    ]
    for stap in stappen:
        cel = ws.cell(row=r, column=2, value=stap)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        cel.font = F_TEKST
        cel.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
        ws.row_dimensions[r].height = 30 if len(stap) > 95 else 17
        r += 1
    r += 1

    r = sectiebalk(ws, r, "BELANGRIJK — DISCLAIMER", "C")
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 3, end_column=3)
    cel = ws.cell(row=r, column=2)
    cel.value = (
        DISCLAIMER + ".\n\n"
        "Rendementen uit het verleden bieden geen garantie voor de toekomst; beleggen brengt risico's "
        "met zich mee en je kunt je inleg verliezen. De berekeningen zijn schattingen op basis van de "
        "aannames die je zelf invult.\n"
        "Raadpleeg voor persoonlijk advies een financieel adviseur."
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
# TABBLAD 4 — Cashflow Overzicht
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
    r = invoerregel(ws, r, "Freelance-inkomsten (netto)", 0, notitie="Netto bedrag, na belasting")
    r = invoerregel(ws, r, "Toeslagen (huur-, zorg-, kinderopvang)", 0)
    r = invoerregel(ws, r, "Huurinkomsten / alimentatie", 0)
    r = invoerregel(ws, r, "Overige inkomsten", 0)
    inkomsten_eind = r - 1
    tot_inkomsten_rij = r
    r = berekendregel(ws, r, "Totale inkomsten",
                      f"=SUM(C{inkomsten_start}:C{inkomsten_eind})")
    sectie_kaart(ws, inkomsten_start, tot_inkomsten_rij, "D")
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
    sectie_kaart(ws, vast_start, tot_vast_rij, "D")
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
    sectie_kaart(ws, var_start, tot_var_rij, "D")
    r += 1

    # --- Resultaat ---
    resultaat_start = r
    r = sectiebalk(ws, r, "RESULTAAT")
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
    resultaat_eind = r
    sectie_kaart(ws, resultaat_start + 1, resultaat_eind, "D")
    r += 1

    # Voorwaardelijke opmaak: groen bij overschot, rood bij tekort.
    ws.conditional_formatting.add(
        f"C{overschot_rij}",
        CellIsRule(operator="lessThan", formula=["0"], fill=FILL_ROOD,
                   font=Font(name=FONT_NAME, size=20, bold=True, color=CLR_ROOD_TEKST)))
    ws.conditional_formatting.add(
        f"C{overschot_rij}",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=FILL_GROEN,
                   font=Font(name=FONT_NAME, size=20, bold=True, color=CLR_GROEN_TEKST)))

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
    return ws, {
        "blad": "Cashflow Overzicht",
        "overschot_rij": overschot_rij,
        "spaarquote_rij": spaarquote_rij,
        "inkomsten_rij": res_inkomsten,
        "uitgaven_rij": res_uitgaven,
    }


# ---------------------------------------------------------------------------
# TABBLAD 5 — Pensioengat Calculator
# ---------------------------------------------------------------------------

def bouw_pensioen(wb, cashflow):
    ws = wb.create_sheet("Pensioengat Calculator")
    blad_basis(ws, "7030A0")
    zet_kolombreedtes(ws, {"A": 2, "B": 44, "C": 16, "D": 52, "E": 2})
    cf = f"'{cashflow['blad']}'"

    r = titelblok(ws, "Pensioengat Calculator", "Hoeveel moet je maandelijks opzijzetten voor het pensioen dat je wilt?", "D")

    situatie_start = r
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
    sectie_kaart(ws, situatie_start + 1, r - 1, "D")
    r += 1

    # --- Koppeling met Cashflow Overzicht ---
    koppel_start = r
    r = sectiebalk(ws, r, "KOPPELING MET CASHFLOW OVERZICHT", accent=True)
    voorstel_rij = r
    r = berekendregel(ws, r, "Voorstel: maandelijks overschot uit Cashflow Overzicht",
                      f"={cf}!C{cashflow['overschot_rij']}",
                      notitie="Wordt automatisch bijgewerkt zodra je dat tabblad invult")
    gebruik_voorstel_rij = r
    r = invoerregel(ws, r, "Gebruik dit voorstel als je huidige pensioeninleg?", "Ja", notatie="General",
                    notitie="Kies 'Nee' om je eigen bedrag hieronder te gebruiken")
    eigen_inleg_rij = r
    r = invoerregel(ws, r, "Eigen invoer (alleen gebruikt bij 'Nee')", 150,
                    notitie="Vul hier je eigen maandelijkse pensioeninleg in")
    inleg_rij = r
    r = berekendregel(ws, r, "Wat leg je nu al maandelijks in?",
                      f'=IF(C{gebruik_voorstel_rij}="Ja",C{voorstel_rij},C{eigen_inleg_rij})',
                      accent=True, notitie="Gebruikt in de berekening hieronder")
    ja_nee_validatie(ws, [gebruik_voorstel_rij])
    sectie_kaart(ws, koppel_start + 1, r - 1, "D")
    r += 1

    aannames_start = r
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
    sectie_kaart(ws, aannames_start + 1, r - 1, "D")
    r += 1

    gat_start = r
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
                      f"=MAX(0,C{kapitaal_rij}-C{groei_rij}-C{groei_inleg_rij})", groot=True,
                      notitie="Nul betekent: je ligt op koers")
    sectie_kaart(ws, gat_start + 1, r - 1, "D")
    r += 1

    doen_start = r
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
    sectie_kaart(ws, doen_start + 1, r - 1, "D")
    r += 1

    ws.conditional_formatting.add(
        f"C{gat_rij}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_ORANJE,
                   font=Font(name=FONT_NAME, size=20, bold=True, color=CLR_ORANJE_TEKST)))
    ws.conditional_formatting.add(
        f"C{gat_rij}",
        CellIsRule(operator="lessThanOrEqual", formula=["0"], fill=FILL_GROEN,
                   font=Font(name=FONT_NAME, size=20, bold=True, color=CLR_GROEN_TEKST)))

    r = notitieregel(ws, r, "Hoe deze berekening werkt", "D", vet=True)
    for tekst in [
        "• Je maandelijkse inleg is gekoppeld aan je overschot in 'Cashflow Overzicht'; kies 'Nee' bij de "
        "koppeling hierboven om je eigen bedrag te gebruiken.",
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
    return ws, {
        "blad": "Pensioengat Calculator",
        "gat_rij": gat_rij,
        "jaren_rij": jaren_rij,
        "extra_rij": extra_rij,
    }


# ---------------------------------------------------------------------------
# TABBLAD 6 — Vermogensgroei
# ---------------------------------------------------------------------------

MAX_JAREN = 40


def bouw_vermogen(wb, cashflow):
    ws = wb.create_sheet("Vermogensgroei")
    blad_basis(ws, "BF8F00")
    zet_kolombreedtes(ws, {"A": 2, "B": 40, "C": 16, "D": 46, "E": 2,
                           "F": 8, "G": 15, "H": 15, "I": 15, "J": 15, "K": 15, "L": 15})
    cf = f"'{cashflow['blad']}'"

    r = titelblok(ws, "Vermogensgroei", "Projectie van je vermogen met samengestelde rente", "D")

    # --- Koppeling met Cashflow Overzicht ---
    koppel_start = r
    r = sectiebalk(ws, r, "KOPPELING MET CASHFLOW OVERZICHT", accent=True)
    voorstel_rij = r
    r = berekendregel(ws, r, "Voorstel: maandelijks overschot uit Cashflow Overzicht",
                      f"={cf}!C{cashflow['overschot_rij']}",
                      notitie="Wordt automatisch bijgewerkt zodra je dat tabblad invult")
    gebruik_voorstel_rij = r
    r = invoerregel(ws, r, "Gebruik dit voorstel als maandelijkse inleg?", "Ja", notatie="General",
                    notitie="Kies 'Nee' om je eigen bedrag hieronder te gebruiken")
    eigen_inleg_rij = r
    r = invoerregel(ws, r, "Eigen invoer (alleen gebruikt bij 'Nee')", 500,
                    notitie="Vul hier je eigen maandelijkse inleg in")
    ja_nee_validatie(ws, [gebruik_voorstel_rij])
    sectie_kaart(ws, koppel_start + 1, r - 1, "D")
    r += 1

    invoer_start = r
    r = sectiebalk(ws, r, "JOUW INVOER")
    start_rij = r
    r = invoerregel(ws, r, "Huidig spaargeld en beleggingen", 25000,
                    notitie="Je startkapitaal")
    maandinleg_rij = r
    r = berekendregel(ws, r, "Maandelijkse inleg (gebruikt in de projectie)",
                      f'=IF(C{gebruik_voorstel_rij}="Ja",C{voorstel_rij},C{eigen_inleg_rij})',
                      accent=True, notitie="Volgt automatisch uit de koppeling hierboven")
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
    sectie_kaart(ws, invoer_start + 1, r - 1, "D")
    r += 1

    jaren_val = DataValidation(type="whole", operator="between",
                               formula1="0", formula2=str(MAX_JAREN), allow_blank=False)
    jaren_val.error = f"Vul een aantal jaren tussen 0 en {MAX_JAREN} in."
    jaren_val.errorTitle = "Ongeldige looptijd"
    ws.add_data_validation(jaren_val)
    jaren_val.add(ws.cell(row=jaren_rij, column=3))

    afgeleid_start = r
    r = sectiebalk(ws, r, "AFGELEIDE WAARDEN")
    netto_rendement_rij = r
    r = berekendregel(ws, r, "Netto rendement na kosten",
                      f"=C{rendement_rij}-C{kosten_rij}", notatie=FMT_PCT2,
                      notitie="Wordt in de projectie gebruikt")
    maandrente_rij = r
    r = berekendregel(ws, r, "Maandelijks rendement (samengesteld)",
                      f"=(1+C{netto_rendement_rij})^(1/12)-1", notatie='0.000%')
    sectie_kaart(ws, afgeleid_start + 1, r - 1, "D")
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
    samenvatting_start = rr
    rr = sectiebalk(ws, rr, "RESULTAAT NA DE GEKOZEN LOOPTIJD")
    eindwaarde_rij = rr
    rr = berekendregel(ws, rr, "Geprojecteerde eindwaarde",
                       f"=INDEX(J{tabel_start}:J{tabel_eind},C{jaren_rij}+1)", groot=True)
    ingelegd_totaal_rij = rr
    rr = berekendregel(ws, rr, "Totaal zelf ingelegd",
                       f"=INDEX(K{tabel_start}:K{tabel_eind},C{jaren_rij}+1)")
    rr = berekendregel(ws, rr, "Waarvan rendement",
                       f"=C{eindwaarde_rij}-C{ingelegd_totaal_rij}",
                       notitie="Het deel dat je geld voor je heeft verdiend")
    reeel_eind_rij = rr
    rr = berekendregel(ws, rr, "Eindwaarde in euro's van nu",
                       f"=INDEX(L{tabel_start}:L{tabel_eind},C{jaren_rij}+1)",
                       notitie="Gecorrigeerd voor inflatie: dit is je koopkracht")
    sectie_kaart(ws, samenvatting_start + 1, rr - 1, "D")

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
    notitieregel(ws, laatste, "Je maandelijkse inleg is gekoppeld aan je overschot in 'Cashflow Overzicht'; "
                              "kies 'Nee' bij de koppeling hierboven om je eigen bedrag te gebruiken.", "D")
    notitieregel(ws, laatste + 1, "De projectie rekent met maandelijkse bijstortingen en samengestelde groei "
                                  "op basis van het rendement ná kosten.", "D")
    notitieregel(ws, laatste + 2, DISCLAIMER + ". Rendementen uit het verleden bieden geen garantie voor de toekomst.", "D")

    ws.freeze_panes = "A5"
    beveilig(ws)
    return ws, {
        "blad": "Vermogensgroei",
        "eindwaarde_rij": eindwaarde_rij,
        "reeel_eind_rij": reeel_eind_rij,
        "jaren_invoer_rij": jaren_rij,
    }


# ---------------------------------------------------------------------------
# Hoofdprogramma
# ---------------------------------------------------------------------------

def bouw_werkmap():
    wb = Workbook()
    wb.remove(wb.active)

    _, cashflow_ref = bouw_cashflow(wb)
    _, pensioen_ref = bouw_pensioen(wb, cashflow_ref)
    _, vermogen_ref = bouw_vermogen(wb, cashflow_ref)
    bouw_dashboard(wb, cashflow_ref, pensioen_ref, vermogen_ref)
    bouw_titelblad(wb)
    bouw_instructies(wb)

    volgorde = ["Titelblad", "Dashboard", "Instructies",
               "Cashflow Overzicht", "Pensioengat Calculator", "Vermogensgroei"]
    wb._sheets = [wb[naam] for naam in volgorde]

    wb.properties.title = "Geldplanner — persoonlijke financien"
    wb.properties.subject = "Dashboard, cashflow, pensioengat en vermogensgroei"
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
