"""
alles_in_een.py
------------------
Alle 5 onderdelen in 1 bestand, met losse subcommando's zodat je elk
onderdeel apart kunt draaien - allemaal met --retailer (zoekt automatisch
het juiste bestand), of --csv <pad> als je een specifiek bestand wilt
forceren.

Gebruik:
    python alles_in_een.py csv --retailer Hoogvliet
    python alles_in_een.py regels --retailer Hoogvliet
    python alles_in_een.py bestand --retailer Hoogvliet
    python alles_in_een.py main --retailer Hoogvliet [--weken 40,41] [--accountmanager Ben]

    # Met een direct pad i.p.v. --retailer (werkt voor csv/regels):
    python alles_in_een.py csv --csv "inzicht/Promo Focus File 2026 Hoogvliet.csv"

Vereist (voor alle onderdelen behalve csv/regels op zichzelf):
Windows, Excel, pywin32 (win32com), psutil, en Actievoorstellen.xlsx in
dezelfde map.
"""

import argparse
import datetime as dt
import glob
import html
import re
import shutil
import sys
import threading
import time
from pathlib import Path

import pandas as pd


# ======================================================================
# ONDERDEEL 1: CSV LEZEN
# ----------------------------------------------------------------------
# Leest en filtert de Promo Focus CSV. Puur Python/pandas, GEEN Excel.
# Los te draaien met: python alles_in_een.py csv --retailer Hoogvliet
# ======================================================================

FOCUS_KOLOMMEN = {
    "week": "A", "week_b": "B", "kolom_c": "C",
    "mech_d": "D", "mech_e": "E",
    "sap_code": "F", "ean": "G", "artikelnaam": "H",
    "kolom_l_bron": "O",
    "volume": "W",
}

KOLOM_C_UITSLUITEN = ["delist", "gesaneerd", "sanering"]


def excel_col_to_index(letter: str) -> int:
    index = 0
    for ch in letter.upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    value = str(value).strip()
    if value.lower() in ("nan", "none"):
        return ""
    return html.unescape(value).strip()


def clean_ean(value) -> str:
    value = clean_text(value)
    if value.endswith(".0"):
        value = value[:-2]
    return value.strip()


def clean_number(value):
    """Punt = decimaalteken (Engelse notatie, deze CSV is komma-gescheiden)."""
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    value = str(value).strip()
    if value == "" or value.lower() in ("nan", "none", "-"):
        return None
    value = value.replace("\u20ac", "").replace(" ", "").replace("\u00a0", "")
    try:
        return float(value)
    except ValueError:
        pass
    try:
        return float(value.replace(",", ""))
    except ValueError:
        return None


def clean_week(value):
    nummer = clean_number(value)
    if nummer is None:
        return None
    try:
        return int(nummer)
    except Exception:
        return None


def clean_volume_number(value):
    return clean_number(value)


def volume_to_excel_value(value):
    nummer = clean_volume_number(value)
    if nummer is None:
        return clean_text(value)
    if abs(nummer - round(nummer)) < 0.000001:
        return int(round(nummer))
    return round(nummer, 2)


def volume_as_float(value) -> float:
    nummer = clean_volume_number(value)
    return float(nummer) if nummer is not None else 0.0


def is_geldige_ean(ean: str) -> bool:
    ean = clean_text(ean)
    if not ean.isdigit() or len(ean) != 13:
        return True
    cijfers = [int(c) for c in ean]
    som = sum(c * (3 if i % 2 else 1) for i, c in enumerate(cijfers[:12]))
    controle = (10 - som % 10) % 10
    return controle == cijfers[12]


def is_npd_code(sap_code_raw: str) -> bool:
    return "npd" in clean_text(sap_code_raw).lower()


def detecteer_csv_separator(pad: Path) -> str:
    with open(pad, "r", encoding="utf-8-sig") as f:
        eerste_regel = f.readline()
    aantal_komma = eerste_regel.count(",")
    aantal_puntkomma = eerste_regel.count(";")
    gekozen = ";" if aantal_puntkomma > aantal_komma else ","
    print(f"CSV-separator gedetecteerd: '{gekozen}' (komma's: {aantal_komma}, puntkomma's: {aantal_puntkomma})")
    return gekozen


def laad_focus_data(pad: Path) -> pd.DataFrame:
    separator = detecteer_csv_separator(pad)
    raw = pd.read_csv(pad, sep=separator, header=None, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    idx = {naam: excel_col_to_index(letter) for naam, letter in FOCUS_KOLOMMEN.items()}
    print(f"Promo Focus CSV: {raw.shape[0]} rijen, {raw.shape[1]} kolommen gelezen.")

    if raw.shape[1] <= max(idx.values()):
        raise ValueError(
            f"Promo Focus CSV heeft maar {raw.shape[1]} kolommen, maar kolom "
            f"{max(FOCUS_KOLOMMEN.values())} is nodig. Controleer of de separator-"
            f"detectie ('{separator}') wel klopt voor dit bestand."
        )

    df = pd.DataFrame({
        "week_int": raw.iloc[:, idx["week"]].apply(clean_week),
        "week_b": raw.iloc[:, idx["week_b"]].apply(clean_week),
        "kolom_c": raw.iloc[:, idx["kolom_c"]].apply(clean_text),
        "mech_d": raw.iloc[:, idx["mech_d"]].apply(clean_text),
        "mech_e": raw.iloc[:, idx["mech_e"]].apply(clean_text),
        "sap_code_raw": raw.iloc[:, idx["sap_code"]].apply(clean_text),
        "ean": raw.iloc[:, idx["ean"]].apply(clean_ean),
        "artikelnaam": raw.iloc[:, idx["artikelnaam"]].apply(clean_text),
        "kolom_l_waarde": raw.iloc[:, idx["kolom_l_bron"]].apply(clean_text),
        "volume_raw": raw.iloc[:, idx["volume"]].apply(clean_text),
    })

    print(f"  na inlezen: {len(df)} rijen")
    df = df[df["week_int"].notna()].copy()
    print(f"  na week-filter (kolom A moet een geldig getal zijn): {len(df)} rijen")

    is_marker = df["week_b"].notna() & (df["week_b"] == df["week_int"])
    print(f"  waarvan {int(is_marker.sum())} mechanisme-markerrij(en) gevonden (kolom B == kolom A op die rij)")

    heeft_ean = df["ean"].ne("") & df["ean"].notna()
    leeg_ean = (~heeft_ean) & (df["sap_code_raw"] != "")
    if leeg_ean.any():
        print(f"WAARSCHUWING: {leeg_ean.sum()} productregel(s) zonder EAN in {pad.name}.")

    ongeldig = heeft_ean & ~df["ean"].apply(is_geldige_ean)
    if ongeldig.any():
        print(f"WAARSCHUWING: {ongeldig.sum()} EAN(s) met ongeldig controlecijfer: {df.loc[ongeldig, 'ean'].tolist()[:20]}")

    dubbel = df[heeft_ean].duplicated(subset=["ean", "week_int"], keep=False)
    if dubbel.any():
        print(f"WAARSCHUWING: {dubbel.sum()} regel(s) met een EAN die dubbel voorkomt binnen dezelfde week.")

    return df.reset_index(drop=True)


def filter_kolom_c(df: pd.DataFrame, uitsluit_keywords=None, toegestane_excepties=None) -> pd.DataFrame:
    uitsluit_keywords = uitsluit_keywords or KOLOM_C_UITSLUITEN
    toegestane_excepties = toegestane_excepties or []
    excepties_lower = [clean_text(e).lower() for e in toegestane_excepties if clean_text(e)]

    def toegestaan(waarde: str) -> bool:
        w = clean_text(waarde).lower()
        if w in excepties_lower:
            return True
        return not any(kw in w for kw in uitsluit_keywords)

    mask = df["kolom_c"].apply(toegestaan)
    uitgesloten = int((~mask).sum())
    if uitgesloten:
        print(f"Kolom C-filter: {uitgesloten} regel(s) uitgesloten (delist/gesaneerd/sanering).")
    return df[mask].reset_index(drop=True)


# ======================================================================
# ONDERDEEL 2: REGELS BOUWEN
# ----------------------------------------------------------------------
# Mechanisme-herkenning, buffering, groepering per week+mechanisme.
# Puur Python, GEEN Excel. Hangt af van onderdeel 1 hierboven.
# Los te draaien met: python alles_in_een.py regels --retailer Hoogvliet
# ======================================================================

ADVIES_AFBEELDING_AANTAL = 2


def normaliseer_mechanisme_basis(mechanisme) -> str:
    tekst = clean_text(mechanisme)
    if not tekst:
        return ""
    tekst = re.sub(r"^\s*\d+\.\s*", "", tekst)
    tekst = re.sub(r"\bSPO\b", "1 voor", tekst, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", tekst).strip()


_MECHANISME_PATRONEN = [
    r"\b\d+\s*\+\s*\d+\b",                    # kale 'N + N' (bv. '1 + 1', zonder 'gratis')
    r"\b\d+\s*\+\s*\d+\s+gratis\b",            # specifieker: 'N+N gratis' - wint bij gelijke startpositie
    r"\b\d{1,3}\s*%\s*korting\b",
    r"\b\d+\s*(?:e|de|ste)\s+gratis\b",
    r"\b\d+\s*(?:e|de|ste)\s+halve\s+prijs\b",
    r"\b\d+\s+voor\s+\u20ac?\s*\d+(?:[,.]\d{1,2})?\b",
    r"\b\d+\s+voor\s+\d+\b",
]


def extraheer_mechanisme_tekst(mechanisme) -> str:
    tekst = normaliseer_mechanisme_basis(mechanisme)
    if not tekst:
        return ""
    gevonden = []
    for patroon in _MECHANISME_PATRONEN:
        for match in re.finditer(patroon, tekst, flags=re.IGNORECASE):
            gevonden.append((match.start(), match.group(0)))
    if not gevonden:
        return ""
    _, resultaat = sorted(gevonden, key=lambda item: item[0])[-1]
    resultaat = re.sub(r"\s*\+\s*", " + ", resultaat)
    resultaat = re.sub(r"\s*%\s*", "% ", resultaat)
    return re.sub(r"\s+", " ", resultaat).strip().lower()


def formatteer_actiemechanisme_geel(mechanisme) -> str:
    """Volledig uitgeschreven tekst voor de gele mechanisme-rij: EXACT de
    brontekst (na het strippen van het leidende volgnummer en SPO->1 voor),
    GEEN automatische merk-prefix meer. Voorbeeld eerder gaf het verkeerde
    idee: daar stond 'Iglo' toevallig al in de brontekst zelf, dat is
    verkeerd geinterpreteerd als 'altijd toevoegen' - bij '20. Vissticks
    1 + 1' hoort het gewoon 'Vissticks 1 + 1' te worden, zonder Iglo ervoor."""
    tekst = normaliseer_mechanisme_basis(mechanisme)
    if not tekst:
        return ""
    tekst = re.sub(r"\b(\d+)\s*\+\s*(\d+)\s+gratis\b", r"\1 + \2 gratis", tekst, flags=re.IGNORECASE)
    tekst = re.sub(r"\b(\d+)\s*\+\s*(\d+)\b", r"\1 + \2", tekst)   # ook spacing fixen zonder 'gratis'
    tekst = re.sub(r"\s*%\s*", "% ", tekst)
    tekst = re.sub(r"\s+", " ", tekst).strip()
    return tekst


def formatteer_actiemechanisme_kolom_p(mechanisme) -> str:
    kort = extraheer_mechanisme_tekst(mechanisme)
    if kort:
        return kort
    tekst = normaliseer_mechanisme_basis(mechanisme)
    if not tekst:
        return ""
    return re.sub(r"^\s*Iglo\s+", "", tekst, flags=re.IGNORECASE).strip()


def is_mechanisme_marker(rij) -> bool:
    """De markerrij wordt herkend doordat kolom B op die rij hetzelfde
    weeknummer bevat als kolom A (terwijl kolom B op gewone productrijen
    leeg is). NIET meer op basis van een lege SAP-code - dat pakte soms per
    ongeluk een gewone productrij met een toevallig lege SAP-code, en las
    daar dan de artikelnaam i.p.v. een mechanismetekst uit kolom D."""
    return pd.notna(rij["week_b"]) and rij["week_b"] == rij["week_int"]


def bouw_outputregels(focus: pd.DataFrame, debug: bool = False) -> list:
    regels, buffer = [], []
    overgeslagen_geen_sap_code = 0
    geen_artikelnaam = 0

    for _, rij in focus.iterrows():
        if is_mechanisme_marker(rij):
            mechanisme_ruw = rij["mech_d"] or rij["mech_e"]
            mechanisme_vol = formatteer_actiemechanisme_geel(mechanisme_ruw)
            mechanisme_kort = formatteer_actiemechanisme_kolom_p(mechanisme_ruw)
            if debug:
                print(f"[DEBUG marker week={rij['week_int']}] bron={mechanisme_ruw!r}")
                print(f"  -> kolom E (geel) = {mechanisme_vol!r}")
                print(f"  -> kolom P (kort) = {mechanisme_kort!r}")
                print(f"  -> {len(buffer)} product(en) krijgen dit toegekend: {[r['productnaam'] for r in buffer]}")
            for regel in buffer:
                regel["mechanisme_vol"] = mechanisme_vol
                regel["mechanisme_kort"] = mechanisme_kort
                regels.append(regel)
            buffer = []
            continue

        sap_raw = rij["sap_code_raw"]
        is_npd = is_npd_code(sap_raw)
        heeft_sap_code = sap_raw != ""

        # Een 'item' is elke rij met een ingevulde SAP-code in kolom F (of
        # een NPD-code), tenzij de kolom C-filters de rij al hebben
        # uitgesloten (dat gebeurt al eerder, in filter_kolom_c). Een lege
        # artikelnaam in kolom H is GEEN reden meer om de regel te
        # negeren - vroeger viel zo'n product onterecht weg.
        if not heeft_sap_code and not is_npd:
            overgeslagen_geen_sap_code += 1
            continue

        productnaam = "NPD" if is_npd else rij["artikelnaam"]
        if not productnaam:
            geen_artikelnaam += 1
            if debug:
                print(f"[DEBUG] SAP-code '{sap_raw}' (EAN {rij['ean'] or '-'}, week {rij['week_int']}) "
                      f"heeft geen artikelnaam in kolom H - regel wordt toch meegenomen.")

        buffer.append({
            "week": rij["week_int"],
            "ean": rij["ean"],
            "productnaam": productnaam,
            "mechanisme_vol": "",
            "mechanisme_kort": "",
            "volume_excel": volume_to_excel_value(rij["volume_raw"]),
            "volume_sorteerwaarde": volume_as_float(rij["volume_raw"]),
            "kolom_l_waarde": rij["kolom_l_waarde"],
        })

    regels.extend(buffer)

    zonder_mechanisme = [r for r in regels if not r["mechanisme_kort"] and r["productnaam"] != "NPD"]
    if zonder_mechanisme:
        print(f"WAARSCHUWING: {len(zonder_mechanisme)} product(en) zonder mechanismetekst. "
              f"Weken: {sorted({r['week'] for r in zonder_mechanisme})}")

    if overgeslagen_geen_sap_code:
        print(f"{overgeslagen_geen_sap_code} regel(s) overgeslagen: geen SAP-code (kolom F) en geen NPD.")
    if geen_artikelnaam:
        print(f"WAARSCHUWING: {geen_artikelnaam} regel(s) hadden wel een SAP-code maar GEEN artikelnaam "
              f"in kolom H - wel meegenomen, maar controleer de CSV (lege productnaam in de output).")
    print(f"bouw_outputregels: {len(regels)} productregel(s) opgebouwd uit {len(focus)} focus-rijen.")

    return regels


def voeg_advies_kruisjes_toe(regels: list, aantal: int = ADVIES_AFBEELDING_AANTAL) -> None:
    per_week = {}
    for r in regels:
        per_week.setdefault(r["week"], []).append(r)
    for week_regels in per_week.values():
        top = sorted(
            week_regels,
            key=lambda r: (-r["volume_sorteerwaarde"], r["productnaam"], r["ean"]),
        )[:aantal]
        top_ids = {id(r) for r in top}
        for r in week_regels:
            r["advies_afbeelding"] = id(r) in top_ids


def groepeer_per_mechanisme(regels: list) -> list:
    groepen, volgorde = {}, []
    for regel in regels:
        sleutel = regel["mechanisme_vol"]
        if sleutel not in groepen:
            groepen[sleutel] = []
            volgorde.append(sleutel)
        groepen[sleutel].append(regel)
    return [(sleutel, groepen[sleutel]) for sleutel in volgorde]


def groepeer_per_week_en_mechanisme(regels: list) -> list:
    per_week, volgorde_weken = {}, []
    for regel in regels:
        week = regel["week"]
        if week not in per_week:
            per_week[week] = []
            volgorde_weken.append(week)
        per_week[week].append(regel)
    return [(week, groepeer_per_mechanisme(per_week[week])) for week in volgorde_weken]


# ======================================================================
# ONDERDEEL 3: EXCEL SCHRIJVEN
# ----------------------------------------------------------------------
# Vult EEN sheet met de geneste week->mechanisme->regels data. GEEN
# CSV-logica. Heeft een echte, geopende Excel-sheet nodig - kan NIET los
# gedraaid worden, alleen via het "main"-subcommando.
# ======================================================================

TEMPLATE_FORMULE_RIJ = 122
START_OUTPUT_RIJ = 7
PRINT_LAATSTE_KOLOM = "R"

TITEL_CEL = "F1"
FROZEN_FOOD_CEL = "B1"
ACCOUNTMANAGER_CEL = "B2"
DATUM_CEL = "B3"
Q_LABEL_CEL = "A6"

KOL_A_DVIP = "A"
KOL_B_WEEK = "B"
KOL_C_ADVIES_AFBEELDING = "C"
KOL_D_EAN = "D"
KOL_E_PRODUCT = "E"
KOL_I = "I"
KOL_L = "L"
KOL_J_ACTIE_INKOOPPRIJS = "J"
KOL_P_MECHANISME = "P"
KOL_Q_VOLUME = "Q"

KOLOM_A_VASTE_WAARDE = "V"
KOLOM_C_VASTE_WAARDE = "X"

EAN_FORMULE_KOLOMMEN = ["F", "G", "H", "J", "K", "M", "R"]

KLEUR_BLAUW = 15257527
KLEUR_GEEL = 10092543

XL_PASTE_ALL = -4104
XL_SHEET_VISIBLE = -1
XL_CENTER = -4108
XL_VCENTER = -4108
XL_TOP = -4160
XL_LEFT = -4131
XL_EDGE_TOP = 8
XL_EDGE_BOTTOM = 9
XL_EDGE_LEFT = 7
XL_EDGE_RIGHT = 10
XL_INSIDE_VERTICAL = 11
XL_INSIDE_HORIZONTAL = 12
XL_CONTINUOUS = 1
XL_THIN = 2
XL_MEDIUM = -4138
XL_SOLID = 1
XL_SHIFT_DOWN = -4121
XL_SHIFT_UP = -4162


def kwartaal_van_week(week: int) -> int:
    if week <= 13:
        return 1
    if week <= 26:
        return 2
    if week <= 39:
        return 3
    return 4


def kwartaal_label(weken: list) -> str:
    kwartalen = sorted({kwartaal_van_week(w) for w in weken})
    return "/".join(f"Q{k}" for k in kwartalen)


def vind_template_sheet(wb, naam):
    if naam:
        for sheet in wb.Sheets:
            if sheet.Name == naam:
                if sheet.Visible != XL_SHEET_VISIBLE:
                    sheet.Visible = XL_SHEET_VISIBLE
                return sheet
        raise ValueError(f"Tabblad '{naam}' niet gevonden. Beschikbaar: {[s.Name for s in wb.Sheets]}")
    sheet = wb.ActiveSheet
    sheet.Visible = XL_SHEET_VISIBLE
    return sheet


def zet_verticale_lijnen_zonder_horizontaal(sheet, rij: int):
    rng = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}")

    # Bewaar de lijndikte die al op deze rij stond (meegekomen via het
    # kopiëren van de sjabloonrij) - dezelfde dikte als de artikelrijen.
    # Die zetten we straks terug, want het zetten van LineStyle hieronder
    # kan 'm resetten naar Excel's eigen standaarddikte. Voorheen werd hier
    # altijd hard XL_THIN gezet, ongeacht wat de sjabloonrij gebruikte -
    # daardoor zagen de gele/witte rijen er dunner uit dan de artikelrijen.
    try:
        sjabloon_dikte = rng.Borders(XL_EDGE_LEFT).Weight
    except Exception:
        sjabloon_dikte = XL_THIN

    try:
        rng.Borders(XL_EDGE_TOP).LineStyle = 0
        rng.Borders(XL_EDGE_BOTTOM).LineStyle = 0
        rng.Borders(XL_INSIDE_HORIZONTAL).LineStyle = 0
    except Exception:
        pass
    try:
        rng.Borders(XL_EDGE_LEFT).LineStyle = XL_CONTINUOUS
        rng.Borders(XL_EDGE_LEFT).Weight = sjabloon_dikte
        rng.Borders(XL_EDGE_RIGHT).LineStyle = XL_CONTINUOUS
        rng.Borders(XL_EDGE_RIGHT).Weight = sjabloon_dikte
        rng.Borders(XL_INSIDE_VERTICAL).LineStyle = XL_CONTINUOUS
        rng.Borders(XL_INSIDE_VERTICAL).Weight = sjabloon_dikte
    except Exception:
        pass


def _schrijf_witte_spacer_rij(sheet, rij: int, zwarte_bovenrand: bool):
    rng = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}")
    rng.ClearContents()
    rng.Interior.ColorIndex = 2

    cel_j = sheet.Range(f"{KOL_J_ACTIE_INKOOPPRIJS}{rij}")
    cel_j.Interior.Pattern = XL_SOLID
    cel_j.Interior.Color = KLEUR_BLAUW

    zet_verticale_lijnen_zonder_horizontaal(sheet, rij)

    if zwarte_bovenrand:
        rng.Borders(XL_EDGE_TOP).LineStyle = XL_CONTINUOUS
        rng.Borders(XL_EDGE_TOP).Weight = XL_MEDIUM

    sheet.Rows(rij).RowHeight = 8


def _schrijf_mechanisme_rij(sheet, rij: int, mechanisme_tekst: str):
    rng = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}")
    rng.ClearContents()
    rng.Interior.ColorIndex = 2

    cel_mechanisme = sheet.Range(f"{KOL_E_PRODUCT}{rij}")
    cel_mechanisme.Value = mechanisme_tekst
    cel_mechanisme.Interior.Pattern = XL_SOLID
    cel_mechanisme.Interior.Color = KLEUR_GEEL
    cel_mechanisme.Font.Bold = True
    cel_mechanisme.HorizontalAlignment = XL_LEFT
    cel_mechanisme.VerticalAlignment = XL_VCENTER

    cel_j = sheet.Range(f"{KOL_J_ACTIE_INKOOPPRIJS}{rij}")
    cel_j.Interior.Pattern = XL_SOLID
    cel_j.Interior.Color = KLEUR_BLAUW

    zet_verticale_lijnen_zonder_horizontaal(sheet, rij)
    sheet.Rows(rij).RowHeight = 18


def _formatteer_artikelrij(sheet, rij: int):
    sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}").VerticalAlignment = XL_TOP
    sheet.Rows(rij).RowHeight = 18
    sheet.Range(f"{KOL_A_DVIP}{rij}").HorizontalAlignment = XL_CENTER
    sheet.Range(f"{KOL_A_DVIP}{rij}").VerticalAlignment = XL_VCENTER
    sheet.Range(f"{KOL_D_EAN}{rij}").HorizontalAlignment = XL_LEFT
    sheet.Range(f"{KOL_E_PRODUCT}{rij}").HorizontalAlignment = XL_LEFT
    sheet.Range(f"{KOL_P_MECHANISME}{rij}").HorizontalAlignment = XL_LEFT
    sheet.Range(f"{KOL_Q_VOLUME}{rij}").HorizontalAlignment = XL_LEFT


def _controleer_ean_formules(sheet, rij: int, sheet_naam: str):
    ontbrekend = [k for k in EAN_FORMULE_KOLOMMEN if not sheet.Range(f"{k}{rij}").HasFormula]
    if ontbrekend:
        print(f"WAARSCHUWING [{sheet_naam}]: kolom(men) {ontbrekend} hebben GEEN formule in rij {rij}.")
    else:
        print(f"  [{sheet_naam}] EAN-formule-check rij {rij}: alle kolommen {EAN_FORMULE_KOLOMMEN} OK.")


def _schrijf_productrij(sheet, rij: int, regel: dict, eerste_rij_van_week: bool, sheet_naam: str = ""):
    if eerste_rij_van_week:
        sheet.Range(f"{KOL_A_DVIP}{rij}").Value = KOLOM_A_VASTE_WAARDE
        sheet.Range(f"{KOL_B_WEEK}{rij}").Value = regel["week"]
        _controleer_ean_formules(sheet, rij, sheet_naam)
    else:
        sheet.Range(f"{KOL_A_DVIP}{rij}").Value = ""
        sheet.Range(f"{KOL_B_WEEK}{rij}").Value = ""

    sheet.Range(f"{KOL_C_ADVIES_AFBEELDING}{rij}").Value = (
        KOLOM_C_VASTE_WAARDE if regel.get("advies_afbeelding") else ""
    )

    ean_cel = sheet.Range(f"{KOL_D_EAN}{rij}")
    ean_waarde = regel["ean"]
    if ean_waarde.isdigit():
        ean_cel.NumberFormat = "0"
        ean_cel.Value = int(ean_waarde)
    else:
        ean_cel.NumberFormat = "@"
        ean_cel.Value = ean_waarde

    sheet.Range(f"{KOL_E_PRODUCT}{rij}").Value = regel["productnaam"]
    sheet.Range(f"{KOL_I}{rij}").Value = 0
    sheet.Range(f"{KOL_L}{rij}").Value = regel["kolom_l_waarde"]
    sheet.Range(f"{KOL_P_MECHANISME}{rij}").Value = regel["mechanisme_kort"]
    sheet.Range(f"{KOL_Q_VOLUME}{rij}").Value = regel["volume_excel"]

    _formatteer_artikelrij(sheet, rij)


def _zet_zwarte_onderrand(sheet, rij: int):
    rng = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}")
    rng.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
    rng.Borders(XL_EDGE_BOTTOM).Weight = XL_MEDIUM


_IGLO_VOETNOOT_PATROON = re.compile(r"\*{1,3}\s*Iglo", re.IGNORECASE)
_DIRK_PATROON = re.compile(r"\bDirk\b", re.IGNORECASE)


def _rij_tekst(sheet, rij: int) -> str:
    """Plakt de inhoud van alle cellen op een rij (kolom A t/m
    PRINT_LAATSTE_KOLOM) aan elkaar, zodat we per rij kunnen checken of er
    een *Iglo-voetnoottekst in staat (ongeacht in welke kolom precies)."""
    waarden = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}").Value
    if waarden is None:
        return ""
    cellen = waarden[0] if isinstance(waarden, tuple) else [waarden]
    return " ".join(str(c) for c in cellen if c not in (None, ""))


def _is_iglo_voetnoot(tekst: str) -> bool:
    # \*{1,3} matcht 1, 2 OF 3 sterretjes - dus *Iglo, **Iglo en ***Iglo
    # worden hier alle drie door herkend, ongeacht volgorde of wat er
    # verder nog tussen de voetnootrijen in staat.
    return bool(_IGLO_VOETNOOT_PATROON.search(tekst))


def _vervang_dirk_in_rij(sheet, rij: int, retailer_naam: str):
    """Het sjabloon is origineel voor supermarkt 'Dirk' gemaakt - in de
    (meestal **) voetnoottekst staat dat woord soms letterlijk. Vervang elk
    voorkomen van 'Dirk' in deze rij door de echte retailernaam."""
    rng = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}")
    waarden = rng.Value
    if waarden is None:
        return
    cellen = list(waarden[0]) if isinstance(waarden, tuple) else [waarden]
    aangepast = False
    for i, cel in enumerate(cellen):
        if isinstance(cel, str) and _DIRK_PATROON.search(cel):
            cellen[i] = _DIRK_PATROON.sub(retailer_naam, cel)
            aangepast = True
    if aangepast:
        rng.Value = (tuple(cellen),)
        print(f"  rij {rij}: 'Dirk' vervangen door '{retailer_naam}' in voetnoottekst.")


def _verwijder_overtollige_rijen(sheet, laatste_output_rij: int, sjabloon_rij: int, retailer_naam: str = None) -> int:
    """Verwijdert de oude placeholder-rijen tussen het laatst geschreven
    artikel en de sjabloon-/formulerij ECHT (met Delete, niet alleen
    leegmaken zoals voorheen). Een rij die een *Iglo / **Iglo / ***Iglo
    voetnoottekst bevat blijft staan; komt diezelfde voetnoottekst meerdere
    keren voor, dan blijft alleen de eerste keer staan en wordt de rest
    verwijderd. In de behouden voetnootrijen wordt 'Dirk' vervangen door de
    echte retailernaam, als die is meegegeven. Geeft de (na verwijdering
    opgeschoven) positie van de sjabloonrij terug."""
    if sjabloon_rij <= laatste_output_rij + 1:
        return sjabloon_rij

    geziene_voetnoten = set()
    te_verwijderen = []
    te_behouden = []

    for rij in range(laatste_output_rij + 1, sjabloon_rij):
        tekst = _rij_tekst(sheet, rij)
        if _is_iglo_voetnoot(tekst):
            sleutel = tekst.strip().lower()
            if sleutel in geziene_voetnoten:
                te_verwijderen.append(rij)  # duplicaat van een voetnoot -> ook weg
            else:
                geziene_voetnoten.add(sleutel)  # eerste keer -> deze rij laten staan
                te_behouden.append(rij)
        else:
            te_verwijderen.append(rij)

    if retailer_naam:
        for rij in te_behouden:
            _vervang_dirk_in_rij(sheet, rij, retailer_naam)

    # Van onder naar boven verwijderen - anders schuiven de rijnummers van
    # de nog te verwijderen rijen mee op tijdens het verwijderen.
    for rij in reversed(te_verwijderen):
        sheet.Rows(rij).Delete(Shift=XL_SHIFT_UP)

    print(f"{len(te_verwijderen)} oude placeholder-rij(en) verwijderd, "
          f"{len(geziene_voetnoten)} *Iglo-voetnoot(en) behouden "
          f"(rijen vóór verschuiving: {te_behouden}).")

    return sjabloon_rij - len(te_verwijderen)


def vul_sheet(sheet, retailer_cfg: dict, weken: list, week_groepen: list, tick=None) -> int:
    sheet.Range(FROZEN_FOOD_CEL).Value = retailer_cfg["categorie"]
    sheet.Range(TITEL_CEL).Value = retailer_cfg["titel"]
    sheet.Range(ACCOUNTMANAGER_CEL).Value = retailer_cfg["accountmanager"]
    datum_cel = sheet.Range(DATUM_CEL)
    datum_cel.Value = dt.datetime.combine(dt.date.today(), dt.time())
    datum_cel.NumberFormat = "dd/mm/yyyy"
    sheet.Range(Q_LABEL_CEL).Value = kwartaal_label(weken)

    aantal_regels = sum(len(groep) for _week, mech_groepen in week_groepen for _tekst, groep in mech_groepen)
    aantal_extra_rijen = sum(
        2 for _week, mech_groepen in week_groepen for tekst, _groep in mech_groepen if tekst
    )
    totaal_rijen = aantal_regels + aantal_extra_rijen

    # Vaste placeholder-ruimte tussen START_OUTPUT_RIJ en de sjabloon-/
    # formulerij (rijen 7 t/m 121 = 115 rijen). Als er meer rijen nodig zijn
    # dan dat (bv. een heel kwartaal met veel artikelen), voegen we het
    # tekort aan lege rijen in vlak vóór de sjabloonrij. Die rij (en alles
    # eronder, zoals een eventuele totalenregel) schuift daardoor automatisch
    # mee naar beneden - Excel past de formules in de meegeschoven rijen
    # daarbij zelf aan. Zo is er geen vaste bovengrens meer en gebeurt dit
    # alleen wanneer het echt nodig is; bij een 'normale' kleinere week-
    # selectie verandert er niets aan het bestaande gedrag.
    vaste_capaciteit = TEMPLATE_FORMULE_RIJ - START_OUTPUT_RIJ
    tekort = totaal_rijen - vaste_capaciteit

    if tekort > 0:
        sheet.Rows(f"{TEMPLATE_FORMULE_RIJ}:{TEMPLATE_FORMULE_RIJ + tekort - 1}").Insert(Shift=XL_SHIFT_DOWN)
        print(f"Let op: {totaal_rijen} rijen nodig voor week(en) {weken}, vaste sjabloonruimte was "
              f"{vaste_capaciteit} - {tekort} extra rij(en) ingevoegd vóór de sjabloonrij (rij "
              f"{TEMPLATE_FORMULE_RIJ}).")

    huidige_template_rij = TEMPLATE_FORMULE_RIJ + max(tekort, 0)
    template_rij = sheet.Rows(huidige_template_rij)

    rij = START_OUTPUT_RIJ

    for week, mech_groepen in week_groepen:
        eerste_product_van_week = True

        for groep_index, (mechanisme_tekst, groep_regels) in enumerate(mech_groepen):
            if mechanisme_tekst:
                template_rij.Copy()
                sheet.Rows(rij).PasteSpecial(Paste=XL_PASTE_ALL)
                _schrijf_witte_spacer_rij(sheet, rij, zwarte_bovenrand=(groep_index == 0))
                rij += 1
                if tick:
                    tick()

                template_rij.Copy()
                sheet.Rows(rij).PasteSpecial(Paste=XL_PASTE_ALL)
                _schrijf_mechanisme_rij(sheet, rij, mechanisme_tekst)
                rij += 1
                if tick:
                    tick()

            for regel in groep_regels:
                template_rij.Copy()
                sheet.Rows(rij).PasteSpecial(Paste=XL_PASTE_ALL)
                _schrijf_productrij(
                    sheet, rij, regel,
                    eerste_rij_van_week=eerste_product_van_week,
                    sheet_naam=sheet.Name,
                )
                eerste_product_van_week = False
                rij += 1
                if tick:
                    tick()

    laatste_rij = rij - 1
    sheet.Application.CutCopyMode = False

    if laatste_rij >= START_OUTPUT_RIJ:
        _zet_zwarte_onderrand(sheet, laatste_rij)
        nieuwe_sjabloon_rij = _verwijder_overtollige_rijen(
            sheet, laatste_rij, huidige_template_rij, retailer_naam=retailer_cfg["weergave_naam"]
        )
        # De behouden *Iglo-voetnootregels staan nu direct onder de
        # laatste artikelregel en horen bij de afdruk - neem ze mee in
        # laatste_rij zodat het printbereik ze ook bevat.
        laatste_rij = nieuwe_sjabloon_rij - 1

    sheet.PageSetup.PrintArea = f"$A$1:${PRINT_LAATSTE_KOLOM}${laatste_rij}"
    return laatste_rij


# ======================================================================
# ONDERDEEL 4: BESTAND BEHEER
# ----------------------------------------------------------------------
# Bronbestand vinden, Excel veilig opstarten, watchdog, veilig opslaan.
# GEEN CSV-logica, GEEN sheet-vul-logica.
# Los te draaien (alleen het bestand-zoek-deel) met:
#   python alles_in_een.py bestand --retailer Hoogvliet
# ======================================================================

BRON_MAP = Path("inzicht")
MSO_AUTOMATION_FORCE_DISABLE = 3
EXCEL_ZICHTBAAR = False


def vind_bestand(prefix: str, extensie: str, map_pad: Path = BRON_MAP) -> Path:
    patroon = str(map_pad / f"{prefix}*{extensie}")
    kandidaten = [Path(p) for p in glob.glob(patroon)]
    kandidaten = [p for p in kandidaten if not p.name.lower().startswith("oud")]

    if not kandidaten:
        raise FileNotFoundError(f"Geen bestand gevonden voor patroon: {patroon}")

    kandidaten.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    if len(kandidaten) > 1:
        print(f"Let op: meerdere bestanden gevonden voor '{prefix}*{extensie}'.")
        print(f"  Gekozen (nieuwste): {kandidaten[0].name}")
        print(f"  Genegeerd: {[p.name for p in kandidaten[1:]]}")
    return kandidaten[0]


def vind_focus_csv(retailer: str, jaar: int) -> Path:
    """Wrapper rond vind_bestand: zoekt de Promo Focus CSV puur op basis van
    --retailer + --jaar, zodat je nooit zelf het pad hoeft te typen."""
    return vind_bestand(f"Promo Focus File {jaar} {retailer}", ".csv")


def kill_orphan_excel():
    import psutil
    gevonden = 0
    for proc in psutil.process_iter(["pid", "name"]):
        if proc.info["name"] == "EXCEL.EXE":
            gevonden += 1
            try:
                proc.kill()
            except Exception:
                pass
    if gevonden:
        print(f"Opgeruimd: {gevonden} achtergebleven EXCEL.EXE-proces(sen).")


def open_excel_veilig():
    import psutil
    import win32com.client as win32

    pids_voor = {p.pid for p in psutil.process_iter(["pid", "name"]) if p.info["name"] == "EXCEL.EXE"}
    app = win32.DispatchEx("Excel.Application")
    app.Visible = EXCEL_ZICHTBAAR
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = MSO_AUTOMATION_FORCE_DISABLE

    pids_na = {p.pid for p in psutil.process_iter(["pid", "name"]) if p.info["name"] == "EXCEL.EXE"}
    nieuwe_pids = pids_na - pids_voor
    excel_pid = nieuwe_pids.pop() if len(nieuwe_pids) == 1 else None
    if excel_pid is None:
        print(f"WAARSCHUWING: kon het PID van het nieuwe Excel-proces niet eenduidig vinden "
              f"({len(nieuwe_pids)} kandidaten).")
    return app, excel_pid


class ExcelWatchdog:
    def __init__(self, excel_pid: int, stilte_timeout: int, check_interval: int = 5):
        self.excel_pid = excel_pid
        self.stilte_timeout = stilte_timeout
        self.check_interval = check_interval
        self.laatste_signaal = time.time()
        self._stop = False
        self._thread = threading.Thread(target=self._bewaak, daemon=True)

    def start(self):
        self._thread.start()
        return self

    def tick(self):
        self.laatste_signaal = time.time()

    def stop(self):
        self._stop = True

    def _bewaak(self):
        import psutil
        while not self._stop:
            time.sleep(self.check_interval)
            if self._stop:
                return
            stilte = time.time() - self.laatste_signaal
            if stilte > self.stilte_timeout:
                try:
                    proc = psutil.Process(self.excel_pid)
                    if proc.is_running():
                        print(f"WATCHDOG: geen voortgang in {int(stilte)}s (PID {self.excel_pid}) - hard afgesloten.")
                        proc.kill()
                except psutil.NoSuchProcess:
                    pass
                return


def is_bestand_vergrendeld(pad: Path) -> bool:
    try:
        with open(pad, "r+b"):
            return False
    except PermissionError:
        return True
    except FileNotFoundError:
        return False


def opslaan_met_retry(wb, output_pad: Path, pogingen: int = 5, wachttijd: float = 2.0):
    laatste_fout = None
    for poging in range(1, pogingen + 1):
        try:
            wb.Save()
            return
        except Exception as fout:
            laatste_fout = fout
            vergrendeld = is_bestand_vergrendeld(output_pad)
            print(f"Opslaan mislukt (poging {poging}/{pogingen}): {fout}")
            print(f"  Direct getest: bestand is {'WEL' if vergrendeld else 'NIET'} vergrendeld door iets anders.")
            if poging < pogingen:
                print(f"  Nieuwe poging over {wachttijd}s...")
                time.sleep(wachttijd)
    raise RuntimeError(
        f"Opslaan van {output_pad} is na {pogingen} pogingen mislukt. Laatste fout: {laatste_fout}."
    ) from laatste_fout


# ======================================================================
# ONDERDEEL 5: MAIN
# ----------------------------------------------------------------------
# Voegt onderdeel 1 t/m 4 samen. ALLEEN de simpele route: 1 doorlopend
# tabblad met alle weken erop (dus NIET Poiesz).
# Los te draaien met: python alles_in_een.py main --retailer Hoogvliet
# ======================================================================

DIRK_TEMPLATE_BESTAND = "Actievoorstellen.xlsx"
DIRK_SHEET_TEMPLATE = "Promoplan Dirk 2026"
OUTPUT_MAP = Path("output")

XL_CALCULATION_MANUAL = -4135
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
EXCEL_WATCHDOG_STILTE_TIMEOUT = 120

RETAILER_OVERRIDES = {
    "hoogvliet": {"titel": "Hoogvliet aktie-overzicht 2026", "accountmanager": "Ben"},
}


def bouw_retailer_cfg(retailer_naam: str, accountmanager_arg) -> dict:
    overrides = RETAILER_OVERRIDES.get(retailer_naam.lower(), {})
    return {
        "weergave_naam": retailer_naam,
        "titel": overrides.get("titel", f"{retailer_naam} aktie-overzicht 2026"),
        "accountmanager": accountmanager_arg or overrides.get("accountmanager", ""),
        "categorie": "Frozen Food",
    }


def laatste_iso_week(jaar: int) -> int:
    return dt.date(jaar, 12, 28).isocalendar()[1]


def bepaal_weken(weken_arg, jaar: int) -> list:
    if weken_arg:
        return sorted(int(w.strip()) for w in weken_arg.split(",") if w.strip())
    return list(range(40, laatste_iso_week(jaar) + 1))


def output_basisnaam(retailer_naam: str, weken: list, jaar: int) -> str:
    periode = f"wk{weken[0]}" if len(weken) == 1 else f"wk{min(weken)}-wk{max(weken)}"
    return f"{retailer_naam}_actievoorstel_{periode}_{jaar}"


def genereer(retailer_naam: str, weken_arg, jaar: int, toegestane_kolom_c, accountmanager_arg, debug: bool = False):
    cfg = bouw_retailer_cfg(retailer_naam, accountmanager_arg)

    template_path = Path(DIRK_TEMPLATE_BESTAND)
    if not template_path.exists():
        raise FileNotFoundError(f"Templatebestand niet gevonden: {template_path}")

    focus_pad = vind_focus_csv(retailer_naam, jaar)
    weken = bepaal_weken(weken_arg, jaar)

    print("Start actievoorstel maken (simpele route, 1 tabblad)")
    print(f"Retailer: {retailer_naam}  |  Jaar: {jaar}  |  Weken: {weken}")
    print(f"Template: {template_path}  |  Focus: {focus_pad.name}")

    focus_raw = laad_focus_data(focus_pad)
    print(f"  unieke weken in CSV: {sorted(focus_raw['week_int'].dropna().unique().tolist())}")

    toegestane_lijst = toegestane_kolom_c.split("|") if toegestane_kolom_c else []
    focus = filter_kolom_c(focus_raw, toegestane_excepties=toegestane_lijst)
    focus = focus[focus["week_int"].isin(weken)]
    print(f"  na week-selectie (gevraagd: {weken}): {len(focus)} rijen")

    alle_regels = bouw_outputregels(focus, debug=debug)
    voeg_advies_kruisjes_toe(alle_regels)
    week_groepen = groepeer_per_week_en_mechanisme(alle_regels)
    print(f"Aantal artikelen: {len(alle_regels)}")
    print(f"Aantal advies-afbeelding kruisjes: {sum(1 for r in alle_regels if r['advies_afbeelding'])}")

    if not alle_regels:
        print("WAARSCHUWING: geen regels gevonden voor deze weken/retailer - geen Excel-bestand aangemaakt.")
        return

    OUTPUT_MAP.mkdir(exist_ok=True)
    output_basis = output_basisnaam(retailer_naam, weken, jaar)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_pad = OUTPUT_MAP / f"{output_basis}_{timestamp}.xlsx"
    shutil.copy2(template_path, output_pad)

    app, excel_pid = open_excel_veilig()
    watchdog = None
    if excel_pid:
        watchdog = ExcelWatchdog(excel_pid, EXCEL_WATCHDOG_STILTE_TIMEOUT).start()
        print(f"Watchdog actief op PID {excel_pid} (grijpt in na {EXCEL_WATCHDOG_STILTE_TIMEOUT}s zonder voortgang).")

    wb = None
    try:
        wb = app.Workbooks.Open(str(output_pad.resolve()), UpdateLinks=0)
        app.Calculation = XL_CALCULATION_MANUAL
        sheet = vind_template_sheet(wb, DIRK_SHEET_TEMPLATE)

        laatste_rij = vul_sheet(sheet, cfg, weken, week_groepen, tick=watchdog.tick if watchdog else None)
        print(f"{len(alle_regels)} regel(s) geschreven naar '{sheet.Name}', t/m rij {laatste_rij}.")

        opslaan_met_retry(wb, output_pad)

        pdf_pad = OUTPUT_MAP / f"{output_basis}_{timestamp}.pdf"
        sheet.ExportAsFixedFormat(
            Type=XL_TYPE_PDF, Filename=str(pdf_pad.resolve()), Quality=XL_QUALITY_STANDARD,
            IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False,
        )
        print(f"PDF: {pdf_pad}")
        print(f"Klaar: {output_pad}")

    finally:
        if watchdog is not None:
            watchdog.stop()
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass
        del app


# ======================================================================
# CLI - subcommando's per onderdeel
# ======================================================================

def _resolve_csv_pad(args) -> Path:
    """Gemeenschappelijke logica: --csv heeft voorrang, anders --retailer +
    --jaar gebruiken om het pad automatisch te vinden."""
    if args.csv:
        return Path(args.csv)
    if not args.retailer:
        print("Geef --retailer (en optioneel --jaar) of --csv <pad> op.")
        sys.exit(1)
    return vind_focus_csv(args.retailer, args.jaar)


def cmd_csv(args):
    pad = _resolve_csv_pad(args)
    df = laad_focus_data(pad)
    print(f"\nUnieke weken in de CSV: {sorted(df['week_int'].dropna().unique().tolist())}")
    df_gefilterd = filter_kolom_c(df)
    print(f"Na kolom-C-filter: {len(df_gefilterd)} rijen")
    print("\nEerste 5 rijen ter controle:")
    print(df_gefilterd.head(5).to_string())


def cmd_regels(args):
    pad = _resolve_csv_pad(args)
    focus = filter_kolom_c(laad_focus_data(pad))
    regels = bouw_outputregels(focus, debug=args.debug)
    voeg_advies_kruisjes_toe(regels)

    print(f"\nAantal artikelen: {len(regels)}")
    print(f"Aantal advies-afbeelding kruisjes: {sum(1 for r in regels if r['advies_afbeelding'])}")

    print("\n=== Geneste groepering (week -> mechanisme -> regels) ===")
    for week, mech_groepen in groepeer_per_week_en_mechanisme(regels):
        print(f"Week {week}:")
        for mech_tekst, groep_regels in mech_groepen:
            namen = [r["productnaam"] for r in groep_regels]
            print(f"  mechanisme={mech_tekst!r} -> {namen}")


def cmd_bestand(args):
    try:
        if args.retailer:
            gevonden = vind_focus_csv(args.retailer, args.jaar)
        else:
            gevonden = vind_bestand(args.prefix, args.extensie)
        print(f"\nGevonden: {gevonden}")
    except FileNotFoundError as e:
        print(f"\n{e}")


def cmd_main(args):
    if not args.skip_cleanup:
        kill_orphan_excel()
    try:
        genereer(args.retailer.strip(), args.weken, args.jaar, args.toegestane_kolom_c,
                 args.accountmanager, debug=args.debug)
    except Exception as error:
        print("\nFOUT:")
        print(error)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="Actievoorstel-script, per onderdeel los te draaien")
    sub = parser.add_subparsers(dest="onderdeel", required=True)

    def voeg_csv_opties_toe(p):
        p.add_argument("--retailer", default=None, help="Bv. Hoogvliet - zoekt het CSV-bestand automatisch.")
        p.add_argument("--jaar", type=int, default=dt.date.today().year)
        p.add_argument("--csv", default=None, help="Direct pad naar de CSV, overschrijft --retailer.")

    p_csv = sub.add_parser("csv", help="Onderdeel 1: CSV inlezen/filteren testen")
    voeg_csv_opties_toe(p_csv)
    p_csv.set_defaults(func=cmd_csv)

    p_regels = sub.add_parser("regels", help="Onderdeel 2: regels bouwen testen")
    voeg_csv_opties_toe(p_regels)
    p_regels.add_argument("--debug", action="store_true",
                           help="Print per markerrij de brontekst naast wat kolom E/P ervan maken.")
    p_regels.set_defaults(func=cmd_regels)

    p_bestand = sub.add_parser("bestand", help="Onderdeel 4: bestand-zoeklogica testen")
    p_bestand.add_argument("--retailer", default=None)
    p_bestand.add_argument("--jaar", type=int, default=dt.date.today().year)
    p_bestand.add_argument("--prefix", default=None, help="Als je geen --retailer gebruikt: prefix direct opgeven.")
    p_bestand.add_argument("--extensie", default=".csv")
    p_bestand.set_defaults(func=cmd_bestand)

    p_main = sub.add_parser("main", help="Onderdeel 5: het hele proces, schrijft echt een Excel-bestand")
    p_main.add_argument("--retailer", required=True)
    p_main.add_argument("--jaar", type=int, default=dt.date.today().year)
    p_main.add_argument("--weken", default=None, help="Komma-gescheiden, bv. 40,41. Standaard: hele Q4.")
    p_main.add_argument("--toegestane-kolom-c", dest="toegestane_kolom_c", default=None)
    p_main.add_argument("--accountmanager", default=None)
    p_main.add_argument("--skip-cleanup", action="store_true")
    p_main.add_argument("--debug", action="store_true",
                         help="Print per markerrij de brontekst naast wat kolom E/P ervan maken.")
    p_main.set_defaults(func=cmd_main)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
