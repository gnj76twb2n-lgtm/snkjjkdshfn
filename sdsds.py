"""
csv_lezen.py
-------------
Onderdeel 1/5: leest en filtert de Promo Focus CSV. Puur Python/pandas,
GEEN Excel, GEEN mechanisme-buffering - dat zit in regels_bouwen.py.

Test dit bestand op zichzelf met:
    python csv_lezen.py "inzicht/Promo Focus File 2026 Hoogvliet.csv"
"""

import sys
import html
from pathlib import Path

import pandas as pd


# Welke kolomletter hoort bij welk veld in de Promo Focus CSV.
FOCUS_KOLOMMEN = {
    "week": "A",
    "kolom_c": "C",          # filter: delist/gesaneerd/sanering
    "mech_d": "D",            # mechanisme-marker, kolom D
    "mech_e": "E",            # mechanisme-marker, kolom E (fallback)
    "sap_code": "F",          # SAP-code, of "NPD" erin
    "ean": "G",
    "artikelnaam": "H",
    "kolom_l_bron": "O",      # bron voor kolom L in de output
    "volume": "W",
}

KOLOM_C_UITSLUITEN = ["delist", "gesaneerd", "sanering"]


def excel_col_to_index(letter: str) -> int:
    """'A' -> 0, 'F' -> 5, 'W' -> 22 (0-indexed kolomnummer)."""
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
        return float(value.replace(",", ""))   # mogelijke duizendtal-komma wegstrippen
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
    """EAN-13-controlecijfer (GS1). Andere lengte -> True (geen vals alarm)."""
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
    """Telt ',' vs ';' in de eerste regel en kiest de meest voorkomende."""
    with open(pad, "r", encoding="utf-8-sig") as f:
        eerste_regel = f.readline()
    aantal_komma = eerste_regel.count(",")
    aantal_puntkomma = eerste_regel.count(";")
    gekozen = ";" if aantal_puntkomma > aantal_komma else ","
    print(f"CSV-separator gedetecteerd: '{gekozen}' (komma's: {aantal_komma}, puntkomma's: {aantal_puntkomma})")
    return gekozen


def laad_focus_data(pad: Path) -> pd.DataFrame:
    """Leest de hele CSV in en geeft een schone DataFrame terug, 1 rij per
    CSV-regel (zowel productregels als mechanisme-markerrijen)."""
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

    is_marker = (df["sap_code_raw"] == "") & ((df["mech_d"] != "") | (df["mech_e"] != ""))
    print(f"  waarvan {int(is_marker.sum())} mechanisme-markerrij(en) gevonden (geen sap-code, tekst in D/E)")

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


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Gebruik: python csv_lezen.py <pad naar Promo Focus CSV>")
        sys.exit(1)

    pad = Path(sys.argv[1])
    df = laad_focus_data(pad)
    print(f"\nUnieke weken in de CSV: {sorted(df['week_int'].dropna().unique().tolist())}")

    df_gefilterd = filter_kolom_c(df)
    print(f"Na kolom-C-filter: {len(df_gefilterd)} rijen")

    print("\nEerste 5 rijen ter controle:")
    print(df_gefilterd.head(5).to_string())


--------------

"""
regels_bouwen.py
------------------
Onderdeel 2/5: bouwt de outputregels uit de gefilterde Promo Focus-data.
Mechanisme-tekstherkenning, buffering (producten koppelen aan hun
markerrij), groepering per week + mechanisme. Puur Python, GEEN Excel.

Hangt af van csv_lezen.py (moet in dezelfde map staan).

Test dit bestand op zichzelf met:
    python regels_bouwen.py "inzicht/Promo Focus File 2026 Hoogvliet.csv"
"""

import re
import sys
from pathlib import Path

from csv_lezen import (
    clean_text,
    is_npd_code,
    laad_focus_data,
    filter_kolom_c,
    volume_to_excel_value,
    volume_as_float,
)


ADVIES_AFBEELDING_AANTAL = 2   # top-N op prognose-volume per week


# ====================================================================
# Mechanisme-tekst formattering
# ====================================================================

def normaliseer_mechanisme_basis(mechanisme) -> str:
    tekst = clean_text(mechanisme)
    if not tekst:
        return ""
    tekst = re.sub(r"^\s*\d+\.\s*", "", tekst)
    tekst = re.sub(r"\bSPO\b", "1 voor", tekst, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", tekst).strip()


_MECHANISME_PATRONEN = [
    r"\b\d+\s*\+\s*\d+\s+gratis\b",
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
    """Volledig uitgeschreven tekst voor de gele mechanisme-rij, met Iglo-prefix."""
    tekst = normaliseer_mechanisme_basis(mechanisme)
    if not tekst:
        return ""
    tekst = re.sub(r"\b(\d+)\s*\+\s*(\d+)\s+gratis\b", r"\1 + \2 gratis", tekst, flags=re.IGNORECASE)
    tekst = re.sub(r"\s*%\s*", "% ", tekst)
    tekst = re.sub(r"\s+", " ", tekst).strip()
    if not tekst.lower().startswith("iglo "):
        tekst = "Iglo " + tekst
    return tekst


def formatteer_actiemechanisme_kolom_p(mechanisme) -> str:
    """Kort mechanisme voor kolom P, bv. '25% korting' of '1 + 1 gratis'."""
    kort = extraheer_mechanisme_tekst(mechanisme)
    if kort:
        return kort
    tekst = normaliseer_mechanisme_basis(mechanisme)
    if not tekst:
        return ""
    return re.sub(r"^\s*Iglo\s+", "", tekst, flags=re.IGNORECASE).strip()


# ====================================================================
# Regels bouwen (mechanisme-marker-buffering)
# ====================================================================

def is_mechanisme_marker(rij) -> bool:
    """Herkent de mechanisme-aankondigingsrij: geen sap-code, maar wel tekst
    in de (oorspronkelijk samengevoegde) D/E-kolommen."""
    return rij["sap_code_raw"] == "" and bool(rij["mech_d"] or rij["mech_e"])


def bouw_outputregels(focus) -> list:
    """Verwerkt de Promo Focus-rijen OP VOLGORDE. Producten worden gebufferd
    tot de afsluitende mechanisme-markerrij van hun sectie verschijnt; dan
    krijgen ze allemaal die mechanismetekst. Een sectie zonder afsluitende
    marker levert producten zonder mechanisme op - die worden niet
    weggegooid, wel gemeld."""
    regels, buffer = [], []
    overgeslagen_geen_titel = 0

    for _, rij in focus.iterrows():
        if is_mechanisme_marker(rij):
            mechanisme_ruw = rij["mech_d"] or rij["mech_e"]
            mechanisme_vol = formatteer_actiemechanisme_geel(mechanisme_ruw)
            mechanisme_kort = formatteer_actiemechanisme_kolom_p(mechanisme_ruw)
            for regel in buffer:
                regel["mechanisme_vol"] = mechanisme_vol
                regel["mechanisme_kort"] = mechanisme_kort
                regels.append(regel)
            buffer = []
            continue

        sap_raw = rij["sap_code_raw"]
        is_npd = is_npd_code(sap_raw)
        productnaam = "NPD" if is_npd else rij["artikelnaam"]

        if not productnaam and not is_npd:
            overgeslagen_geen_titel += 1
            continue   # geen titel + geen NPD = niet meenemen

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

    regels.extend(buffer)   # sectie(s) zonder afsluitende marker: toch meenemen

    zonder_mechanisme = [r for r in regels if not r["mechanisme_kort"] and r["productnaam"] != "NPD"]
    if zonder_mechanisme:
        print(f"WAARSCHUWING: {len(zonder_mechanisme)} product(en) zonder mechanismetekst "
              f"(geen afsluitende markerrij gevonden voor hun sectie). Weken: "
              f"{sorted({r['week'] for r in zonder_mechanisme})}")

    if overgeslagen_geen_titel:
        print(f"{overgeslagen_geen_titel} regel(s) overgeslagen: geen productnaam en geen NPD.")
    print(f"bouw_outputregels: {len(regels)} productregel(s) opgebouwd uit {len(focus)} focus-rijen.")

    return regels


def voeg_advies_kruisjes_toe(regels: list, aantal: int = ADVIES_AFBEELDING_AANTAL) -> None:
    """Annoteert elke regel met advies_afbeelding (True/False): top-N op
    volume, per week. Bij gelijk volume: productnaam dan EAN als tiebreaker."""
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
    """Groepeert EERST per week (volgorde van eerste voorkomen), DAARBINNEN
    per mechanisme. Geeft: [(week, [(mechanisme_tekst, [regels]), ...]), ...]"""
    per_week, volgorde_weken = {}, []
    for regel in regels:
        week = regel["week"]
        if week not in per_week:
            per_week[week] = []
            volgorde_weken.append(week)
        per_week[week].append(regel)
    return [(week, groepeer_per_mechanisme(per_week[week])) for week in volgorde_weken]


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Gebruik: python regels_bouwen.py <pad naar Promo Focus CSV>")
        sys.exit(1)

    pad = Path(sys.argv[1])
    focus = filter_kolom_c(laad_focus_data(pad))
    regels = bouw_outputregels(focus)
    voeg_advies_kruisjes_toe(regels)

    print(f"\nAantal artikelen: {len(regels)}")
    print(f"Aantal advies-afbeelding kruisjes: {sum(1 for r in regels if r['advies_afbeelding'])}")

    print("\n=== Geneste groepering (week -> mechanisme -> regels) ===")
    for week, mech_groepen in groepeer_per_week_en_mechanisme(regels):
        print(f"Week {week}:")
        for mech_tekst, groep_regels in mech_groepen:
            namen = [r["productnaam"] for r in groep_regels]
            print(f"  mechanisme={mech_tekst!r} -> {namen}")




--------------

"""
excel_schrijven.py
--------------------
Onderdeel 3/5: vult EEN Excel-sheet met de al-voorbereide data uit
regels_bouwen.py (de geneste week->mechanisme->regels structuur). Dit
bestand doet GEEN CSV-lezen, GEEN mechanisme-logica - puur Excel/COM.

Layout per week-sectie:
- Witte spacer-rij (hoogte 8, blauwe J-cel), zwarte bovenrand bij de EERSTE
  groep van DIE WEEK (markeert het begin van een nieuwe week-sectie).
- Gele mechanisme-rij: volledige tekst alleen in kolom E geel, J blauw.
- "V" (kolom A) + weeknummer (kolom B): ALLEEN op de eerste productrij van
  ELKE WEEK, niet alleen de allereerste rij van de hele sheet.
- Kolom D: EAN als ECHT GETAL (platte weergave), niet als tekst - anders
  matchen de EAN-XLOOKUP-formules in de template niet.
- Kolommen F/G/H/J/K/M/R: NIET aangeraakt (EAN-formule-gedreven).
- Kolom I: vast leeg/0. Kolom L: uit regel["kolom_l_waarde"].
- Kolom P: kort mechanisme. Kolom Q: volume. Kolom C: "X" bij advies-afbeelding.
- Restanten van de template na de laatste rij: LEEGGEMAAKT (ClearContents),
  NIET fysiek verwijderd (Rows().Delete() bleek save-fouten te veroorzaken,
  vermoedelijk door interactie met merged cells/named ranges elders in de
  template - ClearContents laat de rijstructuur intact, geen verschuiving).
- Printgebied wordt hard gezet op A:R, t/m de laatste echte outputrij.

Dit bestand heeft GEEN __main__-testblok dat los kan draaien, want het
heeft een echte Excel-installatie + een geopende workbook/sheet nodig.
Test het via main_simpel.py (bestand 5).
"""

import datetime as dt


# ====================================================================
# CONFIG - celposities en kolomletters
# ====================================================================

TEMPLATE_FORMULE_RIJ = 122    # rij met de echte opmaak + EAN-formules
START_OUTPUT_RIJ = 7
PRINT_LAATSTE_KOLOM = "R"

TITEL_CEL = "F1"
FROZEN_FOOD_CEL = "B1"
ACCOUNTMANAGER_CEL = "B2"
DATUM_CEL = "B3"
Q_LABEL_CEL = "A6"             # al gemerged in de template (A6:R6)

KOL_A_DVIP = "A"
KOL_B_WEEK = "B"
KOL_C_ADVIES_AFBEELDING = "C"
KOL_D_EAN = "D"
KOL_E_PRODUCT = "E"
KOL_I = "I"
KOL_L = "L"
KOL_J_ACTIE_INKOOPPRIJS = "J"   # alleen voor blauwe styling op spacer/mechanisme-rijen
KOL_P_MECHANISME = "P"
KOL_Q_VOLUME = "Q"

KOLOM_A_VASTE_WAARDE = "V"
KOLOM_C_VASTE_WAARDE = "X"

EAN_FORMULE_KOLOMMEN = ["F", "G", "H", "J", "K", "M", "R"]

KLEUR_BLAUW = 15257527
KLEUR_GEEL = 10092543

# Excel/COM-constanten (letterlijke waarden, geen gencache/EnsureDispatch nodig)
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


# ====================================================================
# Helpers
# ====================================================================

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
    try:
        rng.Borders(XL_EDGE_TOP).LineStyle = 0
        rng.Borders(XL_EDGE_BOTTOM).LineStyle = 0
        rng.Borders(XL_INSIDE_HORIZONTAL).LineStyle = 0
    except Exception:
        pass
    try:
        rng.Borders(XL_EDGE_LEFT).LineStyle = XL_CONTINUOUS
        rng.Borders(XL_EDGE_LEFT).Weight = XL_THIN
        rng.Borders(XL_EDGE_RIGHT).LineStyle = XL_CONTINUOUS
        rng.Borders(XL_EDGE_RIGHT).Weight = XL_THIN
        rng.Borders(XL_INSIDE_VERTICAL).LineStyle = XL_CONTINUOUS
        rng.Borders(XL_INSIDE_VERTICAL).Weight = XL_THIN
    except Exception:
        pass


def _schrijf_witte_spacer_rij(sheet, rij: int, zwarte_bovenrand: bool):
    rng = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}")
    rng.ClearContents()
    rng.Interior.ColorIndex = 2   # wit

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
    rng.Interior.ColorIndex = 2   # wit, alleen E wordt geel

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
    """Checkt of de verwachte EAN-gedreven formules nog aanwezig zijn na het
    kopieren van TEMPLATE_FORMULE_RIJ. Puur een waarschuwing."""
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
        ean_cel.NumberFormat = "0"   # plat getal, geen wetenschappelijke notatie
        ean_cel.Value = int(ean_waarde)
    else:
        ean_cel.NumberFormat = "@"
        ean_cel.Value = ean_waarde

    sheet.Range(f"{KOL_E_PRODUCT}{rij}").Value = regel["productnaam"]
    # Kolommen F/G/H/J/K/M/R NIET aanraken: EAN-formule-gedreven.
    sheet.Range(f"{KOL_I}{rij}").Value = 0
    sheet.Range(f"{KOL_L}{rij}").Value = regel["kolom_l_waarde"]
    sheet.Range(f"{KOL_P_MECHANISME}{rij}").Value = regel["mechanisme_kort"]
    sheet.Range(f"{KOL_Q_VOLUME}{rij}").Value = regel["volume_excel"]

    _formatteer_artikelrij(sheet, rij)


def _zet_zwarte_onderrand(sheet, rij: int):
    rng = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}")
    rng.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
    rng.Borders(XL_EDGE_BOTTOM).Weight = XL_MEDIUM


def _maak_overtollige_rijen_leeg(sheet, laatste_output_rij: int):
    """Maakt restanten van de template LEEG (ClearContents), verwijdert ze
    NIET fysiek - dat laatste bleek save-fouten te veroorzaken."""
    used_range = sheet.UsedRange
    laatste_gebruikte_rij = used_range.Row + used_range.Rows.Count - 1
    bovengrens = min(laatste_gebruikte_rij, TEMPLATE_FORMULE_RIJ)
    if bovengrens > laatste_output_rij:
        bereik = sheet.Range(f"A{laatste_output_rij + 1}:{PRINT_LAATSTE_KOLOM}{bovengrens}")
        bereik.ClearContents()
        bereik.Interior.ColorIndex = 2
        try:
            bereik.Borders(XL_EDGE_TOP).LineStyle = 0
            bereik.Borders(XL_EDGE_BOTTOM).LineStyle = 0
            bereik.Borders(XL_EDGE_LEFT).LineStyle = 0
            bereik.Borders(XL_EDGE_RIGHT).LineStyle = 0
            bereik.Borders(XL_INSIDE_HORIZONTAL).LineStyle = 0
            bereik.Borders(XL_INSIDE_VERTICAL).LineStyle = 0
        except Exception:
            pass


# ====================================================================
# Hoofdfunctie: vult 1 sheet met de geneste week->mechanisme->regels data
# ====================================================================

def vul_sheet(sheet, retailer_cfg: dict, weken: list, week_groepen: list, tick=None) -> int:
    """week_groepen = output van regels_bouwen.groepeer_per_week_en_mechanisme().
    tick: optionele callback, aangeroepen na elke geschreven rij (voor een
    watchdog - zie bestand_beheer.py)."""
    sheet.Range(FROZEN_FOOD_CEL).Value = retailer_cfg["categorie"]
    sheet.Range(TITEL_CEL).Value = retailer_cfg["titel"]
    sheet.Range(ACCOUNTMANAGER_CEL).Value = retailer_cfg["accountmanager"]
    datum_cel = sheet.Range(DATUM_CEL)
    datum_cel.Value = dt.datetime.combine(dt.date.today(), dt.time())
    datum_cel.NumberFormat = "dd/mm/yyyy"
    sheet.Range(Q_LABEL_CEL).Value = kwartaal_label(weken)

    template_rij = sheet.Rows(TEMPLATE_FORMULE_RIJ)

    aantal_regels = sum(len(groep) for _week, mech_groepen in week_groepen for _tekst, groep in mech_groepen)
    aantal_extra_rijen = sum(
        2 for _week, mech_groepen in week_groepen for tekst, _groep in mech_groepen if tekst
    )
    totaal_rijen = aantal_regels + aantal_extra_rijen

    if START_OUTPUT_RIJ + totaal_rijen - 1 >= TEMPLATE_FORMULE_RIJ:
        raise ValueError(
            f"Te veel rijen ({totaal_rijen}) voor week(en) {weken} - "
            f"dit zou rij {TEMPLATE_FORMULE_RIJ} overschrijven."
        )

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
        _maak_overtollige_rijen_leeg(sheet, laatste_rij)

    sheet.PageSetup.PrintArea = f"$A$1:${PRINT_LAATSTE_KOLOM}${laatste_rij}"
    return laatste_rij

-------------------

"""
bestand_beheer.py
-------------------
Onderdeel 4/5: alles rond bestanden en het Excel-PROCES zelf - bronbestand
vinden, Excel veilig opstarten, een hartslag-watchdog, veilig opslaan met
retry. GEEN CSV-logica, GEEN sheet-vul-logica (dat zit in excel_schrijven.py).

Dit bestand heeft een Windows-omgeving met Excel/pywin32 nodig om echt te
draaien. De bestand-zoek-functie (vind_bestand) is wel los te testen.

Test het bestand-zoek-deel met:
    python bestand_beheer.py "Promo Focus File 2026 Hoogvliet" ".csv"
"""

import glob
import sys
import threading
import time
from pathlib import Path

import psutil
import win32com.client as win32


BRON_MAP = Path("inzicht")

MSO_AUTOMATION_FORCE_DISABLE = 3
EXCEL_ZICHTBAAR = False


# ====================================================================
# Bronbestand vinden
# ====================================================================

def vind_bestand(prefix: str, extensie: str, map_pad: Path = BRON_MAP) -> Path:
    """Zoekt naar bestanden die beginnen met `prefix`. Bestanden die beginnen
    met 'oud' worden genegeerd. Bij meerdere kandidaten wordt de meest recent
    gewijzigde gekozen, met melding welke gekozen/genegeerd zijn."""
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


# ====================================================================
# Excel-proces veilig starten/opruimen
# ====================================================================

def kill_orphan_excel():
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
    """Start een verse, geisoleerde Excel-instantie. Macro's staan
    gedwongen uit (Workbook_Open vuurt nooit af). Geeft (app, pid) terug -
    het pid wordt gevonden door de procestabel voor/na te vergelijken,
    betrouwbaarder dan app.Hwnd (dat is een vensterhandle, geen PID)."""
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
    """Hartslag-watchdog: doodt het Excel-proces alleen als er STILTE_TIMEOUT
    seconden GEEN voortgang is gemeld via tick(), niet na een vaste totale
    looptijd. Voorkomt dat een legitiem trage (maar voortgaande) run
    halverwege wordt afgebroken."""

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


# ====================================================================
# Veilig opslaan
# ====================================================================

def is_bestand_vergrendeld(pad: Path) -> bool:
    try:
        with open(pad, "r+b"):
            return False
    except PermissionError:
        return True
    except FileNotFoundError:
        return False


def opslaan_met_retry(wb, output_pad: Path, pogingen: int = 5, wachttijd: float = 2.0):
    """wb.Save() kan transient falen (bv. OneDrive-sync vlak na het kopieren).
    Probeert een paar keer opnieuw, met een directe lock-test voor concrete
    diagnose bij elke mislukte poging."""
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


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Gebruik: python bestand_beheer.py <prefix> <extensie>")
        print('Voorbeeld: python bestand_beheer.py "Promo Focus File 2026 Hoogvliet" ".csv"')
        sys.exit(1)

    prefix, extensie = sys.argv[1], sys.argv[2]
    try:
        gevonden = vind_bestand(prefix, extensie)
        print(f"\nGevonden: {gevonden}")
    except FileNotFoundError as e:
        print(f"\n{e}")

---------------------------
"""
main_simpel.py
----------------
Onderdeel 5/5: voegt csv_lezen.py, regels_bouwen.py, excel_schrijven.py en
bestand_beheer.py samen. ALLEEN de eenvoudige route: 1 doorlopend tabblad
met alle weken erop (dus NIET Poiesz - die heeft een apart tabblad per
week, dat bouwen we apart als dit fundament eenmaal goed werkt).

Gebruik:
    python main_simpel.py --retailer Hoogvliet
    python main_simpel.py --retailer Hoogvliet --weken 40,41
    python main_simpel.py --retailer Hoogvliet --jaar 2026 --accountmanager Ben

Vereist in dezelfde map: csv_lezen.py, regels_bouwen.py, excel_schrijven.py,
bestand_beheer.py, en het bestand Actievoorstellen.xlsx.
"""

import argparse
import datetime as dt
import shutil
import sys
from pathlib import Path

from csv_lezen import laad_focus_data, filter_kolom_c
from regels_bouwen import bouw_outputregels, voeg_advies_kruisjes_toe, groepeer_per_week_en_mechanisme
from excel_schrijven import vul_sheet, vind_template_sheet
from bestand_beheer import (
    vind_bestand, kill_orphan_excel, open_excel_veilig,
    ExcelWatchdog, opslaan_met_retry,
)


DIRK_TEMPLATE_BESTAND = "Actievoorstellen.xlsx"
DIRK_SHEET_TEMPLATE = "Promoplan Dirk 2026"
OUTPUT_MAP = Path("output")

XL_CALCULATION_MANUAL = -4135
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0

EXCEL_WATCHDOG_STILTE_TIMEOUT = 120

# Bekende winkels: bevestigde titel/accountmanager. Onbekende winkelnamen
# krijgen automatisch een generieke titel en kunnen --accountmanager
# meekrijgen.
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
    return list(range(40, laatste_iso_week(jaar) + 1))   # standaard: hele Q4


def output_basisnaam(retailer_naam: str, weken: list, jaar: int) -> str:
    periode = f"wk{weken[0]}" if len(weken) == 1 else f"wk{min(weken)}-wk{max(weken)}"
    return f"{retailer_naam}_actievoorstel_{periode}_{jaar}"


def genereer(retailer_naam: str, weken_arg, jaar: int, toegestane_kolom_c, accountmanager_arg):
    cfg = bouw_retailer_cfg(retailer_naam, accountmanager_arg)

    template_path = Path(DIRK_TEMPLATE_BESTAND)
    if not template_path.exists():
        raise FileNotFoundError(f"Templatebestand niet gevonden: {template_path}")

    focus_pad = vind_bestand(f"Promo Focus File {jaar} {retailer_naam}", ".csv")
    weken = bepaal_weken(weken_arg, jaar)

    print("Start actievoorstel maken (simpele route, 1 tabblad)")
    print(f"Retailer: {retailer_naam}  |  Jaar: {jaar}  |  Weken: {weken}")
    print(f"Template: {template_path}  |  Focus: {focus_pad.name}")

    # ---- Stap 1+2: data inlezen en regels bouwen (geen Excel) ----
    focus_raw = laad_focus_data(focus_pad)
    print(f"  unieke weken in CSV: {sorted(focus_raw['week_int'].dropna().unique().tolist())}")

    toegestane_lijst = toegestane_kolom_c.split("|") if toegestane_kolom_c else []
    focus = filter_kolom_c(focus_raw, toegestane_excepties=toegestane_lijst)
    focus = focus[focus["week_int"].isin(weken)]
    print(f"  na week-selectie (gevraagd: {weken}): {len(focus)} rijen")

    alle_regels = bouw_outputregels(focus)
    voeg_advies_kruisjes_toe(alle_regels)
    week_groepen = groepeer_per_week_en_mechanisme(alle_regels)
    print(f"Aantal artikelen: {len(alle_regels)}")
    print(f"Aantal advies-afbeelding kruisjes: {sum(1 for r in alle_regels if r['advies_afbeelding'])}")

    if not alle_regels:
        print("WAARSCHUWING: geen regels gevonden voor deze weken/retailer - geen Excel-bestand aangemaakt.")
        return

    # ---- Stap 3+4: Excel openen, vullen, opslaan ----
    OUTPUT_MAP.mkdir(exist_ok=True)
    output_basis = output_basisnaam(retailer_naam, weken, jaar)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_pad = OUTPUT_MAP / f"{output_basis}_{timestamp}.xlsx"   # uniek per run, voorkomt naamsbotsingen
    shutil.copy2(template_path, output_pad)

    app, excel_pid = open_excel_veilig()
    watchdog = None
    if excel_pid:
        watchdog = ExcelWatchdog(excel_pid, EXCEL_WATCHDOG_STILTE_TIMEOUT).start()
        print(f"Watchdog actief op PID {excel_pid} (grijpt in na {EXCEL_WATCHDOG_STILTE_TIMEOUT}s zonder voortgang).")

    wb = None
    try:
        wb = app.Workbooks.Open(str(output_pad.resolve()), UpdateLinks=0)
        app.Calculation = XL_CALCULATION_MANUAL   # voorkomt recalc-cascade per EAN-write
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


def main():
    parser = argparse.ArgumentParser(description="Genereer actievoorstel (simpele route, 1 tabblad)")
    parser.add_argument("--retailer", required=True)
    parser.add_argument("--jaar", type=int, default=dt.date.today().year)
    parser.add_argument("--weken", default=None, help="Komma-gescheiden, bv. 40,41. Standaard: hele Q4.")
    parser.add_argument("--toegestane-kolom-c", dest="toegestane_kolom_c", default=None)
    parser.add_argument("--accountmanager", default=None)
    parser.add_argument("--skip-cleanup", action="store_true")
    args = parser.parse_args()

    if not args.skip_cleanup:
        kill_orphan_excel()

    try:
        genereer(args.retailer.strip(), args.weken, args.jaar, args.toegestane_kolom_c, args.accountmanager)
    except Exception as error:
        print("\nFOUT:")
        print(error)
        sys.exit(1)


if __name__ == "__main__":
    main()
