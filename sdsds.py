import argparse
import datetime as dt
import glob
import html
import re
import shutil
import sys
import time
from pathlib import Path

import pandas as pd
import psutil
import win32com.client as win32


# ====================================================================
# CONFIG
# ====================================================================

BRON_MAP = Path("inzicht")
OUTPUT_MAP = Path("output")

DIRK_TEMPLATE_BESTAND = "Actievoorstellen.xlsm"   # zelfde bestand voor elke retailer

# BEVESTIGDE REGEL: alle retailers gebruiken Actievoorstellen.xlsm als
# startbestand. Alleen Poiesz krijgt een apart tabblad per week; elke
# andere (huidige of toekomstige) retailer gebruikt 1 doorlopend tabblad.
RETAILERS = {
    "poiesz": {
        "weergave_naam": "Poiesz",
        "sheet_template": "Promoplan Dirk 2026",
        "titel": "Poiesz aktie-overzicht 2026",
        "accountmanager": "Laura",
        "categorie": "Frozen Food",
        "meerdere_sheets_per_week": True,
        "gele_mechanisme_headers": True,
    },
    "hoogvliet": {
        "weergave_naam": "Hoogvliet",
        "sheet_template": None,                          # None = ActiveSheet van het template-bestand
        "titel": "Hoogvliet aktie-overzicht 2026",        # TODO: check exacte gewenste tekst
        "accountmanager": "Ben",
        "categorie": "Frozen Food",
        "meerdere_sheets_per_week": False,                # bevestigd: 1 doorlopend tabblad
        "gele_mechanisme_headers": False,                 # TODO: check of dit ook moet
    },
}

FOCUS_KOLOMMEN = {
    "week": "A", "kolom_c": "C",
    "mech_d": "D", "mech_e": "E",
    "sap_code": "F", "ean": "G", "artikelnaam": "H",
    "volume": "W",
}

KOLOM_C_UITSLUITEN = ["delist", "gesaneerd", "sanering"]

TEMPLATE_FORMULE_RIJ = 122    # formule-/opmaakrij, NOOIT verwijderen uit de PRISTINE template
                               # (was 129, maar die bleek leeg - 122 heeft de echte EAN-XLOOKUP-formules)
START_OUTPUT_RIJ = 7
Q4_RIJ = 6
PRINT_LAATSTE_KOLOM = "R"     # NOOIT automatisch via UsedRange bepalen

TOP_BLUE_RANGE = "A1:R1"
Q4_MERGE_RANGE = "A6:R6"
TITEL_CEL = "F1"
FROZEN_FOOD_CEL = "B1"
ACCOUNTMANAGER_CEL = "B2"

KOL_A_DVIP = "A"
KOL_B_WEEK = "B"
KOL_C_ADVIES_AFBEELDING = "C"
KOL_D_EAN = "D"
KOL_E_PRODUCT = "E"
KOL_J_ACTIE_INKOOPPRIJS = "J"   # alleen nog gebruikt voor de blauwe styling op spacer/mechanisme-rijen
KOL_P_MECHANISME = "P"
KOL_Q_VOLUME = "Q"
# Kolommen F/G/H/I/J/K/L/M/R bewust NIET als "script schrijft hier waarde"
# behandeld: allemaal EAN-formule-gedreven, blijven onaangeraakt.

KOLOM_A_VASTE_WAARDE = "V"
KOLOM_C_VASTE_WAARDE = "X"
ADVIES_AFBEELDING_AANTAL = 2   # top-N op prognose-volume per week

# Kolommen die na het kopieren van TEMPLATE_FORMULE_RIJ een EAN-gedreven
# (X)LOOKUP-formule moeten bevatten. Wordt 1x per sheet gecontroleerd, puur
# als waarschuwing.
EAN_FORMULE_KOLOMMEN = ["F", "G", "H", "I", "J", "K", "L", "M", "R"]

KLEUR_BLAUW = 15257527
KLEUR_GEEL = 10092543

EXCEL_ZICHTBAAR = False

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
MSO_AUTOMATION_FORCE_DISABLE = 3
XL_CALCULATION_AUTOMATIC = -4105
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0


# ====================================================================
# Opruimen van verweesde Excel-processen (zombie-fix)
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


# ====================================================================
# Tekst/getal-cleaning
# ====================================================================

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
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    value = str(value).strip()
    if value == "" or value.lower() in ("nan", "none", "-"):
        return None
    value = value.replace("\u20ac", "").replace(" ", "").replace("\u00a0", "")
    if "." in value and "," in value:
        value = value.replace(".", "").replace(",", ".")
    else:
        value = value.replace(",", ".")
    try:
        return float(value)
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
    if pd.isna(value):
        return None
    value = str(value).strip()
    if value == "" or value.lower() in ("nan", "none", "-"):
        return None
    value = value.replace(" ", "").replace("\u00a0", "")

    if "." in value and "," not in value:
        delen = value.split(".")
        if len(delen) == 2 and len(delen[1]) == 3 and delen[0].isdigit() and delen[1].isdigit():
            return float(delen[0] + delen[1])   # 3.000 -> 3000

    if "," in value and "." not in value:
        delen = value.split(",")
        if len(delen) == 2 and len(delen[1]) == 3 and delen[0].isdigit() and delen[1].isdigit():
            return float(delen[0] + delen[1])   # 3,000 -> 3000

    if "." in value and "," in value:
        value = value.replace(".", "").replace(",", ".")
    else:
        value = value.replace(",", ".")
    try:
        return float(value)
    except ValueError:
        return None


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


def excel_col_to_index(letter: str) -> int:
    """'A' -> 0, 'F' -> 5, 'W' -> 22 (0-indexed kolomnummer)."""
    index = 0
    for ch in letter.upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def is_npd_code(sap_code_raw: str) -> bool:
    return "npd" in clean_text(sap_code_raw).lower()


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
    _, resultaat = sorted(gevonden, key=lambda item: item[0])[-1]   # laatste match in de tekst
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
# Bestand vinden / periode-logica
# ====================================================================

def vind_bestand(prefix: str, extensie: str) -> Path:
    patroon = str(BRON_MAP / f"{prefix}*{extensie}")
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


def laatste_iso_week(jaar: int) -> int:
    return dt.date(jaar, 12, 28).isocalendar()[1]


def kwartaal_van_week(week: int) -> int:
    if week <= 13:
        return 1
    if week <= 26:
        return 2
    if week <= 39:
        return 3
    return 4


def bepaal_weken(weken_arg, jaar: int) -> list:
    if weken_arg:
        return sorted(int(w.strip()) for w in weken_arg.split(",") if w.strip())
    return list(range(40, laatste_iso_week(jaar) + 1))   # standaard: hele Q4


def kwartaal_label(weken: list) -> str:
    kwartalen = sorted({kwartaal_van_week(w) for w in weken})
    return "/".join(f"Q{k}" for k in kwartalen)


def output_basisnaam(retailer_naam: str, weken: list, jaar: int) -> str:
    periode = f"wk{weken[0]}" if len(weken) == 1 else f"wk{min(weken)}-wk{max(weken)}"
    return f"{retailer_naam}_actievoorstel_{periode}_{jaar}"


# ====================================================================
# Brondata inladen (alleen nog Promo Focus CSV)
# ====================================================================

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
        "volume_raw": raw.iloc[:, idx["volume"]].apply(clean_text),
    })

    print(f"  na inlezen: {len(df)} rijen")
    df = df[df["week_int"].notna()].copy()
    print(f"  na week-filter (kolom A moet een geldig getal zijn): {len(df)} rijen")
    # GEEN filter meer op een lege sap-code: de mechanisme-markerrij heeft
    # bewust GEEN sap-code en moet bewaard blijven om herkend te worden.

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


def filter_kolom_c(df: pd.DataFrame, uitsluit_keywords: list, toegestane_excepties: list) -> pd.DataFrame:
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


# ====================================================================
# Outputregels bouwen (sequentieel, met mechanisme-marker-buffering)
# ====================================================================

def is_mechanisme_marker(rij) -> bool:
    """Herkent de mechanisme-aankondigingsrij: geen sap-code, maar wel tekst
    in de (oorspronkelijk samengevoegde) D/E-kolommen."""
    return rij["sap_code_raw"] == "" and bool(rij["mech_d"] or rij["mech_e"])


def bouw_outputregels(focus: pd.DataFrame) -> list:
    """Verwerkt de Promo Focus-rijen OP VOLGORDE (niet als ongeordende set).
    Producten worden gebufferd tot de afsluitende mechanisme-markerrij van
    hun sectie verschijnt; dan krijgen ze allemaal die mechanismetekst. Een
    sectie zonder afsluitende marker (bv. de laatste in het bestand) levert
    producten zonder mechanisme op - die worden niet weggegooid."""
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
        })

    regels.extend(buffer)   # sectie(s) zonder afsluitende marker: toch meenemen

    if overgeslagen_geen_titel:
        print(f"{overgeslagen_geen_titel} regel(s) overgeslagen: geen productnaam en geen NPD.")
    print(f"  bouw_outputregels: {len(regels)} productregel(s) opgebouwd uit {len(focus)} focus-rijen.")

    return regels


def voeg_advies_kruisjes_toe(regels: list, aantal: int = ADVIES_AFBEELDING_AANTAL) -> None:
    """Annoteert elke regel met advies_afbeelding (True/False): top-N op volume,
    per week. Bij gelijk volume: productnaam dan EAN als tiebreaker, voor een
    voorspelbare uitkomst."""
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
    """Groepeert regels per mechanisme_vol, met behoud van volgorde van eerste
    voorkomen. Regels zonder mechanismetekst krijgen geen gele header."""
    groepen, volgorde = {}, []
    for regel in regels:
        sleutel = regel["mechanisme_vol"]
        if sleutel not in groepen:
            groepen[sleutel] = []
            volgorde.append(sleutel)
        groepen[sleutel].append(regel)
    return [(sleutel, groepen[sleutel]) for sleutel in volgorde]


def schrijf_controlebestand(regels: list, output_dir: Path, retailer: str) -> Path:
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    pad = output_dir / f"controle_input_actievoorstel_{retailer}_{timestamp}.xlsx"
    actie_df = pd.DataFrame(regels)
    with pd.ExcelWriter(pad, engine="openpyxl") as writer:
        actie_df.to_excel(writer, sheet_name="Input_voor_actievoorstel", index=False)
    print(f"Controlebestand opgeslagen als: {pad}")
    return pad


# ====================================================================
# Excel: layout-helpers
# ====================================================================

def open_excel_veilig():
    app = win32.DispatchEx("Excel.Application")
    app.Visible = EXCEL_ZICHTBAAR
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = MSO_AUTOMATION_FORCE_DISABLE   # Workbook_Open vuurt nooit af
    return app


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


def wacht_op_excel_berekening(app, max_seconden=90):
    try:
        start = time.time()
        while app.CalculationState != 0:
            if time.time() - start > max_seconden:
                print("Waarschuwing: Excel blijft lang berekenen, script gaat verder.")
                break
            time.sleep(1)
    except Exception:
        pass


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
    """Checkt 1x per sheet of de verwachte EAN-gedreven formules nog
    aanwezig zijn na het kopieren van TEMPLATE_FORMULE_RIJ. Puur een
    waarschuwing."""
    ontbrekend = [k for k in EAN_FORMULE_KOLOMMEN if not sheet.Range(f"{k}{rij}").HasFormula]
    if ontbrekend:
        print(
            f"WAARSCHUWING [{sheet_naam}]: kolom(men) {ontbrekend} hebben GEEN formule in rij {rij} "
            f"(verwacht een EAN-gedreven (X)LOOKUP, gekopieerd vanuit rij {TEMPLATE_FORMULE_RIJ})."
        )
    else:
        print(f"  [{sheet_naam}] EAN-formule-check: alle kolommen {EAN_FORMULE_KOLOMMEN} hebben een formule. OK.")


def _schrijf_productrij(sheet, rij: int, regel: dict, eerste_rij_van_sheet: bool, sheet_naam: str = ""):
    if eerste_rij_van_sheet:
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
        ean_cel.NumberFormat = "@"   # val terug op tekst als het geen zuiver cijfer is
        ean_cel.Value = ean_waarde

    sheet.Range(f"{KOL_E_PRODUCT}{rij}").Value = regel["productnaam"]
    # Kolommen F/G/H/I/J/K/L/M/R NIET aanraken: allemaal EAN-formule-
    # gedreven in de template, gevoed door de EAN die hierboven in D staat.
    sheet.Range(f"{KOL_P_MECHANISME}{rij}").Value = regel["mechanisme_kort"]
    sheet.Range(f"{KOL_Q_VOLUME}{rij}").Value = regel["volume_excel"]

    _formatteer_artikelrij(sheet, rij)


def _zet_zwarte_onderrand(sheet, rij: int):
    rng = sheet.Range(f"A{rij}:{PRINT_LAATSTE_KOLOM}{rij}")
    rng.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
    rng.Borders(XL_EDGE_BOTTOM).Weight = XL_MEDIUM


def _verwijder_overtollige_rijen(sheet, laatste_output_rij: int):
    """Verwijdert restanten van de gekopieerde template die na de laatste
    echte outputrij nog in deze sheet-KOPIE staan. Gaat NOOIT verder dan
    TEMPLATE_FORMULE_RIJ - voor het geval er verderop in de sheet nog iets
    staat (bv. een lokale opzoektabel) dat niet zomaar verwijderd mag worden."""
    used_range = sheet.UsedRange
    laatste_gebruikte_rij = used_range.Row + used_range.Rows.Count - 1
    bovengrens = min(laatste_gebruikte_rij, TEMPLATE_FORMULE_RIJ)
    if bovengrens > laatste_output_rij:
        sheet.Rows(f"{laatste_output_rij + 1}:{bovengrens}").Delete()


def vul_week_sheet(sheet, retailer_cfg: dict, weken: list, regels: list) -> int:
    sheet.Range(TOP_BLUE_RANGE).Interior.Pattern = XL_SOLID
    sheet.Range(TOP_BLUE_RANGE).Interior.Color = KLEUR_BLAUW
    sheet.Range(TOP_BLUE_RANGE).Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
    sheet.Range(TOP_BLUE_RANGE).Borders(XL_EDGE_BOTTOM).Weight = XL_THIN

    sheet.Range(FROZEN_FOOD_CEL).Value = retailer_cfg["categorie"]
    sheet.Range(FROZEN_FOOD_CEL).Font.Bold = True
    sheet.Range(FROZEN_FOOD_CEL).Font.Color = 16777215
    sheet.Range(FROZEN_FOOD_CEL).Font.Underline = True

    sheet.Range(TITEL_CEL).Value = retailer_cfg["titel"]
    sheet.Range(TITEL_CEL).Font.Bold = True
    sheet.Range(TITEL_CEL).Font.Size = 18
    sheet.Range(TITEL_CEL).HorizontalAlignment = XL_CENTER

    sheet.Range(ACCOUNTMANAGER_CEL).Value = retailer_cfg["accountmanager"]

    try:
        sheet.Range(Q4_MERGE_RANGE).UnMerge()
    except Exception:
        pass
    q_rng = sheet.Range(Q4_MERGE_RANGE)
    q_rng.Merge()
    q_rng.Value = kwartaal_label(weken)
    q_rng.HorizontalAlignment = XL_CENTER
    q_rng.VerticalAlignment = XL_VCENTER
    q_rng.Font.Bold = True
    q_rng.Interior.Pattern = XL_SOLID
    q_rng.Interior.Color = KLEUR_BLAUW
    q_rng.Borders(XL_EDGE_TOP).LineStyle = XL_CONTINUOUS
    q_rng.Borders(XL_EDGE_TOP).Weight = XL_THIN
    q_rng.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
    q_rng.Borders(XL_EDGE_BOTTOM).Weight = XL_THIN
    sheet.Rows(Q4_RIJ).RowHeight = 15

    gegroepeerd = retailer_cfg.get("gele_mechanisme_headers", False)
    template_rij = sheet.Rows(TEMPLATE_FORMULE_RIJ)

    if gegroepeerd:
        groepen = groepeer_per_mechanisme(regels)
        aantal_extra_rijen = sum(2 for tekst, _ in groepen if tekst)   # spacer + mechanisme-rij per groep
        totaal_rijen = len(regels) + aantal_extra_rijen
    else:
        groepen = [("", regels)]
        totaal_rijen = len(regels)

    if START_OUTPUT_RIJ + totaal_rijen - 1 >= TEMPLATE_FORMULE_RIJ:
        raise ValueError(
            f"Te veel rijen ({totaal_rijen}) voor week(en) {weken} - "
            f"dit zou rij {TEMPLATE_FORMULE_RIJ} overschrijven."
        )

    rij = START_OUTPUT_RIJ
    eerste_product_geschreven = False

    for groep_index, (mechanisme_tekst, groep_regels) in enumerate(groepen):
        if gegroepeerd and mechanisme_tekst:
            template_rij.Copy()
            sheet.Rows(rij).PasteSpecial(Paste=XL_PASTE_ALL)
            _schrijf_witte_spacer_rij(sheet, rij, zwarte_bovenrand=(groep_index == 0))
            rij += 1

            template_rij.Copy()
            sheet.Rows(rij).PasteSpecial(Paste=XL_PASTE_ALL)
            _schrijf_mechanisme_rij(sheet, rij, mechanisme_tekst)
            rij += 1

        for regel in groep_regels:
            template_rij.Copy()
            sheet.Rows(rij).PasteSpecial(Paste=XL_PASTE_ALL)
            _schrijf_productrij(
                sheet, rij, regel,
                eerste_rij_van_sheet=not eerste_product_geschreven,
                sheet_naam=sheet.Name,
            )
            eerste_product_geschreven = True
            rij += 1

    laatste_rij = rij - 1
    sheet.Application.CutCopyMode = False

    if laatste_rij >= START_OUTPUT_RIJ:
        _zet_zwarte_onderrand(sheet, laatste_rij)
        _verwijder_overtollige_rijen(sheet, laatste_rij)

    sheet.PageSetup.PrintArea = f"$A$1:${PRINT_LAATSTE_KOLOM}${laatste_rij}"
    return laatste_rij


# ====================================================================
# Hoofdflow per retailer
# ====================================================================

def genereer_voor_retailer(retailer_key: str, weken_arg, jaar: int, toegestane_kolom_c):
    cfg = RETAILERS[retailer_key]
    weergave_naam = cfg["weergave_naam"]

    template_path = Path(DIRK_TEMPLATE_BESTAND)
    if not template_path.exists():
        raise FileNotFoundError(f"Templatebestand niet gevonden: {template_path}")

    focus_pad = vind_bestand(f"Promo Focus File {jaar} {weergave_naam}", ".csv")

    weken = bepaal_weken(weken_arg, jaar)

    print("Start actievoorstel maken")
    print(f"Retailer: {weergave_naam}  |  Jaar: {jaar}  |  Weken: {weken}")
    print(f"Template: {template_path}  |  Focus: {focus_pad.name}")

    focus_raw = laad_focus_data(focus_pad)
    print(f"  unieke weken in Promo Focus CSV: {sorted(focus_raw['week_int'].dropna().unique().tolist())}")

    toegestane_lijst = toegestane_kolom_c.split("|") if toegestane_kolom_c else []
    focus = filter_kolom_c(focus_raw, KOLOM_C_UITSLUITEN, toegestane_lijst)
    print(f"  na kolom-C-filter: {len(focus)} rijen")
    focus = focus[focus["week_int"].isin(weken)]
    print(f"  na week-selectie (gevraagd: {weken}): {len(focus)} rijen")

    alle_regels = bouw_outputregels(focus)
    voeg_advies_kruisjes_toe(alle_regels)
    print(f"Aantal artikelen voor actievoorstel: {len(alle_regels)}")
    print(f"Aantal advies-afbeelding kruisjes: {sum(1 for r in alle_regels if r['advies_afbeelding'])}")

    OUTPUT_MAP.mkdir(exist_ok=True)
    schrijf_controlebestand(alle_regels, OUTPUT_MAP, weergave_naam)

    output_basis = output_basisnaam(weergave_naam, weken, jaar)
    output_pad = OUTPUT_MAP / f"{output_basis}.xlsm"
    shutil.copy2(template_path, output_pad)   # master blijft altijd schoon

    app = open_excel_veilig()
    wb = None
    try:
        wb = app.Workbooks.Open(str(output_pad.resolve()), UpdateLinks=0)
        app.Calculation = XL_CALCULATION_AUTOMATIC
        pristine = vind_template_sheet(wb, cfg["sheet_template"])

        if cfg["meerdere_sheets_per_week"]:
            # Fase 1: eerst ALLE benodigde sheets aanmaken (kopie van de
            # schone pristine), nog niets vullen.
            week_naar_sheet = {}
            for week in weken:
                regels_week = [r for r in alle_regels if r["week"] == week]
                if not regels_week:
                    print(f"Week {week}: geen regels, sheet wordt overgeslagen.")
                    continue
                pristine.Copy(After=wb.Sheets(wb.Sheets.Count))
                nieuw = wb.Sheets(wb.Sheets.Count)
                nieuw.Name = f"wk{week} {jaar}"
                week_naar_sheet[week] = (nieuw, regels_week)

            if week_naar_sheet:
                # Harde veiligheidscheck, niet alleen vertrouwen op de logica
                # hierboven: weiger de delete als er toch geen ander
                # tabblad blijkt te bestaan (een workbook moet altijd
                # minstens 1 zichtbaar tabblad hebben).
                if wb.Sheets.Count <= 1:
                    raise RuntimeError(
                        "Veiligheidscheck gefaald: pristine zou het enige tabblad "
                        "zijn na verwijderen. Delete wordt NIET uitgevoerd."
                    )
                pristine.Delete()
            else:
                print(
                    f"WAARSCHUWING: geen enkele week had regels - de lege "
                    f"'{cfg['sheet_template']}'-template blijft staan (een "
                    "workbook moet minimaal 1 zichtbaar tabblad hebben)."
                )

            # Fase 2: nu pas de inhoud vullen.
            gemaakte_sheets = []
            for week, (sheet, regels_week) in week_naar_sheet.items():
                vul_week_sheet(sheet, cfg, [week], regels_week)
                gemaakte_sheets.append(sheet)
                print(f"Week {week}: {len(regels_week)} regel(s) -> tabblad '{sheet.Name}'")
        else:
            vul_week_sheet(pristine, cfg, weken, alle_regels)
            gemaakte_sheets = [pristine]
            print(f"{len(alle_regels)} regel(s) geschreven naar '{pristine.Name}'.")

        app.CalculateFullRebuild()
        wacht_op_excel_berekening(app)
        wb.Save()

        for sheet in gemaakte_sheets:
            pdf_naam = f"{output_basis}_{sheet.Name}.pdf".replace(" ", "_")
            pdf_pad = OUTPUT_MAP / pdf_naam
            sheet.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=str(pdf_pad.resolve()),
                Quality=XL_QUALITY_STANDARD,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            print(f"PDF: {pdf_pad}")

        print(f"Klaar: {output_pad}")

    finally:
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


# ====================================================================
# Hoofdscript
# ====================================================================

def main():
    parser = argparse.ArgumentParser(description="Genereer actievoorstel per retailer")
    parser.add_argument("--retailer", default="Hoogvliet")
    parser.add_argument("--jaar", type=int, default=dt.date.today().year)
    parser.add_argument("--weken", default=None,
                         help="Komma-gescheiden weeknummers, bv. 37 of 37,38,40. Standaard: hele Q4.")
    parser.add_argument("--toegestane-kolom-c", dest="toegestane_kolom_c", default=None,
                         help="Pipe-gescheiden teksten die ondanks delist/gesaneerd/sanering toch mogen, "
                              "bv. \"sanering wk40 2026|sanering\"")
    parser.add_argument("--skip-cleanup", action="store_true")
    args = parser.parse_args()

    retailer_key = args.retailer.strip().lower()
    if retailer_key not in RETAILERS:
        print(f"Onbekende retailer '{args.retailer}'. Bekend: {list(RETAILERS.keys())}")
        sys.exit(1)

    if not args.skip_cleanup:
        kill_orphan_excel()

    try:
        genereer_voor_retailer(retailer_key, args.weken, args.jaar, args.toegestane_kolom_c)
    except Exception as error:
        print("")
        print("FOUT:")
        print(error)
        sys.exit(1)


if __name__ == "__main__":
    main()
