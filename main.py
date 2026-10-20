#!/usr/bin/env python3
"""
Extraction des factures Chargemap via pdfplumber.

Stratégie : découpage spatial en colonnes (par x) et lignes logiques (par y).
Une description qui wrappe sur 2 lignes visuelles reste dans la même cellule
car la "ligne logique" est délimitée par les marqueurs "Charge effectuée".

Jointure des lignes visuelles : on colle deux lignes quand la ligne du haut est
JUSTIFIÉE (son dernier mot atteint la marge droite de la colonne), signe d'une
coupure de mot. Sinon on insère un espace. La casse est ignorée, car elle ne
porte aucune information fiable (ex : "Mc" + "Donald" = coupure avec majuscule).

EVOLUTION
- extraire date et heure de description pour les mettre dans des colonnes séparées
"""
from datetime import datetime
from os import listdir
from pathlib import Path
from re import compile, fullmatch, search, sub
from sys import argv, exit

import pandas as pd
import pdfplumber

BOTTOM = "bottom"
COERCE = "coerce"
COLUMN = "column"
DATE_HEURE = "date_heure"
DESCRIPTION = "description"
PU_HT = "pu_ht"
QUANTITE = "quantite"
TEXT = "text"
TOP = "top"
TVA_MONTANT = "tva_montant"
TVA_PCT = "tva_pct"
TOTAL_HT = "total_ht"
X0 = "x0"
X1 = "x1"

# ─── Configuration des colonnes ───────────────────────────────────────────────
# Plages X pour chaque colonne (à ajuster selon le PDF, mais stables pour Chargemap).
COLUMNS = {
    DESCRIPTION: (0, 355),
    PU_HT:       (355, 420),
    QUANTITE:    (420, 480),
    TVA_PCT:     (480, 540),
    TVA_MONTANT: (540, 600),
    TOTAL_HT:    (600, 700),
}
MOIS = {
    "janvier": 1, "février": 2, "mars": 3, "avril": 4,"mai": 5, "juin": 6,
    "juillet": 7, "août": 8, "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12
}

# Borne droite de la zone de rendu de la colonne description, utilisée pour
# détecter la justification. À calibrer via le mode debug si besoin.
DESC_RIGHT = COLUMNS[DESCRIPTION][1]
JUSTIFY_TOL = 2.0  # tolérance en points pour considérer que la marge est atteinte


def main() -> None:
    pdf_path = argv[1] if len(argv) > 1 else "/home/kobaya/Téléchargements/chargemap"
    output = argv[2] if len(argv) > 2 else "facture_chargemap.ods"
    recharges = []

    if not Path(pdf_path).exists():
        print(f"❌ Fichier introuvable : {pdf_path}")
        exit(1)

    pattern = compile(r'^\d{8}-\d{8}\.pdf$')

    files = [f for f in listdir(pdf_path) if pattern.match(f)]

    for f in files:
        file = Path(pdf_path, f)
        recharges.extend(parse_pdf(file))

    if not recharges:
        print("⚠️  Aucune ligne extraite. Ajustez les bornes COLUMNS / DESC_RIGHT.")
        debug_positions(pdf_path)
    else:
        export(recharges, output)


def extract_words_by_column(page: pdfplumber.page.Page) -> list[dict]:
    """Extrait tous les mots de la page et les assigne à une colonne selon leur x.
    Retourne une liste de dicts : {"text", "x0", "x1", "top", "bottom", "column"}

    :param page: page du pdf
    :return: liste des mots avec leur position et colonne assignée
    """
    words = page.extract_words(
        x_tolerance=2,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )

    result = []
    for w in words:
        cx = (w[X0] + w[X1]) / 2  # centre du mot pour l'assignation de colonne
        col = None
        for name, (x_min, x_max) in COLUMNS.items():
            if x_min <= cx < x_max:
                col = name
                break
        if col:
            result.append({
                TEXT: w[TEXT],
                X0: w[X0],
                X1: w[X1],
                TOP: w[TOP],
                BOTTOM: w[BOTTOM],
                COLUMN: col,
            })

    return result


def find_row_boundaries(words: list[dict]) -> list[tuple[float, float]]:
    """Identifie les lignes logiques en cherchant les mots "Charge" qui marquent
    le début de chaque entrée. La borne basse est bornée par la fin de la zone
    de données (pied de page exclu). Retourne une liste de (y_start, y_end).

    :param words: liste des mots avec leur position et colonne assignée
    :return: liste de (y_start, y_end) pour chaque ligne logique
    """
    charge_tops = sorted(
        w[TOP] for w in words
        if w[COLUMN] == DESCRIPTION and w[TEXT] == "Charge"
    )
    if not charge_tops:
        return []

    # Fin de zone de données : premier marqueur de total / pied de page
    END_MARKERS = {"Total", "Conditions", "Chargemap", "Net"}
    end_candidates = [
        w[TOP] for w in words
        if w[TEXT] in END_MARKERS and w[TOP] > charge_tops[-1]
    ]
    data_end = min(end_candidates) if end_candidates else max(w[BOTTOM] for w in words)

    boundaries = []
    for i, top in enumerate(charge_tops):
        next_charge = charge_tops[i + 1] if i + 1 < len(charge_tops) else data_end
        boundaries.append((top, min(next_charge, data_end)))

    return boundaries


def group_words_by_visual_line(words_in_col: list[dict]) -> list[tuple[str, float, float]]:
    """Retourne une liste de (texte, x0_premier_mot, x1_dernier_mot), une entrée
    par ligne visuelle (même 'top' arrondi).

    :param words_in_col: liste des mots d'une colonne (ex: description) pour une ligne logique
    :return : liste des lignes visuelles avec leur texte et positions x
    """
    by_top: dict[int, list[tuple[float, float, str]]] = {}
    for w in words_in_col:
        by_top.setdefault(round(w[TOP]), []).append(
            (w[X0], w[X1], w[TEXT])
        )
    out = []
    for _, cells in sorted(by_top.items()):
        cells.sort()
        text = " ".join(t for _, _, t in cells)
        out.append((text, cells[0][0], cells[-1][1]))

    return out


def join_visual_lines(lines: list[tuple[str, float, float]],
                      col_right: float,
                      tol: float = 2.0) -> str:
    """Joint des lignes visuelles en une chaîne.
    Colle (sans espace) si la ligne précédente est justifiée, c.-à-d. que son
    dernier mot atteint la marge droite de la colonne (= mot coupé).
    Sinon, insère un espace.

    :param lines: liste de (texte_ligne, x0_premier_mot, x1_dernier_mot)
    :param col_right: borne droite de la zone de rendu de la colonne
    :param tol: tolérance en points pour considérer que la marge est atteinte
    """
    if not lines:
        return ""
    result = lines[0][0]
    for text, _x0, x1 in lines[1:]:
        if x1 >= col_right - tol:
            result += text          # ligne haute justifiée → mot coupé → on colle
        else:
            result += " " + text    # frontière de mot → espace

    return result


def clean_description(description: str) -> str:
    """Normalise les artefacts restants (apostrophe perdue, espaces multiples).

    :param description: description brute extraite du PDF
    :return: description normalisée
    """
    description = sub(r"\s+", " ", description).strip()
    description = sub(r"McDonald\s+s\b", "McDonald's", description)
    description = sub(r"McDonald's\s+s\b", "McDonald's", description)

    return description


def numeric_value(text: str) -> str | None:
    """Tokens numériques de la zone droite, en retirant les parenthèses.

    :param text: texte brut
    :return: texte nettoyé si numérique, sinon None
    """
    cleaned = text.strip("()")

    return cleaned if fullmatch(r"[\d.,]+", cleaned) else None


def parse_page(page: pdfplumber.page.Page) -> list[dict]:
    """Parse une page et retourne les lignes de charges.

    :param page: page du pdf
    :return: liste des lignes de charges
    """
    words = extract_words_by_column(page)
    boundaries = find_row_boundaries(words)

    rows = []
    for y_start, y_end in boundaries:
        row_words = [w for w in words if y_start <= w[TOP] < y_end]

        # Description : regroupement par ligne visuelle puis jointure par justification
        desc_lines = group_words_by_visual_line(
            [w for w in row_words if w[COLUMN] == DESCRIPTION]
        )
        description = clean_description(
            join_visual_lines(desc_lines, col_right=DESC_RIGHT, tol=JUSTIFY_TOL)
        )

        # récupération date/heure de la recharge
        pattern = r', le\s+(\d+)\s+(\w+)\s+(\d{4})\s+à\s+(\d{2})h(\d{2})'
        date_heure_brute = search(pattern, description)
        if date_heure_brute:
            jour, mois, annee, heure, minute = date_heure_brute.groups()
            dt = datetime(
                year=int(annee),
                month=MOIS[mois.lower()],
                day=int(jour),
                hour=int(heure),
                minute=int(minute)
            )

        # Autres colonnes : identification par CONTENU (robuste au décalage de bornes)
        cells: dict[str, list[str]] = {col: [] for col in COLUMNS}
        for w in row_words:
            cells[w[COLUMN]].append(w[TEXT])

        # Regroupement par ligne visuelle (même 'top' arrondi), dans l'ordre de lecture
        by_top: dict[int, list[tuple[float, float]]] = {}  # top -> [(x0, valeur), ...]
        for w in row_words:
            if w[X0] < COLUMNS[PU_HT][0]:
                continue
            val = numeric_value(w[TEXT])
            if val is not None:
                by_top.setdefault(round(w[TOP]), []).append((w[X0], val))

        visual_lines = [
            [v for _, v in sorted(cells)]
            for _, cells in sorted(by_top.items())
        ]

        # Mapping par structure physique de la facture :
        #   ligne 1 = [PU, quantité, taux]   ligne 2 = [TVA montant]   ligne 3 = [Total HT]
        pu_ht = quantite = tva_pct = tva_montant = total_ht = ""
        if len(visual_lines) >= 1:
            l1 = visual_lines[0]
            pu_ht    = l1[0] if len(l1) >= 1 else ""
            quantite = l1[1] if len(l1) >= 2 else ""
            tva_pct  = l1[2] if len(l1) >= 3 else ""
        if len(visual_lines) >= 2:
            tva_montant = visual_lines[1][0] if visual_lines[1] else ""
        if len(visual_lines) >= 3:
            total_ht = visual_lines[2][0] if visual_lines[2] else ""

        # Garde-fou de cohérence : si total HT manquant, il égale le PU
        # (sur cette facture Total HT == PU car quantité = 1) ; sinon on garde
        # la valeur extraite. TVA montant = PU × taux/100 en secours.
        try:
            if pu_ht and tva_pct and not tva_montant:
                tva_montant = f"{round(float(pu_ht) * float(tva_pct) / 100, 2):.2f}"
            if pu_ht and not total_ht:
                total_ht = pu_ht
        except ValueError:
            pass

        rows.append({
            DESCRIPTION: description,
            DATE_HEURE: dt if date_heure_brute else "",
            PU_HT: pu_ht,
            QUANTITE: quantite,
            TVA_PCT: tva_pct,
            TVA_MONTANT: tva_montant,
            TOTAL_HT: total_ht,
        })

    return rows


def parse_pdf(pdf_path: Path) -> list[dict]:
    """Parse l'ensemble du PDF.

    :param pdf_path: chemin vers le fichier PDF
    :return: liste des lignes de charges
    """
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            all_rows.extend(parse_page(page))

    return all_rows


def to_float(val: str) -> float | None:
    """Convertit '6.25 €' ou '11,50' en float.

    :param val: valeur à convertir
    :return: valeur convertie ou None
    """
    if not val:
        return None
    cleaned = sub(r"[€\s]", "", str(val)).replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def export(recharges: list[dict[str, str]], output: str = "facture_chargemap.ods") -> None:
    """Exporte vers ODS ou XLSX.

    :param recharges: liste des recharges
    :param output: nom du fichier de sortie,  par défaut "facture_chargemap.ods"
    """
    df = pd.DataFrame(recharges)

    for col in [PU_HT, TVA_MONTANT, TVA_PCT, TOTAL_HT]:
        df[col] = df[col].map(to_float)

    df[QUANTITE] = pd.to_numeric(df[QUANTITE], errors=COERCE)
    df[DATE_HEURE] = pd.to_datetime(df[DATE_HEURE], errors=COERCE)

    # création de la colonne total_ttc
    df["total_ttc"] = df[TOTAL_HT] + df[TVA_MONTANT]
    # création de la colonne mois
    df["mois"] = df[DATE_HEURE].dt.to_period("M") 
    ordre = ["description", "date_heure", "mois",
            "pu_ht", "quantite", "tva_pct", "tva_montant",
            "total_ht", "total_ttc"]
    df = df[ordre]

    engine = "odf" if output.endswith(".ods") else "openpyxl"
    with pd.ExcelWriter(output, engine=engine) as writer:
        df.to_excel(writer, sheet_name="Recharges", index=False)

    print(f"✅ {len(df)} lignes → {output}")


def debug_positions(pdf_path: str) -> None:
    """Affiche les positions x des mots pour calibrer COLUMNS et DESC_RIGHT.

    :param pdf_path: chemin vers le fichier PDF
    """
    with pdfplumber.open(pdf_path) as pdf:
        words = pdf.pages[0].extract_words(x_tolerance=2, y_tolerance=3)
        with open("positions_debug.txt", "w", encoding="utf-8") as f:
            for w in words:
                f.write(f"  x0={w[X0]:05.1f}  x1={w[X1]:05.1f}  top={w[TOP]:05.1f}  '{w[TEXT]}'\n")


if __name__ == "__main__":
    main()
