#!/usr/bin/env python3
"""Extraction des factures Chargemap via pdfplumber.

Stratégie : découpage spatial en colonnes (par x) et lignes logiques (par y).
Une description qui wrappe sur 2 lignes visuelles reste dans la même cellule
car la "ligne logique" est délimitée par les marqueurs "Charge effectuée".

Jointure des lignes visuelles : on colle deux lignes quand la seconde reprend à
la MARGE GAUCHE de la colonne description (= mot coupé par le retour à la ligne).
Sinon on insère un espace. La marge est le x0 du mot "Charge" en tête de ligne.
La casse est ignorée, car elle ne porte aucune information fiable.
"""
from datetime import datetime
from openpyxl.styles import numbers as xl_numbers
from os import listdir
from pathlib import Path
from re import compile, fullmatch, search, sub
from sys import exit

import pandas as pd
import pdfplumber

BOTTOM = "bottom"
CHARGEMAP_PATH = "/home/kobaya/Téléchargements/chargemap"
COERCE = "coerce"
COLUMN = "column"
DATE_HEURE = "date_heure"
DESCRIPTION = "description"
MOIS = "mois"
PU_HT = "pu_ht"
QUANTITE = "quantite"
TEXT = "text"
TOP = "top"
TVA_MONTANT = "tva_montant"
TVA_PCT = "tva_pct"
TOTAL_HT = "total_ht"
TOTAL_TTC = "total_ttc"
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
MOIS_DICT = {
    "janvier": 1, "février": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
    "juillet": 7, "août": 8, "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12
}

# Tolérance en points pour décider qu'une ligne reprend à la marge gauche.
LEFT_MARGIN_TOL = 2.0


def main() -> None:
    pdf_path = CHARGEMAP_PATH
    output = Path(CHARGEMAP_PATH, "factures_chargemap.xlsx")
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
        print("⚠️  Aucune ligne extraite. Ajustez les bornes COLUMNS.")
    else:
        export(recharges, output)


def extract_words_by_column(page: pdfplumber.page.Page) -> list[dict]: # type: ignore
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
    """Identifie les lignes logiques via le mot "Charge" en TÊTE de ligne
    (le mot le plus à gauche de sa ligne visuelle), ce qui exclut les
    occurrences internes comme "Easy Charge Services". La borne basse est
    bornée par la fin de la zone de données (pied de page exclu).

    :param words: liste des mots avec leur position et colonne assignée
    :return: liste de (y_start, y_end) pour chaque ligne logique
    """
    # x0 minimal par ligne visuelle (toutes colonnes confondues)
    leftmost_by_top: dict[int, float] = {}
    for w in words:
        t = round(w[TOP])
        if t not in leftmost_by_top or w[X0] < leftmost_by_top[t]:
            leftmost_by_top[t] = w[X0]

    # Un "Charge" valide est le mot le plus à gauche de sa ligne
    charge_tops = sorted(
        w[TOP] for w in words
        if w[COLUMN] == DESCRIPTION
        and w[TEXT] == "Charge"
        and abs(w[X0] - leftmost_by_top[round(w[TOP])]) < 1.0
    )
    if not charge_tops:
        return []

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

    :param words_in_col: liste des mots d'une colonne pour une ligne logique
    :return: liste des lignes visuelles avec leur texte et positions x
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
                      left_margin: float,
                      tol: float = LEFT_MARGIN_TOL) -> str:
    """Colle deux lignes si la seconde reprend à la marge gauche de la colonne
    (= mot coupé par le retour à la ligne), sinon insère un espace.

    :param lines: liste de (texte, x0_premier_mot, x1_dernier_mot)
    :param left_margin: x0 de la marge gauche (le "Charge" en tête de ligne)
    :param tol: tolérance en points
    :return: description sur une seule ligne
    """
    if not lines:
        return ""
    result = lines[0][0]
    for text, x0, _x1 in lines[1:]:
        if x0 <= left_margin + tol:
            result += text          # reprise à la marge → mot coupé → on colle
        else:
            result += " " + text    # sinon → frontière de mot → espace

    return result


def clean_description(description: str) -> str:
    """Normalise les artefacts restants (apostrophe perdue, espaces multiples).

    :param description: description brute jointe
    :return: description normalisée
    """
    description = sub(r"\s+", " ", description).strip()
    description = sub(r"McDonald\s+s\b", "McDonald's", description)
    description = sub(r"McDonald's\s+s\b", "McDonald's", description)

    return description


def numeric_value(text: str) -> str | None:
    """Retourne la valeur numérique d'un token en retirant les parenthèses
    (pdfplumber renvoie "(1.25"), ou None si le token n'est pas numérique.

    :param text: texte brut
    :return: texte nettoyé si numérique, sinon None
    """
    cleaned = text.strip("()")

    return cleaned if fullmatch(r"[\d.,]+", cleaned) else None


def parse_page(page: pdfplumber.page.Page) -> list[dict]: # type: ignore
    """Parse une page et retourne les lignes de charges.

    :param page: page du pdf
    :return: liste des lignes de charges
    """
    words = extract_words_by_column(page)
    boundaries = find_row_boundaries(words)

    rows = []
    for y_start, y_end in boundaries:
        row_words = [w for w in words if y_start <= w[TOP] < y_end]

        # Description : marge gauche = x0 du mot le plus à gauche (le "Charge"),
        # puis regroupement par ligne visuelle et jointure par marge gauche.
        desc_words = [w for w in row_words if w[COLUMN] == DESCRIPTION]
        left_margin = min(w[X0] for w in desc_words) if desc_words else 0.0
        desc_lines = group_words_by_visual_line(desc_words)
        description = clean_description(
            join_visual_lines(desc_lines, left_margin=left_margin)
        )

        # Récupération date/heure de la recharge
        dt = jour = mois = annee = heure = minute = lieu = None
        pattern = r', le\s+(\d+)\s+(\w+)\s+(\d{4})\s+à\s+(\d{2})h(\d{2})\sà\s(.+)$'
        donnees_brutes = search(pattern, description)
        if donnees_brutes:
            jour, mois, annee, heure, minute, lieu = donnees_brutes.groups()
            dt = datetime(
                year=int(annee),
                month=MOIS_DICT[mois.lower()],
                day=int(jour),
                hour=int(heure),
                minute=int(minute),
            )

        # Autres colonnes : identification par CONTENU (robuste au décalage de bornes)
        # Regroupement par ligne visuelle (même 'top' arrondi), dans l'ordre de lecture
        by_top: dict[int, list[tuple[float, str]]] = {}
        for w in row_words:
            if w[X0] < COLUMNS[PU_HT][0]:
                continue
            val = numeric_value(w[TEXT])
            if val is not None:
                by_top.setdefault(round(w[TOP]), []).append((w[X0], str(val)))

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
            DESCRIPTION: lieu,
            DATE_HEURE: dt,
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


def export(recharges: list[dict], output: Path) -> None:
    """Exporte vers ODS ou XLSX.

    :param recharges: liste des recharges
    :param output: nom du fichier de sortie
    """
    df = pd.DataFrame(recharges)

    for col in [PU_HT, TVA_MONTANT, TVA_PCT, TOTAL_HT]:
        df[col] = df[col].map(to_float)

    df[QUANTITE] = pd.to_numeric(df[QUANTITE], errors=COERCE)
    df[DATE_HEURE] = pd.to_datetime(df[DATE_HEURE], errors=COERCE)

    # création de la colonne total_ttc
    df[TOTAL_TTC] = df[TOTAL_HT] + df[TVA_MONTANT]
    # création de la colonne mois
    df[MOIS] = df[DATE_HEURE].dt.to_period("M")
    ordre = [DESCRIPTION, DATE_HEURE, MOIS,
             PU_HT, QUANTITE, TVA_PCT, TVA_MONTANT,
             TOTAL_HT, TOTAL_TTC]
    df = df[ordre]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Recharges", index=False)
        apply_formats(writer, "Recharges", df)

    print(f"✅ {len(df)} lignes → {output}")


def apply_formats(writer: "pd.ExcelWriter", sheet: str, df: pd.DataFrame) -> None:
    """Applique les formats de nombre sur les colonnes concernées.

    - date_heure  -> jj/mm/aaaa hh:mm
    - pu_ht, tva_montant, total_ht, total_ttc -> monnaie (2 décimales, €)

    :param writer: le contexte ExcelWriter encore ouvert
    :param sheet: nom de la feuille
    :param df: le DataFrame écrit (pour retrouver l'index des colonnes)
    """
    date_cols = [DATE_HEURE]
    money_cols = [PU_HT, TVA_MONTANT, TOTAL_HT, TOTAL_TTC]

    ws = writer.sheets[sheet]
    col_index = {name: i + 1 for i, name in enumerate(df.columns)}
    # +1 car la ligne 1 = en-têtes
    for name in date_cols:
        ci = col_index[name]
        for row in range(2, len(df) + 2):
            ws.cell(row=row, column=ci).number_format = "DD/MM/YYYY HH:MM"

    for name in money_cols:
        ci = col_index[name]
        for row in range(2, len(df) + 2):
            ws.cell(row=row, column=ci).number_format = '#,##0.00\\ "€"'


if __name__ == "__main__":
    main()
