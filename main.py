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
"""

from pathlib import Path
import re
import sys

import pandas as pd
import pdfplumber


# ─── Configuration des colonnes ───────────────────────────────────────────────
# Plages X pour chaque colonne (à ajuster selon le PDF, mais stables pour Chargemap).
COLUMNS = {
    "description": (0, 340),
    "pu_ht":       (340, 400),
    "quantite":    (400, 435),
    "tva_pct":     (435, 490),
    "tva_montant": (490, 545),
    "total_ht":    (545, 620),
}

# Borne droite de la zone de rendu de la colonne description, utilisée pour
# détecter la justification. À calibrer via le mode debug si besoin.
DESC_RIGHT = COLUMNS["description"][1]
JUSTIFY_TOL = 2.0  # tolérance en points pour considérer que la marge est atteinte


def extract_words_by_column(page: pdfplumber.page.Page) -> list[dict]:
    """
    Extrait tous les mots de la page et les assigne à une colonne selon leur x.
    Retourne une liste de dicts : {"text", "x0", "x1", "top", "bottom", "column"}
    """
    words = page.extract_words(
        x_tolerance=2,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )

    result = []
    for w in words:
        cx = (w["x0"] + w["x1"]) / 2  # centre du mot pour l'assignation de colonne
        col = None
        for name, (x_min, x_max) in COLUMNS.items():
            if x_min <= cx < x_max:
                col = name
                break
        if col:
            result.append({
                "text": w["text"],
                "x0": w["x0"],
                "x1": w["x1"],
                "top": w["top"],
                "bottom": w["bottom"],
                "column": col,
            })
    return result


def find_row_boundaries(words: list[dict]) -> list[tuple[float, float]]:
    """
    Identifie les lignes logiques en cherchant les mots "Charge" qui marquent
    le début de chaque entrée. Retourne une liste de (y_start, y_end).
    """
    charge_tops = sorted(
        w["top"] for w in words
        if w["column"] == "description" and w["text"] == "Charge"
    )

    if not charge_tops:
        return []

    boundaries = []
    for i, top in enumerate(charge_tops):
        if i + 1 < len(charge_tops):
            bottom_limit = charge_tops[i + 1]
        else:
            bottom_limit = max(w["bottom"] for w in words)
        boundaries.append((top, bottom_limit))

    return boundaries


def group_words_by_visual_line(words_in_col: list[dict]) -> list[tuple[str, float, float]]:
    """
    Retourne une liste de (texte, x0_premier_mot, x1_dernier_mot), une entrée
    par ligne visuelle (même 'top' arrondi).
    """
    by_top: dict[int, list[tuple[float, float, str]]] = {}
    for w in words_in_col:
        by_top.setdefault(round(w["top"]), []).append(
            (w["x0"], w["x1"], w["text"])
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
    """
    Joint des lignes visuelles en une chaîne.
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
    """Normalise les artefacts restants (apostrophe perdue, espaces multiples)."""
    description = re.sub(r"\s+", " ", description).strip()
    description = re.sub(r"McDonald\s+s\b", "McDonald's", description)
    description = re.sub(r"McDonald's\s+s\b", "McDonald's", description)
    return description


def parse_page(page: pdfplumber.page.Page) -> list[dict]:
    """Parse une page et retourne les lignes de charges."""
    words = extract_words_by_column(page)
    boundaries = find_row_boundaries(words)

    rows = []
    for y_start, y_end in boundaries:
        row_words = [w for w in words if y_start <= w["top"] < y_end]

        # Description : regroupement par ligne visuelle puis jointure par justification
        desc_lines = group_words_by_visual_line(
            [w for w in row_words if w["column"] == "description"]
        )
        description = clean_description(
            join_visual_lines(desc_lines, col_right=DESC_RIGHT, tol=JUSTIFY_TOL)
        )

        # Autres colonnes : simple concaténation (pas de wrap de mot problématique)
        cells: dict[str, list[str]] = {col: [] for col in COLUMNS}
        for w in row_words:
            cells[w["column"]].append(w["text"])

        pu_ht = next((t for t in cells["pu_ht"] if re.match(r"[\d.,]+\s*€?", t)), "")
        quantite = next((t for t in cells["quantite"] if re.match(r"^\d+$", t)), "")
        tva_pct = next((t for t in cells["tva_pct"] if re.match(r"[\d.,]+\s*%?", t)), "")

        tva_montant = ""
        for t in cells["tva_montant"]:
            cleaned = t.strip("()")
            if re.match(r"[\d.,]+\s*€?", cleaned):
                tva_montant = cleaned
                break

        total_ht = next((t for t in cells["total_ht"] if re.match(r"[\d.,]+\s*€?", t)), "")

        rows.append({
            "description": description,
            "pu_ht": pu_ht,
            "quantite": quantite,
            "tva_pct": tva_pct,
            "tva_montant": tva_montant,
            "total_ht": total_ht,
        })

    return rows


def parse_pdf(pdf_path: str) -> list[dict]:
    """Parse l'ensemble du PDF."""
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            all_rows.extend(parse_page(page))
    return all_rows


def to_float(val: str) -> float | None:
    """Convertit '6.25 €' ou '11,50' en float."""
    if not val:
        return None
    cleaned = re.sub(r"[€\s]", "", str(val)).replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def export(recharges: list[dict], output: str = "facture_chargemap.ods"):
    """Exporte vers ODS ou XLSX."""
    df = pd.DataFrame(recharges)

    for col in ["pu_ht", "tva_montant", "total_ht"]:
        df[col] = df[col].map(to_float)
    df["quantite"] = pd.to_numeric(df["quantite"], errors="coerce")

    engine = "odf" if output.endswith(".ods") else "openpyxl"
    with pd.ExcelWriter(output, engine=engine) as writer:
        df.to_excel(writer, sheet_name="Recharges", index=False)

    print(f"✅ {len(df)} lignes → {output}")


def debug_positions(pdf_path: str) -> None:
    """Affiche les positions x des mots pour calibrer COLUMNS et DESC_RIGHT."""
    with pdfplumber.open(pdf_path) as pdf:
        words = pdf.pages[0].extract_words(x_tolerance=2, y_tolerance=3)
        print("\nPositions X (échantillon) :")
        for w in words[:60]:
            print(f"  x0={w['x0']:.1f}  x1={w['x1']:.1f}  top={w['top']:.1f}  '{w['text']}'")


if __name__ == "__main__":
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "chargemap_20260925-07896428.pdf"
    output = sys.argv[2] if len(sys.argv) > 2 else "facture_chargemap.ods"

    if not Path(pdf_path).exists():
        print(f"❌ Fichier introuvable : {pdf_path}")
        sys.exit(1)

    recharges = parse_pdf(pdf_path)
    if not recharges:
        print("⚠️  Aucune ligne extraite. Ajustez les bornes COLUMNS / DESC_RIGHT.")
        debug_positions(pdf_path)
        sys.exit(1)

    export(recharges, output)
