#!/usr/bin/env python3
"""Extraction des factures Chargemap via pdfplumber.

Stratégie : découpage spatial en colonnes (par x) et lignes logiques (par y).
Une description qui wrappe sur 2 lignes visuelles reste dans la même cellule
car la "ligne logique" est délimitée par les marqueurs "Charge effectuée",
pas par les sauts de ligne visuels.
"""

import re
import sys

import pandas as pd
import pdfplumber


# ─── Configuration des colonnes ───────────────────────────────────────────────
# Plages X pour chaque colonne (à ajuster selon le PDF, mais stable pour Chargemap)
# Ces valeurs correspondent au layout standard des factures Chargemap.

COLUMNS = {
    "description": (0, 340),
    "pu_ht":       (340, 400),
    "quantite":    (400, 435),
    "tva_pct":     (435, 490),
    "tva_montant": (490, 545),
    "total_ht":    (545, 620),
}
MATCH_PRIX = r"[\d.,]+\s*€?"


def extract_words_by_column(page: pdfplumber.page.Page) -> list[dict]:
    """Extrait tous les mots de la page et les assigne à une colonne selon leur x.
    Retourne une liste de dicts : {"text", "x0", "x1", "top", "bottom", "column"}.

    :param page: page du pdf
    :return: liste des dictionnaires décrivant la ligne de facturation
    """
    words = page.extract_words(
        x_tolerance=2,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )

    result = []
    for w in words:
        # Centre du mot pour l'assignation de colonne
        cx = (w["x0"] + w["x1"]) / 2
        col = None
        for name, (x_min, x_max) in COLUMNS.items():
            if x_min <= cx < x_max:
                col = name
                break
        if col:
            result.append({
                "text": w["text"],
                "x0": w["x0"],
                "top": w["top"],
                "bottom": w["bottom"],
                "column": col,
            })
    return result


def find_row_boundaries(words: list[dict]) -> list[tuple[float, float]]:
    """Identifie les lignes logiques en cherchant les mots "Charge" qui marquent
    le début de chaque entrée. Retourne une liste de (y_start, y_end) pour chaque ligne.

    :param words: liste des mots
    :return: liste des tuples (y_start, y_end) pour chaque ligne logique
    """
    # Trouver les y des mots "Charge" dans la colonne description
    charge_tops = sorted(
        w["top"] for w in words
        if w["column"] == "description" and w["text"] == "Charge"
    )

    if not charge_tops:
        return []

    boundaries = []
    for i, top in enumerate(charge_tops):
        # La ligne logique s'étend du "Charge" jusqu'au prochain "Charge" (ou fin)
        if i + 1 < len(charge_tops):
            bottom_limit = charge_tops[i + 1]
        else:
            bottom_limit = max(w["bottom"] for w in words)
        boundaries.append((top, bottom_limit))

    return boundaries


def parse_page(page: pdfplumber.page.Page) -> list[dict]:
    """Parse une page et retourne les lignes de charges.

    :param page: page du pdf
    :return: list des dictionnaires décrivant la ligne de facturation
    """
    words = extract_words_by_column(page)
    boundaries = find_row_boundaries(words)

    rows = []
    for y_start, y_end in boundaries:
        # Filtrer les mots dans cette plage Y
        row_words = [
            w for w in words
            if y_start <= w["top"] < y_end
        ]

        # Regrouper par colonne
        cells: dict[str, list[str]] = {col: [] for col in COLUMNS}
        for w in row_words:
            cells[w["column"]].append(w["text"])

        # La description : tous les mots concaténés (les 2 lignes sont jointes)
        description = " ".join(cells["description"])

        # PU HT : premier montant trouvé
        pu_ht = ""
        for t in cells["pu_ht"]:
            if re.match(MATCH_PRIX, t):
                pu_ht = t
                break

        # Quantité : entier seul
        quantite = ""
        for t in cells["quantite"]:
            if re.match(r"^\d+$", t):
                quantite = t
                break

        # TVA % : ex "20.00"
        tva_pct = ""
        for t in cells["tva_pct"]:
            if re.match(r"[\d.,]+\s*%?", t):
                tva_pct = t
                break

        # TVA montant : le montant entre parenthèses ou isolé
        tva_montant = ""
        for t in cells["tva_montant"]:
            cleaned = t.strip("()")
            if re.match(MATCH_PRIX, cleaned):
                tva_montant = cleaned
                break

        # Total HT
        total_ht = ""
        for t in cells["total_ht"]:
            if re.match(r"[\d.,]+\s*€?", t):
                total_ht = t
                break

        # Nettoyage description : corriger les coupures OCR/PDF
        description = re.sub(r"\s+", " ", description).strip()
        description = re.sub(r"McDonald\s+s\b", "McDonald's", description)
        description = re.sub(r"Mc\s*D", "McD", description)
        description = re.sub(r"McDonald's\s+s\b", "McDonald's", description)

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
    """Parse l'ensemble du PDF.

    :param pdf_path: chemin vers le fichier PDF
    :return: liste des lignes de facturation
    """
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            all_rows.extend(parse_page(page))
    return all_rows


def to_float(val: str) -> float | None:
    """Convertit '6.25 €' ou '11,50' en float.

    :param val: valeur numérique brute
    :return: valeur nuémrique nettoyée
    """
    if not val:
        return None
    cleaned = re.sub(r"[€\s]", "", str(val))
    cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def export(recharges: list[dict], output: str = "factures_chargemap.ods"):
    """Exporte vers ODS/XLSX.

    :param recharges: liste des recharges
    :param output: nom du fichier de sortie, par défaut "factures_chargemap.ods"
    """
    df = pd.DataFrame(recharges)

    # Typage numérique
    for col in ["pu_ht", "tva_montant", "total_ht"]:
        df[col] = df[col].map(to_float)
    df["quantite"] = pd.to_numeric(df["quantite"], errors="coerce")

    # Écriture
    if output.endswith(".ods"):
        with pd.ExcelWriter(output, engine="odf") as writer:
            df.to_excel(writer, sheet_name="Recharges", index=False)
    else:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Recharges", index=False)

    print(f"✅ {len(df)} lignes → {output}")


if __name__ == "__main__":
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "chargemap_20260925-07896428.pdf"
    output = sys.argv[2] if len(sys.argv) > 2 else "facture_chargemap.ods"

    recharges = parse_pdf(pdf_path)
    if not recharges:
        print("⚠️  Aucune ligne extraite. Ajustez les bornes COLUMNS.")
        # Debug : afficher les positions x des mots pour calibrer
        with pdfplumber.open(pdf_path) as pdf:
            words = pdf.pages[0].extract_words(x_tolerance=2, y_tolerance=3)
            print("\nPositions X (échantillon) :")
            for w in words[:40]:
                print(f"  x0={w['x0']:.1f}  x1={w['x1']:.1f}  top={w['top']:.1f}  '{w['text']}'")
        sys.exit(1)

    export(recharges, output)
