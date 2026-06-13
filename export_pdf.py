# -*- coding: utf-8 -*-
"""
Export PDF des emplois du temps — pensé pour l'impression papier :
  - A4 paysage, noir et blanc uniquement
  - une page par classe (pdf_classes) ou par professeur (pdf_profs)
  - séparateur visuel matin / après-midi, identique à l'export Excel
"""
import io
from collections import defaultdict
from datetime import date

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak,
)

from moteur_edt import JOURS, N_JOURS, N_SLOTS, SLOT_LABELS, IDX_DEBUT_APMIDI

GRIS_FONCE = colors.HexColor("#444444")
GRIS_CLAIR = colors.HexColor("#DDDDDD")
GRIS_FOND = colors.HexColor("#F2F2F2")

STYLE_CELLULE = ParagraphStyle(
    "cellule", fontName="Helvetica", fontSize=8.5, leading=10.5,
    alignment=TA_CENTER,
)
STYLE_MATIERE = ParagraphStyle(
    "matiere", fontName="Helvetica-Bold", fontSize=8.5, leading=10.5,
    alignment=TA_CENTER,
)
STYLE_TITRE = ParagraphStyle(
    "titre", fontName="Helvetica-Bold", fontSize=15, leading=18,
    alignment=TA_CENTER, spaceAfter=2,
)
STYLE_SOUS_TITRE = ParagraphStyle(
    "soustitre", fontName="Helvetica", fontSize=9.5, leading=12,
    alignment=TA_CENTER, textColor=GRIS_FONCE,
)


def _esc(txt):
    return (txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _cellule(ligne1, ligne2):
    """Paragraphe à deux lignes : matière en gras, détail dessous."""
    return Paragraph(
        f"<b>{_esc(ligne1)}</b><br/><font size=8>{_esc(ligne2)}</font>",
        STYLE_CELLULE,
    )


def _table_grille(get_contenu):
    """Construit la Table reportlab d'une grille hebdomadaire.
    get_contenu(d, t) -> (ligne1, ligne2) ou None."""
    lignes = [[""] + JOURS]                       # en-tête
    index_separateur = None

    for t in range(N_SLOTS):
        if t == IDX_DEBUT_APMIDI:
            index_separateur = len(lignes)
            lignes.append(["APRÈS-MIDI"] + [""] * N_JOURS)
        rang = [Paragraph(f"<b>{SLOT_LABELS[t]}</b>",
                          ParagraphStyle("h", parent=STYLE_CELLULE, fontSize=8))]
        for d in range(N_JOURS):
            contenu = get_contenu(d, t)
            rang.append(_cellule(*contenu) if contenu else "")
        lignes.append(rang)

    largeur_h = 28 * mm
    largeur_j = (267 * mm - largeur_h) / N_JOURS
    hauteurs = []
    for i in range(len(lignes)):
        if i == 0:
            hauteurs.append(8 * mm)
        elif i == index_separateur:
            hauteurs.append(5 * mm)
        else:
            hauteurs.append(16.2 * mm)

    table = Table(lignes, colWidths=[largeur_h] + [largeur_j] * N_JOURS,
                  rowHeights=hauteurs)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#AAAAAA")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        # En-tête jours
        ("BACKGROUND", (0, 0), (-1, 0), GRIS_FONCE),
        ("TEXTCOLOR", (1, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        # Colonne horaires
        ("BACKGROUND", (0, 1), (0, -1), GRIS_FOND),
    ]
    if index_separateur is not None:
        style += [
            ("SPAN", (0, index_separateur), (-1, index_separateur)),
            ("BACKGROUND", (0, index_separateur), (-1, index_separateur), GRIS_CLAIR),
            ("FONTNAME", (0, index_separateur), (-1, index_separateur), "Helvetica-Bold"),
            ("FONTSIZE", (0, index_separateur), (-1, index_separateur), 7),
            ("LINEABOVE", (0, index_separateur), (-1, index_separateur), 1.1, GRIS_FONCE),
        ]
    table.setStyle(TableStyle(style))
    return table


def _pied_de_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(GRIS_FONCE)
    canvas.drawString(15 * mm, 8 * mm,
                      f"Généré le {date.today().strftime('%d/%m/%Y')}")
    canvas.drawRightString(282 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()


def _document(buffer):
    return SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
        title="Emplois du temps",
    )


def _page(story, titre, sous_titre, table, derniere=False):
    story.append(Paragraph(_esc(titre), STYLE_TITRE))
    story.append(Paragraph(_esc(sous_titre), STYLE_SOUS_TITRE))
    story.append(Spacer(1, 4 * mm))
    story.append(table)
    if not derniere:
        story.append(PageBreak())


def pdf_classes(emplois, classes, etablissement="", annee=""):
    """PDF multi-pages : une page A4 paysage par classe. Retourne bytes."""
    buffer = io.BytesIO()
    doc = _document(buffer)
    story = []
    sous_titre = " — ".join(x for x in (etablissement, f"Année scolaire {annee}" if annee else "") if x)

    for i, cl in enumerate(classes):
        def contenu(d, t, _cl=cl):
            info = emplois.get((_cl, d, t))
            if not info:
                return None
            return info["matiere"], info["prof"].replace("M. ", "")

        _page(story, f"Emploi du temps — {cl}",
              sous_titre or " ", _table_grille(contenu),
              derniere=(i == len(classes) - 1))

    doc.build(story, onFirstPage=_pied_de_page, onLaterPages=_pied_de_page)
    return buffer.getvalue()


def pdf_profs(emplois, etablissement="", annee=""):
    """PDF multi-pages : une page A4 paysage par professeur. Retourne bytes."""
    edt_prof = defaultdict(dict)
    for (cl, d, t), info in emplois.items():
        edt_prof[info["prof"]][(d, t)] = (info["matiere"], cl)

    buffer = io.BytesIO()
    doc = _document(buffer)
    story = []
    sous_titre = " — ".join(x for x in (etablissement, f"Année scolaire {annee}" if annee else "") if x)
    profs = sorted(edt_prof.keys())

    for i, prof in enumerate(profs):
        def contenu(d, t, _p=prof):
            return edt_prof[_p].get((d, t))

        total_h = len(edt_prof[prof])
        _page(story, f"Emploi du temps — {prof}",
              (sous_titre + " — " if sous_titre else "") + f"{total_h}h / semaine",
              _table_grille(contenu),
              derniere=(i == len(profs) - 1))

    doc.build(story, onFirstPage=_pied_de_page, onLaterPages=_pied_de_page)
    return buffer.getvalue()
