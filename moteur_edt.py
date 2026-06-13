#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔════════════════════════════════════════════════════════════╗
║  MOTEUR EDT — Générateur d'emplois du temps scolaires      ║
║                                                            ║
║  Usage :  python3 moteur_edt.py donnees.xlsx resultat.xlsx ║
║                                                            ║
║  Entrée : le template Excel rempli (onglets Classes,       ║
║  Professeurs, Services, Disponibilités).                   ║
║  Sortie : un Excel avec une grille par classe + la vue     ║
║  par professeur. Zéro conflit garanti (OR-Tools CP-SAT).   ║
╚════════════════════════════════════════════════════════════╝

Règles gérées :
  H1  Volume horaire exact pour chaque service
  H2  Une classe = 1 seul cours par créneau
  H3  Un professeur = 1 seul cours par créneau
  H4  Indisponibilités des professeurs (onglet Disponibilités)
  H5  Permanents bloqués le mercredi après-midi (réunion)
  H6  Maximum 2h de la même matière par jour (hors blocs)
  H7  Jour imposé (optionnel) : le bloc / une séance tombe ce jour-là
  Blocs de 2h ou 3h consécutives (colonne optionnelle ;
  vide = l'outil décide librement). Les blocs peuvent
  traverser la pause déjeuner (choix de l'établissement).

Qualité optimisée :
  + matières scientifiques le matin
  − EPS à 14h30 (chaleur)   − matière lourde en fin de journée
  − trous dans la journée (classes ET professeurs)
"""

import sys
from collections import defaultdict

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model

# ════════════════════════ CONFIG ════════════════════════
JOURS   = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]
N_JOURS = len(JOURS)
N_SLOTS = 8

SLOT_LABELS = [
    "07h00 – 08h00",
    "08h00 – 09h00",
    "09h00 – 10h00",
    "10h15 – 11h15",
    "11h15 – 12h15",
    "14h30 – 15h30",
    "15h30 – 16h30",
    "16h45 – 17h45",
]
SLOTS_MATIN      = list(range(5))
SLOTS_APMIDI     = list(range(5, 8))
IDX_DEBUT_APMIDI = 5
JOUR_INDEX = {j: i for i, j in enumerate(JOURS)}

SCIENCES = {"Mathématiques", "Physique-Chimie", "SVT"}
LOURDES  = {"Mathématiques", "Philosophie"}


# ════════════════════ 1. LECTURE EXCEL ════════════════════
def _val(row, i):
    """Valeur texte propre d'une cellule pandas, '' si vide."""
    if i >= len(row) or pd.isna(row.iloc[i]):
        return ""
    return str(row.iloc[i]).strip()


def lire_excel(chemin):
    xls = pd.ExcelFile(chemin)

    # ── Classes ──
    df = pd.read_excel(xls, "Classes", skiprows=2, header=0)
    df = df.dropna(subset=[df.columns[0]])
    classes = [_val(r, 0) for _, r in df.iterrows() if _val(r, 0)]

    # ── Professeurs (statut) ──
    df = pd.read_excel(xls, "Professeurs", skiprows=2, header=0)
    df = df.dropna(subset=[df.columns[0]])
    permanents = {_val(r, 0) for _, r in df.iterrows()
                  if _val(r, 0) and _val(r, 3) == "Permanent"}

    # ── Services ──
    # Colonnes : A prof | B classe | C matière | D heures
    #            E taille du bloc (optionnel : vide/Indifférent/2h/3h)
    #            F jour imposé   (optionnel : vide ou Lundi…Vendredi)
    df = pd.read_excel(xls, "Services", skiprows=2, header=0)
    df = df.dropna(subset=[df.columns[0]])
    services = []
    for _, r in df.iterrows():
        prof, classe, matiere = _val(r, 0), _val(r, 1), _val(r, 2)
        try:
            heures = int(float(_val(r, 3)))
        except ValueError:
            continue
        if not (prof and classe and matiere and heures > 0):
            continue
        bv = _val(r, 4)
        bloc_size = 3 if bv == "3h" else (2 if bv in ("2h", "Oui") else 1)
        jour_impose = JOUR_INDEX.get(_val(r, 5))   # None si vide
        services.append({"prof": prof, "classe": classe,
                         "matiere": matiere, "heures": heures,
                         "bloc_size": bloc_size,
                         "jour_impose": jour_impose})

    # ── Disponibilités (demi-journées NON = indisponible) ──
    df = pd.read_excel(xls, "Disponibilités", skiprows=2, header=0)
    df = df.dropna(subset=[df.columns[0]])
    mapping = [(j, sl) for j in range(N_JOURS)
               for sl in (SLOTS_MATIN, SLOTS_APMIDI)]
    indispos = defaultdict(set)
    for _, r in df.iterrows():
        prof = _val(r, 0)
        for ci, (jour, slots) in enumerate(mapping, start=1):
            if _val(r, ci).upper() == "NON":
                for s in slots:
                    indispos[prof].add((jour, s))

    return classes, permanents, services, indispos


# ════════════════════ 2. SOLVEUR ════════════════════
def resoudre(classes, permanents, services, indispos, temps_max=120):
    model = cp_model.CpModel()

    # ── Variable principale : x[s,d,t] = service s placé jour d, créneau t
    x = {}
    for s in range(len(services)):
        for d in range(N_JOURS):
            for t in range(N_SLOTS):
                x[s, d, t] = model.new_bool_var(f"x_{s}_{d}_{t}")

    # ── Blocs de B heures consécutives (B = 2 ou 3) ──
    # NOTE : les blocs PEUVENT traverser la pause déjeuner (choix assumé).
    bloc_st = {}
    for s, svc in enumerate(services):
        B = svc["bloc_size"]
        if B < 2:
            continue
        n_blocs = svc["heures"] // B
        n_extra = svc["heures"] % B

        for d in range(N_JOURS):
            for t in range(N_SLOTS - B + 1):
                bloc_st[s, d, t] = model.new_bool_var(f"b{B}_{s}_{d}_{t}")

        extra = {}
        if n_extra:
            for d in range(N_JOURS):
                for t in range(N_SLOTS):
                    extra[d, t] = model.new_bool_var(f"e_{s}_{d}_{t}")
            model.add(sum(extra.values()) == n_extra)

        model.add(sum(bloc_st[s, d, t]
                      for d in range(N_JOURS)
                      for t in range(N_SLOTS - B + 1)) == n_blocs)

        # Blocs sans chevauchement
        for d in range(N_JOURS):
            for t in range(N_SLOTS - B + 1):
                for off in range(1, B):
                    if t + off <= N_SLOTS - B:
                        model.add(bloc_st[s, d, t]
                                  + bloc_st[s, d, t + off] <= 1)

        # Lien x = couverture des blocs + heures isolées
        for d in range(N_JOURS):
            for t in range(N_SLOTS):
                cov = [bloc_st[s, d, ts]
                       for ts in range(max(0, t - B + 1),
                                       min(N_SLOTS - B, t) + 1)]
                ext = [extra[d, t]] if n_extra else []
                model.add(x[s, d, t] == sum(cov) + sum(ext))

        # ≤ B heures de ce service par jour
        for d in range(N_JOURS):
            model.add(sum(x[s, d, t] for t in range(N_SLOTS)) <= B)

    # ── H1 : volume horaire exact ──
    for s, svc in enumerate(services):
        model.add(sum(x[s, d, t]
                      for d in range(N_JOURS)
                      for t in range(N_SLOTS)) == svc["heures"])

    # ── Index classes / profs ──
    cl_svcs, pr_svcs = defaultdict(list), defaultdict(list)
    for s, svc in enumerate(services):
        cl_svcs[svc["classe"]].append(s)
        pr_svcs[svc["prof"]].append(s)

    # ── H2 / H3 : pas de doublon classe ni prof ──
    for groupe in list(cl_svcs.values()) + list(pr_svcs.values()):
        for d in range(N_JOURS):
            for t in range(N_SLOTS):
                model.add(sum(x[s, d, t] for s in groupe) <= 1)

    # ── H4 : indisponibilités ──
    for s, svc in enumerate(services):
        for (d, t) in indispos.get(svc["prof"], set()):
            model.add(x[s, d, t] == 0)

    # ── H5 : permanents bloqués mercredi après-midi ──
    for s, svc in enumerate(services):
        if svc["prof"] in permanents:
            for t in SLOTS_APMIDI:
                model.add(x[s, 2, t] == 0)

    # ── H6 : ≤ 2h par jour pour les services sans bloc ──
    for s, svc in enumerate(services):
        if svc["bloc_size"] == 1:
            for d in range(N_JOURS):
                model.add(sum(x[s, d, t] for t in range(N_SLOTS)) <= 2)

    # ── H7 : jour imposé (optionnel) ──
    for s, svc in enumerate(services):
        d = svc["jour_impose"]
        if d is None:
            continue
        B = svc["bloc_size"]
        if B >= 2 and svc["heures"] >= B:
            # le bloc doit commencer ce jour-là
            model.add(sum(bloc_st[s, d, t]
                          for t in range(N_SLOTS - B + 1)) >= 1)
        else:
            # au moins une séance ce jour-là
            model.add(sum(x[s, d, t] for t in range(N_SLOTS)) >= 1)

    # ── Objectif : qualité pédagogique + compacité ──
    bonus, malus = [], []
    for s, svc in enumerate(services):
        for d in range(N_JOURS):
            if svc["matiere"] in SCIENCES:
                for t in SLOTS_MATIN:
                    bonus.append(x[s, d, t])
            if svc["matiere"] == "EPS":
                malus.append(x[s, d, IDX_DEBUT_APMIDI])
            if svc["matiere"] in LOURDES:
                malus.append(x[s, d, N_SLOTS - 1])

    # Compacité : pénaliser chaque "trou" (créneau vide entre deux cours)
    # dans une demi-journée, pour les classes ET pour les profs.
    def penaliser_trous(groupes, prefixe):
        for g_idx, s_list in enumerate(groupes):
            for d in range(N_JOURS):
                def occ(t):
                    return sum(x[s, d, t] for s in s_list)
                for plage in (range(1, 5), range(6, 8)):
                    for t in plage:
                        j = model.new_bool_var(f"{prefixe}_{g_idx}_{d}_{t}")
                        model.add(j >= occ(t) - occ(t - 1))
                        model.add(j <= occ(t))
                        model.add(j <= 1 - occ(t - 1))
                        malus.append(2 * j)

    penaliser_trous(list(cl_svcs.values()), "tc")
    penaliser_trous(list(pr_svcs.values()), "tp")

    model.maximize(sum(bonus) - sum(malus))

    # ── Résolution ──
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = temps_max
    solver.parameters.num_search_workers  = 4

    print(f"  Résolution en cours (max {temps_max}s)…")
    status = solver.solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print(f"  ✗ Aucune solution ({solver.status_name(status)}).")
        print("    → Vérifiez disponibilités, volumes et jours imposés.")
        return None
    print(f"  ✓ Solution {solver.status_name(status)} "
          f"(score qualité = {int(solver.objective_value)})")

    emplois = {}
    for s, svc in enumerate(services):
        for d in range(N_JOURS):
            for t in range(N_SLOTS):
                if solver.value(x[s, d, t]):
                    emplois[(svc["classe"], d, t)] = {
                        "prof": svc["prof"], "matiere": svc["matiere"]}
    return emplois


# ════════════════ 3. VÉRIFICATEUR INTÉGRÉ ════════════════
def verifier(emplois, classes, permanents, services, indispos):
    ok = True

    # Volumes horaires exacts
    placed = defaultdict(int)
    for (cl, d, t), info in emplois.items():
        placed[(info["prof"], cl, info["matiere"])] += 1
    nb_ko = sum(1 for s in services
                if placed[(s["prof"], s["classe"], s["matiere"])] != s["heures"])
    if nb_ko:
        print(f"  ✗ {nb_ko} volumes horaires incorrects !")
        ok = False
    else:
        print(f"  ✓ Volumes horaires : {len(services)}/{len(services)} exacts")

    # Conflits classe / prof
    seen_p, conflits = set(), 0
    for (cl, d, t), info in emplois.items():
        k = (info["prof"], d, t)
        if k in seen_p:
            conflits += 1
        seen_p.add(k)
    if conflits:
        print(f"  ✗ {conflits} conflits professeur !")
        ok = False
    else:
        print("  ✓ Zéro conflit (classes et professeurs)")

    # Indisponibilités
    nb = sum(1 for (cl, d, t), i in emplois.items()
             if (d, t) in indispos.get(i["prof"], set()))
    print(f"  {'✓' if nb == 0 else '✗'} Disponibilités : "
          f"{'toutes respectées' if nb == 0 else str(nb) + ' violations !'}")
    ok = ok and nb == 0

    # Mercredi PM permanents
    nb = sum(1 for (cl, d, t), i in emplois.items()
             if d == 2 and t in SLOTS_APMIDI and i["prof"] in permanents)
    print(f"  {'✓' if nb == 0 else '✗'} Mercredi après-midi : "
          f"{'aucun permanent' if nb == 0 else str(nb) + ' permanents placés !'}")
    ok = ok and nb == 0

    # Trous dans les grilles classes
    trous = 0
    for cl in classes:
        for d in range(N_JOURS):
            occ = [(cl, d, t) in emplois for t in range(N_SLOTS)]
            for plage in (range(1, 5), range(6, 8)):
                for t in plage:
                    if occ[t] and not occ[t - 1] \
                       and any(occ[tp] for tp in range(plage.start - 1, t)):
                        trous += 1
    print(f"  {'✓' if trous == 0 else '!'} Trous dans les grilles : {trous}")

    # Jours imposés
    for s in services:
        d = s["jour_impose"]
        if d is None:
            continue
        slots = sorted(t for t in range(N_SLOTS)
                       if emplois.get((s["classe"], d, t), {}).get("prof") == s["prof"]
                       and emplois[(s["classe"], d, t)]["matiere"] == s["matiere"])
        B = s["bloc_size"]
        if B >= 2 and s["heures"] >= B:
            # chercher B créneaux consécutifs
            run, best = 1, 1 if slots else 0
            for i in range(1, len(slots)):
                run = run + 1 if slots[i] == slots[i - 1] + 1 else 1
                best = max(best, run)
            respecte = best >= B
        else:
            respecte = len(slots) >= 1
        sym = "✓" if respecte else "✗"
        print(f"  {sym} Jour imposé : {s['matiere']} {s['classe']} "
              f"→ {JOURS[d]} ({len(slots)}h placées"
              f"{', bloc de ' + str(B) + 'h présent' if B >= 2 and respecte else ''})")
        ok = ok and respecte

    return ok


# ════════════════════ 4. EXPORT EXCEL ════════════════════
def _brd(top=False):
    t = Side(style="thin",   color="AAAAAA")
    m = Side(style="medium", color="444444")
    return Border(left=t, right=t, top=m if top else t, bottom=t)


def _hdr(cell, texte):
    cell.value = texte
    cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", start_color="444444")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _brd()


def _grille_jours(ws, row0, get_contenu):
    """Dessine une grille (jours en colonnes, créneaux en lignes)
    à partir de la ligne row0. get_contenu(d, t) → texte ou None."""
    ws.cell(row=row0, column=1).fill = PatternFill("solid", start_color="444444")
    ws.cell(row=row0, column=1).border = _brd()
    for j, jour in enumerate(JOURS, start=2):
        _hdr(ws.cell(row=row0, column=j), jour)
    ws.row_dimensions[row0].height = 19

    row = row0 + 1
    for t_idx, label in enumerate(SLOT_LABELS):
        if t_idx == IDX_DEBUT_APMIDI:
            ws.merge_cells(start_row=row, start_column=1,
                           end_row=row, end_column=6)
            sep = ws.cell(row=row, column=1, value="APRES-MIDI")
            sep.font = Font(name="Arial", size=8, bold=True)
            sep.fill = PatternFill("solid", start_color="DDDDDD")
            sep.alignment = Alignment(horizontal="center", vertical="center")
            sep.border = _brd(top=True)
            ws.row_dimensions[row].height = 12
            row += 1

        ws.row_dimensions[row].height = 36
        h = ws.cell(row=row, column=1, value=label)
        h.font = Font(name="Arial", size=8, bold=True)
        h.fill = PatternFill("solid", start_color="F2F2F2")
        h.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        h.border = _brd()

        for d_idx in range(N_JOURS):
            c = ws.cell(row=row, column=d_idx + 2)
            contenu = get_contenu(d_idx, t_idx)
            if contenu:
                c.value = contenu
                c.font = Font(name="Arial", size=9)
                c.fill = PatternFill("solid", start_color="FFFFFF")
            else:
                c.fill = PatternFill("solid", start_color="FAFAFA")
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
            c.border = _brd()
        row += 1
    return row


def ecrire_excel(emplois, classes, chemin):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Une feuille par classe ──
    for cl in classes:
        ws = wb.create_sheet(title=cl[:31])
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 15
        for j in range(2, 7):
            ws.column_dimensions[get_column_letter(j)].width = 22

        ws.merge_cells("A1:F1")
        t = ws["A1"]
        t.value = f"Emploi du temps  —  {cl}"
        t.font = Font(name="Arial", size=13, bold=True)
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        def contenu(d, ti, _cl=cl):
            info = emplois.get((_cl, d, ti))
            if not info:
                return None
            nom = info["prof"].replace("M. ", "")
            return f"{info['matiere']}\n{nom}"

        _grille_jours(ws, 2, contenu)

    # ── Vue par professeur (texte simple, AUCUN emoji) ──
    ws = wb.create_sheet(title="Vue Professeurs")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 15
    for j in range(2, 7):
        ws.column_dimensions[get_column_letter(j)].width = 22

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = "Emplois du temps  —  Vue par professeur"
    t.font = Font(name="Arial", size=13, bold=True)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    edt_prof = defaultdict(dict)
    for (cl, d, ti), info in emplois.items():
        edt_prof[info["prof"]][(d, ti)] = f"{info['matiere']}\n{cl}"

    cur = 2
    for prof in sorted(edt_prof.keys()):
        ws.merge_cells(start_row=cur, start_column=1,
                       end_row=cur, end_column=6)
        c = ws.cell(row=cur, column=1, value=f"  {prof}")
        c.font = Font(name="Arial", size=10, bold=True)
        c.fill = PatternFill("solid", start_color="CCCCCC")
        c.alignment = Alignment(vertical="center")
        c.border = _brd()
        ws.row_dimensions[cur].height = 18
        cur += 1

        def contenu_p(d, ti, _p=prof):
            return edt_prof[_p].get((d, ti))

        cur = _grille_jours(ws, cur, contenu_p) + 1

    wb.save(chemin)
    print(f"  ✓ Fichier écrit : {chemin}")
    print(f"    Onglets : {', '.join(wb.sheetnames)}")


# ════════════════════ 5. POINT D'ENTRÉE ════════════════════
if __name__ == "__main__":
    entree = sys.argv[1] if len(sys.argv) > 1 else "donnees.xlsx"
    sortie = sys.argv[2] if len(sys.argv) > 2 else "emplois_du_temps.xlsx"

    print("=" * 55)
    print("  MOTEUR EDT — Génération des emplois du temps")
    print("=" * 55)

    print("\n[1/4] Lecture des données…")
    classes, permanents, services, indispos = lire_excel(entree)
    n_b2 = sum(1 for s in services if s["bloc_size"] == 2)
    n_b3 = sum(1 for s in services if s["bloc_size"] == 3)
    n_ji = sum(1 for s in services if s["jour_impose"] is not None)
    print(f"  {len(classes)} classes | {len(services)} services | "
          f"{sum(s['heures'] for s in services)}h à placer")
    print(f"  Blocs 2h : {n_b2} | Blocs 3h : {n_b3} | Jours imposés : {n_ji}")

    print("\n[2/4] Résolution…")
    emplois = resoudre(classes, permanents, services, indispos)
    if not emplois:
        sys.exit(1)

    print("\n[3/4] Vérification automatique…")
    verifier(emplois, classes, permanents, services, indispos)

    print("\n[4/4] Export Excel…")
    ecrire_excel(emplois, classes, sortie)
    print("\n✓ Terminé.")
