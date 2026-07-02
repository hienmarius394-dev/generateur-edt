# -*- coding: utf-8 -*-
"""
Vérifications structurées pour l'interface :
  - validation des données importées (avant génération)
  - contrôle des déplacements manuels de cours (écran Ajuster)
  - bilan qualité d'un emploi du temps (trous, conflits)

Ne modifie jamais le moteur : réutilise ses constantes.
"""
from collections import defaultdict

from moteur_edt import (
    JOURS, N_JOURS, N_SLOTS, SLOT_LABELS,
    SLOTS_MATIN, SLOTS_APMIDI, IDX_DEBUT_APMIDI, JOUR_INDEX,
)

CRENEAUX_SEMAINE = N_JOURS * N_SLOTS  # 40


def _decoupage_possible(H, smin, smax):
    """True s'il existe une façon de découper H heures en séances dont chaque
    durée est comprise entre smin et smax (inclus). Sert à prévenir l'école
    AVANT de lancer le solveur quand un volume est incompatible avec la durée
    de séance imposée (ex: 5h de physique en séances de 2h pile → impossible)."""
    if H <= 0:
        return True
    smin = max(1, smin)
    smax = max(smin, smax)
    # Programmation dynamique : atteignable[h] = on peut composer h heures.
    atteignable = [False] * (H + 1)
    atteignable[0] = True
    for h in range(1, H + 1):
        for k in range(smin, min(smax, h) + 1):
            if atteignable[h - k]:
                atteignable[h] = True
                break
    return atteignable[H]


# ════════════════ VALIDATION DES DONNÉES IMPORTÉES ════════════════
def valider_donnees(classes, permanents, services, indispos,
                    profs_declares=None, heures_max=None,
                    jours_imposes_bruts=None):
    """Retourne (erreurs, avertissements) : listes de messages en français.
    Les erreurs sont bloquantes, les avertissements non.

    profs_declares : ensemble des noms de l'onglet Professeurs (tous statuts)
    heures_max     : dict prof -> heures max déclarées (optionnel)
    jours_imposes_bruts : liste de (ligne, valeur) de la colonne F non reconnues
    """
    erreurs, avertissements = [], []

    if not classes:
        erreurs.append("Aucune classe trouvée dans l'onglet « Classes ».")
    if not services:
        erreurs.append("Aucun service trouvé dans l'onglet « Services ».")
    if erreurs:
        return erreurs, avertissements

    # Doublons de classes
    doublons = {c for c in classes if classes.count(c) > 1}
    for c in sorted(doublons):
        erreurs.append(f"La classe « {c} » apparaît plusieurs fois dans l'onglet Classes.")

    # Cohérence des noms entre onglets
    classes_set = set(classes)
    for s in services:
        if s["classe"] not in classes_set:
            erreurs.append(
                f"Onglet Services : la classe « {s['classe']} » "
                f"(ligne {s['prof']} / {s['matiere']}) n'existe pas dans l'onglet Classes. "
                "Vérifiez l'orthographe exacte."
            )
    if profs_declares is not None:
        for s in services:
            if s["prof"] not in profs_declares:
                erreurs.append(
                    f"Onglet Services : le professeur « {s['prof']} » "
                    f"({s['matiere']} en {s['classe']}) n'existe pas dans l'onglet Professeurs. "
                    "Vérifiez l'orthographe exacte."
                )
        for p in sorted(indispos):
            if p not in profs_declares:
                avertissements.append(
                    f"Onglet Disponibilités : « {p} » n'existe pas dans l'onglet "
                    "Professeurs — cette ligne sera ignorée par le moteur si le nom "
                    "ne correspond à aucun service."
                )

    # Doublons de services (même prof + classe + matière)
    vus = defaultdict(int)
    for s in services:
        vus[(s["prof"], s["classe"], s["matiere"])] += 1
    for (p, c, m), n in vus.items():
        if n > 1:
            erreurs.append(
                f"Service en double : {m} en {c} avec {p} apparaît {n} fois. "
                "Fusionnez les lignes (additionnez les heures)."
            )

    # Charge par classe vs créneaux RÉELLEMENT disponibles.
    # La capacité brute est de 40 créneaux (5j × 8h), mais le mercredi
    # après-midi (3 créneaux) est interdit aux professeurs permanents. Pour
    # une classe donnée, ces 3 créneaux ne sont utilisables que si elle a au
    # moins un professeur VACATAIRE pouvant les couvrir. On calcule donc la
    # capacité réelle classe par classe pour éviter un angle mort : une classe
    # à 38h tous permanents passerait « 38 ≤ 40 » alors que sa vraie capacité
    # est 37 → le solveur échouerait sans message clair.
    h_classe = defaultdict(int)
    profs_de_classe = defaultdict(set)
    for s in services:
        h_classe[s["classe"]] += s["heures"]
        profs_de_classe[s["classe"]].add(s["prof"])

    for c in classes:
        # Créneaux mercredi après-midi réellement exploitables pour CETTE classe
        merc_pm_utilisables = 0
        for t in SLOTS_APMIDI:
            # un créneau merc PM est utilisable s'il existe un vacataire de la
            # classe disponible à ce créneau
            dispo_vacataire = any(
                (p not in permanents) and ((2, t) not in indispos.get(p, set()))
                for p in profs_de_classe[c]
            )
            if dispo_vacataire:
                merc_pm_utilisables += 1
        capacite_classe = CRENEAUX_SEMAINE - (len(SLOTS_APMIDI) - merc_pm_utilisables)

        if h_classe[c] > capacite_classe:
            if capacite_classe < CRENEAUX_SEMAINE:
                erreurs.append(
                    f"Classe {c} : {h_classe[c]}h demandées mais seulement "
                    f"{capacite_classe} créneaux réellement disponibles "
                    f"(le mercredi après-midi est réservé aux vacataires, et "
                    f"cette classe n'en a pas assez pour le couvrir). "
                    f"Réduisez les volumes ou affectez un vacataire au "
                    f"mercredi après-midi."
                )
            else:
                erreurs.append(
                    f"Classe {c} : {h_classe[c]}h demandées pour seulement "
                    f"{capacite_classe} créneaux par semaine. Réduisez les volumes."
                )
        elif h_classe[c] == 0:
            avertissements.append(f"Classe {c} : aucun service ne lui est affecté.")

    # Charge par professeur vs créneaux réellement disponibles
    h_prof = defaultdict(int)
    for s in services:
        h_prof[s["prof"]] += s["heures"]
    for p, h in sorted(h_prof.items()):
        dispo = CRENEAUX_SEMAINE - len(indispos.get(p, set()))
        if p in permanents:
            # mercredi après-midi bloqué (réunion), sauf créneaux déjà indisponibles
            deja = sum(1 for t in SLOTS_APMIDI if (2, t) in indispos.get(p, set()))
            dispo -= (len(SLOTS_APMIDI) - deja)
        if h > dispo:
            erreurs.append(
                f"{p} : {h}h de service mais seulement {dispo} créneaux "
                "disponibles (disponibilités et réunion du mercredi déduites)."
            )
        if heures_max and p in heures_max and heures_max[p] and h > heures_max[p]:
            avertissements.append(
                f"{p} : {h}h de service pour un maximum déclaré de "
                f"{heures_max[p]}h dans l'onglet Professeurs."
            )

    # Blocs : reste d'heures isolées (information utile, non bloquant)
    for s in services:
        b = s["bloc_size"]
        if b >= 2 and s["heures"] % b:
            avertissements.append(
                f"{s['matiere']} en {s['classe']} ({s['prof']}) : {s['heures']}h "
                f"en blocs de {b}h → {s['heures'] // b} bloc(s) + "
                f"{s['heures'] % b}h isolée(s)."
            )

    # Durée de séance imposée incompatible avec le volume horaire (BLOQUANT).
    # Ex : Physique 5h avec séances de 2h pile → 5 ne se découpe pas en 2.
    for s in services:
        smin = s.get("seance_min")
        smax = s.get("seance_max")
        if smin is None or smax is None:
            continue
        # Ne signaler que si une vraie contrainte est posée (pas le cas libre).
        contrainte_posee = not (smin == 1 and smax >= s["heures"])
        if contrainte_posee and not _decoupage_possible(s["heures"], smin, smax):
            if smin == smax:
                detail = f"séances de {smin}h exactement"
            else:
                detail = f"séances de {smin}h à {smax}h"
            erreurs.append(
                f"{s['matiere']} en {s['classe']} ({s['prof']}) : "
                f"{s['heures']}h ne peut pas se découper en {detail}. "
                f"Ajustez le volume horaire ou la durée de séance de cette "
                f"matière dans l'onglet Paramètres."
            )
        # Volume trop élevé pour tenir en une séance par jour sur la semaine.
        # (Le moteur place au plus UNE séance de la matière par jour pour
        #  garder les heures consécutives ; au-delà de seance_max × nb_jours,
        #  c'est mathématiquement impossible.)
        elif smax * N_JOURS < s["heures"]:
            erreurs.append(
                f"{s['matiere']} en {s['classe']} ({s['prof']}) : "
                f"{s['heures']}h impossible avec des séances d'au plus {smax}h. "
                f"Une matière n'a qu'une séance par jour, donc au maximum "
                f"{smax * N_JOURS}h par semaine ({smax}h × {N_JOURS} jours). "
                f"Augmentez la durée max de séance ou réduisez le volume."
            )

    # Jours imposés mal orthographiés (détectés à la lecture du fichier)
    if jours_imposes_bruts:
        for ligne, valeur in jours_imposes_bruts:
            erreurs.append(
                f"Onglet Services, ligne {ligne} : jour imposé « {valeur} » "
                f"non reconnu. Valeurs possibles : {', '.join(JOURS)} (ou vide)."
            )

    # Jour imposé incompatible avec une indisponibilité totale ce jour-là
    for s in services:
        d = s["jour_impose"]
        if d is None:
            continue
        bloque = indispos.get(s["prof"], set())
        slots_ok = [t for t in range(N_SLOTS) if (d, t) not in bloque]
        if s["prof"] in permanents and d == 2:
            slots_ok = [t for t in slots_ok if t not in SLOTS_APMIDI]
        besoin = s["bloc_size"] if s["bloc_size"] >= 2 and s["heures"] >= s["bloc_size"] else 1
        if len(slots_ok) < besoin:
            erreurs.append(
                f"{s['matiere']} en {s['classe']} : jour imposé {JOURS[d]} "
                f"mais {s['prof']} n'a pas assez de créneaux disponibles ce jour-là."
            )

    return erreurs, avertissements


# ════════════════ CONTRÔLE D'UN DÉPLACEMENT MANUEL ════════════════
def _prof_occupe_ailleurs(emplois, prof, classe_exclue, d, t):
    """Retourne la classe où `prof` est déjà en cours en (d, t), hors classe_exclue."""
    for (cl, dd, tt), info in emplois.items():
        if dd == d and tt == t and cl != classe_exclue and info["prof"] == prof:
            return cl
    return None


def _conflits_placement(emplois, classe, cours, d, t, permanents, indispos):
    """Conflits si on place `cours` (dict prof/matiere) en (classe, d, t).
    Le créneau (classe, d, t) est supposé libéré pour la classe."""
    conflits = []
    prof = cours["prof"]
    cible = f"{JOURS[d]} {SLOT_LABELS[t]}"

    autre = _prof_occupe_ailleurs(emplois, prof, classe, d, t)
    if autre:
        conflits.append(
            f"{prof} a déjà cours en {autre} le {cible}."
        )
    if (d, t) in indispos.get(prof, set()):
        conflits.append(
            f"{prof} est indisponible le {cible} (onglet Disponibilités)."
        )
    if prof in permanents and d == 2 and t in SLOTS_APMIDI:
        conflits.append(
            f"{prof} est permanent : réunion pédagogique le mercredi après-midi."
        )
    return conflits


def verifier_deplacement(emplois, classe, src, dst, permanents, indispos):
    """Contrôle le déplacement du cours de (classe, *src) vers (classe, *dst).
    Si la destination est occupée, contrôle l'ÉCHANGE des deux cours.

    Retourne (conflits, avertissements, est_echange).
    conflits vide = déplacement autorisé.
    """
    d1, t1 = src
    d2, t2 = dst
    cours_a = emplois.get((classe, d1, t1))
    if cours_a is None:
        return ["Aucun cours sur le créneau d'origine."], [], False
    if src == dst:
        return ["Le créneau d'arrivée est identique au créneau de départ."], [], False

    cours_b = emplois.get((classe, d2, t2))
    est_echange = cours_b is not None

    # Vue de travail sans les deux cours concernés
    travail = dict(emplois)
    travail.pop((classe, d1, t1), None)
    travail.pop((classe, d2, t2), None)

    conflits = _conflits_placement(travail, classe, cours_a, d2, t2,
                                   permanents, indispos)
    if est_echange:
        conflits += _conflits_placement(travail, classe, cours_b, d1, t1,
                                        permanents, indispos)

    # Avertissements non bloquants : casse d'un bloc consécutif
    avertissements = []
    for (dd, tt), c in ((src, cours_a),) + (((d2, t2), cours_b),) if est_echange else ((src, cours_a),):
        voisins = [(dd, tt - 1), (dd, tt + 1)]
        for (vd, vt) in voisins:
            v = emplois.get((classe, vd, vt))
            if v and v["prof"] == c["prof"] and v["matiere"] == c["matiere"]:
                avertissements.append(
                    f"Attention : {c['matiere']} ({JOURS[dd]}) faisait partie "
                    "d'un bloc d'heures consécutives — le déplacement le sépare."
                )
                break
    return conflits, avertissements, est_echange


def appliquer_deplacement(emplois, classe, src, dst):
    """Applique le déplacement (ou l'échange) et retourne le nouveau dict."""
    d1, t1 = src
    d2, t2 = dst
    nouveau = dict(emplois)
    a = nouveau.pop((classe, d1, t1))
    b = nouveau.pop((classe, d2, t2), None)
    nouveau[(classe, d2, t2)] = a
    if b is not None:
        nouveau[(classe, d1, t1)] = b
    return nouveau


# ════════════════ PERMUTATION EN CHAÎNE (déplacements groupés) ════════════════
def creneaux_libres_prof(emplois, prof, permanents, indispos, classe=None):
    """Liste les créneaux où `prof` peut réellement donner cours :
    disponible (onglet Disponibilités), pas en réunion (permanent le mercredi
    après-midi) et pas déjà en cours dans une classe à ce moment-là.

    Retourne une liste de (jour, slot, classe_libre) triée, où classe_libre
    indique si `classe` (si fournie) est elle aussi libre sur ce créneau :
      True  → déplacement direct possible,
      False → le créneau de la classe est occupé (ce serait un échange).
    """
    occupe = {(d, t) for (cl, d, t), info in emplois.items()
              if info["prof"] == prof}
    indispo_prof = indispos.get(prof, set())
    libres = []
    for d in range(N_JOURS):
        for t in range(N_SLOTS):
            if (d, t) in indispo_prof or (d, t) in occupe:
                continue
            if prof in permanents and d == 2 and t in SLOTS_APMIDI:
                continue
            cl_libre = (classe is None
                        or (classe, d, t) not in emplois)
            libres.append((d, t, cl_libre))
    return libres


def verifier_chaine(emplois, classe, etapes, permanents, indispos):
    """Vérifie une CHAÎNE de déplacements à appliquer dans l'ordre, sur une
    copie de travail (l'emploi du temps réel n'est jamais modifié ici).

    etapes : liste de (src, dst), chacun un (jour, slot).
    Chaque étape déplace le cours qui se trouve EN CE MOMENT sur `src`
    (donc après les étapes précédentes) vers `dst`. Si `dst` est occupé,
    c'est un échange — exactement la même sémantique qu'un déplacement seul.

    Retourne (conflits, avertissements, etapes_resolues) :
      - conflits : liste de messages ; si non vide, RIEN ne doit être appliqué.
        Chaque message précise le numéro de l'étape en cause.
      - avertissements : messages non bloquants (ex: bloc consécutif cassé).
      - etapes_resolues : liste de (src, dst, est_echange), dans le même
        ordre, prête à être rejouée par appliquer_chaine si conflits est vide.
    """
    travail = dict(emplois)
    conflits, avertissements, resolues = [], [], []

    for i, (src, dst) in enumerate(etapes, start=1):
        c, a, ech = verifier_deplacement(
            travail, classe, src, dst, permanents, indispos,
        )
        if c:
            conflits.extend(f"Étape {i} ({JOURS[src[0]]} {SLOT_LABELS[src[1]]} "
                             f"→ {JOURS[dst[0]]} {SLOT_LABELS[dst[1]]}) : {m}"
                             for m in c)
            # On arrête dès la première étape en échec : les étapes suivantes
            # dépendent de l'état produit par celle-ci et n'ont pas de sens
            # tant qu'elle n'est pas résolue.
            break
        avertissements.extend(a)
        resolues.append((src, dst, ech))
        travail = appliquer_deplacement(travail, classe, src, dst)

    return conflits, avertissements, resolues


def appliquer_chaine(emplois, classe, etapes_resolues):
    """Rejoue dans l'ordre les étapes déjà validées par verifier_chaine
    (etapes_resolues = sortie de verifier_chaine, conflits vide requis).
    Retourne le nouveau dict d'emplois."""
    courant = dict(emplois)
    for src, dst, _ech in etapes_resolues:
        courant = appliquer_deplacement(courant, classe, src, dst)
    return courant


# ════════════════ BILAN QUALITÉ DE L'ÉTAT COURANT ════════════════
def bilan_etat(emplois, classes, permanents, services, indispos):
    """Re-vérifie l'état courant (après modifications manuelles).
    Retourne une liste de (statut, message) avec statut ∈ {ok, erreur, info}."""
    bilan = []

    # Volumes horaires
    placés = defaultdict(int)
    for (cl, d, t), info in emplois.items():
        placés[(info["prof"], cl, info["matiere"])] += 1
    manquants = [s for s in services
                 if placés[(s["prof"], s["classe"], s["matiere"])] != s["heures"]]
    if manquants:
        bilan.append(("erreur", f"{len(manquants)} volume(s) horaire(s) incorrect(s)."))
    else:
        bilan.append(("ok", f"Volumes horaires exacts ({len(services)}/{len(services)} services)."))

    # Conflits professeurs
    vus, n_conf = set(), 0
    for (cl, d, t), info in emplois.items():
        k = (info["prof"], d, t)
        if k in vus:
            n_conf += 1
        vus.add(k)
    bilan.append(("ok", "Aucun conflit professeur.") if n_conf == 0
                 else ("erreur", f"{n_conf} conflit(s) professeur."))

    # Disponibilités
    n = sum(1 for (cl, d, t), i in emplois.items()
            if (d, t) in indispos.get(i["prof"], set()))
    bilan.append(("ok", "Disponibilités toutes respectées.") if n == 0
                 else ("erreur", f"{n} cours sur des créneaux indisponibles."))

    # Mercredi après-midi permanents
    n = sum(1 for (cl, d, t), i in emplois.items()
            if d == 2 and t in SLOTS_APMIDI and i["prof"] in permanents)
    bilan.append(("ok", "Mercredi après-midi : aucun permanent.") if n == 0
                 else ("erreur", f"{n} permanent(s) placé(s) le mercredi après-midi."))

    # Trous (information, non bloquant) — sur la journée entière : un créneau
    # vide entouré de cours avant ET après compte, y compris à l'heure du midi.
    trous = 0
    for cl in classes:
        for d in range(N_JOURS):
            occ = [(cl, d, t) in emplois for t in range(N_SLOTS)]
            for t in range(1, N_SLOTS):
                if not occ[t] and any(occ[:t]) and any(occ[t:]):
                    trous += 1
    bilan.append(("ok", "Aucun trou dans les grilles des classes.") if trous == 0
                 else ("info", f"{trous} trou(s) dans les grilles (créneau vide entre deux cours)."))

    return bilan
