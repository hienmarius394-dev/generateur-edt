# -*- coding: utf-8 -*-
"""
EDT — Générateur d'emplois du temps scolaires (interface web)
Lancement :  streamlit run app.py

4 écrans : 1 Importer · 2 Générer · 3 Ajuster · 4 Exporter
Tout en français, léger (connexion lente), utilisable au téléphone.
"""
import io
import os
import time
import hashlib
import tempfile
import threading
from datetime import datetime
from contextlib import redirect_stdout

import pandas as pd
import streamlit as st

import moteur_edt as moteur
from moteur_edt import (
    JOURS, N_JOURS, N_SLOTS, SLOT_LABELS, IDX_DEBUT_APMIDI, SLOTS_APMIDI,
)
import verification as verif
import export_pdf
import persistance as persist

# Composant de stockage navigateur (survit aux coupures réseau / fermeture
# d'onglet). Import protégé : si le paquet manque, l'app marche quand même
# (sans sauvegarde auto) au lieu de planter.
try:
    from streamlit_local_storage import LocalStorage
    _STORAGE_DISPO = True
except Exception:
    _STORAGE_DISPO = False

DOSSIER = os.path.dirname(os.path.abspath(__file__))
FICHIER_EXEMPLE = os.path.join(DOSSIER, "donnees_exemple.xlsx")

# ════════════════════════ PAGE & STYLE ════════════════════════
st.set_page_config(
    page_title="EDT — Emplois du temps",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 4rem; max-width: 1150px;}
div.stButton > button, div.stDownloadButton > button {
    width: 100%; padding: 0.6rem 0.8rem; font-weight: 600;
}
/* ── Uploader en français ──
   Streamlit affiche « Drag and drop file here » / « Browse files » en dur.
   On masque ces textes et on les remplace en CSS. Si Streamlit change son
   DOM un jour, l'anglais réapparaît simplement (aucune casse possible). */
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small {display: none;}
[data-testid="stFileUploaderDropzoneInstructions"] > div::before {
    content: "Glissez votre fichier Excel ici"; font-weight: 600;
    display: block;
}
[data-testid="stFileUploaderDropzoneInstructions"] > div::after {
    content: "Fichier .xlsx uniquement"; font-size: 0.8rem; opacity: 0.7;
    display: block; margin-top: 2px;
}
[data-testid="stFileUploader"] section button {font-size: 0 !important;}
[data-testid="stFileUploader"] section button::after {
    content: "Parcourir…"; font-size: 0.9rem; font-weight: 600;
}
/* Séparateur « ou » de l'écran d'accueil */
.edt-ou {text-align: center; color: #999; font-size: 0.85rem;
         margin: 1.3rem 0 0.9rem; letter-spacing: 0.02em;}
.edt-wrap {overflow-x: auto; border: 1px solid #ccc; border-radius: 6px;}
table.edt {border-collapse: collapse; width: 100%; min-width: 620px;
           font-family: Arial, sans-serif;}
table.edt th, table.edt td {border: 1px solid #bbb; text-align: center;
                            padding: 4px 3px; font-size: 12.5px;}
table.edt th {background: #444; color: #fff; font-size: 13px; padding: 7px 3px;}
table.edt td.heure {background: #f2f2f2; font-weight: 700; font-size: 11px;
                    white-space: nowrap; width: 92px;}
table.edt td.vide {background: #fafafa;}
table.edt tr.sep td {background: #ddd; font-weight: 700; font-size: 10px;
                     padding: 2px; border-top: 2px solid #444;}
table.edt .mat {font-weight: 700;}
table.edt .qui {font-size: 11px; color: #333;}
table.edt td.src {outline: 3px solid #1f77d0; outline-offset: -3px;}
table.edt td.dst {outline: 3px dashed #e07b00; outline-offset: -3px;}
@media (max-width: 640px) {
    table.edt th, table.edt td {font-size: 10.5px; padding: 3px 2px;}
    table.edt .qui {font-size: 9.5px;}
    table.edt td.heure {font-size: 9px; width: 70px;}
}
</style>
""", unsafe_allow_html=True)

# ════════════════════════ ÉTAT ════════════════════════
defauts = {
    "donnees": None,        # (classes, permanents, services, indispos)
    "regles": {},           # règles particulières lues dans l'Excel
    "meta": {},             # etablissement, annee, heures_max…
    "fichier_nom": None,
    "erreurs": [], "avertissements": [],
    "emplois": None,        # {(classe, jour, slot): {prof, matiere}}
    "log_moteur": "",
    "duree_generation": None,
    "historique": [],       # pile pour Annuler
    "nb_modifs": 0,
    "chaine": [],           # liste de (src, dst) en attente de validation groupée
    "versions": [],         # versions nommées enregistrées par l'utilisateur
}
for k, v in defauts.items():
    st.session_state.setdefault(k, v)


# ════════════════════════ SAUVEGARDE AUTOMATIQUE (navigateur) ════════════════════════
# Objectif : ne JAMAIS perdre le travail si la connexion saute ou si l'onglet
# se ferme. L'état des 4 écrans est stocké dans le localStorage du navigateur,
# puis restauré automatiquement au retour.
def _init_storage():
    if not _STORAGE_DISPO:
        return None
    if "_ls" not in st.session_state:
        # instanceId fixe pour viser TOUJOURS la même clé entre les sessions
        st.session_state._ls = LocalStorage()
    return st.session_state._ls


def restaurer_si_besoin():
    """Au tout premier passage de la session, tente de recharger l'état
    sauvegardé dans le navigateur. Sans effet les fois suivantes."""
    if st.session_state.get("_restauration_faite"):
        return
    ls = _init_storage()
    if ls is None:
        st.session_state._restauration_faite = True
        return
    try:
        texte = ls.getItem(persist.CLE_STORAGE)
    except Exception:
        texte = None
    if texte:
        etat = persist.deserialiser_etat(texte)
        if etat and (etat.get("donnees") or etat.get("emplois")):
            for cle, valeur in etat.items():
                st.session_state[cle] = valeur
            st.session_state._sauvegarde_restauree = True
            # mémoriser l'empreinte pour ne pas réécrire inutilement juste après
            st.session_state._derniere_empreinte = persist.empreinte(texte)
    st.session_state._restauration_faite = True


def sauvegarder_si_change():
    """Réécrit l'état dans le navigateur UNIQUEMENT s'il a changé depuis la
    dernière sauvegarde (empreinte). Appelé en fin de script."""
    ls = _init_storage()
    if ls is None:
        return
    texte = persist.serialiser_etat(st.session_state)
    if texte is None:
        return
    emp = persist.empreinte(texte)
    if emp == st.session_state.get("_derniere_empreinte"):
        return  # rien n'a changé : on n'écrit pas (évite de ralentir)
    try:
        ls.setItem(persist.CLE_STORAGE, texte, key="edt_save")
        st.session_state._derniere_empreinte = emp
    except Exception:
        pass  # un échec d'écriture ne doit jamais casser l'app


def effacer_sauvegarde():
    """Supprime la sauvegarde navigateur (bouton « repartir de zéro »)."""
    ls = _init_storage()
    if ls is not None:
        try:
            ls.deleteItem(persist.CLE_STORAGE)
        except Exception:
            pass
    st.session_state._derniere_empreinte = None
    st.session_state._sauvegarde_restauree = False


# ── Versions nommées (historique) : chargement + écriture navigateur ──
def charger_versions():
    """Lit la liste des versions enregistrées depuis le navigateur.
    Met à jour st.session_state.versions (une seule lecture par session)."""
    if "versions" in st.session_state and st.session_state.get("_versions_chargees"):
        return st.session_state.versions
    ls = _init_storage()
    versions = []
    if ls is not None:
        try:
            texte = ls.getItem(persist.CLE_VERSIONS)
            versions = persist.texte_vers_versions(texte)
        except Exception:
            versions = []
    st.session_state.versions = versions
    st.session_state._versions_chargees = True
    return versions


def ecrire_versions(versions):
    """Écrit la liste des versions dans le navigateur et en session."""
    st.session_state.versions = versions
    ls = _init_storage()
    if ls is not None:
        try:
            ls.setItem(persist.CLE_VERSIONS, persist.versions_vers_texte(versions),
                       key="edt_versions_save")
        except Exception:
            pass


restaurer_si_besoin()


# ════════════════════════ LECTURES COMPLÉMENTAIRES ════════════════════════
def lire_meta(chemin):
    """Nom d'établissement, année, heures max par prof, profs déclarés,
    valeurs douteuses des colonnes optionnelles (jours / blocs)."""
    meta = {"etablissement": "", "annee": "", "heures_max": {},
            "profs_declares": set(), "jours_bruts": [], "blocs_bruts": []}
    xls = pd.ExcelFile(chemin)

    if "Paramètres" in xls.sheet_names:
        df = pd.read_excel(xls, "Paramètres", header=None)
        for i in range(len(df)):
            for j in range(len(df.columns) - 1):
                v = df.iat[i, j]
                if pd.isna(v):
                    continue
                texte = str(v)
                droite = df.iat[i, j + 1]
                droite = "" if pd.isna(droite) else str(droite).strip()
                if "Nom de l'établissement" in texte:
                    meta["etablissement"] = droite
                elif "Année scolaire" in texte:
                    meta["annee"] = droite

    if "Professeurs" in xls.sheet_names:
        df = pd.read_excel(xls, "Professeurs", skiprows=2, header=0)
        df = df.dropna(subset=[df.columns[0]])
        for _, r in df.iterrows():
            nom = str(r.iloc[0]).strip()
            if not nom:
                continue
            meta["profs_declares"].add(nom)
            try:
                meta["heures_max"][nom] = int(float(r.iloc[4]))
            except (ValueError, TypeError, IndexError):
                pass

    if "Services" in xls.sheet_names:
        df = pd.read_excel(xls, "Services", skiprows=2, header=0)
        valides_bloc = {"", "Indifférent", "2h", "3h", "Oui", "Non"}
        for i, r in df.iterrows():
            if pd.isna(r.iloc[0]):
                continue
            ligne_excel = i + 4
            bloc = "" if len(r) < 5 or pd.isna(r.iloc[4]) else str(r.iloc[4]).strip()
            jour = "" if len(r) < 6 or pd.isna(r.iloc[5]) else str(r.iloc[5]).strip()
            if bloc not in valides_bloc:
                meta["blocs_bruts"].append((ligne_excel, bloc))
            if jour and jour not in JOURS:
                meta["jours_bruts"].append((ligne_excel, jour))
    return meta


def charger_fichier(chemin, nom_affiche):
    """Lit + valide le fichier ; remplit st.session_state. Retourne True si OK."""
    try:
        classes, permanents, services, indispos, regles = moteur.lire_excel(chemin)
        meta = lire_meta(chemin)
    except Exception as e:
        st.session_state.erreurs = [
            f"Impossible de lire le fichier : {e}. "
            "Vérifiez qu'il s'agit bien du template Excel rempli "
            "(onglets Classes, Professeurs, Services, Disponibilités)."
        ]
        st.session_state.avertissements = []
        st.session_state.donnees = None
        return False

    erreurs, avert = verif.valider_donnees(
        classes, permanents, services, indispos,
        profs_declares=meta["profs_declares"] or None,
        heures_max=meta["heures_max"],
        jours_imposes_bruts=meta["jours_bruts"],
    )
    for ligne, valeur in meta["blocs_bruts"]:
        avert.append(
            f"Onglet Services, ligne {ligne} : taille de bloc « {valeur} » "
            "non reconnue (valeurs possibles : 2h, 3h ou vide) — elle sera ignorée."
        )

    st.session_state.erreurs = erreurs
    st.session_state.avertissements = avert
    st.session_state.meta = meta
    st.session_state.regles = regles
    st.session_state.fichier_nom = nom_affiche
    if erreurs:
        st.session_state.donnees = None
        return False

    st.session_state.donnees = (classes, permanents, services, indispos)
    # Nouveau fichier → on repart de zéro
    st.session_state.emplois = None
    st.session_state.log_moteur = ""
    st.session_state.historique = []
    st.session_state.nb_modifs = 0
    return True


# ════════════════════════ GRILLE HTML ════════════════════════
def grille_html(get_contenu, src=None, dst=None):
    """Tableau HTML d'une grille hebdomadaire.
    get_contenu(d, t) -> (ligne1, ligne2) ou None.
    src / dst : (jour, slot) à surligner."""
    h = ['<div class="edt-wrap"><table class="edt">']
    h.append("<tr><th></th>" + "".join(f"<th>{j}</th>" for j in JOURS) + "</tr>")
    for t in range(N_SLOTS):
        if t == IDX_DEBUT_APMIDI:
            h.append(f'<tr class="sep"><td colspan="{N_JOURS + 1}">APRÈS-MIDI</td></tr>')
        h.append(f'<tr><td class="heure">{SLOT_LABELS[t]}</td>')
        for d in range(N_JOURS):
            contenu = get_contenu(d, t)
            classes_css = []
            if src == (d, t):
                classes_css.append("src")
            if dst == (d, t):
                classes_css.append("dst")
            if contenu:
                l1, l2 = contenu
                h.append(f'<td class="{" ".join(classes_css)}">'
                         f'<div class="mat">{l1}</div>'
                         f'<div class="qui">{l2}</div></td>')
            else:
                h.append(f'<td class="vide {" ".join(classes_css)}"></td>')
        h.append("</tr>")
    h.append("</table></div>")
    return "".join(h)


def contenu_classe(emplois, classe):
    def f(d, t):
        info = emplois.get((classe, d, t))
        if not info:
            return None
        return info["matiere"], info["prof"].replace("M. ", "")
    return f


def contenu_prof(emplois, prof):
    def f(d, t):
        for (cl, dd, tt), info in emplois.items():
            if dd == d and tt == t and info["prof"] == prof:
                return info["matiere"], cl
        return None
    return f


# ════════════════════════ EXPORTS (avec cache) ════════════════════════
def _cle_emplois(emplois):
    return tuple(sorted((cl, d, t, i["prof"], i["matiere"])
                        for (cl, d, t), i in emplois.items()))


@st.cache_data(show_spinner=False)
def fabriquer_excel(cle, classes):
    emplois = {(cl, d, t): {"prof": p, "matiere": m} for (cl, d, t, p, m) in cle}
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        chemin = tmp.name
    with redirect_stdout(io.StringIO()):
        moteur.ecrire_excel(emplois, list(classes), chemin)
    with open(chemin, "rb") as f:
        data = f.read()
    os.unlink(chemin)
    return data


@st.cache_data(show_spinner=False)
def fabriquer_pdf_classes(cle, classes, etablissement, annee):
    emplois = {(cl, d, t): {"prof": p, "matiere": m} for (cl, d, t, p, m) in cle}
    return export_pdf.pdf_classes(emplois, list(classes), etablissement, annee)


@st.cache_data(show_spinner=False)
def fabriquer_pdf_profs(cle, etablissement, annee):
    emplois = {(cl, d, t): {"prof": p, "matiere": m} for (cl, d, t, p, m) in cle}
    return export_pdf.pdf_profs(emplois, etablissement, annee)


# ════════════════════════ EN-TÊTE ════════════════════════
st.title("📅 Générateur d'emplois du temps")
meta = st.session_state.meta
if meta.get("etablissement"):
    st.caption(f"**{meta['etablissement']}**"
               + (f" — Année scolaire {meta['annee']}" if meta.get("annee") else ""))
else:
    st.caption("Tous les emplois du temps de votre établissement, "
               "générés sans conflit et prêts à imprimer.")

FICHIER_TEMPLATE = os.path.join(DOSSIER, "template_vierge.xlsx")

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["**1 · Importer**", "**2 · Générer**", "**3 · Ajuster**",
     "**4 · Exporter**", "**❓ Aide**"]
)

# ════════════════════════ ÉCRAN 1 — IMPORTER ════════════════════════
with tab1:
    st.subheader("Importer le fichier de données")

    # ── Bandeau « travail restauré » (sauvegarde navigateur) ──
    if st.session_state.get("_sauvegarde_restauree"):
        c_info, c_btn = st.columns([4, 1])
        with c_info:
            quoi = []
            if st.session_state.get("emplois"):
                quoi.append("emploi du temps")
            if st.session_state.get("nb_modifs"):
                quoi.append(f"{st.session_state['nb_modifs']} ajustement(s)")
            detail = " et ".join(quoi) if quoi else "votre travail"
            st.success(f"✅ Travail restauré automatiquement ({detail}). "
                       "Vous pouvez reprendre où vous en étiez.")
        with c_btn:
            if st.button("🗑 Effacer", help="Supprime la sauvegarde du navigateur "
                         "et repart de zéro."):
                effacer_sauvegarde()
                for k in list(defauts.keys()):
                    st.session_state[k] = defauts[k]
                st.session_state._sauvegarde_restauree = False
                st.session_state.pop("_fichier_traite", None)
                st.rerun()
    elif _STORAGE_DISPO and (st.session_state.get("donnees")
                             or st.session_state.get("emplois")):
        st.caption("💾 Sauvegarde automatique active — votre travail est "
                   "conservé dans ce navigateur même en cas de coupure.")

    # ── Accueil : le chemin le plus simple d'abord ──
    st.markdown(
        "Créez l'emploi du temps **complet** de votre établissement en quelques "
        "minutes : zéro conflit, volumes horaires exacts, et professeurs "
        "vacataires regroupés sur un minimum de jours."
    )
    if os.path.exists(FICHIER_EXEMPLE):
        bouton_exemple = st.button(
            "🚀 Essayer maintenant avec une école d'exemple",
            type="primary", key="btn_demo",
        )
        st.caption(
            "Lycée réaliste : 10 classes, 20 professeurs, 262 h de cours — "
            "rien à préparer, le résultat arrive en moins d'une minute."
        )
    else:
        bouton_exemple = False

    st.markdown('<div class="edt-ou">— ou avec les données de votre '
                'établissement —</div>', unsafe_allow_html=True)

    st.markdown(
        "**①** Téléchargez le modèle Excel&ensp;→&ensp;**②** Remplissez-le avec "
        "vos classes, professeurs et matières&ensp;→&ensp;**③** Déposez-le "
        "ci-dessous : il est vérifié automatiquement."
    )

    # ── Boutons téléchargement template vierge + exemple rempli ──
    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        if os.path.exists(FICHIER_TEMPLATE):
            with open(FICHIER_TEMPLATE, "rb") as _f:
                _data_template = _f.read()
            st.download_button(
                label="📥 Modèle vierge à remplir",
                data=_data_template,
                file_name="template_emplois_du_temps.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Téléchargez ce fichier, remplissez-le avec les données de votre "
                     "école, puis déposez-le ci-dessous.",
            )
    with dl_col2:
        if os.path.exists(FICHIER_EXEMPLE):
            with open(FICHIER_EXEMPLE, "rb") as _f:
                _data_exemple = _f.read()
            st.download_button(
                label="📄 Exemple rempli (pour voir comment faire)",
                data=_data_exemple,
                file_name="exemple_template_rempli.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Un exemple complet et rempli, à garder ouvert comme référence "
                     "pendant que vous remplissez le vôtre.",
            )

    fichier = st.file_uploader(
        "Fichier Excel (.xlsx)", type=["xlsx"],
        help="Le fichier de collecte rempli par l'établissement.",
        label_visibility="collapsed",
    )

    # ── Validation AUTOMATIQUE dès le dépôt (plus aucun clic à faire) ──
    # L'empreinte du contenu évite de re-valider à chaque interaction tant
    # que le fichier déposé n'a pas changé.
    if fichier is not None:
        empreinte_fichier = hashlib.md5(fichier.getvalue()).hexdigest()
        if st.session_state.get("_fichier_traite") != empreinte_fichier:
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(fichier.getvalue())
                chemin_tmp = tmp.name
            with st.spinner("Lecture et vérification des données…"):
                charger_fichier(chemin_tmp, fichier.name)
            os.unlink(chemin_tmp)
            st.session_state._fichier_traite = empreinte_fichier

    if bouton_exemple:
        with st.spinner("Chargement de l'exemple…"):
            charger_fichier(FICHIER_EXEMPLE, "Exemple — Lycée Moderne de Cocody")
        # La démo devient la source active : si l'utilisateur redépose ensuite
        # son propre fichier (même inchangé), il sera bien re-validé.
        st.session_state._fichier_traite = "exemple"

    # ── Résultat de la validation ──
    if st.session_state.erreurs:
        st.error("**Le fichier ne peut pas être utilisé en l'état :**")
        for e in st.session_state.erreurs:
            st.markdown(f"- ❌ {e}")
        st.info("Corrigez le fichier Excel puis importez-le à nouveau.")

    if st.session_state.donnees:
        classes, permanents, services, indispos = st.session_state.donnees
        total_h = sum(s["heures"] for s in services)
        n_b2 = sum(1 for s in services if s["bloc_size"] == 2)
        n_b3 = sum(1 for s in services if s["bloc_size"] == 3)
        n_ji = sum(1 for s in services if s["jour_impose"] is not None)
        profs = sorted({s["prof"] for s in services})

        st.success(f"**Fichier validé : {st.session_state.fichier_nom}**")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Classes", len(classes))
        c2.metric("Professeurs", len(profs))
        c3.metric("Services", len(services))
        c4.metric("Heures / semaine", total_h)
        st.caption(
            f"Blocs 2h : {n_b2} · Blocs 3h : {n_b3} · Jours imposés : {n_ji} · "
            f"Permanents : {len(permanents & set(profs))} · "
            f"Vacataires : {len(set(profs) - permanents)}"
        )

        # Règles particulières détectées dans le fichier
        if st.session_state.get("regles", {}).get("eps_heures_chaudes"):
            n_cr = len(st.session_state.get("regles", {})["eps_heures_chaudes"])
            st.caption(f"🌞 Règle active : **pas d'EPS aux heures chaudes** "
                       f"({n_cr} créneau(x) bloqué(s) pour l'EPS). "
                       "Désactivable dans l'onglet Paramètres du fichier Excel.")

        if st.session_state.avertissements:
            with st.expander(f"⚠️ {len(st.session_state.avertissements)} point(s) "
                             "d'attention (non bloquants)"):
                for a in st.session_state.avertissements:
                    st.markdown(f"- {a}")

        with st.expander("Détail des charges par classe et par professeur"):
            h_cl = {c: sum(s["heures"] for s in services if s["classe"] == c)
                    for c in classes}
            st.dataframe(
                pd.DataFrame({"Classe": list(h_cl), "Heures / semaine": list(h_cl.values())}),
                hide_index=True, width="stretch",
            )
            h_pr = {p: sum(s["heures"] for s in services if s["prof"] == p)
                    for p in profs}
            st.dataframe(
                pd.DataFrame({
                    "Professeur": list(h_pr),
                    "Heures / semaine": list(h_pr.values()),
                    "Statut": ["Permanent" if p in permanents else "Vacataire"
                               for p in h_pr],
                }),
                hide_index=True, width="stretch",
            )

        st.info("➡️ Passez à l'onglet **2 · Générer**.")

# ════════════════════════ ÉCRAN 2 — GÉNÉRER ════════════════════════
with tab2:
    st.subheader("Générer les emplois du temps")

    if not st.session_state.donnees:
        st.warning("Commencez par importer et valider un fichier (onglet **1 · Importer**).")
    else:
        classes, permanents, services, indispos = st.session_state.donnees
        st.write(
            f"Prêt à placer **{sum(s['heures'] for s in services)} heures** "
            f"de cours pour **{len(classes)} classes**. "
            "Le moteur garantit : zéro conflit, volumes exacts, disponibilités "
            "respectées, mercredi après-midi libéré pour les permanents, et "
            "**journées compactes** (pas de cours isolé à 16h après une matinée vide)."
        )
        temps_max = st.slider(
            "Temps de calcul maximum (secondes)",
            min_value=30, max_value=300, value=90, step=30,
            help="Le moteur s'arrête dès qu'il trouve la meilleure solution. "
                 "Augmentez si votre établissement est très contraint.",
        )
        matin_prefere = st.checkbox(
            "Privilégier les cours le matin (libérer les après-midis)",
            value=False,
            help="À cocher pour les établissements qui font cours surtout le "
                 "matin. Le moteur remplit alors les matinées en priorité. "
                 "Sans cette option, il regroupe les cours sans préférence "
                 "matin/après-midi, mais toujours sans trou.",
        )

        deja = st.session_state.emplois is not None
        libelle = "🔄 Régénérer (efface les ajustements manuels)" if deja \
            else "🚀 Générer les emplois du temps"
        if st.button(libelle, type="primary"):
            resultat = {}

            def travail():
                buf = io.StringIO()
                try:
                    with redirect_stdout(buf):
                        emplois = moteur.resoudre(
                            classes, permanents, services, indispos,
                            temps_max=temps_max,
                            matin_prefere=matin_prefere,
                            regles=st.session_state.get("regles", {}),
                        )
                    resultat["emplois"] = emplois
                except Exception as e:           # garde-fou
                    resultat["exception"] = str(e)
                resultat["log"] = buf.getvalue()

            th = threading.Thread(target=travail, daemon=True)
            debut = time.time()
            th.start()
            barre = st.progress(0.0, text="Démarrage du moteur…")
            while th.is_alive():
                ecoule = time.time() - debut
                frac = min(ecoule / max(temps_max, 1), 0.97)
                barre.progress(frac, text=f"Résolution en cours… {int(ecoule)} s "
                                          f"(maximum {temps_max} s)")
                time.sleep(0.4)
            th.join()
            duree = time.time() - debut
            barre.progress(1.0, text=f"Calcul terminé en {duree:.0f} s")

            if resultat.get("exception"):
                st.error(f"Erreur pendant la génération : {resultat['exception']}")
            elif resultat.get("emplois") is None:
                # Diagnostic : est-ce la règle EPS qui rend tout impossible ?
                coupable_eps = False
                if st.session_state.get("regles", {}).get("eps_heures_chaudes"):
                    with st.spinner("Analyse de la cause du blocage…"):
                        try:
                            with redirect_stdout(io.StringIO()):
                                test = moteur.resoudre(
                                    classes, permanents, services, indispos,
                                    temps_max=min(temps_max, 45),
                                    matin_prefere=matin_prefere,
                                    regles={},   # sans la règle EPS
                                )
                            coupable_eps = test is not None
                        except Exception:
                            coupable_eps = False

                if coupable_eps:
                    st.error(
                        "**Impossible avec la règle « Pas d'EPS aux heures "
                        "chaudes ».** Cette règle, telle que réglée, ne laisse "
                        "pas assez de créneaux pour caser toutes les heures "
                        "d'EPS — souvent parce qu'un seul professeur d'EPS doit "
                        "couvrir beaucoup de classes."
                    )
                    st.info(
                        "**Que faire ?**\n"
                        "- Dans le fichier Excel, onglet **Paramètres**, mettez "
                        "la règle EPS sur **Non** (ou élargissez la plage, par "
                        "exemple 12h00–15h00 au lieu de 11h00–16h00), puis "
                        "réimportez le fichier.\n"
                        "- Ou ajoutez un second professeur d'EPS si possible.\n\n"
                        "Sans cette règle, l'emploi du temps se génère "
                        "normalement."
                    )
                else:
                    st.error(
                        "**Aucune solution trouvée.** Les contraintes sont "
                        "incompatibles entre elles. Pistes : assouplir les jours "
                        "imposés, vérifier les disponibilités des vacataires, "
                        "réduire les volumes des professeurs les plus chargés, "
                        "ou augmenter le temps de calcul."
                    )
                with st.expander("Journal du moteur"):
                    st.code(resultat.get("log", ""), language=None)
            else:
                st.session_state.emplois = resultat["emplois"]
                st.session_state.duree_generation = duree
                st.session_state.historique = []
                st.session_state.nb_modifs = 0
                buf = io.StringIO()
                with redirect_stdout(buf):
                    moteur.verifier(resultat["emplois"], classes,
                                    permanents, services, indispos)
                st.session_state.log_moteur = (resultat.get("log", "")
                                               + "\n--- Vérification ---\n"
                                               + buf.getvalue())

        # ── Résultat ──
        if st.session_state.emplois is not None:
            st.success(
                f"**Emplois du temps générés** "
                f"({len(st.session_state.emplois)} heures placées"
                + (f" en {st.session_state.duree_generation:.0f} s"
                   if st.session_state.duree_generation else "") + ")."
            )
            bilan = verif.bilan_etat(st.session_state.emplois, classes,
                                     permanents, services, indispos)
            for statut, msg in bilan:
                icone = {"ok": "✅", "erreur": "❌", "info": "ℹ️"}[statut]
                st.markdown(f"{icone} {msg}")
            with st.expander("Journal complet du moteur"):
                st.code(st.session_state.log_moteur, language=None)
            st.info("➡️ Visualisez et ajustez dans l'onglet **3 · Ajuster**, "
                    "ou téléchargez directement dans **4 · Exporter**.")

# ════════════════════════ ÉCRAN 3 — AJUSTER ════════════════════════
with tab3:
    st.subheader("Visualiser et ajuster")

    if st.session_state.emplois is None:
        st.warning("Générez d'abord les emplois du temps (onglet **2 · Générer**).")
    else:
        classes, permanents, services, indispos = st.session_state.donnees
        emplois = st.session_state.emplois
        profs = sorted({i["prof"] for i in emplois.values()})

        # ════════════ VERSIONS ENREGISTRÉES (historique nommé) ════════════
        versions = charger_versions()
        with st.expander(f"🗂️ Versions enregistrées ({len(versions)})",
                         expanded=False):
            st.caption(
                "Enregistrez l'emploi du temps actuel sous un nom pour le "
                "retrouver plus tard. Pratique pour comparer plusieurs "
                "organisations sans rien perdre."
            )

            # ── Enregistrer la version courante ──
            c_nom, c_btn = st.columns([3, 1])
            with c_nom:
                nom_version = st.text_input(
                    "Nom de la version",
                    placeholder="ex : Version sans trous le mardi",
                    label_visibility="collapsed",
                    key="saisie_nom_version",
                )
            with c_btn:
                enregistrer = st.button("💾 Enregistrer", use_container_width=True)

            if enregistrer:
                nom = (nom_version or "").strip()
                if not nom:
                    nom = "Version du " + datetime.now().strftime("%d/%m à %Hh%M")
                horod = datetime.now().strftime("%d/%m/%Y %H:%M")
                nouvelle = persist.creer_version(emplois, nom, horod)
                ecrire_versions(persist.ajouter_version(versions, nouvelle))
                st.session_state.pop("saisie_nom_version", None)
                st.success(f"Version « {nom} » enregistrée.")
                time.sleep(0.8)
                st.rerun()

            # ── Liste des versions ──
            if not versions:
                st.info("Aucune version enregistrée pour l'instant.")
            else:
                st.markdown("**Vos versions** (de la plus récente à la plus ancienne) :")
                for v in versions:
                    c_info, c_rest, c_suppr = st.columns([3, 1, 1])
                    with c_info:
                        st.markdown(
                            f"**{v['nom']}**  \n"
                            f"<span style='color:#888;font-size:0.85em'>"
                            f"{v['horodatage']} · {v['nb_cours']} cours</span>",
                            unsafe_allow_html=True,
                        )
                    with c_rest:
                        if st.button("📂 Restaurer", key=f"rest_{v['id']}",
                                     use_container_width=True):
                            # Filet de sécurité : on enregistre d'abord le travail
                            # EN COURS comme version auto, pour ne rien perdre.
                            deja = any(
                                persist.version_vers_emplois(x) == emplois
                                for x in versions
                            )
                            if not deja:
                                auto = persist.creer_version(
                                    emplois,
                                    "Travail en cours (avant restauration)",
                                    datetime.now().strftime("%d/%m/%Y %H:%M"),
                                )
                                liste_maj = persist.ajouter_version(versions, auto)
                            else:
                                liste_maj = versions
                            # Restaurer la version choisie
                            st.session_state.historique.append(dict(emplois))
                            st.session_state.emplois = persist.version_vers_emplois(v)
                            st.session_state.chaine = []
                            ecrire_versions(liste_maj)
                            st.success(f"Version « {v['nom']} » restaurée.")
                            time.sleep(0.8)
                            st.rerun()
                    with c_suppr:
                        if st.button("🗑", key=f"suppr_{v['id']}",
                                     use_container_width=True,
                                     help="Supprimer cette version"):
                            ecrire_versions(persist.retirer_version(versions, v['id']))
                            st.rerun()

        vue = st.radio("Vue", ["Par classe", "Par professeur"],
                       horizontal=True, label_visibility="collapsed")

        if vue == "Par professeur":
            prof = st.selectbox("Professeur", profs)
            total = sum(1 for i in emplois.values() if i["prof"] == prof)
            st.caption(f"{total} heure(s) de cours par semaine. "
                       "Pour déplacer un cours, passez par la vue **Par classe**.")
            st.markdown(grille_html(contenu_prof(emplois, prof)),
                        unsafe_allow_html=True)
        else:
            classe = st.selectbox("Classe", classes)

            def lib_cours(c):
                d, t, info = c
                return (f"{JOURS[d]} {SLOT_LABELS[t]} — {info['matiere']} "
                        f"({info['prof'].replace('M. ', '')})")

            with st.container(border=True):
                st.markdown("**Déplacer plusieurs cours d'un coup** — "
                            "ajoutez chaque déplacement à la liste ci-dessous "
                            "(le créneau d'arrivée peut être occupé : les cours "
                            "seront alors échangés). Quand la liste est prête, "
                            "appliquez tout en une seule fois — rien n'est "
                            "modifié si un seul maillon pose problème.")

                # ── Aperçu : grille « après chaîne actuelle » pour choisir
                # le prochain mouvement en connaissance de cause.
                etapes_en_cours = st.session_state.chaine
                emplois_apercu = emplois
                if etapes_en_cours:
                    _c, _a, _resolues = verif.verifier_chaine(
                        emplois, classe, etapes_en_cours, permanents, indispos,
                    )
                    if not _c:
                        emplois_apercu = verif.appliquer_chaine(
                            emplois, classe, _resolues,
                        )

                cours_apercu = sorted(
                    [(d, t, info) for (cl, d, t), info in emplois_apercu.items()
                     if cl == classe],
                    key=lambda x: (x[0], x[1]),
                )

                sel_src = st.selectbox(
                    "Cours à déplacer", cours_apercu, format_func=lib_cours,
                    index=None, placeholder="Choisir un cours…",
                    key=f"src_{classe}",
                )
                destinations = [(d, t) for d in range(N_JOURS)
                                for t in range(N_SLOTS)]

                def lib_dest_apercu(dt):
                    d, t = dt
                    occ = emplois_apercu.get((classe, d, t))
                    etat = f"occupé : {occ['matiere']}" if occ else "libre"
                    return f"{JOURS[d]} {SLOT_LABELS[t]} — {etat}"

                sel_dst = st.selectbox(
                    "Nouveau créneau", destinations, format_func=lib_dest_apercu,
                    index=None, placeholder="Choisir le créneau d'arrivée…",
                    key=f"dst_{classe}",
                )

                c_add, c_app, c_vid, c_undo = st.columns(4)
                with c_add:
                    src_eq_dst = sel_src is not None and (sel_src[0], sel_src[1]) == sel_dst \
                        if (sel_src and sel_dst) else False
                    ajouter = st.button(
                        "➕ Ajouter à la chaîne",
                        disabled=(sel_src is None or sel_dst is None or src_eq_dst),
                    )
                with c_app:
                    appliquer = st.button(
                        f"✅ Appliquer la chaîne ({len(etapes_en_cours)})",
                        type="primary", disabled=not etapes_en_cours,
                    )
                with c_vid:
                    vider = st.button(
                        "🗑 Vider la chaîne", disabled=not etapes_en_cours,
                    )
                with c_undo:
                    annuler = st.button(
                        f"↩️ Annuler la dernière modification "
                        f"({len(st.session_state.historique)})",
                        disabled=not st.session_state.historique,
                    )

                if ajouter and sel_src and sel_dst:
                    src = (sel_src[0], sel_src[1])
                    st.session_state.chaine.append((src, sel_dst))
                    for k in (f"src_{classe}", f"dst_{classe}"):
                        st.session_state.pop(k, None)
                    st.rerun()

                if vider:
                    st.session_state.chaine = []
                    st.rerun()

                # ── Liste des étapes en attente, avec retrait individuel ──
                if etapes_en_cours:
                    st.markdown("**Chaîne en attente :**")
                    for i, (src, dst) in enumerate(etapes_en_cours):
                        col_txt, col_del = st.columns([5, 1])
                        with col_txt:
                            st.markdown(
                                f"{i + 1}. {JOURS[src[0]]} {SLOT_LABELS[src[1]]} "
                                f"→ {JOURS[dst[0]]} {SLOT_LABELS[dst[1]]}"
                            )
                        with col_del:
                            if st.button("✖", key=f"del_etape_{classe}_{i}"):
                                st.session_state.chaine.pop(i)
                                st.rerun()

                    # Vérification en direct de la chaîne actuelle
                    conflits_apercu, _a, _r = verif.verifier_chaine(
                        emplois, classe, etapes_en_cours, permanents, indispos,
                    )
                    if conflits_apercu:
                        st.error("**Cette chaîne ne peut pas être appliquée "
                                 "telle quelle :**")
                        for c in conflits_apercu:
                            st.markdown(f"- ❌ {c}")
                        st.caption("Retirez ou réordonnez l'étape en cause "
                                   "ci-dessus avant d'appliquer.")
                    else:
                        st.caption("✅ Chaîne valide — prête à être appliquée.")

                if appliquer and etapes_en_cours:
                    conflits, averts, resolues = verif.verifier_chaine(
                        emplois, classe, etapes_en_cours, permanents, indispos,
                    )
                    if conflits:
                        st.error("**Application impossible :**")
                        for c in conflits:
                            st.markdown(f"- ❌ {c}")
                    else:
                        st.session_state.historique.append(dict(emplois))
                        if len(st.session_state.historique) > 25:
                            st.session_state.historique.pop(0)
                        st.session_state.emplois = verif.appliquer_chaine(
                            emplois, classe, resolues,
                        )
                        st.session_state.nb_modifs += 1
                        st.session_state.chaine = []
                        for a in averts:
                            st.warning(a)
                        n_ech = sum(1 for _, _, e in resolues if e)
                        msg = f"{len(resolues)} déplacement(s) appliqué(s)"
                        if n_ech:
                            msg += f" (dont {n_ech} échange(s))"
                        st.success(msg + ".")
                        for k in (f"src_{classe}", f"dst_{classe}"):
                            st.session_state.pop(k, None)
                        time.sleep(0.9)
                        st.rerun()

                if annuler and st.session_state.historique:
                    st.session_state.emplois = st.session_state.historique.pop()
                    st.session_state.nb_modifs = max(
                        0, st.session_state.nb_modifs - 1)
                    st.session_state.chaine = []
                    for k in (f"src_{classe}", f"dst_{classe}"):
                        st.session_state.pop(k, None)
                    st.rerun()

            # Grille avec surlignage de la sélection ; affiche l'état RÉEL +
            # la chaîne en attente, pour visualiser le résultat avant validation.
            src_hl = (sel_src[0], sel_src[1]) if sel_src else None
            st.markdown(
                grille_html(contenu_classe(emplois_apercu, classe),
                            src=src_hl, dst=sel_dst),
                unsafe_allow_html=True,
            )
            if st.session_state.chaine:
                st.caption("🟦 cours sélectionné · 🟧 créneau d'arrivée · "
                           "grille = état réel + chaîne en attente "
                           "(non encore appliquée)")
            else:
                st.caption("🟦 cours sélectionné · 🟧 créneau d'arrivée")

        # ── État des vérifications après modifications ──
        if st.session_state.nb_modifs:
            st.caption(f"{st.session_state.nb_modifs} modification(s) manuelle(s) "
                       "depuis la génération.")
        with st.expander("État des vérifications"):
            for statut, msg in verif.bilan_etat(
                    st.session_state.emplois, classes,
                    permanents, services, indispos):
                icone = {"ok": "✅", "erreur": "❌", "info": "ℹ️"}[statut]
                st.markdown(f"{icone} {msg}")

# ════════════════════════ ÉCRAN 4 — EXPORTER ════════════════════════
with tab4:
    st.subheader("Exporter et imprimer")

    if st.session_state.emplois is None:
        st.warning("Générez d'abord les emplois du temps (onglet **2 · Générer**).")
    else:
        classes, permanents, services, indispos = st.session_state.donnees
        etab = st.session_state.meta.get("etablissement", "")
        annee = st.session_state.meta.get("annee", "")
        cle = _cle_emplois(st.session_state.emplois)

        if st.session_state.nb_modifs:
            st.caption(f"Les exports incluent vos "
                       f"{st.session_state.nb_modifs} modification(s) manuelle(s).")

        with st.spinner("Préparation des fichiers…"):
            data_xlsx = fabriquer_excel(cle, tuple(classes))
            data_pdf_cl = fabriquer_pdf_classes(cle, tuple(classes), etab, annee)
            data_pdf_pr = fabriquer_pdf_profs(cle, etab, annee)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button(
                "📄 PDF — toutes les classes", data_pdf_cl,
                file_name="emplois_du_temps_classes.pdf",
                mime="application/pdf",
                help="Une page A4 par classe, prête à imprimer et afficher.",
            )
            st.caption(f"{len(classes)} pages — affichage dans les salles")
        with c2:
            st.download_button(
                "📄 PDF — tous les professeurs", data_pdf_pr,
                file_name="emplois_du_temps_professeurs.pdf",
                mime="application/pdf",
                help="Une page A4 par professeur, à distribuer en salle des profs.",
            )
            n_profs = len({i["prof"] for i in st.session_state.emplois.values()})
            st.caption(f"{n_profs} pages — à remettre aux enseignants")
        with c3:
            st.download_button(
                "📊 Excel complet (.xlsx)", data_xlsx,
                file_name="emplois_du_temps.xlsx",
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet",
                help="Un onglet par classe + la vue par professeur.",
            )
            st.caption(f"{len(classes) + 1} onglets — archive et retouches")

        st.divider()
        st.markdown(
            "💡 **Conseil impression** : les PDF sont en **noir et blanc**, "
            "au format **A4 paysage**, sans couleur — adaptés à toutes les "
            "imprimantes et photocopieuses."
        )

# ════════════════════════ ÉCRAN 5 — AIDE ════════════════════════
with tab5:
    st.subheader("Guide d'utilisation")
    st.caption("Tout ce qu'il faut savoir pour générer un emploi du temps "
               "avec cet outil, étape par étape.")

    # ── C'est quoi cet outil ? ──
    with st.expander("📌 C'est quoi cet outil ?", expanded=True):
        st.markdown("""
Cet outil génère automatiquement les emplois du temps d'une école à partir
d'un fichier Excel que vous remplissez une seule fois.

**Ce qu'il fait :**
- Place chaque cours dans la semaine sans aucun conflit (un prof ne peut pas être dans deux classes en même temps)
- Respecte les disponibilités des professeurs vacataires
- Libère le mercredi après-midi pour les permanents (réunion pédagogique)
- Regroupe les heures de chaque prof sur le moins de jours possible (utile pour les vacataires qui viennent de loin)
- Produit des journées compactes sans trous pour les élèves

**Ce qu'il ne fait pas :**
- Il ne connaît pas les préférences personnelles des profs (sauf ce que vous renseignez dans le fichier)
- Le résultat peut nécessiter quelques ajustements manuels à l'écran 3
""")

    # ── Les 4 étapes ──
    with st.expander("🔢 Les 4 étapes en un coup d'œil"):
        st.markdown("""
**Étape 1 — Importer**
Téléchargez le template vierge, remplissez-le avec les données de votre école,
puis importez-le ici. L'outil vérifie vos données et vous signale les erreurs
avant de commencer.

**Étape 2 — Générer**
Cliquez sur « Générer ». Le moteur calcule automatiquement la meilleure
grille possible. Selon la taille de l'école, cela prend entre 30 secondes
et 3 minutes.

**Étape 3 — Ajuster**
Visualisez la grille par classe ou par professeur. Si un cours ne vous convient
pas, déplacez-le à la main. L'outil vérifie les conflits automatiquement.

**Étape 4 — Exporter**
Téléchargez les PDF (un par classe, un par prof) et l'Excel complet.
Les PDF sont prêts à imprimer et afficher dans l'école.
""")

    # ── Comment remplir le template ──
    with st.expander("📋 Comment remplir le template Excel ?"):
        st.markdown("""
Le template contient plusieurs onglets. Voici ce qu'il faut remplir dans chacun :

**Onglet « Classes »**
Une ligne par classe (ex : 6ème A, 1ère D, Tle D). Remplissez le nom exact,
le niveau et la série si c'est une classe de lycée.

**Onglet « Professeurs »**
Un ligne par professeur. Précisez le nom exact (ex : M. KONÉ Mamadou),
le statut (Permanent ou Vacataire) et éventuellement le nombre d'heures max
par semaine.

**Onglet « Services »**
C'est l'onglet le plus important. Une ligne par cours :
- Professeur (nom exact, identique à l'onglet Professeurs)
- Classe (nom exact, identique à l'onglet Classes)
- Matière
- Heures par semaine
- Taille du bloc (optionnel : 2h ou 3h pour forcer des heures consécutives)
- Jour imposé (optionnel : Lundi, Mardi… pour forcer un jour précis)

**Onglet « Disponibilités »**
Indiquez les créneaux où chaque professeur n'est PAS disponible
(matin ou après-midi, par jour). Laissez « Oui » si disponible, « Non » sinon.
Crucial pour les vacataires qui ont un autre établissement.

**Onglet « Paramètres »**
Réglages de l'établissement : nom, année scolaire, créneaux horaires,
règle EPS heures chaudes, et durée des séances par matière.
""")

    # ── Les règles importantes ──
    with st.expander("⚙️ Les règles du moteur (ce qu'il respecte automatiquement)"):
        st.markdown("""
**Zéro conflit**
Un professeur ne peut jamais être placé dans deux classes en même temps.
Une classe ne peut jamais avoir deux cours en même temps.

**Disponibilités des vacataires**
Les créneaux marqués « Non » dans l'onglet Disponibilités sont strictement
interdits pour le prof concerné.

**Mercredi après-midi libéré**
Les professeurs permanents ne sont jamais placés le mercredi après-midi
(réunion pédagogique). Les vacataires peuvent l'être.

**Journées compactes**
Le moteur essaie de ne pas laisser de trous dans les journées des élèves.
Un élève ne devrait pas avoir cours à 8h, puis à 16h seulement.

**Regroupement des profs**
Le moteur essaie de regrouper les heures de chaque prof sur le moins de jours
possible — utile pour les vacataires qui font des trajets.

**Durée des séances**
Réglable par matière dans l'onglet Paramètres. Exemple : EPS = 2h exactement,
SVT = 2h à 3h, Philosophie = 1h à 2h. Une séance est toujours en heures
consécutives (jamais coupée).

**EPS aux heures chaudes**
Activable dans l'onglet Paramètres : interdit l'EPS entre 12h et 15h
(heures de forte chaleur). Les plages horaires sont réglables.
""")

    # ── FAQ ──
    with st.expander("❓ Questions fréquentes"):
        st.markdown("""
**« Mon emploi du temps est infaisable, pourquoi ? »**
L'outil vous dit précisément pourquoi dans l'écran 1 (section erreurs).
Les causes les plus courantes :
- Un professeur a plus d'heures que de créneaux disponibles dans la semaine
- Un volume horaire ne peut pas se découper avec la durée de séance imposée
  (ex : 5h de physique avec des séances de 2h exactement → impossible)
- Un jour imposé est incompatible avec les indisponibilités du prof ce jour-là

**« Le prof vacataire a encore des journées éclatées »**
Le moteur fait de son mieux selon les contraintes. Si un prof a beaucoup de
classes, il peut être impossible de le regrouper sur 2 jours. Vous pouvez
aussi limiter ses jours de disponibilité dans l'onglet Disponibilités.

**« J'ai déplacé un cours et maintenant il y a un conflit »**
L'outil vérifie les conflits avant d'appliquer chaque déplacement. Si le
déplacement est refusé, il vous explique pourquoi. Si vous avez fait une
erreur, le bouton « Annuler » revient à la situation précédente.

**« Puis-je utiliser cet outil pour une école qui a cours le samedi ? »**
Oui, modifiez les jours de cours dans l'onglet Paramètres du template.

**« Le résultat change à chaque génération »**
Oui, c'est normal. Il peut exister des milliers de grilles valides. Le moteur
en trouve une bonne dans le temps imparti. Si vous n'êtes pas satisfait,
vous pouvez relancer la génération ou ajuster manuellement à l'écran 3.

**« Mes données sont-elles confidentielles ? »**
Le fichier Excel que vous importez est traité par l'application le temps de la
session. Votre travail (emploi du temps, ajustements) est sauvegardé
automatiquement dans **votre propre navigateur** pour que vous le retrouviez si
la connexion se coupe ou si vous fermez l'onglet. Ces données restent sur votre
appareil ; vous pouvez les effacer à tout moment avec le bouton « Effacer la
sauvegarde » de l'écran 1.
""")

    # ── Contact / signalement ──
    st.divider()
    st.caption(
        "Un problème ? Une suggestion ? Contactez le développeur ou signalez "
        "le bug directement sur "
        "[GitHub](https://github.com/hienmarius394-dev/generateur-edt)."
    )


# ════════════════════════ SAUVEGARDE AUTO (fin de script) ════════════════════════
# Appelé en TOUT dernier : l'état complet des 4 écrans est à jour ici.
sauvegarder_si_change()
