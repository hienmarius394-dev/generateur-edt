# -*- coding: utf-8 -*-
"""Persistance de l'état dans le navigateur (localStorage).

But : qu'une coupure réseau ou une fermeture d'onglet ne fasse PAS perdre le
travail. Au retour sur l'app, l'emploi du temps, les ajustements et l'historique
sont restaurés automatiquement.

Stratégie :
  - On sérialise tout st.session_state utile en JSON compact, compressé (zlib),
    encodé base64 pour tenir dans une chaîne localStorage.
  - On n'écrit QUE quand l'état a changé (empreinte), pour ne pas ralentir.
  - La restauration n'a lieu qu'une fois par session (drapeau).

Les structures Python non-JSON (tuples en clés, sets, defaultdict) sont
converties explicitement, puis reconstruites à l'identique au chargement.
"""
import json
import zlib
import base64
from collections import defaultdict

CLE_STORAGE = "edt_sauvegarde_auto_v1"
CLE_VERSIONS = "edt_versions_v1"   # liste de versions nommées (séparée de l'auto)
MAX_VERSIONS = 10                   # au-delà, la plus ancienne est retirée

# Clés de session_state que l'on sauvegarde (couvre les 4 écrans).
CLES_PERSISTEES = [
    "donnees", "regles", "meta", "fichier_nom",
    "erreurs", "avertissements",
    "emplois", "log_moteur", "duree_generation",
    "historique", "nb_modifs", "chaine",
]


# ───────────────────────── Sérialisation des structures ─────────────────────────
def _emplois_vers_json(emplois):
    """{(classe, jour, slot): info} -> {"classe|jour|slot": info}."""
    if not emplois:
        return None
    return {f"{cl}|{d}|{t}": info for (cl, d, t), info in emplois.items()}


def _json_vers_emplois(d):
    if not d:
        return None
    out = {}
    for k, info in d.items():
        cl, dd, tt = k.rsplit("|", 2)
        out[(cl, int(dd), int(tt))] = info
    return out


def _indispos_vers_json(indispos):
    """defaultdict(set) avec valeurs = set de tuples (jour, slot)
    -> {prof: [[jour, slot], ...]}."""
    if not indispos:
        return None
    return {prof: [list(x) for x in sorted(slots)]
            for prof, slots in indispos.items()}


def _json_vers_indispos(d):
    out = defaultdict(set)
    if d:
        for prof, slots in d.items():
            out[prof] = {tuple(x) for x in slots}
    return out


def _donnees_vers_json(donnees):
    """(classes, permanents, services, indispos) -> dict JSON."""
    if not donnees:
        return None
    classes, permanents, services, indispos = donnees
    return {
        "classes": list(classes),
        "permanents": sorted(permanents),          # set -> liste triée
        "services": services,                      # liste de dicts (déjà JSON-ok)
        "indispos": _indispos_vers_json(indispos),
    }


def _json_vers_donnees(d):
    if not d:
        return None
    return (
        d["classes"],
        set(d["permanents"]),
        d["services"],
        _json_vers_indispos(d["indispos"]),
    )


def _chaine_vers_json(chaine):
    """liste de ((d1,t1),(d2,t2)) -> liste de [[d1,t1],[d2,t2]]."""
    return [[list(src), list(dst)] for (src, dst) in chaine]


def _json_vers_chaine(d):
    return [(tuple(src), tuple(dst)) for (src, dst) in d] if d else []


# ───────────────────────── État complet <-> texte ─────────────────────────
def serialiser_etat(session_state):
    """Construit la chaîne compressée représentant l'état à sauvegarder.
    Retourne None s'il n'y a rien d'utile à sauvegarder (pas de fichier chargé)."""
    if not session_state.get("donnees") and not session_state.get("emplois"):
        return None

    emplois = session_state.get("emplois")
    historique = session_state.get("historique") or []

    paquet = {
        "donnees": _donnees_vers_json(session_state.get("donnees")),
        "regles": _regles_vers_json(session_state.get("regles") or {}),
        "meta": _meta_vers_json(session_state.get("meta") or {}),
        "fichier_nom": session_state.get("fichier_nom"),
        "erreurs": session_state.get("erreurs") or [],
        "avertissements": session_state.get("avertissements") or [],
        "emplois": _emplois_vers_json(emplois),
        "log_moteur": session_state.get("log_moteur") or "",
        "duree_generation": session_state.get("duree_generation"),
        "historique": [_emplois_vers_json(e) for e in historique],
        "nb_modifs": session_state.get("nb_modifs", 0),
        "chaine": _chaine_vers_json(session_state.get("chaine") or []),
    }
    brut = json.dumps(paquet, ensure_ascii=False).encode("utf-8")
    return base64.b64encode(zlib.compress(brut, 9)).decode("ascii")


def deserialiser_etat(texte):
    """Reconstruit le dict de valeurs session_state depuis la chaîne stockée.
    Retourne None si le texte est vide ou illisible (corruption, ancien format)."""
    if not texte:
        return None
    try:
        brut = zlib.decompress(base64.b64decode(texte)).decode("utf-8")
        paquet = json.loads(brut)
    except Exception:
        return None

    return {
        "donnees": _json_vers_donnees(paquet.get("donnees")),
        "regles": _json_vers_regles(paquet.get("regles") or {}),
        "meta": _json_vers_meta(paquet.get("meta") or {}),
        "fichier_nom": paquet.get("fichier_nom"),
        "erreurs": paquet.get("erreurs") or [],
        "avertissements": paquet.get("avertissements") or [],
        "emplois": _json_vers_emplois(paquet.get("emplois")),
        "log_moteur": paquet.get("log_moteur") or "",
        "duree_generation": paquet.get("duree_generation"),
        "historique": [_json_vers_emplois(e) for e in (paquet.get("historique") or [])],
        "nb_modifs": paquet.get("nb_modifs", 0),
        "chaine": _json_vers_chaine(paquet.get("chaine")),
    }


def _meta_vers_json(meta):
    """meta peut contenir un set (profs_declares) -> liste."""
    m = dict(meta)
    if isinstance(m.get("profs_declares"), set):
        m["profs_declares"] = sorted(m["profs_declares"])
    return m


def _json_vers_meta(meta):
    m = dict(meta)
    if isinstance(m.get("profs_declares"), list):
        m["profs_declares"] = set(m["profs_declares"])
    return m


def _regles_vers_json(regles):
    """`eps_heures_chaudes` est un set de tuples (jour, slot) -> liste de listes.
    Les autres clés de regles sont des types JSON simples (bool, etc.)."""
    if not regles:
        return {}
    r = dict(regles)
    chaud = r.get("eps_heures_chaudes")
    if chaud is not None:
        r["eps_heures_chaudes"] = [list(x) if isinstance(x, (tuple, list)) else x
                                   for x in sorted(chaud)]
    return r


def _json_vers_regles(regles):
    if not regles:
        return {}
    r = dict(regles)
    chaud = r.get("eps_heures_chaudes")
    if chaud is not None:
        r["eps_heures_chaudes"] = {tuple(x) if isinstance(x, list) else x
                                   for x in chaud}
    return r


def empreinte(texte):
    """Petite empreinte pour détecter un changement sans tout recomparer."""
    if texte is None:
        return None
    # longueur + checksum zlib : suffisant pour repérer une modification.
    return f"{len(texte)}:{zlib.adler32(texte.encode('ascii'))}"


# ═══════════════════════ VERSIONS NOMMÉES (historique) ═══════════════════════
# Une « version » = un instantané d'emploi du temps que l'utilisateur nomme et
# peut retrouver/restaurer plus tard. Stockées en JSON compressé sous une clé
# dédiée du navigateur, indépendantes de la sauvegarde automatique.
#
# Format d'une version :
#   {"id": str, "nom": str, "horodatage": str, "nb_cours": int,
#    "emplois": {<clés sérialisées>}}

def versions_vers_texte(versions):
    """Liste de versions (dicts) -> chaîne compressée pour le navigateur."""
    if not versions:
        return ""
    brut = json.dumps(versions, ensure_ascii=False).encode("utf-8")
    return base64.b64encode(zlib.compress(brut, 9)).decode("ascii")


def texte_vers_versions(texte):
    """Chaîne du navigateur -> liste de versions. [] si vide/illisible."""
    if not texte:
        return []
    try:
        brut = zlib.decompress(base64.b64decode(texte)).decode("utf-8")
        return json.loads(brut)
    except Exception:
        return []


def creer_version(emplois, nom, horodatage):
    """Construit un dict-version à partir d'un emploi du temps en mémoire."""
    import uuid
    return {
        "id": uuid.uuid4().hex[:12],
        "nom": nom,
        "horodatage": horodatage,
        "nb_cours": len(emplois) if emplois else 0,
        "emplois": _emplois_vers_json(emplois),
    }


def version_vers_emplois(version):
    """Reconstruit l'emploi du temps (clés tuples) depuis une version stockée."""
    return _json_vers_emplois(version.get("emplois"))


def ajouter_version(versions, nouvelle, max_versions=MAX_VERSIONS):
    """Ajoute `nouvelle` en tête de liste, tronque au-delà de max_versions.
    Retourne la nouvelle liste (la plus récente en premier)."""
    liste = [nouvelle] + list(versions or [])
    return liste[:max_versions]


def retirer_version(versions, version_id):
    """Retourne la liste sans la version d'identifiant `version_id`."""
    return [v for v in (versions or []) if v.get("id") != version_id]


def renommer_version(versions, version_id, nouveau_nom):
    """Retourne une nouvelle liste où la version `version_id` porte
    `nouveau_nom` (les autres versions sont inchangées)."""
    nouveau_nom = (nouveau_nom or "").strip()
    if not nouveau_nom:
        return list(versions or [])
    resultat = []
    for v in (versions or []):
        if v.get("id") == version_id:
            v = dict(v)
            v["nom"] = nouveau_nom
        resultat.append(v)
    return resultat


# ─────────────── Export / import d'une version en FICHIER ───────────────
# Filet de sécurité : le localStorage peut disparaître (cache vidé, autre
# navigateur, autre appareil). Une version exportée en fichier .json se
# réimporte partout.

FORMAT_FICHIER_VERSION = "edt-version-1"   # marqueur de format (compatibilité)


def version_vers_fichier(version):
    """Version -> texte JSON lisible, prêt à télécharger."""
    contenu = {"format": FORMAT_FICHIER_VERSION}
    contenu.update({k: version.get(k) for k in
                    ("nom", "horodatage", "nb_cours", "emplois")})
    return json.dumps(contenu, ensure_ascii=False, indent=1)


def fichier_vers_version(texte):
    """Texte JSON d'un fichier exporté -> version (avec un id NEUF).
    Retourne None si le fichier est illisible ou d'un autre format."""
    import uuid
    try:
        contenu = json.loads(texte)
    except Exception:
        return None
    if not isinstance(contenu, dict):
        return None
    if contenu.get("format") != FORMAT_FICHIER_VERSION:
        return None
    emplois = contenu.get("emplois")
    if not isinstance(emplois, dict) or not emplois:
        return None
    # Validation minimale : les clés doivent se désérialiser en tuples 3 champs
    try:
        if _json_vers_emplois(emplois) is None:
            return None
    except Exception:
        return None
    return {
        "id": uuid.uuid4().hex[:12],
        "nom": str(contenu.get("nom") or "Version importée"),
        "horodatage": str(contenu.get("horodatage") or ""),
        "nb_cours": int(contenu.get("nb_cours") or len(emplois)),
        "emplois": emplois,
    }

