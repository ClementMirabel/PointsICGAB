"""
Valeurs des joueurs figées par week-end d'interclubs.

Les classements évoluent chaque semaine : recalculer la valeur d'une équipe
d'un week-end passé avec les classements du jour la ferait varier d'un export
à l'autre. On fige donc, au premier export qui suit chaque week-end (celui
du jeudi), les classements, cotes et valeurs ICN/ICR des joueurs qui y ont
joué. L'historique est un fichier JSON conservé d'une semaine sur l'autre
(branche git "donnees" du dépôt en CI, voir le workflow ; fichier local
sinon) :

    {"v": 1, "weekends": {"2026-11-07": {"date": "2026-11-12",
        "joueurs": {"06826828": {"nom", "classements", "cotes", "meilleur_tableau", "icn", "icr"}}}},
     "exports": {"2026-11-19T01:12Z": {"cotes": {licence: [S, D, M]},
                                       "forme": {licence: [globale, S, D, M]}}}}

"exports" : un instantané par export (un par jour au plus) - cotes de tout
l'effectif (cotes "J-10" des week-ends à venir, voir cotes_j10) et indices de
forme des joueurs suivis (tendances sur la saison).
"""
import json
import os
from datetime import date, datetime, timedelta, timezone

FICHIER = "historique_valeurs.json"
# Le CPPH (cote) est republié chaque jeudi dans la journée ; hypothèse :
# publié avant 12:00 UTC. Un instantané pris entre le jeudi J-10 12:00 et le
# jeudi J-3 12:00 (UTC) contient donc les cotes de J-10 du week-end (règle de
# hiérarchie : CPPH du jeudi précédant la semaine de la journée).
HEURE_PUBLICATION_CPPH_UTC = 12


def charger(path):
    """Historique existant, ou un historique vide (fichier absent ou illisible)."""
    try:
        with open(path, encoding="utf-8") as f:
            hist = json.load(f)
        if isinstance(hist, dict) and isinstance(hist.get("weekends"), dict):
            return hist
    except (OSError, ValueError):
        pass
    return {"v": 1, "weekends": {}}


def sauver(hist, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1, sort_keys=True)


def ajouter_export(hist, all_stats, effectif, maintenant=None):
    """Instantané de cet export : cotes de tout l'effectif et indices de forme
    (global, S, D, M) des joueurs suivis. Un seul par jour (le dernier gagne)."""
    maintenant = maintenant or datetime.now(timezone.utc)
    exports = hist.setdefault("exports", {})
    jour = maintenant.strftime("%Y-%m-%d")
    for cle in [k for k in exports if k.startswith(jour)]:
        del exports[cle]
    arrondi = lambda v: round(v, 1) if isinstance(v, (int, float)) else None
    exports[maintenant.strftime("%Y-%m-%dT%H:%MZ")] = {
        "cotes": {p["licence"]: p["cotes"] for p in (effectif or [])},
        # classements S/D/M : montées et descentes de la semaine (actualités)
        "classements": {p["licence"]: p["classements"] for p in (effectif or [])},
        "forme": {s["Licence"]: [arrondi(s["global"].get("indice_forme"))]
                  + [arrondi(s["par_tableau"][t].get("indice_forme")) for t in ("Simple", "Double", "Mixte")]
                  for s in all_stats},
    }


def _instant(cle):
    return datetime.strptime(cle, "%Y-%m-%dT%H:%MZ").replace(tzinfo=timezone.utc)


def cotes_j10(hist, samedi):
    """(clé de l'export, {licence: cotes}) contenant le CPPH de J-10 du
    week-end (jeudi qui précède la semaine du week-end), ou None si aucun
    instantané ne tombe dans la fenêtre (week-end lointain : pas encore publié)."""
    dimanche = date.fromisoformat(samedi) + timedelta(days=1)
    jeudi = dimanche - timedelta(days=10)
    debut = datetime(jeudi.year, jeudi.month, jeudi.day, HEURE_PUBLICATION_CPPH_UTC, tzinfo=timezone.utc)
    fin = debut + timedelta(days=7)
    candidats = [k for k in hist.get("exports", {}) if debut <= _instant(k) < fin]
    if not candidats:
        return None
    cle = max(candidats)
    return cle, hist["exports"][cle]["cotes"]


def samedi_de(jour):
    """Samedi du week-end d'une date (le samedi lui-même, ou le samedi précédent)."""
    return jour - timedelta(days=(jour.weekday() - 5) % 7)


def mettre_a_jour(hist, interclubs, effectif, aujourd_hui=None):
    """Fige les week-ends terminés qui ne le sont pas encore : joueurs ayant
    joué au moins un match (feuilles icbad) -> leurs classements/cotes/valeurs
    actuels (effectif du jour = semaine qui suit le week-end). Un week-end est
    terminé quand toutes ses rencontres sont jouées ou que son dimanche est
    passé. Renvoie la liste des week-ends ajoutés."""
    if not interclubs or not effectif:
        return []
    aujourd_hui = aujourd_hui or date.today()
    par_licence = {p["licence"]: p for p in effectif}
    rencontres_par_weekend = {}
    for e in interclubs.get("equipes", []):
        for r in e["rencontres"]:
            if r.get("date"):
                sam = samedi_de(date.fromisoformat(r["date"])).isoformat()
                rencontres_par_weekend.setdefault(sam, []).append(r)
    ajoutes = []
    for sam, rencontres in sorted(rencontres_par_weekend.items()):
        if sam in hist["weekends"]:
            continue
        dimanche = date.fromisoformat(sam) + timedelta(days=1)
        if not (all(r["jouee"] for r in rencontres) or dimanche < aujourd_hui):
            continue
        joueurs = {}
        for r in rencontres:
            for m in r["matchs"]:
                for p in m["club"]:
                    lic = p.get("licence")
                    if lic and lic in par_licence and lic not in joueurs:
                        x = par_licence[lic]
                        joueurs[lic] = {k: x.get(k) for k in ("nom", "classements", "cotes", "meilleur_tableau", "icn", "icr")}
        if joueurs:
            hist["weekends"][sam] = {"date": aujourd_hui.isoformat(), "joueurs": joueurs}
            ajoutes.append(sam)
    return ajoutes
