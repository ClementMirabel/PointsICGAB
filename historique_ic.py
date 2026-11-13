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
        "joueurs": {"06826828": {"nom", "classements", "cotes", "meilleur_tableau", "icn", "icr"}}}}}
"""
import json
import os
from datetime import date, timedelta

FICHIER = "historique_valeurs.json"


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
