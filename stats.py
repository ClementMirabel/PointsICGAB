"""
Traitement des résultats scrapés (results.py) en statistiques par joueur :
victoires/défaites par tableau, split partenaire club/hors club, indices,
et dédoublonnage des tournois/interclubs.
"""
from datetime import date, timedelta

import holidays as holidays_lib

import results
import roster

FR_HOLIDAYS = holidays_lib.France()

MOIS_FR = {
    "janvier": 1, "février": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
    "juillet": 7, "août": 8, "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12,
}


def parse_date_fr(text):
    """'18 janvier 2026' -> date(2026, 1, 18)."""
    jour, mois, annee = text.split()
    return date(int(annee), MOIS_FR[mois.lower()], int(jour))


def _est_pont(d):
    """Lundi collé à un mardi férié, ou vendredi collé à un jeudi férié."""
    if d.weekday() == 0 and (d + timedelta(days=1)) in FR_HOLIDAYS:
        return True
    if d.weekday() == 4 and (d - timedelta(days=1)) in FR_HOLIDAYS:
        return True
    return False


def type_jour(d):
    """'weekend' (samedi/dimanche, jour férié ou pont) ou 'soiree' (semaine)."""
    if d.weekday() >= 5 or d in FR_HOLIDAYS or _est_pont(d):
        return "weekend"
    return "soiree"


def _agg_matchs(matchs):
    joues = len(matchs)
    victoires = sum(1 for m in matchs if m["victoire"])
    pct = (victoires / joues * 100) if joues else 0.0
    return {"matchs_joues": joues, "victoires": victoires, "pct_victoire": pct}


def discipline_stats(events):
    """events d'un seul tableau (Simple/Double/Mixte) -> stats agrégées
    (matchs joués, victoires, % victoire)."""
    return _agg_matchs([m for e in events for m in e["matchs"]])


def discipline_stats_partenaire(events):
    """Double/Mixte uniquement : split selon que le partenaire est du club
    (GAB38) ou non, + différence de % de victoire entre les deux."""
    matchs = [m for e in events for m in e["matchs"]]
    avec_club = [m for m in matchs if m["partenaire_club"] == results.MY_CLUB]
    hors_club = [m for m in matchs
                 if m["partenaire_club"] is not None and m["partenaire_club"] != results.MY_CLUB]

    avec = _agg_matchs(avec_club)
    hors = _agg_matchs(hors_club)
    return {
        "avec_partenaire_club": avec,
        "sans_partenaire_club": hors,
        "delta_pct": avec["pct_victoire"] - hors["pct_victoire"],
    }


def indice_performance(stats):
    """victoires x taux de victoire : récompense le volume de victoires ET
    le taux de réussite en même temps (une valeur brute, normalisée plus
    tard en % du max observé dans le club)."""
    if stats["matchs_joues"] == 0:
        return 0.0
    return stats["victoires"] * (stats["victoires"] / stats["matchs_joues"])


def indice_global(indice_perf, indice_niv):
    """Un joueur qui gagne beaucoup ET à un niveau élevé doit ressortir en
    tête sur les deux facteurs à la fois (valeur brute, normalisée après
    coup en % du max du club)."""
    return indice_perf * indice_niv


def normaliser(items, cle_brute, cle_normalisee, echelle=100):
    """Ajoute item[cle_normalisee] = item[cle_brute] en % du max observé
    dans `items`. Renvoie ce maximum."""
    valeurs = [item[cle_brute] for item in items if item.get(cle_brute) is not None]
    maximum = max(valeurs) if valeurs else 0
    for item in items:
        brute = item.get(cle_brute)
        item[cle_normalisee] = (brute / maximum * echelle) if maximum and brute is not None else 0.0
    return maximum


def build_tournois(events_par_tableau):
    """events_par_tableau : dict {"Simple": events, "Double": events,
    "Mixte": events} (issus de trois appels à results.parse_results).
    Fusionne et dédoublonne : un tournoi individuel = un tournoi_id (compté
    une fois même joué dans plusieurs tableaux/jours) ; un interclub = une
    date de rencontre (peut regrouper plusieurs ties/adversaires le même
    jour)."""
    tournois = {}  # tournoi_id -> {date, nom, type}
    interclubs_dates = {}  # date -> nom (le premier rencontré)

    for events in events_par_tableau.values():
        for e in events:
            d = parse_date_fr(e["date"])
            if e["est_interclub"]:
                interclubs_dates.setdefault(d, e["evenement"])
                continue
            tid = next((m["tournoi_id"] for m in e["matchs"] if m["tournoi_id"]), None)
            if tid is None:
                continue
            if tid not in tournois:
                tournois[tid] = {"date": d, "nom": e["evenement"], "type": type_jour(d)}

    liste_tournois = list(tournois.values())
    return {
        "tournois": liste_tournois,
        "interclubs_dates": sorted(interclubs_dates.keys()),
        "nb_tournois": len(liste_tournois),
        "nb_tournois_weekend": sum(1 for t in liste_tournois if t["type"] == "weekend"),
        "nb_tournois_soiree": sum(1 for t in liste_tournois if t["type"] == "soiree"),
        "nb_interclubs": len(interclubs_dates),
    }


def build_player_stats(player, events_par_tableau):
    """player : entrée roster (Nom/Sexe/Classement Actuel/Points Actuel),
    complétée par le "Meilleur tableau"/"Valeur IC" déjà calculés
    (pointsIC.compute_IC_points). events_par_tableau : voir build_tournois.

    Renvoie un dict prêt à être normalisé au niveau du club puis écrit dans
    l'Excel."""
    par_tableau = {}
    for tableau in ("Simple", "Double", "Mixte"):
        events = events_par_tableau.get(tableau, [])
        stats = discipline_stats(events)
        classement = player["Classement Actuel"][tableau]
        points = {"Simple": 0.0, "Double": 0.0, "Mixte": 0.0}
        points[tableau] = player["Points Actuel"][tableau]
        niveau = roster.valeur_ic(classement, player["Sexe"],
                                   points["Simple"], points["Double"], points["Mixte"])
        perf = indice_performance(stats)
        entry = {
            **stats,
            "classement": classement,
            "indice_niveau": niveau,
            "indice_performance_brut": perf,
            "indice_global_brut": indice_global(perf, niveau),
        }
        if tableau in ("Double", "Mixte"):
            entry["partenaire"] = discipline_stats_partenaire(events)
        par_tableau[tableau] = entry

    tous_matchs = [m for events in events_par_tableau.values() for e in events for m in e["matchs"]]
    stats_globales = _agg_matchs(tous_matchs)
    perf_globale = indice_performance(stats_globales)
    niveau_global = player["Valeur IC"]  # meilleur tableau des 3, déjà calculé (pointsIC)

    tournois = build_tournois(events_par_tableau)

    return {
        "Licence": player["Licence"],
        "Nom": player["Nom"],
        "Sexe": player["Sexe"],
        "par_tableau": par_tableau,
        "global": {
            **stats_globales,
            "indice_niveau": niveau_global,
            "indice_performance_brut": perf_globale,
            "indice_global_brut": indice_global(perf_globale, niveau_global),
        },
        "tournois": tournois,
    }
