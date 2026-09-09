"""
Traitement des résultats scrapés (results.py) en statistiques par joueur :
victoires/défaites par tableau, split partenaire club/hors club, indices,
et dédoublonnage des tournois/interclubs.
"""
from collections import defaultdict
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


def _split_partenaire(matchs):
    avec_club = [m for m in matchs if m["partenaire_club"] == results.MY_CLUB]
    hors_club = [m for m in matchs
                 if m["partenaire_club"] is not None and m["partenaire_club"] != results.MY_CLUB]
    avec, hors = _agg_matchs(avec_club), _agg_matchs(hors_club)
    return {
        "avec_partenaire_club": avec,
        "sans_partenaire_club": hors,
        "delta_pct": avec["pct_victoire"] - hors["pct_victoire"],
    }


def discipline_stats_partenaire(events):
    """Double ou Mixte : split selon que le partenaire est du club (GAB38)
    ou non, + différence de % de victoire entre les deux."""
    return _split_partenaire([m for e in events for m in e["matchs"]])


def discipline_stats_partenaire_combinee(events_double, events_mixte):
    """Même chose mais Double + Mixte fusionnés (le partenaire est un
    partenaire, peu importe le tableau)."""
    matchs = [m for events in (events_double, events_mixte) for e in events for m in e["matchs"]]
    return _split_partenaire(matchs)


def meilleur_partenaire(events):
    """Double ou Mixte : le partenaire avec lequel le joueur a le meilleur
    indice de performance (victoires x taux de victoire - même formule que
    indice_performance, pas de seuil arbitraire : un partenaire joué 1 fois
    et gagné donne un indice de 1, tandis qu'un partenaire joué 5 fois et
    gagné 5 fois donne 5 - le volume l'emporte naturellement). Regroupe par
    licence quand connue, par nom sinon (adversaire jamais identifié en
    double, mais un partenaire l'est presque toujours). None si le joueur
    n'a aucun match dans ce tableau."""
    par_partenaire = defaultdict(list)
    for e in events:
        for m in e["matchs"]:
            cle = m["partenaire_licence"] or m["partenaire_nom"]
            if cle is None:
                continue
            par_partenaire[cle].append(m)

    meilleur = None
    meilleur_indice = -1
    for matchs in par_partenaire.values():
        agg = _agg_matchs(matchs)
        indice = indice_performance(agg)
        if indice > meilleur_indice:
            meilleur_indice = indice
            meilleur = {
                "nom": matchs[0]["partenaire_nom"],
                "matchs_joues": agg["matchs_joues"],
                "victoires": agg["victoires"],
                "pct_victoire": agg["pct_victoire"],
                "indice_performance": indice,
            }
    return meilleur


def diff_classement(player):
    """Progression entre le 1er septembre (player["Classement 1er
    septembre"], rempli par statsJoueurs.scrape_player via
    results.parse_classement_evolution + find_september_1 : {"Simple":
    {"rang","classement","points"}, "Double": {...}, "Mixte": {...}}) et
    l'actuel (roster), par tableau + cumulé sur les 3.

    - diff_points = cote actuelle - cote de septembre, diff_relatif = ce
      diff / cote de septembre (progression relative depuis le début de
      saison).
    - gain_places = rang de septembre - rang actuel : positif = a gagné des
      places (un rang plus PETIT est meilleur, donc rang qui baisse)."""
    sept = player.get("Classement 1er septembre") or {}
    rang_actuel = player.get("Rang Actuel") or {}

    resultat = {}
    total_diff_points = total_sept_points = 0.0
    for tableau in ("Simple", "Double", "Mixte"):
        actuel_points = player["Points Actuel"][tableau]
        actuel_rang = rang_actuel.get(tableau)
        s = sept.get(tableau)  # {"rang","classement","points"} ou None

        diff_points = diff_rel = None
        if s is not None and s.get("points") is not None:
            diff_points = actuel_points - s["points"]
            diff_rel = (diff_points / s["points"]) if s["points"] else None

        gain_places = None
        if s is not None and s.get("rang") is not None and actuel_rang is not None:
            gain_places = s["rang"] - actuel_rang

        resultat[tableau] = {
            "septembre_classement": s.get("classement") if s else None,
            "septembre_rang": s.get("rang") if s else None,
            "septembre_points": s.get("points") if s else None,
            "actuel_rang": actuel_rang,
            "actuel_points": actuel_points,
            "diff_points": diff_points,
            "diff_relatif": diff_rel,
            "gain_places": gain_places,
        }
        if diff_points is not None:
            total_diff_points += diff_points
            total_sept_points += s["points"]

    resultat["cumule"] = {
        "diff_points": total_diff_points,
        "diff_relatif": (total_diff_points / total_sept_points) if total_sept_points else None,
    }
    return resultat


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
            "cote": player["Points Actuel"][tableau],
            "indice_niveau": niveau,
            "indice_performance_brut": perf,
            "indice_global_brut": indice_global(perf, niveau),
        }
        if tableau in ("Double", "Mixte"):
            entry["partenaire"] = discipline_stats_partenaire(events)
            entry["meilleur_partenaire"] = meilleur_partenaire(events)
        par_tableau[tableau] = entry

    partenaire_double_mixte = discipline_stats_partenaire_combinee(
        events_par_tableau.get("Double", []), events_par_tableau.get("Mixte", []))

    tous_matchs = [m for events in events_par_tableau.values() for e in events for m in e["matchs"]]
    stats_globales = _agg_matchs(tous_matchs)
    perf_globale = indice_performance(stats_globales)
    niveau_global = player["Valeur IC"]  # meilleur tableau des 3, déjà calculé (pointsIC)

    tournois = build_tournois(events_par_tableau)

    # un enregistrement par match (tableau, mois, victoire), pour la
    # ventilation mensuelle club-wide (onglet Stats club) - inutile de
    # stocker le sexe ici, on l'a déjà au niveau du joueur.
    match_log = [
        {"tableau": tableau, "mois": parse_date_fr(e["date"]).strftime("%Y-%m"), "victoire": m["victoire"]}
        for tableau, events in events_par_tableau.items()
        for e in events
        for m in e["matchs"]
    ]

    return {
        "Licence": player["Licence"],
        "Nom": player["Nom"],
        "Sexe": player["Sexe"],
        "par_tableau": par_tableau,
        "partenaire_double_mixte": partenaire_double_mixte,
        "progression": diff_classement(player),
        "global": {
            **stats_globales,
            "meilleur_tableau": player["Meilleur tableau"],
            "rang_meilleur_tableau": player.get("Rang meilleur tableau"),
            "points_meilleur_tableau": player["Points meilleur tableau"],
            "indice_niveau": niveau_global,
            "indice_performance_brut": perf_globale,
            "indice_global_brut": indice_global(perf_globale, niveau_global),
        },
        "tournois": tournois,
        "match_log": match_log,
    }


def normaliser_club(all_stats):
    """Normalise les indices bruts en % du max observé dans le club (une
    fois tous les joueurs traités), par tableau et au global. Modifie
    all_stats en place (ajoute indice_performance/indice_global à côté des
    valeurs _brut) et renvoie les maxima bruts observés."""
    maxima = {}
    for tableau in ("Simple", "Double", "Mixte"):
        entries = [s["par_tableau"][tableau] for s in all_stats]
        maxima[f"{tableau} - performance"] = normaliser(entries, "indice_performance_brut", "indice_performance")
        maxima[f"{tableau} - global"] = normaliser(entries, "indice_global_brut", "indice_global")

    globaux = [s["global"] for s in all_stats]
    maxima["Global - performance"] = normaliser(globaux, "indice_performance_brut", "indice_performance")
    maxima["Global - global"] = normaliser(globaux, "indice_global_brut", "indice_global")
    return maxima


def compute_ordre_tableau(all_stats):
    """Classe tous les joueurs du club : meilleur tableau (N1 > N2 > ...),
    puis en cas d'égalité leur position dans ce tableau (rang club), puis en
    cas d'égalité (rare) leur cote. Ajoute s["global"]["ordre_tableau"] =
    position 1..N (modifie all_stats en place)."""
    def key(s):
        g = s["global"]
        tier = roster.TABLEAUX.index(g["meilleur_tableau"])
        rang = g["rang_meilleur_tableau"] if g["rang_meilleur_tableau"] is not None else 10 ** 9
        return (tier, rang, -g["points_meilleur_tableau"])

    for position, s in enumerate(sorted(all_stats, key=key), start=1):
        s["global"]["ordre_tableau"] = position
