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
    points = [m["points_cote"] for m in matchs if m["points_cote"] is not None]
    return {
        "matchs_joues": joues,
        "victoires": victoires,
        "pct_victoire": pct,
        # somme des points de cote gagnés/perdus sur ces matchs - déjà
        # ajustée à la force de l'adversaire par le barème FFBad officiel
        # (battre quelqu'un de mieux classé rapporte plus que battre
        # quelqu'un de plus faible, et inversement pour une défaite) : sert
        # de signal de qualité, complémentaire du simple compte de
        # victoires (voir indice_qualite).
        "points_cote_total": sum(points) if points else 0.0,
    }


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


def _partenaires_stats(matchs):
    """matchs -> dict clé_partenaire -> agrégat (nom, matchs/victoires/%,
    indice de performance, total des points de cote gagnés/perdus avec ce
    partenaire). Regroupe par licence quand connue, par nom sinon."""
    par_partenaire = defaultdict(list)
    for m in matchs:
        cle = m["partenaire_licence"] or m["partenaire_nom"]
        if cle is None:
            continue
        par_partenaire[cle].append(m)

    resultat = {}
    for cle, ms in par_partenaire.items():
        agg = _agg_matchs(ms)
        resultat[cle] = {
            "nom": ms[0]["partenaire_nom"],
            "matchs_joues": agg["matchs_joues"],
            "victoires": agg["victoires"],
            "pct_victoire": agg["pct_victoire"],
            "indice_performance": indice_performance(agg),
            "points_cote_total": agg["points_cote_total"],
        }
    return resultat


def _meilleur_partenaire_parmi(matchs, club_uniquement=False):
    """Le partenaire qui ressort en tête, classé D'ABORD par points de cote
    cumulés (déjà ajustés à la force de l'adversaire par le barème FFBad -
    voir _agg_matchs), et à égalité par l'indice de performance (volume x
    taux de victoire) comme départage.

    Pas que l'indice de performance seul : constaté qu'un partenaire avec
    beaucoup de victoires mais un bilan de cote négatif ensemble (donc des
    victoires contre des adversaires plus faibles et des défaites contre
    des adversaires plus forts que la moyenne du joueur) ressortait quand
    même "meilleur partenaire" sur le seul critère du volume - alors que le
    résultat réel ensemble était EN DESSOUS du standard habituel du joueur.

    club_uniquement=True restreint aux partenaires du club (GAB38) - pour
    distinguer "meilleur partenaire au club" de "meilleur partenaire, club
    ou pas". None si `matchs` est vide."""
    if club_uniquement:
        matchs = [m for m in matchs if m["partenaire_club"] == results.MY_CLUB]
    partenaires = _partenaires_stats(matchs)
    if not partenaires:
        return None
    return max(partenaires.values(), key=lambda v: (v["points_cote_total"], v["indice_performance"]))


def meilleur_partenaire(events, club_uniquement=False):
    """Double ou Mixte uniquement (un seul tableau)."""
    return _meilleur_partenaire_parmi([m for e in events for m in e["matchs"]], club_uniquement)


def meilleur_partenaire_combine(events_double, events_mixte, club_uniquement=False):
    """Double + Mixte fusionnés (le partenaire est un partenaire, peu
    importe le tableau)."""
    matchs = [m for events in (events_double, events_mixte) for e in events for m in e["matchs"]]
    return _meilleur_partenaire_parmi(matchs, club_uniquement)


def meilleur_partenaire_club_si_different(overall, club):
    """`club` (résultat de meilleur_partenaire(..., club_uniquement=True))
    seulement si c'est une personne différente de `overall` (le meilleur
    partenaire sans restriction) - évite d'afficher deux fois la même
    info quand le meilleur partenaire overall est déjà au club."""
    if not club:
        return None
    if overall and overall["nom"] == club["nom"]:
        return None
    return club


def ordre_disciplines(par_tableau):
    """Les 3 tableaux (Simple/Double/Mixte) triés du plus fort au moins fort
    pour ce joueur, selon indice_global (déjà normalisé club-wide, voir
    normaliser_club - il faut donc appeler cette fonction APRÈS
    normaliser_club, pas depuis build_player_stats).

    indice_global et pas indice_performance_brut (l'ancienne version) :
    c'est le même indice composite utilisé pour classer/trier les joueurs
    partout ailleurs dans le classeur - sinon "Ordre tableau" pouvait
    afficher un ordre qui contredisait l'indice global affiché juste à
    côté (ex: le tableau avec l'indice global le plus haut du joueur pas
    en tête de son propre "Ordre tableau")."""
    tableaux = ("Simple", "Double", "Mixte")
    return sorted(tableaux, key=lambda t: -par_tableau[t]["indice_global"])


def diff_classement(player):
    """Progression entre le 1er septembre (player["Classement 1er
    septembre"], rempli par statsJoueurs.scrape_player via
    results.parse_classement_evolution + find_september_1 : {"Simple":
    {"rang","classement","points"}, "Double": {...}, "Mixte": {...}}) et
    l'actuel (player["Rang National Actuel"], même source - la ligne la plus
    récente de l'historique de classement), par tableau + cumulé sur les 3.

    Le rang utilisé ici est le rang NATIONAL (même échelle que le rang de
    septembre) - pas player["Rang Actuel"] (roster), qui est la position du
    joueur DANS LE CLUB pour ce tableau et n'a donc rien de comparable avec
    un rang de septembre à plusieurs centaines/milliers.

    - diff_points = cote actuelle - cote de septembre. Pas de version
      "relative" (diff / cote de septembre) : sur cette échelle, gagner 50
      points en partant de 1000 n'est pas plus facile que gagner 50 points
      en partant de 4000 - un même diff_points représente un effort
      comparable quel que soit le niveau de départ, donc en faire un
      pourcentage du niveau de départ n'a pas de sens (ça ferait paraître
      un joueur bas niveau plus "progressif" qu'un joueur haut niveau pour
      un progrès équivalent).
    - gain_places = rang de septembre - rang actuel : positif = a gagné des
      places (un rang plus PETIT est meilleur, donc rang qui baisse).
    - gain_tableau = de combien de tableaux (N1/N2/.../P12) le joueur a
      progressé depuis septembre (positif = monté, ex: R4 -> N3 = +1).
    - tendance_hebdo = diff_points / nombre de semaines écoulées depuis le
      1er septembre - progression moyenne par semaine (None si moins d'une
      semaine s'est écoulée, la valeur serait trop instable pour être
      lisible)."""
    sept = player.get("Classement 1er septembre") or {}
    rang_actuel = player.get("Rang National Actuel") or {}
    sept_date = sept.get("date")
    semaines_ecoulees = None
    if sept_date is not None:
        jours = (date.today() - sept_date).days
        if jours >= 7:
            semaines_ecoulees = jours / 7

    def _index_tableau(classement):
        return roster.TABLEAUX.index(classement) if classement in roster.TABLEAUX else None

    resultat = {}
    total_diff_points = 0.0
    for tableau in ("Simple", "Double", "Mixte"):
        actuel_points = player["Points Actuel"][tableau]
        actuel_rang = rang_actuel.get(tableau)
        actuel_classement = player["Classement Actuel"][tableau]
        s = sept.get(tableau)  # {"rang","classement","points"} ou None

        diff_points = None
        if s is not None and s.get("points") is not None:
            diff_points = actuel_points - s["points"]

        tendance_hebdo = None
        if diff_points is not None and semaines_ecoulees is not None:
            tendance_hebdo = diff_points / semaines_ecoulees

        gain_places = None
        if s is not None and s.get("rang") is not None and actuel_rang is not None:
            gain_places = s["rang"] - actuel_rang

        gain_tableau = None
        i_sept = _index_tableau(s.get("classement")) if s else None
        i_actuel = _index_tableau(actuel_classement)
        if i_sept is not None and i_actuel is not None:
            gain_tableau = i_sept - i_actuel

        resultat[tableau] = {
            "septembre_classement": s.get("classement") if s else None,
            "septembre_rang": s.get("rang") if s else None,
            "septembre_points": s.get("points") if s else None,
            "actuel_classement": actuel_classement,
            "actuel_rang": actuel_rang,
            "actuel_points": actuel_points,
            "diff_points": diff_points,
            "tendance_hebdo": tendance_hebdo,
            "gain_places": gain_places,
            "gain_tableau": gain_tableau,
        }
        if diff_points is not None:
            total_diff_points += diff_points

    resultat["cumule"] = {
        "diff_points": total_diff_points,
        "tendance_hebdo": (total_diff_points / semaines_ecoulees) if semaines_ecoulees is not None else None,
    }
    return resultat


def indice_performance(stats):
    """victoires x taux de victoire : récompense le volume de victoires ET
    le taux de réussite en même temps (une valeur brute, normalisée plus
    tard en % du max observé dans le club)."""
    if stats["matchs_joues"] == 0:
        return 0.0
    return stats["victoires"] * (stats["victoires"] / stats["matchs_joues"])


def indice_global(indice_perf_normalise, indice_niveau_normalise, indice_qualite_normalise):
    """Moyenne de indice_performance, indice_niveau et indice_qualite, une
    fois tous les trois normalisés (0-100, % du club - voir
    normaliser_club) : un joueur à 100 est le meilleur du club sur les
    trois plans à la fois, 50 = dans la moyenne partout.

    indice_qualite (solde de points de cote gagnés/perdus sur la saison,
    déjà ajusté à la force des adversaires par le barème FFBad officiel -
    battre quelqu'un de mieux classé rapporte plus de points que battre
    quelqu'un de plus faible, et l'inverse pour une défaite) complète
    indice_performance, qui ne regarde que le nombre de victoires et le
    taux sans tenir compte de la force des adversaires - un joueur qui
    gagne beaucoup mais contre des adversaires plus faibles que lui ne doit
    pas dominer un joueur au taux de victoire supérieur mais aux victoires
    plus dures à obtenir.

    Une moyenne de valeurs déjà normalisées plutôt qu'un produit de valeurs
    brutes (l'ancienne approche, sans indice_qualite) : les trois indices
    bruts n'ont ni la même échelle ni la même distribution - les combiner
    directement ferait dominer celui qui a la plus grande variance brute
    plutôt que celui qui est réellement le meilleur sur les trois plans."""
    return (indice_perf_normalise + indice_niveau_normalise + indice_qualite_normalise) / 3


def normaliser(items, cle_brute, cle_normalisee, echelle=100):
    """Ajoute item[cle_normalisee] = item[cle_brute] en % du max observé
    dans `items`. Suppose cle_brute >= 0 (sinon utiliser
    normaliser_min_max). Renvoie ce maximum."""
    valeurs = [item[cle_brute] for item in items if item.get(cle_brute) is not None]
    maximum = max(valeurs) if valeurs else 0
    for item in items:
        brute = item.get(cle_brute)
        item[cle_normalisee] = (brute / maximum * echelle) if maximum and brute is not None else 0.0
    return maximum


def normaliser_min_max(items, cle_brute, cle_normalisee, echelle=100):
    """Comme normaliser(), mais ramène [min, max] observé sur [0, echelle]
    plutôt que value/max - nécessaire pour indice_qualite (un solde de
    points de cote qui peut être négatif : diviser par le max donnerait des
    valeurs négatives et casserait l'hypothèse "0 à 100" utilisée partout
    ailleurs). Si tout le monde a la même valeur (ex: tout le monde à 0 en
    tout début de saison), tout le monde reçoit 50 (ni pénalisé ni
    avantagé) plutôt qu'une division par zéro. Renvoie (min, max)."""
    valeurs = [item[cle_brute] for item in items if item.get(cle_brute) is not None]
    minimum = min(valeurs) if valeurs else 0
    maximum = max(valeurs) if valeurs else 0
    etendue = maximum - minimum
    for item in items:
        brute = item.get(cle_brute)
        if brute is None:
            item[cle_normalisee] = 0.0
        elif not etendue:
            item[cle_normalisee] = echelle / 2
        else:
            item[cle_normalisee] = (brute - minimum) / etendue * echelle
    return minimum, maximum


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


def _cote_saison_stats(journal_entries):
    """journal_entries : player["Journal cote"][tableau] (voir
    statsJoueurs.scrape_player via results.parse_journal_complet) -> min/max/
    moyenne de la cote observée cette saison (le journal de suivi repart à
    zéro à chaque "Initialisation saison", donc pas besoin de filtrer par
    date ici), + un indice de stabilité.

    stabilite = 1 - (amplitude / moyenne), borné à 0 : proche de 100% = cote
    qui a peu varié sur la saison, proche de 0% = grosses variations.
    Nécessite au moins 2 points de mesure (sinon amplitude=0 donnerait
    trivialement 100% dès la première mesure, ce qui ne voudrait rien dire -
    "stable" doit vouloir dire "observé stable dans le temps", pas "une
    seule observation")."""
    cotes = [e["cote"] for e in journal_entries if e.get("cote") is not None]
    if not cotes:
        return {"cote_min": None, "cote_max": None, "cote_moyenne": None,
                "nb_points_mesure": 0, "stabilite": None}
    cote_min, cote_max = min(cotes), max(cotes)
    cote_moyenne = sum(cotes) / len(cotes)
    stabilite = None
    if len(cotes) >= 2 and cote_moyenne:
        stabilite = max(0.0, 1 - (cote_max - cote_min) / cote_moyenne) * 100
    return {
        "cote_min": cote_min,
        "cote_max": cote_max,
        "cote_moyenne": cote_moyenne,
        "nb_points_mesure": len(cotes),
        "stabilite": stabilite,
    }


def evolution_mensuelle_cote(journal_entries):
    """journal_entries (voir _cote_saison_stats) -> dict {"YYYY-MM": somme}
    net de variation de cote par mois, pour visualiser une tendance mois par
    mois plutôt qu'un seul chiffre sur toute la saison.

    Le journal de suivi ne couvre qu'UNE saison (repart à zéro à chaque
    "Initialisation saison") et est trié du plus récent au plus ancien - la
    dernière entrée de la liste est donc toujours cette ligne
    d'initialisation, dont "points" est la cote de départ (pas un delta) :
    on l'exclut pour ne pas fausser le mois de septembre."""
    if len(journal_entries) < 2:
        return {}
    par_mois = defaultdict(float)
    for e in journal_entries[:-1]:  # tout sauf la ligne d'initialisation
        if e.get("points") is None:
            continue
        mois = e["date"].strftime("%Y-%m")
        par_mois[mois] += e["points"]
    return dict(par_mois)


def build_player_stats(player, events_par_tableau):
    """player : entrée roster (Nom/Sexe/Classement Actuel/Points Actuel),
    complétée par le "Meilleur tableau"/"Valeur IC" déjà calculés
    (pointsIC.compute_IC_points). events_par_tableau : voir build_tournois.

    Renvoie un dict prêt à être normalisé au niveau du club puis écrit dans
    l'Excel."""
    journal = player.get("Journal cote") or {}
    par_tableau = {}
    for tableau in ("Simple", "Double", "Mixte"):
        events = events_par_tableau.get(tableau, [])
        stats = discipline_stats(events)
        classement = player["Classement Actuel"][tableau]
        points = {"Simple": 0.0, "Double": 0.0, "Mixte": 0.0}
        points[tableau] = player["Points Actuel"][tableau]
        niveau = roster.valeur_ic(classement, player["Sexe"],
                                   points["Simple"], points["Double"], points["Mixte"])
        entry = {
            **stats,
            "classement": classement,
            "cote": player["Points Actuel"][tableau],
            "indice_niveau_brut": niveau,
            "indice_performance_brut": indice_performance(stats),
            "indice_qualite_brut": stats["points_cote_total"],
            "cote_saison": _cote_saison_stats(journal.get(tableau, [])),
            "evolution_mensuelle_cote": evolution_mensuelle_cote(journal.get(tableau, [])),
        }
        if tableau in ("Double", "Mixte"):
            entry["partenaire"] = discipline_stats_partenaire(events)
            entry["meilleur_partenaire"] = meilleur_partenaire(events)
            entry["meilleur_partenaire_club"] = meilleur_partenaire_club_si_different(
                entry["meilleur_partenaire"], meilleur_partenaire(events, club_uniquement=True))
        par_tableau[tableau] = entry

    partenaire_double_mixte = discipline_stats_partenaire_combinee(
        events_par_tableau.get("Double", []), events_par_tableau.get("Mixte", []))
    meilleur_partenaire_dm = meilleur_partenaire_combine(
        events_par_tableau.get("Double", []), events_par_tableau.get("Mixte", []))
    meilleur_partenaire_dm_club = meilleur_partenaire_club_si_different(
        meilleur_partenaire_dm,
        meilleur_partenaire_combine(events_par_tableau.get("Double", []),
                                     events_par_tableau.get("Mixte", []), club_uniquement=True))

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
        "meilleur_partenaire_double_mixte": meilleur_partenaire_dm,
        "meilleur_partenaire_double_mixte_club": meilleur_partenaire_dm_club,
        # "ordre_disciplines" pas rempli ici : ordre_disciplines() se base
        # sur indice_global, qui n'existe qu'après normaliser_club (voir
        # plus bas) - rempli par normaliser_club, pas par build_player_stats.
        "progression": diff_classement(player),
        "global": {
            **stats_globales,
            "meilleur_tableau": player["Meilleur tableau"],
            "rang_meilleur_tableau": player.get("Rang meilleur tableau"),
            "points_meilleur_tableau": player["Points meilleur tableau"],
            "indice_niveau_brut": niveau_global,
            "indice_performance_brut": perf_globale,
            "indice_qualite_brut": stats_globales["points_cote_total"],
        },
        "tournois": tournois,
        "match_log": match_log,
    }


def normaliser_club(all_stats):
    """Normalise indice_performance, indice_niveau et indice_qualite en %
    (ou position min-max pour indice_qualite, qui peut être négatif - voir
    normaliser_min_max) du club (une fois tous les joueurs traités), par
    tableau et au global, PUIS calcule indice_global à partir des trois
    valeurs déjà normalisées (voir indice_global - une moyenne n'a de sens
    que si les trois côtés sont sur la même échelle). Modifie all_stats en
    place et renvoie les bornes brutes observées."""
    bornes = {}
    for tableau in ("Simple", "Double", "Mixte"):
        entries = [s["par_tableau"][tableau] for s in all_stats]
        bornes[f"{tableau} - performance"] = normaliser(entries, "indice_performance_brut", "indice_performance")
        bornes[f"{tableau} - niveau"] = normaliser(entries, "indice_niveau_brut", "indice_niveau")
        bornes[f"{tableau} - qualite"] = normaliser_min_max(entries, "indice_qualite_brut", "indice_qualite")
        for entry in entries:
            entry["indice_global"] = indice_global(
                entry["indice_performance"], entry["indice_niveau"], entry["indice_qualite"])

    # les 3 indice_global par tableau sont prêts : c'est seulement
    # maintenant qu'on peut classer les 3 tableaux d'un joueur entre eux.
    for s in all_stats:
        s["ordre_disciplines"] = ordre_disciplines(s["par_tableau"])

    globaux = [s["global"] for s in all_stats]
    bornes["Global - performance"] = normaliser(globaux, "indice_performance_brut", "indice_performance")
    bornes["Global - niveau"] = normaliser(globaux, "indice_niveau_brut", "indice_niveau")
    bornes["Global - qualite"] = normaliser_min_max(globaux, "indice_qualite_brut", "indice_qualite")
    for g in globaux:
        g["indice_global"] = indice_global(g["indice_performance"], g["indice_niveau"], g["indice_qualite"])
    return bornes
