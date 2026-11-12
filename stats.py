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
        "nb_matchs_cotes": len(points),
    }


def discipline_stats(events):
    """events d'un seul tableau (Simple/Double/Mixte) -> stats agrégées
    (matchs joués, victoires, % victoire)."""
    return _agg_matchs([m for e in events for m in e["matchs"]])


def cote_adverse_moyenne(matchs):
    """Cote moyenne des adversaires affrontés (force du calendrier) - un
    joueur qui croise surtout des adversaires forts peut avoir un
    points_cote_total plus faible qu'un autre qui en croise peu, sans que
    ça veuille dire qu'il performe moins bien à niveau d'adversaire égal."""
    valeurs = [m["adversaire_cote"] for m in matchs if m["adversaire_cote"] is not None]
    return sum(valeurs) / len(valeurs) if valeurs else None


def _est_forfait(m):
    """Un forfait (adversaire absent) rapporte 0 point de cote - le barème
    FFBad ne note pas un résultat qui n'a pas été réellement joué, même si
    le site affiche un score générique du genre 21-0/21-0. À exclure des
    stats basées sur le score (meilleure victoire/pire défaite, profil de
    score) : un 21-0 par forfait n'est pas un "exploit".

    Le score générique sert aussi de signal, pas seulement les points de
    cote : quand la cote du match est inconnue (points_cote None), le test
    "== 0" seul laissait passer le forfait - constaté sur un "Score max
    infligé" à 21-0 en mixte, impossible en vrai depuis le passage aux sets
    de 15 points (et quasi inexistant avant)."""
    if m["points_cote"] == 0:
        return True
    sets = m.get("sets_detail") or []
    return bool(sets) and all(sorted(s) == [0, 21] for s in sets)


def _fait_marquant(m):
    return {
        "adversaire_nom": " / ".join(m["noms_adverses"]),
        "adversaire_cote": m["adversaire_cote"],
        "ma_cote": m["ma_cote"],
        "score": " / ".join(f"{mon}-{son}" for mon, son in m["sets_detail"]),
        "points_cote": m["points_cote"],
    }


def meilleure_victoire(matchs):
    """Parmi les matchs gagnés (hors forfaits), celui contre l'adversaire à
    la cote la plus haute - "l'exploit" de la saison. None si aucun match
    gagné avec une cote adverse connue."""
    candidats = [m for m in matchs
                 if m["victoire"] and m["adversaire_cote"] is not None and not _est_forfait(m)]
    if not candidats:
        return None
    return _fait_marquant(max(candidats, key=lambda m: m["adversaire_cote"]))


def pire_defaite(matchs):
    """Parmi les matchs perdus (hors forfaits), celui contre l'adversaire à
    la cote la plus basse."""
    candidats = [m for m in matchs
                 if not m["victoire"] and m["adversaire_cote"] is not None and not _est_forfait(m)]
    if not candidats:
        return None
    return _fait_marquant(min(candidats, key=lambda m: m["adversaire_cote"]))


def plus_grosse_perf(matchs):
    """Le match qui a rapporté le plus de points de cote (forfaits exclus)
    - la plus grosse perf de la saison en termes de gain de classement, pas
    forcément contre l'adversaire le mieux coté (voir meilleure_victoire
    pour ça - un match qui rapporte gros peut être contre un adversaire
    "seulement" un peu mieux coté mais avec un écart de score qui compte
    plus dans le barème que prévu, ou l'inverse)."""
    candidats = [m for m in matchs
                 if m["victoire"] and m["points_cote"] is not None and not _est_forfait(m)]
    if not candidats:
        return None  # jamais gagné : pas de "perf" à afficher (sinon la défaite la moins coûteuse ressortait)
    return _fait_marquant(max(candidats, key=lambda m: m["points_cote"]))


def plus_grosse_defaite_points(matchs):
    """Le match qui a coûté le plus de points de cote (forfaits exclus),
    parmi les défaites uniquement - None si le joueur n'a jamais perdu
    (sinon la victoire la moins rentable ressortait comme "contre-perf")."""
    candidats = [m for m in matchs
                 if not m["victoire"] and m["points_cote"] is not None and not _est_forfait(m)]
    if not candidats:
        return None
    return _fait_marquant(min(candidats, key=lambda m: m["points_cote"]))


def plus_grande_difference_cote(matchs):
    """Le match avec le plus grand écart de cote entre les deux joueurs,
    peu importe le résultat - la confrontation la plus déséquilibrée sur le
    papier de la saison (forfaits exclus, l'écart n'y reflète rien de
    réel : l'adversaire absent n'a jamais vraiment affronté ce niveau).

    ecart_cote est SIGNÉ (ma_cote - adversaire_cote), pas une valeur
    absolue : négatif = j'étais l'outsider dans ce match, positif = le
    favori - le signe dit directement dans quel sens penchait le papier,
    pas seulement l'ampleur de l'écart."""
    candidats = [m for m in matchs
                 if m["ma_cote"] is not None and m["adversaire_cote"] is not None and not _est_forfait(m)]
    if not candidats:
        return None
    m = max(candidats, key=lambda m: abs(m["ma_cote"] - m["adversaire_cote"]))
    resultat = _fait_marquant(m)
    resultat["ecart_cote"] = m["ma_cote"] - m["adversaire_cote"]
    resultat["victoire"] = m["victoire"]
    return resultat


def david_vs_goliath(matchs):
    """Parmi les matchs où l'adversaire avait une meilleure cote (outsider,
    "David"), % de victoires ; parmi ceux où on avait la meilleure cote
    (favori, "Goliath"), % de victoires. Forfaits exclus (le barème
    n'ayant pas tourné, l'écart de cote n'y a jamais été mis à l'épreuve)."""
    utiles = [m for m in matchs
              if m["ma_cote"] is not None and m["adversaire_cote"] is not None and not _est_forfait(m)]
    outsider = [m for m in utiles if m["adversaire_cote"] > m["ma_cote"]]
    favori = [m for m in utiles if m["ma_cote"] > m["adversaire_cote"]]
    return {
        "nb_outsider": len(outsider),
        "taux_outsider": (sum(1 for m in outsider if m["victoire"]) / len(outsider) * 100) if outsider else None,
        "nb_favori": len(favori),
        "taux_favori": (sum(1 for m in favori if m["victoire"]) / len(favori) * 100) if favori else None,
    }


SEUIL_SETS_SERRES = 3

# nombre minimum de matchs à la base d'un taux de dynamique (taux propre,
# après 1er set gagné, comeback) pour qu'il soit affiché : en dessous, un
# 0% ou 100% sur 1-2 matchs n'apprend rien et attire l'oeil pour rien.
SEUIL_DYNAMIQUE = 3


def indice_clutch(matchs):
    """Le joueur performe-t-il mieux ou moins bien que sa moyenne quand le
    set est serré (écart final <= 2 points - le seul écart possible en fin
    de set, quelle que soit la règle en vigueur : victoire "normale" par 2
    points, ou par 1 seul point au plafond de prolongation - 30-29 avant le
    1er septembre 2026, 21-20 depuis le passage aux sets de 15 points) ?

    delta_clutch = % de sets serrés gagnés - % de sets gagnés au global :
    positif = monte en niveau dans les moments chauds, négatif = craque un
    peu plus que d'habitude. None (comme taux_clutch) si moins de
    SEUIL_SETS_SERRES sets serrés joués - trop peu pour être lisible."""
    tous_sets = [(mon, son) for m in matchs for mon, son in m["sets_detail"]]
    serres = [(mon, son) for mon, son in tous_sets if abs(mon - son) <= 2]
    resultat = {
        "sets_serres_joues": len(serres),
        "sets_serres_gagnes": sum(1 for mon, son in serres if mon > son),
        "taux_clutch": None,
        "delta_clutch": None,
    }
    if len(serres) < SEUIL_SETS_SERRES or not tous_sets:
        return resultat
    taux_global = sum(1 for mon, son in tous_sets if mon > son) / len(tous_sets) * 100
    taux_clutch = resultat["sets_serres_gagnes"] / len(serres) * 100
    resultat["taux_clutch"] = taux_clutch
    resultat["delta_clutch"] = taux_clutch - taux_global
    return resultat


def points_cote_moyens(matchs):
    """Moyenne des points de cote gagnés sur les matchs remportés, et
    perdus sur les matchs perdus - deux moyennes séparées (les mélanger
    donnerait juste points_cote_total/nb, qui dit moins sur l'ampleur
    typique d'un gain ou d'une perte). Sert notamment à comparer
    l'interclub au reste : les matchs y rapportent/coûtent-ils plus de
    points en moyenne (barème FFBad différent en compétition par équipe) ?"""
    gagnes = [m["points_cote"] for m in matchs if m["victoire"] and m["points_cote"] is not None]
    perdus = [m["points_cote"] for m in matchs if not m["victoire"] and m["points_cote"] is not None]
    return {
        "moyenne_gagne": sum(gagnes) / len(gagnes) if gagnes else None,
        "moyenne_perdu": sum(perdus) / len(perdus) if perdus else None,
        "nb_gagnes": len(gagnes),
        "nb_perdus": len(perdus),
    }


def profil_score(matchs):
    """Profil de score set par set (mon score, celui de l'adversaire),
    forfaits exclus (voir _est_forfait - leur score 21-0/21-0 générique
    fausserait score_max_inflige en particulier) : score moyen des deux
    côtés, plus gros écart infligé (set gagné avec la plus grande marge) et
    reçu (set perdu avec la plus grande marge), et points marqués en
    moyenne selon que le set est gagné ou perdu - un set gagné à l'arraché
    (22-20) et un autre plié (21-5) comptent pareil pour le nombre de sets,
    mais pas du tout pareil ici."""
    tous_sets = [(mon, son) for m in matchs if not _est_forfait(m) for mon, son in m["sets_detail"]]
    if not tous_sets:
        return {
            "score_moyen_mien": None, "score_moyen_adverse": None,
            "score_max_inflige": None, "score_max_recu": None,
            "points_moyens_victoire": None, "points_moyens_defaite": None,
        }
    gagnes = [(mon, son) for mon, son in tous_sets if mon > son]
    perdus = [(mon, son) for mon, son in tous_sets if son > mon]
    max_inflige = max(gagnes, key=lambda s: s[0] - s[1]) if gagnes else None
    max_recu = max(perdus, key=lambda s: s[1] - s[0]) if perdus else None
    return {
        "score_moyen_mien": sum(mon for mon, son in tous_sets) / len(tous_sets),
        "score_moyen_adverse": sum(son for mon, son in tous_sets) / len(tous_sets),
        "score_max_inflige": f"{max_inflige[0]}-{max_inflige[1]}" if max_inflige else None,
        "score_max_recu": f"{max_recu[0]}-{max_recu[1]}" if max_recu else None,
        "points_moyens_victoire": sum(mon for mon, son in gagnes) / len(gagnes) if gagnes else None,
        "points_moyens_defaite": sum(mon for mon, son in perdus) / len(perdus) if perdus else None,
    }


def sets_extremes(matchs):
    """Sets joués au plafond de prolongation exact (écart final d'1 seul
    point - la seule façon d'avoir un écart de 1, quelle que soit la règle
    en vigueur : 30-29 avant le 1er septembre 2026, 21-20 depuis le passage
    aux sets de 15 points) - les plus tendus qui puissent exister. Nombre
    joués/gagnés, et le détail des scores pour référence."""
    extremes = [(mon, son) for m in matchs for mon, son in m["sets_detail"] if abs(mon - son) == 1]
    return {
        "nb_joues": len(extremes),
        "nb_gagnes": sum(1 for mon, son in extremes if mon > son),
        "scores": [f"{mon}-{son}" for mon, son in extremes],
    }


def dynamique_match(matchs):
    """Comment tourne un match selon ce qui s'est passé au 1er set (matchs à
    au moins 2 sets, forfaits exclus - voir _est_forfait, leur schéma de
    sets ne dit rien d'une vraie dynamique de match) :
    - taux_propre : parmi les matchs GAGNÉS, % gagnés "propre" (tous les
      sets remportés, sans en concéder un) plutôt qu'accrochés à 3 sets.
    - taux_apres_premier_set : parmi les matchs où le 1er set est gagné, %
      de matchs remportés au final - "ferme-t-on derrière" ?
    - taux_comeback : parmi les matchs où le 1er set est PERDU, % de
      matchs gagnés quand même - capacité à revenir de 0-1."""
    matchs_utiles = [m for m in matchs if len(m["sets_detail"]) >= 2 and not _est_forfait(m)]
    gagnes = [m for m in matchs_utiles if m["victoire"]]
    propres = [m for m in gagnes if all(mon > son for mon, son in m["sets_detail"])]
    premier_gagne = [m for m in matchs_utiles if m["sets_detail"][0][0] > m["sets_detail"][0][1]]
    premier_perdu = [m for m in matchs_utiles if m["sets_detail"][0][1] > m["sets_detail"][0][0]]

    def _taux(nb, base):
        # None sous SEUIL_DYNAMIQUE matchs de base (voir sa définition)
        return (nb / len(base) * 100) if len(base) >= SEUIL_DYNAMIQUE else None

    return {
        "nb_gagnes": len(gagnes),
        "nb_propres": len(propres),
        "taux_propre": _taux(len(propres), gagnes),
        "nb_premier_set_gagne": len(premier_gagne),
        "taux_apres_premier_set": _taux(sum(1 for m in premier_gagne if m["victoire"]), premier_gagne),
        "nb_premier_set_perdu": len(premier_perdu),
        "taux_comeback": _taux(sum(1 for m in premier_perdu if m["victoire"]), premier_perdu),
    }


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
    taux de victoire lissé, total des points de cote gagnés/perdus avec ce
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
            "taux_lisse": indice_performance(agg),
            "points_cote_total": agg["points_cote_total"],
        }
    return resultat


def _meilleur_partenaire_parmi(matchs, club_uniquement=False):
    """Le partenaire qui ressort en tête, classé D'ABORD par points de cote
    cumulés (déjà ajustés à la force de l'adversaire par le barème FFBad -
    voir _agg_matchs), et à égalité par le taux de victoire lissé (voir
    indice_performance) comme départage.

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
    return max(partenaires.values(), key=lambda v: (v["points_cote_total"], v["taux_lisse"]))


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
    """Les tableaux JOUÉS cette saison (au moins un match), triés du plus
    fort au moins fort pour ce joueur selon indice_forme (déjà calculé
    club-wide, voir normaliser_club - à appeler APRÈS normaliser_club).

    Les tableaux sans aucun match sont omis plutôt que classés : un
    tableau non joué n'a pas d'indice de forme (None), et l'ancienne
    version lui en attribuait un quand même (niveau + qualité par défaut),
    ce qui faisait sortir en tête un tableau jamais joué - constaté sur
    plusieurs joueurs ("Simple > ..." avec 0 match en simple)."""
    tableaux = [t for t in ("Simple", "Double", "Mixte") if par_tableau[t]["indice_forme"] is not None]
    return sorted(tableaux, key=lambda t: -par_tableau[t]["indice_forme"])


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


# lissage des indices : on fait comme si chaque joueur avait joué
# LISSAGE_MATCHS matchs "fictifs" en plus, à 50% de victoire et 0 point de
# cote. Effet : sur 2 matchs, un 2/2 ne vaut pas plus qu'un 12/15 (le peu
# de matchs tire vers la moyenne), alors qu'au-delà de ~15 matchs le
# lissage ne pèse presque plus. Remplace l'ancien "victoires x taux", qui
# mesurait surtout le volume (corrélé à 0,9 au nombre de matchs).
LISSAGE_MATCHS = 5


def indice_performance(stats):
    """Taux de victoire lissé, en % (voir LISSAGE_MATCHS) - valeur brute,
    convertie plus tard en rang percentile dans le club. None si aucun
    match joué (pas de performance à mesurer, plutôt qu'un 0 qui
    classerait le joueur comme "le plus faible")."""
    if stats["matchs_joues"] == 0:
        return None
    return (stats["victoires"] + LISSAGE_MATCHS * 0.5) / (stats["matchs_joues"] + LISSAGE_MATCHS) * 100


def indice_qualite(stats):
    """Points de cote gagnés/perdus PAR MATCH, lissés (voir
    LISSAGE_MATCHS) - valeur brute, convertie plus tard en rang
    percentile. Par match et non plus en total : le total (solde de cote
    de la saison) faisait doublon avec "Diff cote" (corrélé à 0,93) et
    dépendait lui aussi surtout du nombre de matchs joués. Les points de
    cote sont déjà ajustés à la force de l'adversaire par le barème FFBad
    (battre plus fort rapporte plus, perdre contre plus faible coûte plus) :
    c'est le signal "qualité des résultats". None si aucun match."""
    if stats["matchs_joues"] == 0:
        return None
    return stats["points_cote_total"] / (stats["nb_matchs_cotes"] + LISSAGE_MATCHS)


def indice_forme(indice_perf, indice_qual):
    """Moyenne des rangs percentile de indice_performance et
    indice_qualite : 50 = forme médiane du groupe, 100 = meilleur du groupe
    sur les deux plans. indice_niveau (le classement) n'en fait plus
    partie : c'est un état, pas une forme - le mélanger faisait remonter un
    joueur bien classé même sans résultats (voire sans aucun match). None
    si l'un des deux manque (aucun match)."""
    if indice_perf is None or indice_qual is None:
        return None
    return (indice_perf + indice_qual) / 2


def rang_percentile(items, cle_brute, cle_normalisee):
    """item[cle_normalisee] = rang percentile (0-100) de item[cle_brute]
    parmi les items qui ont une valeur (None -> None) : 0 = le plus bas du
    groupe, 100 = le plus haut, 50 = la médiane, ex aequo au rang moyen.
    Plutôt que "% du max" (l'ancienne normalisation) : un seul joueur
    exceptionnel écrasait toute l'échelle vers le bas, et "50" ne voulait
    pas dire "dans la moyenne"."""
    valeurs = sorted(item[cle_brute] for item in items if item.get(cle_brute) is not None)
    n = len(valeurs)
    for item in items:
        brute = item.get(cle_brute)
        if brute is None:
            item[cle_normalisee] = None
        elif n == 1:
            item[cle_normalisee] = 50.0
        else:
            dessous = sum(1 for v in valeurs if v < brute)
            egaux = sum(1 for v in valeurs if v == brute)
            item[cle_normalisee] = (dessous + (egaux - 1) / 2) / (n - 1) * 100


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
    date ici), + l'amplitude (max - min, en points).

    Amplitude en points plutôt que l'ancienne "stabilité" (1 - amplitude /
    moyenne) : rapportée à une cote de ~2000, l'amplitude d'une saison
    donnait 87-100% pour tout le club, sans rien distinguer. Nécessite au
    moins 2 points de mesure (une seule observation donnerait trivialement
    0, ce qui ne voudrait rien dire)."""
    cotes = [e["cote"] for e in journal_entries if e.get("cote") is not None]
    if not cotes:
        return {"cote_min": None, "cote_max": None, "cote_moyenne": None,
                "nb_points_mesure": 0, "amplitude": None}
    cote_min, cote_max = min(cotes), max(cotes)
    cote_moyenne = sum(cotes) / len(cotes)
    return {
        "cote_min": cote_min,
        "cote_max": cote_max,
        "cote_moyenne": cote_moyenne,
        "nb_points_mesure": len(cotes),
        "amplitude": (cote_max - cote_min) if len(cotes) >= 2 else None,
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


def journal_depuis_matchs(cote_septembre, date_septembre, matchs_avec_date):
    """Reconstitue un journal de cote (même format que results.
    parse_journal_complet : liste de {date, points, cote}, du plus récent au
    plus ancien, terminée par la ligne d'initialisation de saison) à partir
    de la cote du 1er septembre + les points de cote de chaque match, cumulés
    jour par jour. Filet de sécurité quand le vrai journal de suivi n'a pas
    pu être lu (constaté : vide pour tout le club sur un run complet) - sans
    lui, Cote min/max/moyenne, Stabilité, Diff cote mensuelle et Pics/chutes
    restaient entièrement vides. Moins exact que le journal (un match dont
    les points ne sont pas encore publiés compte pour 0), mais suffisant pour
    ces stats. [] si la cote de septembre est inconnue.

    matchs_avec_date : liste de (date, match)."""
    if cote_septembre is None or date_septembre is None:
        return []
    par_jour = defaultdict(float)
    for d, m in matchs_avec_date:
        par_jour[d] += m["points_cote"] or 0.0
    entries = [{"date": date_septembre, "points": cote_septembre, "cote": cote_septembre}]
    cote = cote_septembre
    for d in sorted(par_jour):
        cote += par_jour[d]
        entries.append({"date": d, "points": par_jour[d], "cote": cote})
    return entries[::-1]


def build_player_stats(player, events_par_tableau):
    """player : entrée roster (Nom/Sexe/Classement Actuel/Points Actuel),
    complétée par le "Meilleur tableau"/"Valeur IC" déjà calculés
    (pointsIC.compute_IC_points). events_par_tableau : voir build_tournois.

    Renvoie un dict prêt à être normalisé au niveau du club puis écrit dans
    l'Excel."""
    journal = dict(player.get("Journal cote") or {})
    sept = player.get("Classement 1er septembre") or {}
    par_tableau = {}
    for tableau in ("Simple", "Double", "Mixte"):
        events = events_par_tableau.get(tableau, [])
        matchs = [m for e in events for m in e["matchs"]]
        if not journal.get(tableau):
            journal[tableau] = journal_depuis_matchs(
                (sept.get(tableau) or {}).get("points"), sept.get("date"),
                [(parse_date_fr(e["date"]), m) for e in events for m in e["matchs"]])
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
            "indice_qualite_brut": indice_qualite(stats),
            "cote_saison": _cote_saison_stats(journal.get(tableau, [])),
            "evolution_mensuelle_cote": evolution_mensuelle_cote(journal.get(tableau, [])),
            # (date, cote) du plus ancien au plus récent - courbe de cote
            # semaine par semaine de la page web (voir export_web)
            "courbe_cote": sorted((e["date"], e["cote"]) for e in journal.get(tableau, [])
                                  if e.get("cote") is not None and e.get("date") is not None),
            "cote_adverse_moyenne": cote_adverse_moyenne(matchs),
            "meilleure_victoire": meilleure_victoire(matchs),
            "pire_defaite": pire_defaite(matchs),
            "plus_grosse_perf": plus_grosse_perf(matchs),
            "plus_grosse_defaite_points": plus_grosse_defaite_points(matchs),
            "plus_grande_difference_cote": plus_grande_difference_cote(matchs),
            "clutch": indice_clutch(matchs),
            "profil_score": profil_score(matchs),
            "sets_extremes": sets_extremes(matchs),
            "dynamique": dynamique_match(matchs),
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

    # un enregistrement par match (tableau, mois, date, evenement, victoire,
    # sets_detail, est_interclub, intra_club, points_cote, noms_adverses),
    # pour la ventilation club-wide (onglet Stats club) et les stats
    # "avancées" - inutile de stocker le sexe/nom ici, on les a déjà au
    # niveau du joueur (voir all_stats côté export_excel). sets_detail
    # permet de rappeler indice_clutch()/dynamique_match() club-wide sans
    # ré-analyser tous les matchs bruts. intra_club = adversaire(s) 100% du
    # club (GAB38 vs GAB38, un joueur du club contre un autre en tournoi
    # individuel) - toujours 1 victoire + 1 défaite pour le club, dilue le %
    # club vers 50% sans rien dire de sa performance face à l'extérieur (mais
    # sert justement à repérer les rivalités internes). evenement + date
    # identifient un tournoi (hors interclubs, où evenement se répète sur
    # toute la saison pour des rencontres différentes - date seule suffit à
    # les distinguer, mais pas les tournois individuels : deux tournois
    # différents pourraient tomber le même jour).
    match_log = [
        {
            "tableau": tableau,
            "mois": parse_date_fr(e["date"]).strftime("%Y-%m"),
            "date": parse_date_fr(e["date"]),
            "evenement": e["evenement"],
            # numéro de l'équipe du club en interclub ("38-GAB-3 contre
            # ..." -> 3), None en tournoi individuel
            "equipe": results.equipe_club(e.get("tableau_ou_rencontre")) if e["est_interclub"] else None,
            "victoire": m["victoire"],
            "sets_detail": m["sets_detail"],
            "est_interclub": e["est_interclub"],
            "intra_club": bool(m["clubs_adverses"]) and all(c == results.MY_CLUB for c in m["clubs_adverses"]),
            "points_cote": m["points_cote"],
            "noms_adverses": m["noms_adverses"],
            "partenaire_nom": m.get("partenaire_nom"),
            "forfait": _est_forfait(m),
            "ma_cote": m["ma_cote"],
            "adversaire_cote": m["adversaire_cote"],
        }
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
        # sur indice_forme, qui n'existe qu'après normaliser_club (voir
        # plus bas) - rempli par normaliser_club, pas par build_player_stats.
        "progression": diff_classement(player),
        "global": {
            **stats_globales,
            "meilleur_tableau": player["Meilleur tableau"],
            "rang_meilleur_tableau": player.get("Rang meilleur tableau"),
            "points_meilleur_tableau": player["Points meilleur tableau"],
            "indice_niveau_brut": niveau_global,
            "indice_performance_brut": perf_globale,
            "indice_qualite_brut": indice_qualite(stats_globales),
        },
        "tournois": tournois,
        "match_log": match_log,
    }


def _groupes_comparaison(all_stats, tableau):
    """Les joueurs comparés entre eux pour un tableau : même découpage que
    les onglets (SH/SD, DH/DD, MX) - Simple et Double par genre, Mixte tous
    ensemble ; tableau=None pour le global (par genre, comme Simple/Double)."""
    if tableau == "Mixte":
        return [all_stats]
    return [[s for s in all_stats if s["Sexe"] == sexe] for sexe in ("H", "F")]


def _normaliser_groupe(entries):
    rang_percentile(entries, "indice_performance_brut", "indice_performance")
    rang_percentile(entries, "indice_qualite_brut", "indice_qualite")
    rang_percentile(entries, "indice_niveau_brut", "indice_niveau")
    for entry in entries:
        entry["indice_forme"] = indice_forme(entry["indice_performance"], entry["indice_qualite"])


def normaliser_club(all_stats):
    """Convertit indice_performance, indice_qualite et indice_niveau en
    rang percentile (voir rang_percentile) au sein de chaque groupe de
    comparaison (voir _groupes_comparaison), par tableau et au global,
    PUIS calcule indice_forme et l'ordre des tableaux de chaque joueur.
    Les joueurs sans match dans un tableau sont exclus de la comparaison
    de ce tableau (indices None), y compris indice_niveau - sinon un
    tableau jamais joué recevait des indices comme s'il l'avait été.
    Modifie all_stats en place."""
    for tableau in ("Simple", "Double", "Mixte"):
        for groupe in _groupes_comparaison(all_stats, tableau):
            entries = [s["par_tableau"][tableau] for s in groupe]
            for entry in entries:
                if entry["matchs_joues"] == 0:
                    entry["indice_niveau_brut"] = None
            _normaliser_groupe(entries)

    for s in all_stats:
        s["ordre_disciplines"] = ordre_disciplines(s["par_tableau"])

    for groupe in _groupes_comparaison(all_stats, None):
        globaux = [s["global"] for s in groupe]
        for g in globaux:
            if g["matchs_joues"] == 0:
                g["indice_niveau_brut"] = None
        _normaliser_groupe(globaux)
