"""Export web des stats joueurs : un fichier HTML autonome (données
incluses, aucun serveur ni dépendance réseau - s'ouvre directement depuis
la release GitHub ou en local), avec une page par joueur et une page club,
à côté de l'Excel. Plus un dump JSON complet de all_stats, qui sert aussi
de cache : `python statsJoueurs.py --from-json statsJoueurs.json`
régénère Excel + HTML sans re-scraper (voir charger_stats_json).

Le gabarit est web/statsJoueurs.template.html ; les données y sont
injectées à la place de DATA_PLACEHOLDER."""
import json
import os
import re

import requests
from collections import defaultdict
from datetime import date, datetime, timezone

import export_excel
import results
import roster
import stats

APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(APP_DIR, "web", "statsJoueurs.template.html")
DATA_PLACEHOLDER = "/*__DATA__*/null"
TABLEAUX = ("Simple", "Double", "Mixte")
# lien fixe vers l'Excel de la dernière release hebdo (même principe que
# pointsIC.xlsx dans le README)
EXCEL_URL = "https://github.com/ClementMirabel/PointsICGAB/releases/download/weekly-data/statsJoueurs.xlsx"
# page publiée (GitHub Pages, voir le workflow) : on y lit le résumé de
# l'export précédent pour afficher les écarts d'une semaine sur l'autre
PAGES_URL = "https://clementmirabel.github.io/PointsICGAB/"
RESUME_NOM = "statsJoueurs_resume.json"
# résumé de l'export précédent utilisé par cet export : conservé avec le dump,
# pour qu'une page régénérée plus tard (--from-json) compare toujours à la
# même référence, et non à son propre résumé
PRECEDENT_NOM = "statsJoueurs_precedent.json"

_DATE_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _json_default(obj):
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    if isinstance(obj, (set, frozenset)):
        return sorted(obj)
    raise TypeError(f"non sérialisable : {type(obj).__name__}")


# ---------------------------------------------------------------- dump / cache

def write_stats_json(all_stats, output_path, interclubs=None, effectif=None):
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump({"genere_le": datetime.now().isoformat(timespec="minutes"), "all_stats": all_stats,
                   "interclubs": interclubs, "effectif": effectif},
                  f, ensure_ascii=False, default=_json_default)
    return output_path


def _reviver(obj):
    """Les dates ont été écrites en ISO "AAAA-MM-JJ" : on les retransforme
    en date (les exports s'en servent - tri, regroupement de jours
    consécutifs, type de jour). "AAAA-MM" (clés de mois) reste une chaîne."""
    if isinstance(obj, dict):
        return {k: _reviver(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_reviver(v) for v in obj]
    if isinstance(obj, str) and _DATE_ISO.match(obj):
        return date.fromisoformat(obj)
    return obj


def charger_stats_json(path):
    return charger_dump(path)[0]


def charger_dump(path):
    """(all_stats, interclubs, effectif) - interclubs/effectif None pour un
    dump antérieur à leur ajout. Les données interclubs (icbad) et l'effectif
    gardent leurs dates en chaînes ISO : seule la page web les lit."""
    with open(path, encoding="utf-8") as f:
        dump = json.load(f)
    return _reviver(dump["all_stats"]), dump.get("interclubs"), dump.get("effectif")


def effectif_payload(players):
    """Tous les licenciés du club (roster.build_players + pointsIC, avant le
    filtre R6 du scraping) pour le planificateur IC : classements et cotes
    S/D/M, meilleur tableau, valeur ICN (équipe 1) et ICR (équipes 2+)."""
    out = []
    for p in players:
        pts, clt = p.get("Points Actuel") or {}, p.get("Classement Actuel") or {}
        out.append({
            "licence": p["Licence"], "nom": p["Nom"], "sexe": p["Sexe"], "prive": bool(p.get("Prive")),
            "categorie": p.get("Categorie"),
            "classements": [clt.get(t) for t in TABLEAUX], "cotes": [_r(pts.get(t), 0) for t in TABLEAUX],
            "meilleur_tableau": p.get("Meilleur tableau"),
            "icn": p.get("Valeur IC"),
            "icr": roster.valeur_icr(p.get("Meilleur tableau")),
        })
    return sorted(out, key=lambda j: j["nom"])


# ---------------------------------------------------------------- payload page

def _r(value, nd=1):
    return round(value, nd) if isinstance(value, (int, float)) else value


def _fait(fm, cle_valeur):
    if not fm:
        return None
    return {"adversaire": fm["adversaire_nom"], "valeur": _r(fm.get(cle_valeur)), "score": fm["score"]}


def _ecart(fm):
    if not fm:
        return None
    return {"adversaire": fm["adversaire_nom"], "valeur": _r(fm["ecart_cote"]),
            "victoire": fm["victoire"], "score": fm["score"]}


def _partenaires(match_log):
    """Tous les partenaires (Double + Mixte), pas seulement le meilleur :
    matchs, victoires, points de cote cumulés ensemble."""
    par_nom = defaultdict(lambda: {"matchs": 0, "victoires": 0, "points": 0.0, "tableaux": set()})
    for m in match_log:
        nom = m.get("partenaire_nom")
        if m["tableau"] == "Simple" or not nom:
            continue
        p = par_nom[nom]
        p["matchs"] += 1
        p["victoires"] += 1 if m["victoire"] else 0
        p["points"] += m["points_cote"] or 0.0
        p["tableaux"].add(m["tableau"])
    return sorted(
        ({"nom": nom, "matchs": p["matchs"], "victoires": p["victoires"], "points": _r(p["points"]),
          "tableaux": sorted(p["tableaux"])} for nom, p in par_nom.items()),
        key=lambda p: (-p["matchs"], -p["points"]))


def _dynamique(dyn):
    return {"propre": _r(dyn["taux_propre"]), "nb_gagnes": dyn["nb_gagnes"],
            "apres_1er_set": _r(dyn["taux_apres_premier_set"]), "nb_1er_gagne": dyn["nb_premier_set_gagne"],
            "comeback": _r(dyn["taux_comeback"]), "nb_1er_perdu": dyn["nb_premier_set_perdu"]}


def _tableau_payload(entry, prog, matchs_tableau):
    clutch, dyn = entry["clutch"], entry["dynamique"]
    dvg = stats.david_vs_goliath(matchs_tableau)
    out = {
        "classement": entry["classement"], "cote": _r(entry["cote"], 0),
        "matchs": entry["matchs_joues"], "victoires": entry["victoires"], "pct": _r(entry["pct_victoire"]),
        "forme": _r(entry["indice_forme"]), "perf": _r(entry["indice_performance"]),
        "qualite": _r(entry["indice_qualite"]), "niveau": _r(entry["indice_niveau"]),
        "points_par_match": _r(entry["points_cote_total"] / entry["nb_matchs_cotes"])
        if entry.get("nb_matchs_cotes") else None,
        "classement_sept": prog["septembre_classement"], "cote_sept": _r(prog["septembre_points"], 0),
        "diff_cote": _r(prog["diff_points"], 0), "tendance": _r(prog["tendance_hebdo"]),
        "rang_sept": prog["septembre_rang"], "rang": prog["actuel_rang"],
        "courbe": [[d, _r(c, 0)] for d, c in entry.get("courbe_cote", [])],
        "cote_adverse_moyenne": _r(entry["cote_adverse_moyenne"], 0),
        "meilleure_victoire": _fait(entry["meilleure_victoire"], "adversaire_cote"),
        "pire_defaite": _fait(entry["pire_defaite"], "adversaire_cote"),
        "plus_grosse_perf": _fait(entry["plus_grosse_perf"], "points_cote"),
        "plus_grosse_contre_perf": _fait(entry["plus_grosse_defaite_points"], "points_cote"),
        "plus_gros_ecart": _ecart(entry["plus_grande_difference_cote"]),
        "clutch": {"joues": clutch["sets_serres_joues"], "gagnes": clutch["sets_serres_gagnes"],
                   "taux": _r(clutch["taux_clutch"]), "delta": _r(clutch["delta_clutch"])},
        "dynamique": _dynamique(dyn),
        "outsider": {"matchs": dvg["nb_outsider"], "taux": _r(dvg["taux_outsider"])},
        "favori": {"matchs": dvg["nb_favori"], "taux": _r(dvg["taux_favori"])},
        "profil_score": {k: _r(v) for k, v in entry["profil_score"].items()},
        "sets_point_en_or": {"joues": entry["sets_extremes"]["nb_joues"],
                             "gagnes": entry["sets_extremes"]["nb_gagnes"]},
    }
    if "partenaire" in entry:
        p = entry["partenaire"]
        out["avec_partenaire_club"] = {"matchs": p["avec_partenaire_club"]["matchs_joues"],
                                       "pct": _r(p["avec_partenaire_club"]["pct_victoire"])}
        out["sans_partenaire_club"] = {"matchs": p["sans_partenaire_club"]["matchs_joues"],
                                       "pct": _r(p["sans_partenaire_club"]["pct_victoire"])}
    return out


def _joueur_payload(s):
    g, prog = s["global"], s["progression"]
    matchs = sorted(s["match_log"], key=lambda m: m["date"], reverse=True)
    return {
        "licence": s["Licence"], "nom": s["Nom"], "sexe": s["Sexe"],
        "ordre": s["ordre_disciplines"],
        "global": {
            "matchs": g["matchs_joues"], "victoires": g["victoires"], "pct": _r(g["pct_victoire"]),
            "forme": _r(g["indice_forme"]), "perf": _r(g["indice_performance"]),
            "qualite": _r(g["indice_qualite"]), "niveau": _r(g["indice_niveau"]),
            "meilleur_tableau": g["meilleur_tableau"],
            "diff_cote": _r(prog["cumule"]["diff_points"], 0),
            # profil de score tous tableaux confondus (filtre "Tous" de
            # Profil de jeu) - mêmes fonctions que par tableau, sur tous les matchs
            "profil_score": {k: _r(v) for k, v in stats.profil_score(s["match_log"]).items()},
            "dynamique": _dynamique(stats.dynamique_match(s["match_log"])),
            "sets_point_en_or": {"joues": sum(s["par_tableau"][t]["sets_extremes"]["nb_joues"] for t in TABLEAUX),
                                 "gagnes": sum(s["par_tableau"][t]["sets_extremes"]["nb_gagnes"] for t in TABLEAUX)},
        },
        "tableaux": {t: _tableau_payload(s["par_tableau"][t], prog[t],
                                         [m for m in s["match_log"] if m["tableau"] == t])
                     for t in TABLEAUX},
        "partenaires": _partenaires(s["match_log"]),
        "tournois": sorted(({"date": t["date"], "nom": t["nom"], "type": t["type"]}
                            for t in s["tournois"]["tournois"]), key=lambda t: t["date"]),
        "nb_interclubs": s["tournois"]["nb_interclubs"],
        "competitions_par_mois": {mois: dict(b) for mois, b in export_excel._joueur_par_mois(s["tournois"]).items()},
        "matchs": [{
            "date": m["date"], "tableau": m["tableau"], "evenement": m["evenement"],
            "interclub": m["est_interclub"], "equipe": m.get("equipe"), "intra_club": m["intra_club"], "forfait": m.get("forfait", False),
            "victoire": m["victoire"], "score": " / ".join(f"{a}-{b}" for a, b in m["sets_detail"]),
            "adversaires": m["noms_adverses"], "partenaire": m.get("partenaire_nom"),
            "ma_cote": _r(m["ma_cote"], 0), "cote_adverse": _r(m["adversaire_cote"], 0),
            "points": _r(m["points_cote"]),
        } for m in matchs],
    }


def _taux(b):
    return _r(b["victoires"] / b["joues"] * 100) if b["joues"] else None


def _cle_match(m):
    """Identifiant d'un match réel : le même dans le match_log de chacun des
    joueurs du club qui y ont joué (même date, compétition, tableau,
    adversaires et score vu de leur côté)."""
    return (m["date"], m["evenement"], m["tableau"], tuple(sorted(m["noms_adverses"])),
            tuple(tuple(st) for st in m["sets_detail"]))


def matchs_club_uniques(all_stats):
    """Matchs du club comptés UNE fois chacun (un double où jouent deux joueurs
    du club figure dans les deux match_log : il ne doit compter qu'une fois
    pour le club), matchs internes exclus. -> [(match, {sexes des joueurs du
    club ayant joué}, [noms])]."""
    vus = {}
    # par nom : le représentant d'un match vu par deux joueurs du club est le
    # premier par ordre alphabétique - le même que côté page web, où les
    # joueurs sont aussi triés par nom (listes détaillées cohérentes)
    for s in sorted(all_stats, key=lambda x: x["Nom"]):
        for m in s["match_log"]:
            if m["intra_club"]:
                continue
            cle = _cle_match(m)
            if cle not in vus:
                vus[cle] = (m, set(), [])
            vus[cle][1].add(s["Sexe"])
            vus[cle][2].append(s["Nom"])
    return list(vus.values())


def _categories_match(tableau, sexes):
    """Catégories club d'un match (SH/SD/DH/DD/MX, H/F tous tableaux, S/D
    unisexe, Club) selon le tableau et les genres des joueurs du club."""
    cats = [export_excel.CATEGORIE_CLUB_TOTAL]
    cats += [f"{g} (tous tableaux)" for g in ("H", "F") if g in sexes]
    if tableau == "Mixte":
        cats.append("MX")
    else:
        cats += [tableau[0] + export_excel._GENRE_CODE[g] for g in ("H", "F") if g in sexes]
        cats.append(f"{tableau[0]} (unisexe)")
    return cats


def _interclubs_par_categorie(interclubs):
    """Bilan interclubs par catégorie d'après les feuilles icbad (source de
    référence : chaque match de chaque rencontre, une fois)."""
    out = defaultdict(lambda: {"joues": 0, "victoires": 0})
    for e in (interclubs or {}).get("equipes", []):
        for r in e["rencontres"]:
            for m in r["matchs"]:
                if m["victoire"] is None:
                    continue
                t = m["code"][:2]
                tableau = {"SH": "Simple", "SD": "Simple", "DH": "Double", "DD": "Double"}.get(t, "Mixte")
                sexes = {"SH": {"H"}, "DH": {"H"}, "SD": {"F"}, "DD": {"F"}}.get(t, {"H", "F"})
                for cat in _categories_match(tableau, sexes):
                    out[cat]["joues"] += 1
                    out[cat]["victoires"] += 1 if m["victoire"] else 0
    return out


def _club_payload(all_stats, interclubs=None, effectif=None):
    """Agrégats club (ex-onglets Stats club, Tournois et Stats avancées de
    l'Excel). Chaque match réel compté une fois (voir matchs_club_uniques),
    matchs internes exclus (toujours 1 victoire + 1 défaite pour le club,
    ils ne disent rien du niveau face à l'extérieur). Colonnes interclubs :
    feuilles icbad quand elles sont disponibles."""
    categories = export_excel.CLUB_CATEGORIES + [export_excel.CATEGORIE_CLUB_TOTAL]
    matchs_cat = defaultdict(list)
    for m, sexes, _ in matchs_club_uniques(all_stats):
        for cat in _categories_match(m["tableau"], sexes):
            matchs_cat[cat].append(m)
    ic_icbad = _interclubs_par_categorie(interclubs) if interclubs else None
    nb_plan = (interclubs or {}).get("equipes_planificateur") or 99

    def _bilan(ms):
        return {"joues": len(ms), "victoires": sum(1 for m in ms if m["victoire"])}

    par_categorie = []
    for cat in categories:
        ms = matchs_cat.get(cat, [])
        # interclubs côté myffbad (points de cote) : équipes de la page Interclubs seulement
        ic = [m for m in ms if m["est_interclub"] and (m.get("equipe") or 99) <= nb_plan]
        par_mois = defaultdict(list)
        for m in ms:
            par_mois[m["mois"]].append(m)
        total, inter = _bilan(ms), (ic_icbad[cat] if ic_icbad is not None else _bilan(ic))
        pts_t, pts_ic = stats.points_cote_moyens(ms), stats.points_cote_moyens(ic)
        dyn, dvg, clutch = stats.dynamique_match(ms), stats.david_vs_goliath(ms), stats.indice_clutch(ms)
        par_categorie.append({
            "categorie": cat,
            "matchs": total["joues"], "victoires": total["victoires"], "pct": _taux(total),
            "matchs_ic": inter["joues"], "pct_ic": _taux(inter),
            "par_mois": {mois: {**_bilan(x), "pct": _taux(_bilan(x))} for mois, x in sorted(par_mois.items())},
            "pts_gagnes": _r(pts_t["moyenne_gagne"]), "pts_perdus": _r(pts_t["moyenne_perdu"]),
            "pts_gagnes_ic": _r(pts_ic["moyenne_gagne"]), "pts_perdus_ic": _r(pts_ic["moyenne_perdu"]),
            "propre": _r(dyn["taux_propre"]), "apres_1er_set": _r(dyn["taux_apres_premier_set"]),
            "comeback": _r(dyn["taux_comeback"]),
            "outsider_matchs": dvg["nb_outsider"], "outsider": _r(dvg["taux_outsider"]),
            "favori_matchs": dvg["nb_favori"], "favori": _r(dvg["taux_favori"]),
            "sets_serres": clutch["sets_serres_joues"], "clutch": _r(clutch["taux_clutch"]),
            "delta_clutch": _r(clutch["delta_clutch"]),
        })

    sets_or = [s["par_tableau"][t]["sets_extremes"] for s in all_stats for t in TABLEAUX]
    joueurs_or = sum(1 for s in all_stats if any(s["par_tableau"][t]["sets_extremes"]["nb_joues"] for t in TABLEAUX))
    tournois = sorted(export_excel._tournois_club(all_stats, min_joueurs=1), key=lambda t: t["date"], reverse=True)
    return {
        "categories": par_categorie,
        "sets_point_en_or": {"joues": sum(x["nb_joues"] for x in sets_or),
                             "gagnes": sum(x["nb_gagnes"] for x in sets_or), "joueurs": joueurs_or},
        "tournois": [{**t, "pct": _r(t["pct"]), "points": _r(t["points"])} for t in tournois],
        "rivalites": rivalites(all_stats, {p["nom"] for p in (effectif or [])} | {s["Nom"] for s in all_stats}),
        "pics_mensuels": [{**p, "delta": _r(p["delta"])} for p in export_excel._pics_mensuels(all_stats)],
    }


def _norm_nom(nom):
    return " ".join(str(nom or "").upper().split())


def rivalites(all_stats, noms_club=None):
    """Confrontations entre joueurs du club, chaque match compté une fois
    (il figure dans le match_log de chacun des joueurs du club présents) :

    - "Tous" : 1 contre 1 entre deux joueurs du club, toutes disciplines,
      même si l'un ou l'autre jouait avec un partenaire extérieur (en double,
      chaque joueur du club d'un côté affronte chaque joueur du club de
      l'autre) ;
    - "Simple" : joueur du club contre joueur du club ;
    - "Double" / "Mixte" : paire contre paire, seulement quand les deux
      paires sont entièrement composées de joueurs du club.

    Appartenance au club : noms de l'effectif complet (noms_club) - le
    drapeau intra_club ne suffit pas (il exige deux adversaires du club et
    ignore le partenaire). -> {tableau: [{a, b, total, va, vb, matchs}]},
    chaque match : date, compétition, tableau, côtés gagnant / perdant
    (noms complets, partenaires extérieurs compris), score vu du gagnant."""
    club = {_norm_nom(n) for n in (noms_club or [])}
    du_club = lambda nom: _norm_nom(nom) in club
    vus = {}
    for s in all_stats:
        for m in s["match_log"]:
            adv = tuple(sorted(m["noms_adverses"]))
            if not adv or not any(du_club(n) for n in adv):
                continue
            partenaire = m.get("partenaire_nom")
            moi = tuple(sorted([s["Nom"]] + ([partenaire] if partenaire else [])))
            cle = (m["date"], m["evenement"], m["tableau"], frozenset([moi, adv]))
            if cle not in vus:
                vus[cle] = {"tableau": m["tableau"], "date": m["date"], "evenement": m["evenement"],
                            "gagnant": moi if m["victoire"] else adv, "perdant": adv if m["victoire"] else moi,
                            "score": " / ".join(f"{a}-{b}" if m["victoire"] else f"{b}-{a}" for a, b in m["sets_detail"]),
                            "forfait": bool(m.get("forfait"))}

    def _ajouter(groupes, x, y, match, x_gagne):
        a, b = sorted([x, y])
        g = groupes.setdefault((a, b), {"a": " / ".join(a), "b": " / ".join(b), "total": 0, "va": 0, "vb": 0, "matchs": []})
        g["total"] += 1
        g["va" if (a == x) == x_gagne else "vb"] += 1
        g["matchs"].append({"date": match["date"], "evenement": match["evenement"], "tableau": match["tableau"],
                            "score": match["score"], "gagnant": " / ".join(match["gagnant"]),
                            "perdant": " / ".join(match["perdant"]), "forfait": match["forfait"]})

    out = {t: {} for t in (*TABLEAUX, "Tous")}
    for v in vus.values():
        g_club = [n for n in v["gagnant"] if du_club(n)]
        p_club = [n for n in v["perdant"] if du_club(n)]
        for x in g_club:
            for y in p_club:
                _ajouter(out["Tous"], (x,), (y,), v, True)
        # par tableau : les deux côtés entièrement du club
        if len(g_club) == len(v["gagnant"]) and len(p_club) == len(v["perdant"]) and g_club and p_club:
            _ajouter(out[v["tableau"]], v["gagnant"], v["perdant"], v, True)
    for groupes in out.values():
        for g in groupes.values():
            g["matchs"].sort(key=lambda m: m["date"], reverse=True)
    return {t: sorted(g.values(), key=lambda r: (-r["total"], r["a"])) for t, g in out.items()}


def resume(payload):
    """Chiffres clés d'un export (comparés d'une semaine sur l'autre dans
    l'Aperçu) - publié à côté de la page (voir le workflow)."""
    club = next((c for c in payload["club"]["categories"] if c["categorie"] == export_excel.CATEGORIE_CLUB_TOTAL), {})
    ic = payload.get("interclubs") or {}
    return {
        "genere_le": payload["genere_le"],
        "joueurs_actifs": sum(1 for j in payload["joueurs"] if j["global"]["matchs"]),
        "matchs": club.get("matchs"), "victoires": club.get("victoires"), "pct": club.get("pct"),
        "tournois": len(payload["club"]["tournois"]),
        "rencontres_ic": sum(1 for e in ic.get("equipes", []) for r in e["rencontres"] if r["jouee"]),
    }


def resume_precedent():
    """Résumé de l'export actuellement publié (la semaine précédente), None
    si indisponible (première publication, pas de réseau...)."""
    try:
        r = requests.get(PAGES_URL + "resume.json", timeout=15)
        return r.json() if r.ok else None
    except Exception:
        return None


def _cotes_j10_payload(historique, interclubs):
    """{samedi: {"export": clé, "cotes": {licence: [S, D, M]}}} pour les
    week-ends à venir dont les cotes de J-10 sont connues (voir
    historique_ic.cotes_j10) - les week-ends passés affichent les feuilles."""
    import historique_ic  # import local : historique_ic est indépendant de la page
    if not historique or not interclubs:
        return {}
    samedis = set()
    for e in interclubs.get("equipes", []):
        for r in e["rencontres"]:
            if r.get("date") and not r["jouee"]:
                samedis.add(historique_ic.samedi_de(date.fromisoformat(r["date"])).isoformat())
    out = {}
    for sam in sorted(samedis):
        trouve = historique_ic.cotes_j10(historique, sam)
        if trouve:
            out[sam] = {"export": trouve[0], "cotes": trouve[1]}
    return out


def calendriers_ics(interclubs, dossier):
    """Un calendrier iCalendar par équipe (calendrier-equipe-N.ics) et un pour
    toutes (calendrier-interclubs.ics) : rencontres datées, à importer ou à
    suivre par abonnement (Google Agenda, Apple Calendrier...). Renvoie les
    chemins écrits."""
    if not interclubs:
        return []
    def _esc(t):
        return str(t or "").replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")
    vtimezone = ["BEGIN:VTIMEZONE", "TZID:Europe/Paris",
                 "BEGIN:DAYLIGHT", "TZOFFSETFROM:+0100", "TZOFFSETTO:+0200", "TZNAME:CEST",
                 "DTSTART:19700329T020000", "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU", "END:DAYLIGHT",
                 "BEGIN:STANDARD", "TZOFFSETFROM:+0200", "TZOFFSETTO:+0100", "TZNAME:CET",
                 "DTSTART:19701025T030000", "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU", "END:STANDARD",
                 "END:VTIMEZONE"]
    horodatage = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    def _evenements(e):
        lignes = []
        for r in e["rencontres"]:
            if not r.get("date"):
                continue
            jour = r["date"].replace("-", "")
            h, mn = (r.get("heure") or "14:00").split(":")[:2]
            debut = f"{jour}T{int(h):02d}{int(mn):02d}00"
            fin_h = min(int(h) + 4, 23)
            titre = f"IC Équipe {e['numero']} ({e['division']}) {'vs' if r['domicile'] else '@'} {r['adversaire']}"
            if r.get("jouee") and r.get("score_club") is not None:
                titre += f" · {r['score_club']}-{r['score_adv']}"
            lignes += ["BEGIN:VEVENT", f"UID:ic-{r['id']}@gab38-stats", f"DTSTAMP:{horodatage}",
                       f"DTSTART;TZID=Europe/Paris:{debut}", f"DTEND;TZID=Europe/Paris:{jour}T{fin_h:02d}{int(mn):02d}00",
                       f"SUMMARY:{_esc(titre)}", f"LOCATION:{_esc(r.get('lieu'))}",
                       f"DESCRIPTION:{_esc(e['division_titre'] + ' · ' + str(r.get('journee') or ''))}",
                       f"URL:https://icbad.ffbad.org/rencontre/{r['id']}", "END:VEVENT"]
        return lignes

    def _ecrire(nom, titre, evenements):
        contenu = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//GAB38//Stats joueurs//FR", "CALSCALE:GREGORIAN",
                   f"X-WR-CALNAME:{_esc(titre)}", "X-WR-TIMEZONE:Europe/Paris", *vtimezone, *evenements, "END:VCALENDAR"]
        chemin = os.path.join(dossier, nom)
        with open(chemin, "w", encoding="utf-8", newline="") as f:
            f.write("\r\n".join(contenu) + "\r\n")
        return chemin

    chemins, tous = [], []
    for e in interclubs.get("equipes", []):
        evs = _evenements(e)
        tous += evs
        chemins.append(_ecrire(f"calendrier-equipe-{e['numero']}.ics", f"GAB38 · Interclubs équipe {e['numero']}", evs))
    chemins.append(_ecrire("calendrier-interclubs.ics", "GAB38 · Interclubs (toutes les équipes)", tous))
    return chemins


def _saison(today=None):
    today = today or date.today()
    debut = today.year if today.month >= 9 else today.year - 1
    return f"{debut}-{debut + 1}"


def _interclubs_payload(interclubs, all_stats):
    """Données icbad (voir icbad.collecter) telles quelles pour la page
    Interclubs, plus les points de cote gagnés/perdus en interclub par
    joueur et par équipe - icbad ne les donne pas, les résultats joueurs
    myffbad si (match_log : est_interclub + equipe)."""
    if not interclubs:
        return None
    points = defaultdict(lambda: defaultdict(float))
    for s in all_stats:
        for m in s["match_log"]:
            if m["est_interclub"] and m.get("equipe") is not None and m["points_cote"] is not None:
                points[s["Licence"]][str(m["equipe"])] += m["points_cote"]
    # joueurs d'interclubs sans fiche complète : scrape allégé de leurs
    # résultats (statsJoueurs.completer_points_cote_ic)
    for lic, par_eq in (interclubs.get("points_cote_complement") or {}).items():
        if lic not in points:
            for eq, v in par_eq.items():
                points[lic][str(eq)] += v
    return {**{k: v for k, v in interclubs.items() if k != "points_cote_complement"},
            "points_cote": {lic: {eq: _r(v) for eq, v in par_eq.items()} for lic, par_eq in points.items()}}


def build_payload(all_stats, interclubs=None, effectif=None, historique=None):
    joueurs = sorted((_joueur_payload(s) for s in all_stats), key=lambda j: j["nom"])
    return {
        "genere_le": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "saison": _saison(),
        "excel_url": EXCEL_URL,
        "joueurs": joueurs,
        "club": _club_payload(all_stats, interclubs, effectif),
        "interclubs": _interclubs_payload(interclubs, all_stats),
        "effectif": effectif,
        # valeurs figées par week-end d'interclubs (voir historique_ic)
        "valeurs_historiques": (historique or {}).get("weekends") or {},
        # indices de forme export par export (tendances sur la saison)
        "historique_forme": {k: v.get("forme", {}) for k, v in sorted(((historique or {}).get("exports") or {}).items())},
        # cotes de J-10 des week-ends à venir (hiérarchie réglementaire)
        "cotes_j10": _cotes_j10_payload(historique, interclubs),
    }


def write_stats_html(all_stats, output_path, template_path=TEMPLATE_PATH, interclubs=None, effectif=None, historique=None,
                     precedent=None):
    with open(template_path, encoding="utf-8") as f:
        template = f.read()
    assert DATA_PLACEHOLDER in template, f"{DATA_PLACEHOLDER} absent du gabarit {template_path}"
    payload = build_payload(all_stats, interclubs, effectif, historique)
    dossier = os.path.dirname(os.path.abspath(output_path))
    # precedent fourni ({} = « aucun » connu) : pas d'appel réseau
    payload["precedent"] = (precedent or None) if precedent is not None else resume_precedent()
    with open(os.path.join(dossier, PRECEDENT_NOM), "w", encoding="utf-8") as f:
        json.dump(payload["precedent"], f, ensure_ascii=False, default=_json_default)
    with open(os.path.join(dossier, RESUME_NOM), "w", encoding="utf-8") as f:
        json.dump(resume(payload), f, ensure_ascii=False, default=_json_default)
    calendriers_ics(interclubs, os.path.dirname(os.path.abspath(output_path)))
    # JSON compact (sans espaces) : ~10 % de moins à charger
    data = json.dumps(payload, ensure_ascii=False, default=_json_default, separators=(",", ":"))
    # "</script>" dans un nom de tournoi fermerait la balise : on échappe "</"
    data = data.replace("</", "<\\/")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(template.replace(DATA_PLACEHOLDER, data))
    return output_path
