"""
Jeu de données ENTIÈREMENT FICTIF (noms, licences, clubs) pour les tests :
le dépôt est public, aucune donnée réelle du club n'y figure.

Passe par le vrai pipeline (stats.build_player_stats -> normaliser_club ->
export_web) avec, en plus d'un volume aléatoire reproductible, des cas
construits à la main pour les règles délicates (voir CAS) :

- un double joué par deux joueurs du club : figure dans les deux fiches, doit
  compter UNE fois pour le club ;
- une confrontation interne paire contre paire (100 % club) ;
- un double contre un joueur du club associé à un partenaire extérieur : pas
  "interne", absent des rivalités Double, présent en "Tous" (1 contre 1) ;
- un forfait (21-0 / 21-0, points de cote inconnus).
"""
import os
import random
import sys
from datetime import date, timedelta

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RACINE not in sys.path:
    sys.path.insert(0, RACINE)

import export_web  # noqa: E402
import historique_ic  # noqa: E402
import pointsIC  # noqa: E402
import stats  # noqa: E402

MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
        "septembre", "octobre", "novembre", "décembre"]
HOMMES = ["Arthur ALPHA", "Basile BRAVO", "Cyril CHARLIE", "Denis DELTA", "Emile ECHO", "Félix FOXTROT",
          "Gaston GOLF", "Hector HOTEL", "Igor INDIA", "Jules JULIETT", "Kevin KILO", "Louis LIMA"]
FEMMES = ["Agnès AMBRE", "Berthe BRUME", "Clara CORAIL", "Delphine DUNE", "Eva ECUME", "Flore FLOCON",
          "Gisèle GIVRE", "Hortense HOULE", "Iris ISARD", "Julie JADE", "Karine KAOLIN", "Lina LAGUNE"]
EXTERIEURS_H = ["Marc MARS", "Noé NEPTUNE", "Oscar ORION", "Paul PLUTON"]
EXTERIEURS_F = ["Maud MERCURE", "Nina NOVA", "Olga OPALE", "Pia PERLE"]
TOURNOIS = ["Tournoi du Lac", "Open des Cimes", "Plumes du Plateau", "Trophée de la Vallée"]
SEPT = date(2026, 9, 1)
FIN = date(2026, 11, 10)

# licence fictive par joueur du club
CLUB = {nom: ("H", f"T{100 + i:06d}") for i, nom in enumerate(HOMMES)}
CLUB.update({nom: ("F", f"T{200 + i:06d}") for i, nom in enumerate(FEMMES)})

# cas construits à la main (date, compétition, tableau, côté gagnant, côté perdant, sets du gagnant)
CAS = {
    "double_deux_club": (date(2026, 10, 3), "Tournoi du Lac", "Double",
                         ["Arthur ALPHA", "Basile BRAVO"], ["Marc MARS", "Noé NEPTUNE"], [(15, 10), (15, 12)]),
    "interne_paires": (date(2026, 10, 4), "Tournoi du Lac", "Double",
                       ["Arthur ALPHA", "Basile BRAVO"], ["Cyril CHARLIE", "Denis DELTA"], [(15, 13), (14, 16), (15, 11)]),
    "partenaire_exterieur": (date(2026, 10, 10), "Open des Cimes", "Double",
                             ["Cyril CHARLIE", "Oscar ORION"], ["Arthur ALPHA", "Paul PLUTON"], [(15, 9), (15, 13)]),
    "interne_simple": (date(2026, 10, 17), "Plumes du Plateau", "Simple",
                       ["Arthur ALPHA"], ["Cyril CHARLIE"], [(15, 12), (21, 20)]),
    "forfait": (date(2026, 10, 24), "Trophée de la Vallée", "Simple",
                ["Basile BRAVO"], ["Marc MARS"], [(21, 0), (21, 0)]),
}


def _fr(d):
    return f"{d.day} {MOIS[d.month - 1]} {d.year}"


def _classement(cote):
    seuils = [(3400, "N1"), (2800, "N2"), (2300, "N3"), (2000, "R4"), (1800, "R5"), (1550, "R6"),
              (1350, "D7"), (1200, "D8"), (1050, "D9"), (900, "P10"), (800, "P11"), (0, "P12")]
    return next(c for s, c in seuils if cote >= s)


class Generateur:
    def __init__(self, graine=11):
        self.rnd = random.Random(graine)
        self.cotes = {nom: {t: round(self.rnd.gauss(2200, 450)) for t in ("Simple", "Double", "Mixte")} for nom in CLUB}
        # events[nom][tableau] = {(date, compétition, rencontre): event}
        self.events = {nom: {t: {} for t in ("Simple", "Double", "Mixte")} for nom in CLUB}

    def jouer(self, d, evenement, tableau, gagnants, perdants, sets, interclub=False, rencontre=None, forfait=False):
        """Ajoute le match dans la fiche de CHAQUE joueur du club présent, vu de son côté."""
        for cote_x, autre, gagne in ((gagnants, perdants, True), (perdants, gagnants, False)):
            for nom in cote_x:
                if nom not in CLUB:
                    continue
                partenaire = next((n for n in cote_x if n != nom), None)
                ma_cote = self.cotes[nom][tableau]
                cote_adv = sum(self.cotes.get(n, {}).get(tableau, 2100) for n in autre) / len(autre)
                pts = None if forfait else (round(12 + (cote_adv - ma_cote) / 60) if gagne else -round(12 + (ma_cote - cote_adv) / 60))
                cle = (d, evenement, rencontre)
                ev = self.events[nom][tableau].setdefault(cle, {
                    "date": _fr(d), "evenement": evenement, "est_interclub": interclub,
                    "tableau_ou_rencontre": rencontre or tableau, "points": "", "matchs": []})
                ev["matchs"].append({
                    "tour": "Poule", "mes_clubs": ["GAB38"],
                    "clubs_adverses": ["GAB38" if n in CLUB else "EXT38" for n in autre],
                    "mes_noms": list(cote_x), "partenaire_nom": partenaire,
                    "partenaire_licence": CLUB[partenaire][1] if partenaire in CLUB else None,
                    "partenaire_club": None if not partenaire else ("GAB38" if partenaire in CLUB else "EXT38"),
                    "noms_adverses": list(autre),
                    "sets_gagnes": len([s for s in sets if (s[0] > s[1]) == gagne]),
                    "sets_perdus": len([s for s in sets if (s[0] > s[1]) != gagne]),
                    "sets_detail": [tuple(s) if gagne else (s[1], s[0]) for s in sets],
                    "victoire": gagne, "points_cote": pts, "ma_cote": ma_cote, "adversaire_cote": round(cote_adv),
                    "tournoi_id": None if interclub else f"T-{evenement}-{d.isoformat()}",
                })

    def remplir(self):
        for cle, (d, ev, t, g, p, sets) in CAS.items():
            self.jouer(d, ev, t, g, p, sets, forfait=(cle == "forfait"))
        # volume aléatoire reproductible : tournois du week-end, quelques interclubs
        d = SEPT + timedelta(days=5)
        while d < FIN:
            for _ in range(10):
                t = self.rnd.choice(["Simple", "Double", "Mixte"])
                if t == "Simple":
                    pool = self.rnd.choice([HOMMES, FEMMES])
                    a = [self.rnd.choice(pool)]
                    b = [self.rnd.choice(EXTERIEURS_H if pool is HOMMES else EXTERIEURS_F)]
                elif t == "Double":
                    pool, ext = self.rnd.choice([(HOMMES, EXTERIEURS_H), (FEMMES, EXTERIEURS_F)])
                    a, b = self.rnd.sample(pool, 2), self.rnd.sample(ext, 2)
                else:
                    a, b = [self.rnd.choice(HOMMES), self.rnd.choice(FEMMES)], [self.rnd.choice(EXTERIEURS_H), self.rnd.choice(EXTERIEURS_F)]
                serre = self.rnd.random() < 0.3
                sets = [(15, self.rnd.randint(3, 12)), (self.rnd.choice([15, 21]) if serre else 15, self.rnd.choice([13, 20]) if serre else self.rnd.randint(3, 12))]
                if self.rnd.random() < 0.5:
                    a, b = b, a
                self.jouer(d, self.rnd.choice(TOURNOIS), t, a, b, sets)
            if d.weekday() == 5 and self.rnd.random() < 0.5:  # une rencontre d'interclub fictive
                rencontre = "38-GAB-2 contre 69-EXT-1"
                for t, a, b in (("Simple", [HOMMES[4]], [EXTERIEURS_H[0]]), ("Mixte", [HOMMES[5], FEMMES[5]], [EXTERIEURS_H[1], EXTERIEURS_F[1]])):
                    self.jouer(d, "Interclubs Régionaux AURA Pré-National", t, a, b, [(15, 11), (15, 9)], interclub=True, rencontre=rencontre)
            d += timedelta(days=self.rnd.randint(4, 9))

    def joueurs(self):
        """(player roster complété, events par tableau) pour build_player_stats."""
        out = []
        for nom, (sexe, licence) in CLUB.items():
            cotes = self.cotes[nom]
            player = {
                "Licence": licence, "Nom": nom, "Sexe": sexe,
                "Classement Actuel": {t: _classement(c) for t, c in cotes.items()},
                "Points Actuel": dict(cotes), "Rang Actuel": {t: 1 for t in cotes},
                "Journal cote": {t: [{"date": SEPT, "points": c, "cote": c}] for t, c in cotes.items()},
                "Classement 1er septembre": {"date": SEPT, **{t: {"rang": 500, "classement": _classement(c), "points": c}
                                                              for t, c in cotes.items()}},
                "Rang National Actuel": {t: 480 for t in cotes},
            }
            pointsIC.compute_IC_points(player)
            events = {t: sorted(self.events[nom][t].values(), key=lambda e: stats.parse_date_fr(e["date"]), reverse=True)
                      for t in ("Simple", "Double", "Mixte")}
            out.append((player, events))
        return out


def interclubs_fictifs(aujourd_hui):
    """Structure icbad.collecter() fictive : 2 équipes, une rencontre jouée
    (feuille complète) et une à venir chacune, adversaires avec composition type."""
    passe = aujourd_hui - timedelta(days=(aujourd_hui.weekday() - 5) % 7 + 14)   # samedi, deux semaines avant
    futur = aujourd_hui + timedelta(days=(5 - aujourd_hui.weekday()) % 7 + 7)    # samedi suivant + 1 semaine
    codes = ["SH1", "SH2", "SD1", "SD2", "DH1", "DD1", "DX1", "DX2"]

    def joueur_adv(nom, sexe, i):
        return {"licence": f"X{i:06d}", "nom": nom, "sexe": sexe, "capitaine": i == 1, "mute": i == 2,
                "classements": [{"classement": "N3", "cote": 2400 - 50 * i}] * 3, "rencontres": 2, "matchs": 4, "pct": 50.0}

    adv_joueurs = [joueur_adv(n, "H", i) for i, n in enumerate(EXTERIEURS_H)] + \
                  [joueur_adv(n, "F", i + 10) for i, n in enumerate(EXTERIEURS_F)]
    lic = {p["nom"]: p["licence"] for p in adv_joueurs}

    def feuille(h, f):
        roles = {"SH1": [h[0]], "SH2": [h[1]], "SD1": [f[0]], "SD2": [f[1]], "DH1": [h[0], h[2]], "DD1": [f[0], f[2]],
                 "DX1": [h[1], f[1]], "DX2": [h[2], f[2]]}
        adv = {"SH1": [EXTERIEURS_H[0]], "SH2": [EXTERIEURS_H[1]], "SD1": [EXTERIEURS_F[0]], "SD2": [EXTERIEURS_F[1]],
               "DH1": EXTERIEURS_H[:2], "DD1": EXTERIEURS_F[:2], "DX1": [EXTERIEURS_H[2], EXTERIEURS_F[2]], "DX2": [EXTERIEURS_H[3], EXTERIEURS_F[3]]}
        return [{"code": c, "club": [{"licence": CLUB[n][1], "nom": n, "classement": "N3", "cote": 2300, "mute": n == h[2]} for n in roles[c]],
                 "adverse": [{"licence": lic[n], "nom": n, "classement": "N3", "cote": 2250, "mute": False} for n in adv[c]],
                 "sets": [[15, 11], [15, 13]] if k % 3 else [[11, 15], [13, 15]], "victoire": bool(k % 3)}
                for k, c in enumerate(codes)]

    equipes, adversaires = [], {}
    for numero, division, h, f in ((1, "N2", HOMMES[:3], FEMMES[:3]), (2, "PN", HOMMES[3:6], FEMMES[3:6])):
        adv_id, mon_id = f"9{numero}00", f"8{numero}00"
        code_adv = f"69-EXT-{numero}"
        classement = [
            {"rang": 1, "nom": f"Club fictif ({code_adv})", "code": code_adv, "equipe_id": adv_id, "club": False,
             "j": 1, "g": 1, "n": 0, "p": 0, "f": 0, "bonus": 0, "penalites": 0, "pts": 5, "diff_matchs": 4, "diff_sets": 6, "diff_points": 40},
            {"rang": 2, "nom": f"GAB fictif (38-GAB-{numero})", "code": f"38-GAB-{numero}", "equipe_id": mon_id, "club": True,
             "j": 1, "g": 1, "n": 0, "p": 0, "f": 0, "bonus": 0, "penalites": 0, "pts": 5, "diff_matchs": 2, "diff_sets": 3, "diff_points": 20},
        ]
        jouee = feuille(h, f)
        rencontres = [
            {"id": f"{numero}001", "journee": "J1", "date": passe.isoformat(), "heure": "15:00", "lieu": "Gymnase fictif", "statut": None,
             "domicile": True, "adversaire": f"Club fictif ({code_adv})", "adversaire_code": code_adv, "adversaire_equipe_id": adv_id,
             "score_club": sum(1 for m in jouee if m["victoire"]), "score_adv": sum(1 for m in jouee if not m["victoire"]),
             "jouee": True, "matchs": jouee},
            {"id": f"{numero}002", "journee": "J2", "date": futur.isoformat(), "heure": "14:00", "lieu": "Salle fictive", "statut": None,
             "domicile": False, "adversaire": f"Club fictif ({code_adv})", "adversaire_code": code_adv, "adversaire_equipe_id": adv_id,
             "score_club": None, "score_adv": None, "jouee": False, "matchs": []},
        ]
        joueurs = [{"licence": CLUB[n][1], "nom": n, "sexe": CLUB[n][0], "capitaine": n == h[0], "mute": n == h[2],
                    "classements": [{"classement": "N3", "cote": 2300}] * 3, "rencontres": 1, "matchs": 2, "pct": 50.0} for n in h + f]
        equipes.append({"equipe_id": mon_id, "nom": f"GAB fictif (38-GAB-{numero})", "code": f"38-GAB-{numero}", "numero": numero,
                        "division": division, "division_titre": f"Division fictive {division}", "division_libelle": division,
                        "rang": 2, "nb_equipes": 2, "classement": classement, "joueurs": joueurs, "rencontres": rencontres})
        adversaires[adv_id] = {"nom": f"Club fictif ({code_adv})", "code": code_adv, "joueurs": adv_joueurs,
                               "composition_type": [{"code": c, "joueurs": [{"licence": lic[n], "nom": n, "classement": "N3"} for n in noms]}
                                                    for c, noms in zip(codes, [[EXTERIEURS_H[0]], [EXTERIEURS_H[1]], [EXTERIEURS_F[0]], [EXTERIEURS_F[1]],
                                                                               EXTERIEURS_H[:2], EXTERIEURS_F[:2], [EXTERIEURS_H[2], EXTERIEURS_F[2]],
                                                                               [EXTERIEURS_H[3], EXTERIEURS_F[3]]])],
                               "rencontres": [{"journee": "J1", "date": passe.isoformat(), "domicile": f"Club fictif ({code_adv})",
                                               "exterieur": "Autre club (01-AUT-1)", "score": [5, 3]}]}
    return {"club": "GAB38", "url": "https://icbad.ffbad.org/instance/GAB38", "equipes": equipes,
            "adversaires": adversaires, "equipes_planificateur": 6}


def construire(aujourd_hui=None):
    """-> (all_stats normalisé, interclubs, effectif, historique), tout fictif."""
    aujourd_hui = aujourd_hui or date.today()
    gen = Generateur()
    gen.remplir()
    joueurs = gen.joueurs()
    all_stats = [stats.build_player_stats(p, e) for p, e in joueurs]
    stats.normaliser_club(all_stats)
    effectif = export_web.effectif_payload([p for p, _ in joueurs])
    interclubs = interclubs_fictifs(aujourd_hui)
    historique = {"v": 1, "weekends": {}}
    historique_ic.mettre_a_jour(historique, interclubs, effectif, aujourd_hui)
    return all_stats, interclubs, effectif, historique


def ecrire_page(dossier, aujourd_hui=None):
    """Écrit la page HTML fictive dans `dossier` (sans réseau) et renvoie son chemin."""
    all_stats, interclubs, effectif, historique = construire(aujourd_hui)
    from datetime import datetime, timezone
    for jours in (14, 7, 0):
        historique_ic.ajouter_export(historique, all_stats, effectif,
                                     datetime.now(timezone.utc) - timedelta(days=jours))
    ancien = export_web.resume_precedent
    export_web.resume_precedent = lambda: None  # pas d'appel réseau en test
    try:
        return export_web.write_stats_html(all_stats, os.path.join(dossier, "statsJoueurs.html"),
                                           interclubs=interclubs, effectif=effectif, historique=historique)
    finally:
        export_web.resume_precedent = ancien
