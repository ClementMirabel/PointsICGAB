"""Règles de calcul (sans réseau, données fictives) :  python -m unittest discover tests"""
import os
import tempfile
import unittest
from datetime import date, datetime, timezone

from tests import donnees_fictives as df

import export_web  # noqa: E402  (chemin ajouté par donnees_fictives)
import historique_ic  # noqa: E402
import results  # noqa: E402
import roster  # noqa: E402
import stats  # noqa: E402


class TestRegles(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.all_stats, cls.ic, cls.effectif, cls.hist = df.construire(date(2026, 11, 13))
        cls.noms_club = {p["nom"] for p in cls.effectif}

    # --- comptage des matchs du club
    def test_double_a_deux_joueurs_du_club_compte_une_fois(self):
        d, ev, tableau, gagnants, perdants, _ = df.CAS["double_deux_club"]
        uniques = [(m, joueurs) for m, _, joueurs in export_web.matchs_club_uniques(self.all_stats)
                   if m["date"] == d and m["evenement"] == ev and sorted(m["noms_adverses"]) == sorted(perdants)]
        self.assertEqual(len(uniques), 1)
        self.assertEqual(sorted(uniques[0][1]), sorted(gagnants))  # vu par les deux joueurs du club

    def test_match_interne_exclu_des_stats_club(self):
        d, ev, *_ = df.CAS["interne_paires"]
        self.assertFalse([m for m, _, _ in export_web.matchs_club_uniques(self.all_stats) if m["date"] == d and m["evenement"] == ev])

    # --- rivalités
    def test_rivalites(self):
        riv = export_web.rivalites(self.all_stats, self.noms_club)
        paires = {(r["a"], r["b"]) for r in riv["Double"]}
        self.assertIn(("Arthur ALPHA / Basile BRAVO", "Cyril CHARLIE / Denis DELTA"), paires)
        # partenaire extérieur : jamais dans une paire, mais compté en 1 contre 1
        self.assertFalse(any("ORION" in a + b or "PLUTON" in a + b for a, b in paires))
        tous = {(r["a"], r["b"]): r for r in riv["Tous"]}
        arthur_cyril = tous[("Arthur ALPHA", "Cyril CHARLIE")]
        self.assertEqual(arthur_cyril["total"], 3)  # simple interne + double interne + double avec partenaires extérieurs
        self.assertEqual(arthur_cyril["va"] + arthur_cyril["vb"], arthur_cyril["total"])
        self.assertFalse(any(n not in self.noms_club for r in riv["Tous"] for n in (r["a"], r["b"])))

    # --- interclubs : feuilles icbad
    def test_colonnes_interclubs_depuis_icbad(self):
        club = next(c for c in export_web._club_payload(self.all_stats, self.ic, self.effectif)["categories"]
                    if c["categorie"] == "Club (tous tableaux)")
        nb_feuilles = sum(1 for e in self.ic["equipes"] for r in e["rencontres"] for m in r["matchs"] if m["victoire"] is not None)
        self.assertEqual(club["matchs_ic"], nb_feuilles)

    def test_equipe_club(self):
        self.assertEqual(results.equipe_club("38-GAB-3 contre 74-ABC-2"), 3)
        self.assertEqual(results.equipe_club("13-BCF-2 contre 38-GAB-12"), 12)
        self.assertIsNone(results.equipe_club("38-HP-1 contre 38-EB-1"))

    # --- indices, forfaits, barèmes
    def test_forfait_sans_points_de_cote(self):
        self.assertTrue(stats._est_forfait({"points_cote": None, "sets_detail": [(21, 0), (0, 21)]}))
        self.assertFalse(stats._est_forfait({"points_cote": None, "sets_detail": [(15, 0), (15, 3)]}))

    def test_rang_percentile(self):
        items = [{"v": 1}, {"v": 2}, {"v": 3}, {"v": None}]
        stats.rang_percentile(items, "v", "p")
        self.assertEqual([i["p"] for i in items], [0.0, 50.0, 100.0, None])

    def test_bareme_icr(self):
        self.assertEqual(roster.valeur_icr("N1"), 57)
        self.assertEqual(roster.valeur_icr("R4"), 30)
        self.assertEqual(roster.valeur_icn("N1", "H", 4500, 0, 0), 93)

    # --- historique : valeurs figées et cotes de J-10
    def test_week_end_passe_fige(self):
        self.assertTrue(self.hist["weekends"])

    def test_cotes_j10(self):
        h = {"v": 1, "weekends": {}}
        for jour in ("2026-11-12", "2026-11-19", "2026-11-26"):
            historique_ic.ajouter_export(h, [], self.effectif, datetime.fromisoformat(jour + "T01:05:00+00:00"))
        self.assertEqual(historique_ic.cotes_j10(h, "2026-11-28")[0], "2026-11-26T01:05Z")
        self.assertIsNone(historique_ic.cotes_j10(h, "2026-12-19"))

    # --- sorties : calendriers et page
    def test_calendriers(self):
        with tempfile.TemporaryDirectory() as tmp:
            chemins = export_web.calendriers_ics(self.ic, tmp)
            self.assertEqual(len(chemins), len(self.ic["equipes"]) + 1)
            with open(chemins[0], encoding="utf-8") as f:
                contenu = f.read()
            self.assertEqual(contenu.count("BEGIN:VEVENT"), len(self.ic["equipes"][0]["rencontres"]))
            self.assertIn("TZID=Europe/Paris", contenu)

    def test_page_generee(self):
        with tempfile.TemporaryDirectory() as tmp:
            chemin = df.ecrire_page(tmp)
            with open(chemin, encoding="utf-8") as f:
                html = f.read()
            self.assertNotIn("/*__DATA__*/null", html)
            # alerte de taille (voir README) : à surveiller si la page grossit
            self.assertLess(os.path.getsize(chemin), 8_000_000, "page HTML > 8 Mo : envisager de découper les données")


if __name__ == "__main__":
    unittest.main()
