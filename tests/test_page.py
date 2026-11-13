"""Parcours de la page web générée (données fictives) dans Chrome headless :
aucune erreur JavaScript, chaque page affiche du contenu, et les actions
principales du planificateur fonctionnent.

    python -m unittest tests.test_page
Nécessite Chrome (présent sur les runners GitHub ubuntu-latest)."""
import os
import pathlib
import tempfile
import time
import unittest

from tests import donnees_fictives as df

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
except ImportError:  # pragma: no cover
    webdriver = None


@unittest.skipIf(webdriver is None, "selenium absent")
class TestPage(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.url = pathlib.Path(df.ecrire_page(cls.tmp.name)).as_uri()
        o = Options()
        for arg in ("--headless=new", "--no-sandbox", "--disable-dev-shm-usage", "--window-size=1360,1000"):
            o.add_argument(arg)
        o.set_capability("goog:loggingPrefs", {"browser": "ALL"})
        cls.d = webdriver.Chrome(options=o)

    @classmethod
    def tearDownClass(cls):
        cls.d.quit()
        cls.tmp.cleanup()

    def aller(self, ancre, attente=0.5):
        self.d.get(self.url + ancre)
        time.sleep(attente)

    def js(self, code):
        return self.d.execute_script("return " + code)

    def verifier_sans_erreur(self, ancre):
        erreurs = [e["message"] for e in self.d.get_log("browser") if e["level"] == "SEVERE"]
        self.assertFalse(erreurs, f"{ancre} : {erreurs}")
        self.assertGreater(len(self.js("document.querySelector('#main').innerText").strip()), 40, f"{ancre} : page vide")

    def test_toutes_les_pages(self):
        joueur = df.CLUB["Arthur ALPHA"][1]
        pages = ["", "#club/classement", "#club/competitions", "#club/categories", "#club/profil", "#club/records",
                 "#club/detail/clutch/SH", "#club/detail/outsider/Club%20(tous%20tableaux)", "#club/detail/or/Club%20(tous%20tableaux)",
                 "#club/tournoi/0", "#club/suivi", "#club/paires", "#ic", "#ic/pilotage", "#ic/1", "#ic/2", "#ic/adv/9100", "#plan", f"#joueur/{joueur}"]
        for ancre in pages:
            with self.subTest(page=ancre):
                self.aller(ancre)
                self.verifier_sans_erreur(ancre)

    def test_planificateur(self):
        self.aller("#plan")
        self.d.execute_script("localStorage.clear()")
        self.aller("#plan")
        futur = self.js("[...document.querySelectorAll('#planWes tbody tr')].map(tr => tr.dataset.href).find(h => h && document.querySelector(`tr[data-href='${h}']`).innerText.match(/prochain|à venir/))")
        self.assertTrue(futur, "aucun week-end à venir")
        self.aller(futur)
        # affecter trois hommes et trois femmes à l'équipe 2 via les menus
        for sexe in ("H", "F"):
            for _ in range(3):
                self.d.execute_script("""const s = [...document.querySelectorAll('#planPool select[data-affecter]')].find(x => !x.value && x.closest('tr').cells[1].innerText.trim() === arguments[0]);
                    if (s) { s.value = '2'; s.dispatchEvent(new Event('change', {bubbles: true})); }""", sexe)
                time.sleep(0.1)
        self.assertEqual(self.js("document.querySelectorAll('.plan-eq[data-equipe=\"2\"] .pj').length"), 6)
        self.verifier_sans_erreur(futur)
        # composition proposée automatiquement : complète et sans erreur réglementaire
        lien = self.js("document.querySelector('.plan-eq[data-equipe=\"2\"] a.renc-lien').getAttribute('href')")
        self.aller(lien)
        self.d.execute_script("document.querySelector('#planProposer').click()")
        time.sleep(1.5)
        self.assertEqual(self.js("[...document.querySelectorAll('select[data-cote=gab]')].filter(s => !s.value).length"), 0)
        self.assertFalse(self.js("[...document.querySelectorAll('#planChecks .chk.ko')].map(c => c.innerText)"))
        self.verifier_sans_erreur(lien)
        # le plan survit au rechargement
        self.d.refresh()
        time.sleep(0.6)
        self.assertEqual(self.js("[...document.querySelectorAll('select[data-cote=gab]')].filter(s => !s.value).length"), 0)


if __name__ == "__main__":
    unittest.main()
