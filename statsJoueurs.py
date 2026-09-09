"""
Bilan de performance par joueur (résultats de la saison, classement au 1er
septembre, indices, tournois) pour tout le club.

    python statsJoueurs.py                  # tout le roster -> statsJoueurs.xlsx
    python statsJoueurs.py 06627061 00123456 # juste ces licences (test rapide)
    python statsJoueurs.py --debug 06627061  # dump HTML brut pour inspection

Nécessite une connexion à myffbad.fr ("Résultats" et l'historique de
classement ne sont pas publics, contrairement à pointsIC.py) : licence et
mot de passe via MYFFBAD_LICENCE/MYFFBAD_PASSWORD, ou saisie interactive.
"""
import os
import sys
import time
from getpass import getpass

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

import export_excel
import pointsIC
import results
import roster
import stats

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEBUG_DIR = os.path.join(APP_DIR, "debug")


def get_credentials():
    licence = os.environ.get("MYFFBAD_LICENCE")
    password = os.environ.get("MYFFBAD_PASSWORD")
    if not licence:
        licence = input("Licence : ")
    if not password:
        password = getpass("Mot de passe : ")
    return licence, password


def new_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1400,1000")
    options.add_argument(f"user-agent={roster.USER_AGENT}")
    return webdriver.Chrome(options=options)


def dismiss_rgpd_banner(driver):
    try:
        button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "[data-testid='rgpd-accept-button']")))
        button.click()
    except TimeoutException:
        pass  # pas de bannière (déjà acceptée / cookie présent) : on continue


def login(driver, licence, password):
    driver.get("https://myffbad.fr/")
    dismiss_rgpd_banner(driver)
    try:
        # attend que le champ soit vraiment cliquable (pas juste présent dans le
        # DOM) : juste après la fermeture de la bannière RGPD, une transition
        # CSS peut encore être en cours et rendre le formulaire non interactif
        # un court instant. Sélecteur scopé au formulaire pour rester sur LE
        # bon champ (il peut y avoir d'autres formulaires/boutons sur la page).
        licence_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "section[data-testid='sign-in-form'] input[name='licence']")))
        licence_input.send_keys(licence)

        form = driver.find_element(By.CSS_SELECTOR, "section[data-testid='sign-in-form']")
        form.find_element(By.CSS_SELECTOR, "input[name='password']").send_keys(password)
        form.find_element(By.CSS_SELECTOR, "input[type='checkbox']").click()
        form.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div[data-testid='widgets-grid']")))
        return True
    except TimeoutException:
        return False


def click(driver, element):
    driver.execute_script("arguments[0].click();", element)
    time.sleep(1)  # laisse le temps au contenu de charger après le clic


def expand_all_sections(driver):
    """Déplie chaque section repliable (Résultats, classement, ...)."""
    count = len(driver.find_elements(By.CSS_SELECTOR, "div[data-testid='collapse-item']"))
    for i in range(count):
        items = driver.find_elements(By.CSS_SELECTOR, "div[data-testid='collapse-item']")
        if i >= len(items):
            break
        try:
            click(driver, items[i])
        except Exception as e:
            print(f"  clic section {i} : {e}")


def click_buttons_by_text(driver, text):
    """Clique tous les boutons portant ce texte (plusieurs groupes Simple/Double/
    Mixte coexistent sur la page : Résultats, Nombre de points/Classement,
    Progression). Renvoie le nombre de boutons cliqués.

    Recherche les boutons un par un, juste avant chaque clic : cliquer sur
    l'un d'eux provoque un re-rendu React qui invalide (stale) les
    références des autres boutons déjà récupérés dans une même liste.
    """
    n = 0
    for _ in range(10):  # garde-fou : jamais plus de 10 groupes sur la page
        target = None
        for btn in driver.find_elements(By.CSS_SELECTOR, "button[data-testid='button']"):
            # .text ne lit que le texte visible à l'écran (souvent vide pour
            # un bouton hors viewport en headless) ; textContent lit le DOM,
            # fiable peu importe le scroll.
            label = (btn.get_attribute("textContent") or "").strip()
            if label == text and btn.get_attribute("data-variant") != "primary":
                target = btn
                break
        if target is None:
            break  # plus aucun bouton "text" non encore actif
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
            click(driver, target)
            n += 1
        except Exception as e:
            print(f"  clic bouton '{text}' : {e}")
            break
    return n


def dump_html(driver, licence, suffix):
    os.makedirs(DEBUG_DIR, exist_ok=True)
    path = os.path.join(DEBUG_DIR, f"joueur_{licence}_{suffix}.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print(f"-> {path}")


def dump_player_debug(driver, licence):
    driver.get(f"https://myffbad.fr/joueur/{licence}")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "section[data-testid='player-layout']")))
    except TimeoutException:
        print(f"  page joueur {licence} : chargement trop long, on continue quand même.")

    expand_all_sections(driver)
    dump_html(driver, licence, "simple")

    # bascule le graphique "Nombre de points / Classement" en vue tableau
    n = click_buttons_by_text(driver, "Journal de suivi")
    print(f"  'Journal de suivi' : {n} bouton(s) cliqué(s)")
    dump_html(driver, licence, "journal")

    # bascule tous les groupes Simple/Double/Mixte de la page (Résultats,
    # classement, progression) sur Double puis Mixte, un dump par état
    for tableau in ("Double", "Mixte"):
        n = click_buttons_by_text(driver, tableau)
        print(f"  '{tableau}' : {n} bouton(s) cliqué(s)")
        dump_html(driver, licence, tableau.lower())


def _soup(driver):
    return BeautifulSoup(driver.page_source, "html.parser")


def scrape_player(driver, player):
    """Scrape la page d'un joueur du roster : résultats des 3 tableaux +
    cote au 1er septembre. Complète player["Cote 1er septembre"] et renvoie
    events_par_tableau (voir stats.build_tournois)."""
    driver.get(f"https://myffbad.fr/joueur/{player['Licence']}")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "section[data-testid='player-layout']")))
    except TimeoutException:
        print("  chargement trop long, on continue quand même.")

    expand_all_sections(driver)
    events = {"Simple": results.parse_results(_soup(driver))}

    click_buttons_by_text(driver, "Journal de suivi")
    cote_sept = {"Simple": results.parse_journal_cote(_soup(driver))}

    click_buttons_by_text(driver, "Double")
    soup = _soup(driver)
    events["Double"] = results.parse_results(soup)
    cote_sept["Double"] = results.parse_journal_cote(soup)

    click_buttons_by_text(driver, "Mixte")
    soup = _soup(driver)
    events["Mixte"] = results.parse_results(soup)
    cote_sept["Mixte"] = results.parse_journal_cote(soup)

    player["Cote 1er septembre"] = cote_sept
    return events


def scrape_all(driver, players):
    all_stats = []
    for i, player in enumerate(players, 1):
        print(f"[{i}/{len(players)}] {player['Nom']}...")
        try:
            events = scrape_player(driver, player)
            all_stats.append(stats.build_player_stats(player, events))
        except Exception as e:
            print(f"  erreur, joueur ignoré : {e}")
    return all_stats


def main_debug(licence_ids):
    """Phase 1 : dump HTML brut pour inspection (voir dump_player_debug)."""
    if not licence_ids:
        licence_ids = [input("Licence du joueur à inspecter : ")]

    my_licence, my_password = get_credentials()
    driver = new_driver()
    try:
        print("Connexion...")
        if not login(driver, my_licence, my_password):
            print("Connexion échouée.")
            sys.exit(1)
        print("Connecté.")
        for licence in licence_ids:
            print(f"Joueur {licence}...")
            dump_player_debug(driver, licence)
    finally:
        driver.quit()


# seuil minimum (meilleur tableau, tous discipline confondues) pour être
# scrapé : chaque joueur nécessite une page authentifiée à plusieurs clics,
# tout le club prendrait beaucoup trop de temps.
SEUIL_TABLEAU_MIN = "R6"


def joueur_est_eligible(player):
    return roster.TABLEAUX.index(player["Meilleur tableau"]) <= roster.TABLEAUX.index(SEUIL_TABLEAU_MIN)


def main(licence_filter):
    print("Récupération du roster du club...")
    with roster.new_session() as session:
        players = roster.build_players(session)
    for player in players:
        pointsIC.compute_IC_points(player)

    if licence_filter:
        players = [p for p in players if p["Licence"] in licence_filter]
        print(f"{len(players)} joueur(s) sélectionné(s) (test).")
    else:
        total = len(players)
        players = [p for p in players if joueur_est_eligible(p)]
        print(f"{len(players)} joueurs avec au moins un classement {SEUIL_TABLEAU_MIN} ou mieux "
              f"(sur {total} au total).")

    my_licence, my_password = get_credentials()
    driver = new_driver()
    try:
        print("Connexion...")
        if not login(driver, my_licence, my_password):
            print("Connexion échouée.")
            sys.exit(1)
        print("Connecté.")
        all_stats = scrape_all(driver, players)
    finally:
        driver.quit()

    if not all_stats:
        print("Aucune donnée récupérée.")
        sys.exit(1)

    stats.normaliser_club(all_stats)
    output_path = os.path.join(APP_DIR, "statsJoueurs.xlsx")
    export_excel.write_stats_excel(all_stats, output_path)
    print(f"\n{len(all_stats)} joueur(s) traité(s) -> {output_path}")


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--debug" in args:
        args.remove("--debug")
        main_debug(args)
    else:
        main(set(args) if args else None)
