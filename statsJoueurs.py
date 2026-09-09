"""
Bilan de performance par joueur (résultats de la saison, classement au 1er
septembre, indices, tournois) pour tout le club.

    python statsJoueurs.py                  # tout le roster -> statsJoueurs.xlsx
    python statsJoueurs.py 06627061 00123456 # juste ces licences (test rapide)
    python statsJoueurs.py --debug 06627061  # dump HTML brut pour inspection
    python statsJoueurs.py --check-collapse 06627061  # les panneaux ont-ils
        # besoin d'un clic pour livrer leurs données, ou sont-elles déjà
        # dans le DOM avant tout clic ? (diagnostic de perf)

Nécessite une connexion à myffbad.fr ("Résultats" et l'historique de
classement ne sont pas publics, contrairement à pointsIC.py) : licence et
mot de passe via MYFFBAD_LICENCE/MYFFBAD_PASSWORD, ou saisie interactive.
"""
import os
import sys
import time
from datetime import date
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


CLICK_PAUSE = 1  # laisse le temps au re-rendu React après un clic


def click(driver, element):
    driver.execute_script("arguments[0].click();", element)
    time.sleep(CLICK_PAUSE)


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


def expand_section_by_text(driver, texte):
    """Déplie UN SEUL panneau repliable dont le libellé contient `texte`,
    sans toucher aux autres (plus rapide que expand_all_sections quand on
    n'a besoin que d'un panneau précis, ex: "Évolution classement" sur la
    page classement-historique, qui contient aussi Résultats/Progression
    qu'on n'a pas besoin d'ouvrir sur cette page)."""
    for item in driver.find_elements(By.CSS_SELECTOR, "div[data-testid='collapse-item']"):
        if texte in (item.get_attribute("textContent") or ""):
            click(driver, item)
            return True
    return False


def click_buttons_by_text(driver, text, container_testid=None):
    """Clique tous les boutons portant ce texte. Par défaut, plusieurs
    groupes Simple/Double/Mixte coexistent sur la page (Résultats, Nombre
    de points/Classement, Progression) et sont tous cliqués - passer
    container_testid pour se limiter à un seul (plus rapide quand on n'a
    besoin de basculer qu'une seule section). Renvoie le nombre de boutons
    cliqués.

    Recherche les boutons un par un, juste avant chaque clic : cliquer sur
    l'un d'eux provoque un re-rendu React qui invalide (stale) les
    références des autres boutons déjà récupérés dans une même liste.
    """
    selector = "button[data-testid='button']"
    if container_testid:
        selector = f"div[data-testid='{container_testid}'] {selector}"

    n = 0
    for _ in range(10):  # garde-fou : jamais plus de 10 groupes sur la page
        target = None
        for btn in driver.find_elements(By.CSS_SELECTOR, selector):
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

    # page séparée /joueur/<licence>/classement-historique (pas un panel de
    # la page principale - lien "classement historique" en haut de page)
    driver.get(f"https://myffbad.fr/joueur/{licence}/classement-historique")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "section[data-testid='player-layout']")))
    except TimeoutException:
        print("  page classement-historique : chargement trop long, on continue quand même.")
    expand_all_sections(driver)
    dump_html(driver, licence, "historique_simple")
    for tableau in ("Double", "Mixte"):
        n = click_buttons_by_text(driver, tableau)
        print(f"  historique '{tableau}' : {n} bouton(s) cliqué(s)")
        dump_html(driver, licence, f"historique_{tableau.lower()}")


def _soup(driver):
    return BeautifulSoup(driver.page_source, "html.parser")


def _current_season_start_year(today=None):
    """La saison FFBad démarre le 1er septembre : avant cette date on est
    encore dans la saison qui a commencé l'année précédente."""
    today = today or date.today()
    return today.year if today.month >= 9 else today.year - 1


def _parse_results_with_retry(driver, tentatives=3, pause=0.7):
    """results.parse_results, avec un filet de sécurité : si des événements
    sont trouvés mais aucun match dedans, c'est probablement que le rendu
    du détail (row-details) n'était pas encore terminé au moment du
    snapshot - on retente avant de conclure "vraiment aucun match"."""
    events = []
    for _ in range(tentatives):
        events = results.parse_results(_soup(driver))
        nb_matchs = sum(len(e["matchs"]) for e in events)
        if not events or nb_matchs > 0:
            return events
        time.sleep(pause)
    return events


def scrape_player(driver, player):
    """Scrape la page d'un joueur du roster : résultats des 3 tableaux
    (page principale) + classement/place/points au 1er septembre (page
    classement-historique, panel "Évolution classement" - identique quel
    que soit l'état des boutons Simple/Double/Mixte). Complète
    player["Classement 1er septembre"] et renvoie events_par_tableau (voir
    stats.build_tournois)."""
    driver.get(f"https://myffbad.fr/joueur/{player['Licence']}")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "section[data-testid='player-layout']")))
    except TimeoutException:
        print("  chargement trop long, on continue quand même.")

    # seul le panneau "Résultats" nous sert sur cette page (Nombre de
    # points/Classement, Ratio victoires/défaites et Progression ne sont
    # plus utilisés depuis le passage à la page classement-historique)
    expand_section_by_text(driver, "Résultats")
    events = {"Simple": _parse_results_with_retry(driver)}

    # limité à player-results : les autres groupes Simple/Double/Mixte de
    # la page (Nombre de points/Classement, Progression) n'ont plus besoin
    # d'être basculés
    click_buttons_by_text(driver, "Double", container_testid="player-results")
    events["Double"] = _parse_results_with_retry(driver)

    click_buttons_by_text(driver, "Mixte", container_testid="player-results")
    events["Mixte"] = _parse_results_with_retry(driver)

    driver.get(f"https://myffbad.fr/joueur/{player['Licence']}/classement-historique")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "section[data-testid='player-layout']")))
    except TimeoutException:
        print("  chargement trop long (classement-historique), on continue quand même.")
    expand_section_by_text(driver, "Évolution classement")
    evolution = results.parse_classement_evolution(_soup(driver))
    player["Classement 1er septembre"] = results.find_september_1(evolution, _current_season_start_year())

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


def check_collapse_hypothesis(licence_id):
    """Diagnostic : les panneaux repliables (Résultats, Évolution
    classement) affichent-ils des données déjà présentes dans le DOM avant
    tout clic (juste cachées en CSS, comme pour le détail d'un match dans
    un tournoi), ou seulement après avoir cliqué sur le panneau (contenu
    monté à la demande) ? Si c'est le premier cas, on peut supprimer les
    clics d'expansion et gagner encore du temps de scraping."""
    my_licence, my_password = get_credentials()
    driver = new_driver()
    try:
        print("Connexion...")
        if not login(driver, my_licence, my_password):
            print("Connexion échouée.")
            return
        print("Connecté.")

        driver.get(f"https://myffbad.fr/joueur/{licence_id}")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "section[data-testid='player-layout']")))
        except TimeoutException:
            pass

        events_avant = results.parse_results(_soup(driver))
        nb_avant = sum(len(e["matchs"]) for e in events_avant)
        print(f"Résultats AVANT clic : {len(events_avant)} événement(s), {nb_avant} match(s)")

        expand_section_by_text(driver, "Résultats")
        events_apres = results.parse_results(_soup(driver))
        nb_apres = sum(len(e["matchs"]) for e in events_apres)
        print(f"Résultats APRES clic : {len(events_apres)} événement(s), {nb_apres} match(s)")

        driver.get(f"https://myffbad.fr/joueur/{licence_id}/classement-historique")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "section[data-testid='player-layout']")))
        except TimeoutException:
            pass

        evo_avant = results.parse_classement_evolution(_soup(driver))
        print(f"Évolution classement AVANT clic : {len(evo_avant)} ligne(s)")

        expand_section_by_text(driver, "Évolution classement")
        evo_apres = results.parse_classement_evolution(_soup(driver))
        print(f"Évolution classement APRES clic : {len(evo_apres)} ligne(s)")
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
    if "--check-collapse" in args:
        args.remove("--check-collapse")
        check_collapse_hypothesis(args[0] if args else input("Licence à tester : "))
    elif "--debug" in args:
        args.remove("--debug")
        main_debug(args)
    else:
        main(set(args) if args else None)
