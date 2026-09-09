"""
Phase 1 : outil de reconnaissance.

Se connecte à myffbad.fr, ouvre la page d'un ou plusieurs joueurs, déplie
les sections repliables (Résultats de la saison, historique de classement,
...) et sauvegarde le HTML obtenu dans debug/joueur_<licence>.html.

Objectif : inspecter la structure réelle de ces sections (elles nécessitent
une connexion, donc invisibles publiquement) avant d'écrire le parsing et le
traitement complet.
"""
import os
import sys
import time
from getpass import getpass

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

import roster

DEBUG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug")


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


def login(driver, licence, password):
    driver.get("https://myffbad.fr/")
    try:
        form = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "section[data-testid='sign-in-form']")))
        form.find_element(By.CSS_SELECTOR, "input[name='licence']").send_keys(licence)
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
    Progression). Renvoie le nombre de boutons cliqués."""
    n = 0
    for btn in driver.find_elements(By.CSS_SELECTOR, "button[data-testid='button']"):
        if btn.text.strip() == text:
            try:
                click(driver, btn)
                n += 1
            except Exception as e:
                print(f"  clic bouton '{text}' : {e}")
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
    click_buttons_by_text(driver, "Journal de suivi")
    dump_html(driver, licence, "journal")

    # bascule tous les groupes Simple/Double/Mixte de la page (Résultats,
    # classement, progression) sur Double puis Mixte, un dump par état
    for tableau in ("Double", "Mixte"):
        click_buttons_by_text(driver, tableau)
        dump_html(driver, licence, tableau.lower())


if __name__ == "__main__":
    licence_ids = sys.argv[1:]
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
