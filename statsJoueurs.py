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
    """Déplie chaque section repliable (Résultats, classement, ...) - sauf
    celles déjà ouvertes (chevron "rotate-180", voir expand_section_by_
    text), pour ne pas les refermer par erreur : au moins un panneau
    ("Nombre de points / Classement") est ouvert par défaut sur la page
    joueur, contrairement aux autres."""
    count = len(driver.find_elements(By.CSS_SELECTOR, "div[data-testid='collapse-item']"))
    for i in range(count):
        items = driver.find_elements(By.CSS_SELECTOR, "div[data-testid='collapse-item']")
        if i >= len(items):
            break
        if items[i].find_elements(By.CSS_SELECTOR, "svg.rotate-180"):
            continue
        try:
            click(driver, items[i])
        except Exception as e:
            print(f"  clic section {i} : {e}")


def expand_section_by_text(driver, texte):
    """Déplie UN SEUL panneau repliable dont le libellé contient `texte`,
    sans toucher aux autres (plus rapide que expand_all_sections quand on
    n'a besoin que d'un panneau précis, ex: "historique classement" sur la
    page classement-historique, qui contient aussi Résultats/Progression
    qu'on n'a pas besoin d'ouvrir sur cette page).

    Comparaison insensible à la casse (défensif - le site n'est pas
    toujours cohérent sur la capitalisation de ses libellés, ex: un É
    majuscule ici, un é minuscule ailleurs pour un concept similaire).

    Piège identifié sur ce site précis : le libellé affiché d'un panneau
    ne correspond pas forcément à son data-testid ni au nom "logique" de
    la donnée qu'il contient - le tableau data-testid='player-classement-
    evolution' (classement/cote au fil du temps) est caché derrière un
    panneau labellisé "historique classement", pas "Évolution classement"
    (un panneau différent, sans rapport, qui existe aussi sur la page) -
    confirmé en inspectant le DOM réel avec l'utilisateur : chercher le
    mauvais texte le laissait fermé/vide en permanence, sans erreur,
    quel que soit le budget de retry alloué. À vérifier au cas par cas
    plutôt que de supposer que le libellé suit le nom de la donnée.

    Idempotent : si le panneau est déjà ouvert (chevron avec la classe
    "rotate-180", constaté sur le DOM réel - tous les autres panneaux de
    la page en sont dépourvus par défaut), on ne clique PAS - cliquer sur
    un panneau déjà ouvert le REFERME au lieu de l'ouvrir. Piège identifié
    sur "Nombre de points / Classement" : contrairement aux autres
    panneaux, celui-ci est ouvert par défaut sur la page joueur - notre
    clic systématique le refermait juste avant qu'on essaie de lire ses
    cartes de cote/classement live, qui ressortaient vides en boucle
    (jamais "en cours de rendu", donc aucun budget de retry n'y changeait
    quoi que ce soit)."""
    cible = texte.lower()
    for item in driver.find_elements(By.CSS_SELECTOR, "div[data-testid='collapse-item']"):
        if cible in (item.get_attribute("textContent") or "").lower():
            deja_ouvert = bool(item.find_elements(By.CSS_SELECTOR, "svg.rotate-180"))
            if not deja_ouvert:
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


def charger_tous_les_resultats(driver, container_testid="player-results", max_clics=30):
    """Le panneau Résultats n'affiche que les ~10 événements les plus
    récents par défaut, avec un bouton "Voir plus" qui charge un lot
    supplémentaire à chaque clic (constaté en le cherchant en direct sur le
    site : un joueur ayant joué près de 200 matchs dans la saison n'en
    montrait qu'une dizaine par tableau sans ça, sans qu'aucune pagination
    ne soit visible dans le HTML brut qu'on avait dumpé jusque-là).

    On NE PEUT PAS réutiliser click_buttons_by_text ici : ce bouton porte
    data-variant="primary", exactement comme le bouton Simple/Double/Mixte
    actif - le filtre "pas primary" de click_buttons_by_text (pensé pour
    ignorer le bouton de tableau déjà sélectionné) l'exclurait aussi.

    Clique tant que le bouton est présent, pas désactivé, ET que le clic
    précédent a effectivement ajouté des lignes (garde-fou contre un bouton
    qui resterait dans le DOM sans plus rien charger)."""
    selector = f"div[data-testid='{container_testid}'] button[data-testid='button']"
    lignes_selector = f"div[data-testid='{container_testid}'] tbody > tr"
    avant = len(driver.find_elements(By.CSS_SELECTOR, lignes_selector))
    for _ in range(max_clics):
        target = None
        for btn in driver.find_elements(By.CSS_SELECTOR, selector):
            label = (btn.get_attribute("textContent") or "").strip()
            if label == "Voir plus" and not btn.get_attribute("disabled"):
                target = btn
                break
        if target is None:
            return
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
            click(driver, target)
        except Exception as e:
            print(f"  clic 'Voir plus' : {e}")
            return
        apres = len(driver.find_elements(By.CSS_SELECTOR, lignes_selector))
        if apres <= avant:
            return  # le clic n'a rien chargé de plus : on s'arrête là
        avant = apres


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


def _filtrer_saison_courante(events):
    """Ne garde que les événements datés à partir du 1er septembre de la
    saison en cours. Filet de sécurité : le panneau Résultats est censé ne
    montrer que la saison en cours (les résultats plus anciens devraient
    être uniquement dans l'historique de classement), mais un bug côté site
    peut laisser passer des résultats de la saison précédente - on filtre
    nous-même plutôt que de faire confiance au site sur ce point."""
    debut_saison = date(_current_season_start_year(), 9, 1)
    return [e for e in events if stats.parse_date_fr(e["date"]) >= debut_saison]


def _stabilise(lire, taille, tentatives=6, pause=0.5, accepter_zero_stable=True):
    """Interroge `lire()` (une capture de l'état actuel du DOM) plusieurs
    fois de suite jusqu'à ce que deux lectures consécutives donnent la même
    `taille(valeur)`, ou jusqu'à `tentatives` lectures - le rendu React a
    fini d'ajouter des lignes.

    Se contenter du premier résultat "non vide" ne suffit pas : le tableau
    peut apparaître avec quelques lignes puis continuer à se remplir en
    plusieurs vagues, pas juste "vide" puis "tout d'un coup". C'est ce qui
    causait des résultats partiels (pas seulement des joueurs entièrement
    vides, déjà corrigé séparément) pour pas mal de joueurs.

    accepter_zero_stable=False : n'accepte jamais "stable à 0" comme fini,
    insiste jusqu'à épuisement de `tentatives` même si ça reste à 0 - "stable
    à 0" n'est pas la même chose que "stable, rendu terminé" : un panneau
    qui n'a pas encore commencé à se remplir renvoie 0 deux lectures de
    suite tout autant qu'un panneau réellement vide. Ce mode existe pour
    des panneaux SANS garde-fou de chargement en amont (voir
    _parse_evolution_with_retry, seul appelant à passer False) : sans lui,
    la boucle sortait dès la 2e lecture (~une seule pause) sans jamais
    utiliser le reste du budget. Par défaut (True) on accepte "stable à 0"
    dès la 2e lecture identique, ce qui reste correct pour les panneaux
    protégés par charger_tous_les_resultats en amont (Résultats, Journal de
    suivi) - une discipline réellement non jouée par un joueur (fréquent,
    ex: pas de Mixte) y est aussi fréquente qu'un panneau pas encore rendu,
    donc insister jusqu'au bout du budget à chaque fois ralentirait
    sensiblement le run complet du club pour rien."""
    valeur = None
    precedente = -1
    for i in range(tentatives):
        valeur = lire()
        actuelle = taille(valeur)
        if actuelle == precedente and (accepter_zero_stable or actuelle != 0):
            return valeur
        precedente = actuelle
        if i < tentatives - 1:
            time.sleep(pause)
    return valeur


def _parse_results_with_retry(driver, mon_nom=None, tentatives=6, pause=0.5):
    """results.parse_results, en attendant que le nombre de matchs se
    stabilise (voir _stabilise) plutôt que de s'arrêter au premier essai
    non vide - constaté : jusqu'à 16 joueurs sur 73 revenaient entièrement
    vides dans Bilan joueur avec l'ancienne version qui ne retentait même
    pas sur "0 événement" ; et au-delà de ça, des joueurs actifs revenaient
    avec des résultats incomplets parce qu'un premier essai non vide n'est
    pas toujours un essai COMPLET.

    `mon_nom` (player["Nom"]) est transmis à parse_results pour identifier
    "moi" de façon fiable dans chaque match (voir results.parse_match_row)
    plutôt que par la seule absence de lien, qui peut se tromper si un
    adversaire/partenaire hors-club n'a pas non plus de profil "linkable"."""
    return _stabilise(
        lambda: results.parse_results(_soup(driver), mon_nom),
        lambda events: sum(len(e["matchs"]) for e in events),
        tentatives, pause)


def _parse_evolution_with_retry(driver, tentatives=10, pause=1):
    """Même principe que _parse_results_with_retry, pour le panel "historique
    classement" (data-testid='player-classement-evolution') - avec un budget
    plus généreux (jusqu'à 10s) ET accepter_zero_stable=False : ce panneau
    n'a pas de garde-fou "Voir plus" en amont (charger_tous_les_resultats)
    comme Résultats/Journal de suivi, donc toute la robustesse contre un
    rendu lent repose uniquement sur cette boucle - contrairement aux
    disciplines Résultats/Journal, "vide" n'est pas un résultat normal ici
    (l'historique de classement existe toujours pour un joueur licencié),
    donc insister jusqu'au bout du budget avant d'accepter 0 est justifié
    (un seul appel par joueur, pas ×3 disciplines - l'impact sur la durée
    totale du run reste limité)."""
    return _stabilise(
        lambda: results.parse_classement_evolution(_soup(driver)),
        len,
        tentatives, pause, accepter_zero_stable=False)


def _parse_ranking_cards_with_retry(driver, tentatives=10, pause=1):
    """Même principe que _parse_evolution_with_retry, pour les cartes
    "Nombre de points / Classement" (voir results.parse_player_ranking_
    cards) - même budget généreux et accepter_zero_stable=False : ces
    cartes viennent d'apparaître suite au clic sur le panel juste avant,
    et n'ont pas de garde-fou "Voir plus" en amont non plus. Oubli commis
    une première fois (lecture unique sans retry) : revenait vide, donc
    "Points Actuel"/"Classement Actuel" retombaient sur la valeur
    évolution/roster au lieu de la valeur live visée."""
    return _stabilise(
        lambda: results.parse_player_ranking_cards(_soup(driver)),
        len,
        tentatives, pause, accepter_zero_stable=False)


def _parse_journal_with_retry(driver, tentatives=6, pause=0.5):
    """Même principe que _parse_results_with_retry, pour le journal de suivi
    (voir results.parse_journal_complet)."""
    return _stabilise(
        lambda: results.parse_journal_complet(_soup(driver)),
        len,
        tentatives, pause)


class JoueurPrive(Exception):
    """Le joueur a rendu ses résultats privés (RGPD) : /joueur/<licence>
    redirige vers /recherche/joueur au lieu de charger sa page. Cas normal
    et attendu, à distinguer d'un timeout réseau ou d'une vraie erreur de
    scraping - inutile de gaspiller des tentatives dessus."""


def _page_joueur_privee(driver):
    """True si la page a redirigé vers la recherche au lieu de charger la
    page joueur - constaté : /joueur/<licence> d'un joueur privé renvoie
    vers /recherche/joueur avec le message "Veuillez spécifier votre
    recherche" (data-testid='no-result')."""
    if "/recherche/joueur" in driver.current_url:
        return True
    return bool(driver.find_elements(By.CSS_SELECTOR, "[data-testid='no-result']"))


def _charger_page_joueur(driver, url, tentatives=3, pause=2):
    """Navigue vers `url` (une page /joueur/...) et attend qu'elle soit
    chargée, en retentant (nouvelle navigation complète, pas juste une
    attente plus longue) en cas de timeout - constaté : des échecs
    intermittents ("server busy" côté site, rien à voir avec le joueur)
    faisaient perdre silencieusement des joueurs qui ont pourtant de vrais
    résultats, sans aucune retentative (on continuait sur une page pas
    forcément chargée).

    Lève JoueurPrive dès qu'une redirection vers la recherche est détectée
    (voir _page_joueur_privee) - vérifié à la première tentative pour ne
    pas gaspiller tout le budget de retentatives à re-timeout dessus (la
    page de recherche, elle, charge très bien - ce n'est pas un timeout)."""
    for tentative in range(tentatives):
        driver.get(url)
        if _page_joueur_privee(driver):
            raise JoueurPrive
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "section[data-testid='player-layout']")))
            return True
        except TimeoutException:
            if tentative < tentatives - 1:
                print(f"  chargement trop long, nouvelle tentative ({tentative + 2}/{tentatives})...")
                time.sleep(pause)
    print("  chargement trop long après plusieurs tentatives, on continue quand même.")
    return False


def scrape_player(driver, player):
    """Scrape la page d'un joueur du roster : résultats des 3 tableaux et
    journal de suivi de la cote (page principale) + classement/place/points
    au 1er septembre (page classement-historique, panel "Évolution
    classement" - identique quel que soit l'état des boutons Simple/Double/
    Mixte). Complète player["Classement 1er septembre"]/["Journal cote"] et
    renvoie (events_par_tableau, nb_matchs_avant_filtre_saison) (voir
    stats.build_tournois). Ce 2e élément permet à scrape_all de distinguer
    "vraiment rien récupéré" (suspect, à retenter) de "des résultats
    existent mais aucun depuis le 1er septembre" (confirmé, pas la peine de
    retenter - constaté : un joueur avec un vrai historique mais 0 match
    CETTE saison déclenchait quand même une retentative complète inutile,
    coûteuse sur l'ensemble du run). Lève JoueurPrive si le joueur a rendu
    ses résultats privés."""
    _charger_page_joueur(driver, f"https://myffbad.fr/joueur/{player['Licence']}")

    # panneau "Résultats" (Ratio victoires/défaites et Progression ne sont
    # plus utilisés depuis le passage à la page classement-historique)
    expand_section_by_text(driver, "Résultats")
    charger_tous_les_resultats(driver)
    events = {"Simple": _parse_results_with_retry(driver, player["Nom"])}

    # limité à player-results : les autres groupes Simple/Double/Mixte de
    # la page (Nombre de points/Classement, Progression) n'ont plus besoin
    # d'être basculés. Chaque bascule de tableau repart avec seulement les
    # ~10 événements les plus récents - il faut recliquer "Voir plus".
    click_buttons_by_text(driver, "Double", container_testid="player-results")
    charger_tous_les_resultats(driver)
    events["Double"] = _parse_results_with_retry(driver, player["Nom"])

    click_buttons_by_text(driver, "Mixte", container_testid="player-results")
    charger_tous_les_resultats(driver)
    events["Mixte"] = _parse_results_with_retry(driver, player["Nom"])

    # panneau "Nombre de points / Classement", basculé en mode tableau
    # ("Journal de suivi") : historique (en général hebdomadaire) de la cote
    # sur la saison - plus fin que l'évolution de classement ci-dessous (qui
    # n'a que quelques points de mesure), sert aux stats min/max/moyenne de
    # cote. Le bouton "Voir plus" y existe potentiellement aussi (même
    # composant que pour Résultats) - même garde-fou par précaution.
    expand_section_by_text(driver, "Nombre de points / Classement")
    # cote/classement/rang fédéral LIVE (cartes "player-ranking-card"),
    # capturés tout de suite après le dépli du panel, avant les bascules
    # Journal de suivi/Double/Mixte ci-dessous (qui ne touchent que le
    # graphique, pas ces cartes) - voir results.parse_player_ranking_cards.
    ranking_live = _parse_ranking_cards_with_retry(driver)
    click_buttons_by_text(driver, "Journal de suivi", container_testid="player-ranking-elo")
    charger_tous_les_resultats(driver, container_testid="player-ranking-elo")
    journal = {"Simple": _parse_journal_with_retry(driver)}

    click_buttons_by_text(driver, "Double", container_testid="player-ranking-elo")
    charger_tous_les_resultats(driver, container_testid="player-ranking-elo")
    journal["Double"] = _parse_journal_with_retry(driver)

    click_buttons_by_text(driver, "Mixte", container_testid="player-ranking-elo")
    charger_tous_les_resultats(driver, container_testid="player-ranking-elo")
    journal["Mixte"] = _parse_journal_with_retry(driver)
    player["Journal cote"] = journal

    _charger_page_joueur(driver, f"https://myffbad.fr/joueur/{player['Licence']}/classement-historique")
    # le panneau repliable qui contient le tableau (data-testid='player-
    # classement-evolution') s'appelle "historique classement" sur le site
    # - PAS "Évolution classement" (un panneau différent et sans rapport,
    # confirmé en inspectant le DOM réel avec l'utilisateur : celui-ci ne
    # matchait jamais, laissant le tableau toujours fermé/vide).
    expand_section_by_text(driver, "historique classement")
    evolution = _parse_evolution_with_retry(driver)
    player["Classement 1er septembre"] = results.find_september_1(evolution, _current_season_start_year())

    # la cote "les tops" (roster, page publique) peut manquer un tableau où
    # le joueur a pourtant joué - filtre non documenté de cette page
    # publique (voir roster.fetch_club_roster), constaté sur un joueur N3 en
    # Simple qui ressortait avec une cote à 0. On corrige avec la valeur la
    # plus fiable disponible : en priorité ranking_live (cartes "Nombre de
    # points/Classement", page principale, valeur live) - la dernière ligne
    # de l'historique de classement (evolution[0]) ne sert qu'en secours,
    # car elle ne se met à jour que sporadiquement (par ex. tous les 6
    # mois, constaté) et peut donc être très en retard sur la valeur
    # actuelle. On récupère aussi le rang fédéral (national) au passage - à
    # ne pas confondre avec player["Rang Actuel"] (roster), qui est la
    # position du joueur DANS LE CLUB pour ce tableau, pas son rang
    # national (utilisé tel quel ailleurs pour pointsIC.py, où c'est le bon
    # niveau de détail).
    evolution_actuel = evolution[0] if evolution else None
    player["Rang National Actuel"] = {}
    for tableau in ("Simple", "Double", "Mixte"):
        bloc_live = ranking_live.get(tableau) or {}
        bloc_evolution = (evolution_actuel or {}).get(tableau) or {}
        rang = bloc_live.get("rang")
        player["Rang National Actuel"][tableau] = rang if rang is not None else bloc_evolution.get("rang")
        points = bloc_live.get("points")
        if points is None:
            points = bloc_evolution.get("points")
        if points is not None:
            player["Points Actuel"][tableau] = points
        classement = bloc_live.get("classement")
        if classement is not None:
            player["Classement Actuel"][tableau] = classement

    nb_avant_filtre = _nb_matchs(events)
    events = {tableau: _filtrer_saison_courante(es) for tableau, es in events.items()}
    return events, nb_avant_filtre


def _nb_matchs(events):
    return sum(len(e["matchs"]) for es in events.values() for e in es)


def scrape_all(driver, players, tentatives_si_vide=2):
    """Scrape tout le monde, avec un log par joueur exploitable dans les
    logs GitHub Actions (statut entre crochets - facile à repérer/grep) et
    un résumé en fin de run.

    Un joueur "accessible" (page chargée) mais 0 match récupéré CETTE
    SAISON est RETENTÉ en entier (nouvelle navigation + nouveau scrape
    complet, pas juste une relecture du DOM déjà chargé - voir _stabilise
    pour ce niveau-là, déjà en place) jusqu'à `tentatives_si_vide` fois
    avant d'accepter : constaté que des joueurs ayant pourtant de vrais
    résultats ressortaient encore vides malgré les filets de sécurité
    existants.

    Sauf si le scrape a bel et bien trouvé des résultats AVANT le filtre
    de saison (voir scrape_player, nb_avant_filtre) : ça veut dire que la
    page a été correctement chargée et lue, et que "0 match depuis le 1er
    septembre" est un résultat exact, pas un raté de chargement - retenter
    ne changerait rien et coûterait une navigation complète pour rien
    (constaté : un joueur avec un vrai historique de la saison précédente
    mais pas encore joué cette saison déclenchait quand même une
    retentative complète)."""
    all_stats = []
    compteurs = {"ok": 0, "vide_confirme": 0, "vide": 0, "prive": 0, "erreur": 0}
    for i, player in enumerate(players, 1):
        print(f"[{i}/{len(players)}] {player['Nom']}...")
        try:
            events = None
            for tentative in range(tentatives_si_vide):
                events, nb_avant_filtre = scrape_player(driver, player)
                if _nb_matchs(events) > 0:
                    break
                if nb_avant_filtre > 0:
                    break  # confirmé vide CETTE saison (historique réel ailleurs) : pas la peine de retenter
                if tentative < tentatives_si_vide - 1:
                    print(f"  [VIDE] 0 match récupéré, nouvelle tentative complète "
                          f"({tentative + 2}/{tentatives_si_vide})...")

            nb = _nb_matchs(events)
            all_stats.append(stats.build_player_stats(player, events))
            if nb > 0:
                print(f"  [OK] {nb} match(s)")
                compteurs["ok"] += 1
            elif nb_avant_filtre > 0:
                print(f"  [VIDE CONFIRMÉ] {player['Nom']} : historique réel mais 0 match depuis "
                      f"le 1er septembre - pas de retentative (résultat exact).")
                compteurs["vide_confirme"] += 1
            else:
                print(f"  [VIDE] {player['Nom']} (licence {player['Licence']}) : 0 match après "
                      f"{tentatives_si_vide} tentative(s) - à vérifier manuellement : "
                      f"https://myffbad.fr/joueur/{player['Licence']}")
                compteurs["vide"] += 1
        except JoueurPrive:
            print(f"  [PRIVÉ] {player['Nom']} : résultats non publics, ignoré.")
            compteurs["prive"] += 1
        except Exception as e:
            print(f"  [ERREUR] {player['Nom']} (licence {player['Licence']}) : {e}")
            compteurs["erreur"] += 1

    print(f"\nRésumé scraping : {compteurs['ok']} OK, {compteurs['vide_confirme']} vide(s) confirmé(s) "
          f"(0 match cette saison, historique réel), {compteurs['vide']} vide(s) après retry (à vérifier), "
          f"{compteurs['prive']} privé(s), {compteurs['erreur']} erreur(s) sur {len(players)} joueur(s).")
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

        print("Panneaux 'collapse-item' présents sur la page (avant tout clic) :")
        for i, item in enumerate(driver.find_elements(By.CSS_SELECTOR, "div[data-testid='collapse-item']")):
            label = (item.get_attribute("textContent") or "").strip()
            print(f"  [{i}] {label[:120]!r}")

        trouve = expand_section_by_text(driver, "historique classement")
        print(f"expand_section_by_text('historique classement') a trouvé/cliqué un panneau : {trouve}")
        evo_apres = _parse_evolution_with_retry(driver)
        print(f"Évolution classement APRES clic (avec retry) : {len(evo_apres)} ligne(s)")
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

    # profil RGPD-privé détecté depuis "les tops" (page publique, voir
    # roster._est_prive) : on sait déjà que /joueur/<licence> redirigera
    # (voir JoueurPrive) - autant s'épargner la navigation Selenium plutôt
    # que de le découvrir à la dure pour chacun d'eux.
    nb_prive_roster = sum(1 for p in players if p.get("Prive"))
    if nb_prive_roster:
        print(f"{nb_prive_roster} joueur(s) déjà repéré(s) comme privé(s) via le roster, "
              f"scraping Selenium sauté pour eux.")
        players = [p for p in players if not p.get("Prive")]

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
