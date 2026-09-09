import requests
from bs4 import BeautifulSoup

LEAGUE = "3304"
COMMITTEE = "110"
CLUB = "604"
CLUB_ACRONYM = "GAB38"
TOPS_URL = "https://myffbad.fr/recherche/les-tops"
SEARCH_URL = "https://myffbad.fr/recherche/joueur"

SEXE_CODE = {"Homme": "H", "Femme": "F"}

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")

# id du dropdown "Discipline" -> (colonne du classement joueur, sexe de la liste)
DISCIPLINES = {
    1: ("Simple", "H"),
    2: ("Simple", "F"),
    3: ("Double", "H"),
    4: ("Double", "F"),
    5: ("Mixte", "H"),
    6: ("Mixte", "F"),
}

# tableau le plus bas, utilisé quand un joueur n'apparaît pas dans une liste
DEFAULT_CLASSEMENT = "P12"
DEFAULT_POINTS = 0.0

# du meilleur (N1) au moins bon (P12)
TABLEAUX = ["N1", "N2", "N3", "R4", "R5", "R6", "D7", "D8", "D9", "P10", "P11", "P12"]

VALEUR_IC_PAR_TABLEAU = {
    "R6": 18, "R5": 24, "R4": 30, "N3": 39, "N2": 48,
    "D7": 12, "D8": 9, "D9": 6, "P10": 3, "P11": 2, "P12": 1,
}

# (seuil Hommes, seuil Femmes en Simple/Double, seuil Femmes en Mixte, valeur)
VALEUR_IC_N1_SEUILS = [
    (4400, 3500, 4400, 93),
    (4100, 3200, 4100, 84),
    (3500, 2800, 3500, 75),
    (3400, 2700, 3400, 66),
]


def valeur_ic(tableau, sexe, points_simple, points_double, points_mixte):
    """Barème IC : plus le tableau est haut, plus la valeur est élevée. Pour
    N1, la valeur dépend en plus du nombre de points (paliers différents
    Hommes/Femmes) dans n'importe lequel des 3 tableaux."""
    if tableau != "N1":
        return VALEUR_IC_PAR_TABLEAU.get(tableau, 0)

    for seuil_h, seuil_f, seuil_f_mixte, valeur in VALEUR_IC_N1_SEUILS:
        if sexe == "H" and (points_simple >= seuil_h or points_double >= seuil_h or points_mixte >= seuil_h):
            return valeur
        if sexe == "F" and (points_simple >= seuil_f or points_double >= seuil_f or points_mixte >= seuil_f_mixte):
            return valeur
    return 57


def new_session():
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session


def fetch_discipline(session, discipline_id):
    """Récupère la liste des joueurs du club pour une discipline (top complet du club)."""
    response = session.get(TOPS_URL, params={
        "league": LEAGUE,
        "committee": COMMITTEE,
        "club": CLUB,
        "isFirstLoad": "false",
        "disciplineId": discipline_id,
        "maxResults": 500,
    }, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    table = soup.find(attrs={"data-testid": "table"})

    players = {}
    for row in table.find("tbody").find_all("tr"):
        cells = row.find_all("td")
        licence = cells[4].get_text(strip=True)
        rang = cells[0].get_text(strip=True)
        classement = cells[8].get_text(strip=True)
        points = cells[9].get_text(strip=True)
        # jeunes joueurs pas encore classés (Minibad, Poussin...) : pas de cote
        if not points:
            classement, points, rang = None, None, None
        players[licence] = {
            "Nom": cells[3].get_text(strip=True),
            "Rang": int(rang) if rang else None,  # position dans le club pour ce tableau (1, 2, 3...)
            "Classement": classement,
            "Points": float(points) if points else None,
        }
    return players


def fetch_club_roster(session):
    """Liste complète des licenciés du club (page /recherche/joueur, triée
    par cote). C'est la référence pour la liste des joueurs et leur
    classement (lettre) par tableau : la page "les tops" a un filtre
    supplémentaire non documenté qui en exclut certains, y compris des
    joueurs bien classés (constaté : un joueur N3 en Simple absent des
    tops). L'API ne filtre pas non plus strictement par club malgré le
    paramètre "club" - on filtre nous-même sur CLUB_ACRONYM."""
    joueurs = {}
    page = 1
    while True:
        response = session.get(SEARCH_URL, params={
            "league": LEAGUE,
            "committee": COMMITTEE,
            "club": CLUB,
            "isFirstLoad": "false",
            "sort": "rating",
            "sortDirection": "desc",
            "page": page,
        }, timeout=30)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        table = soup.find(attrs={"data-testid": "search-players-results-table"})
        rows = table.find("tbody").find_all("tr") if table else []
        if not rows:
            break

        for row in rows:
            cells = row.find_all("td")
            if cells[2].get_text(strip=True) != CLUB_ACRONYM:
                continue
            tableaux = cells[4].get_text(" ", strip=True).split()  # "N3 N3 N2" -> [S, D, M]
            if len(tableaux) != 3:
                continue
            licence = cells[1].get_text(strip=True)
            # "NC" (non classé) ou "-" -> comme un joueur absent d'un tableau
            classement = [t if t in TABLEAUX else DEFAULT_CLASSEMENT for t in tableaux]
            joueurs[licence] = {
                "Nom": cells[0].get_text(strip=True),
                "Sexe": SEXE_CODE.get(cells[5].get_text(strip=True), "H"),
                "Classement": dict(zip(("Simple", "Double", "Mixte"), classement)),
            }

        if len(rows) < 50:  # dernière page
            break
        page += 1

    return joueurs


def build_players(session):
    """Union de /recherche/joueur et des 6 listes /recherche/les-tops :
    chacune des deux sources exclut des joueurs que l'autre a (constaté :
    des joueurs bien classés absents des tops, et d'autres absents de la
    recherche) - aucune des deux n'est complète seule."""
    joueurs = fetch_club_roster(session)
    by_discipline = {d: fetch_discipline(session, d) for d in DISCIPLINES}

    players = {}

    # 1) référence pour Nom/Sexe/lettre : tous les licenciés du club
    for licence, infos in joueurs.items():
        players[licence] = {
            "Licence": licence,
            "Nom": infos["Nom"],
            "Sexe": infos["Sexe"],
            "Classement Actuel": dict(infos["Classement"]),
            "Points Actuel": {},
            "Rang Actuel": {},
        }

    # 2) points/rang numériques (+ lettre exacte, cohérente avec la
    # recherche quand les deux sont dispo) là où les tops en ont, y compris
    # pour des joueurs absents de /recherche/joueur
    for discipline_id, (tableau, sexe) in DISCIPLINES.items():
        for licence, infos in by_discipline[discipline_id].items():
            if infos["Classement"] is None:
                continue
            player = players.setdefault(licence, {
                "Licence": licence,
                "Nom": infos["Nom"],
                "Sexe": sexe,
                "Classement Actuel": {},
                "Points Actuel": {},
                "Rang Actuel": {},
            })
            player["Classement Actuel"][tableau] = infos["Classement"]
            player["Points Actuel"][tableau] = infos["Points"]
            player["Rang Actuel"][tableau] = infos["Rang"]

    # 3) tableaux toujours pas renseignés (absent des deux, ou absent d'un
    # seul des 3 tableaux) -> valeurs par défaut (pas de rang connu)
    for player in players.values():
        for tableau in ("Simple", "Double", "Mixte"):
            player["Classement Actuel"].setdefault(tableau, DEFAULT_CLASSEMENT)
            player["Points Actuel"].setdefault(tableau, DEFAULT_POINTS)
            player["Rang Actuel"].setdefault(tableau, None)

    return list(players.values())
