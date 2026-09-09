import requests
from bs4 import BeautifulSoup

LEAGUE = "3304"
COMMITTEE = "110"
CLUB = "604"
TOPS_URL = "https://myffbad.fr/recherche/les-tops"

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
        classement = cells[8].get_text(strip=True)
        points = cells[9].get_text(strip=True)
        # jeunes joueurs pas encore classés (Minibad, Poussin...) : pas de cote
        if not points:
            classement, points = None, None
        players[licence] = {
            "Nom": cells[3].get_text(strip=True),
            "Classement": classement,
            "Points": float(points) if points else None,
        }
    return players


def build_players(session):
    by_discipline = {d: fetch_discipline(session, d) for d in DISCIPLINES}

    players = {}
    for discipline_id, (_, sexe) in DISCIPLINES.items():
        for licence, infos in by_discipline[discipline_id].items():
            player = players.setdefault(licence, {
                "Licence": licence,
                "Nom": infos["Nom"],
                "Sexe": sexe,
                "Classement Actuel": {},
                "Points Actuel": {},
            })
            if infos["Classement"] is None:
                continue
            colonne, _ = DISCIPLINES[discipline_id]
            player["Classement Actuel"][colonne] = infos["Classement"]
            player["Points Actuel"][colonne] = infos["Points"]

    for player in players.values():
        for colonne in ("Simple", "Double", "Mixte"):
            player["Classement Actuel"].setdefault(colonne, DEFAULT_CLASSEMENT)
            player["Points Actuel"].setdefault(colonne, DEFAULT_POINTS)

    return list(players.values())
