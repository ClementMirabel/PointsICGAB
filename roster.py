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
