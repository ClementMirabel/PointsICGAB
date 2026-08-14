import requests
from bs4 import BeautifulSoup
import openpyxl

LEAGUE = "3304"
COMMITTEE = "110"
CLUB = "604"
TOPS_URL = "https://myffbad.fr/recherche/les-tops"

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


def compute_IC_points(player):
    t = ["N1", "N2", "N3", "R4", "R5", "R6", "D7", "D8", "D9", "P10", "P11", "P12"]
    best_t = 11
    best_t_pts = 0

    for k in player["Classement Actuel"].keys():
        val = t.index(player["Classement Actuel"][k])
        if val <= best_t:
            best_t = val
            best_t_pts = max(best_t_pts, player["Points Actuel"][k])

    player["Meilleur tableau"] = t[best_t]
    player["Points meilleur tableau"] = best_t_pts

    if player["Meilleur tableau"] == "D7":
        player["Valeur IC"] = 12
    elif player["Meilleur tableau"] == "R6":
        player["Valeur IC"] = 18
    elif player["Meilleur tableau"] == "R5":
        player["Valeur IC"] = 24
    elif player["Meilleur tableau"] == "R4":
        player["Valeur IC"] = 30
    elif player["Meilleur tableau"] == "N3":
        player["Valeur IC"] = 39
    elif player["Meilleur tableau"] == "N2":
        player["Valeur IC"] = 48
    elif player["Meilleur tableau"] == "D8":
        player["Valeur IC"] = 9
    elif player["Meilleur tableau"] == "D9":
        player["Valeur IC"] = 6
    elif player["Meilleur tableau"] == "P10":
        player["Valeur IC"] = 3
    elif player["Meilleur tableau"] == "P11":
        player["Valeur IC"] = 2
    elif player["Meilleur tableau"] == "P12":
        player["Valeur IC"] = 1
    elif player["Meilleur tableau"] == "N1":
        if (player["Sexe"] == "F" and (player["Points Actuel"]["Simple"]>=3500 or player["Points Actuel"]["Double"]>=3500 or player["Points Actuel"]["Mixte"]>=4400)) or (player["Sexe"] == "H" and (player["Points Actuel"]["Simple"]>=4400 or player["Points Actuel"]["Double"]>=4400 or player["Points Actuel"]["Mixte"]>=4400)):
            player["Valeur IC"] = 93
        elif (player["Sexe"] == "F" and (player["Points Actuel"]["Simple"]>=3200 or player["Points Actuel"]["Double"]>=3200 or player["Points Actuel"]["Mixte"]>=4100)) or (player["Sexe"] == "H" and (player["Points Actuel"]["Simple"]>=4100 or player["Points Actuel"]["Double"]>=4100 or player["Points Actuel"]["Mixte"]>=4100)):
            player["Valeur IC"] = 84
        elif (player["Sexe"] == "F" and (player["Points Actuel"]["Simple"]>=2800 or player["Points Actuel"]["Double"]>=2800 or player["Points Actuel"]["Mixte"]>=3500)) or (player["Sexe"] == "H" and (player["Points Actuel"]["Simple"]>=3500 or player["Points Actuel"]["Double"]>=3500 or player["Points Actuel"]["Mixte"]>=3500)):
            player["Valeur IC"] = 75
        elif (player["Sexe"] == "F" and (player["Points Actuel"]["Simple"]>=2700 or player["Points Actuel"]["Double"]>=2700 or player["Points Actuel"]["Mixte"]>=3400)) or (player["Sexe"] == "H" and (player["Points Actuel"]["Simple"]>=3400 or player["Points Actuel"]["Double"]>=3400 or player["Points Actuel"]["Mixte"]>=3400)):
            player["Valeur IC"] = 66
        else:
            player["Valeur IC"] = 57

    return player


def write_summaries(data):
    data = [player for player in data if player["Valeur IC"] >= 24]
    data.sort(key=lambda player: (
        -player["Valeur IC"],
        -player["Points meilleur tableau"],
        player["Nom"]
    ))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([
        "Sexe",
        "Nom",
        "S",
        "D",
        "M",
        "Pts S",
        "Pts D",
        "Pts M",
        "Meilleur tableau",
        "Points meilleur tableau",
        "Valeur IC"
    ])

    for player in data:
        ws.append([
            player["Sexe"],
            player["Nom"],
            player["Classement Actuel"]["Simple"],
            player["Classement Actuel"]["Double"],
            player["Classement Actuel"]["Mixte"],
            player["Points Actuel"]["Simple"],
            player["Points Actuel"]["Double"],
            player["Points Actuel"]["Mixte"],
            player["Meilleur tableau"],
            player["Points meilleur tableau"],
            player["Valeur IC"]
        ])
    wb.save("pointsIC.xlsx")
    return len(data)


if __name__ == "__main__":
    try:
        with requests.Session() as session:
            session.headers.update({
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                              "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
            })
            players = build_players(session)

        results = [compute_IC_points(player) for player in players]
        written = write_summaries(results)
        print(f"{len(results)} joueurs traités, {written} écrits (Valeur IC >= 24) -> pointsIC.xlsx")
    except Exception as e:
        print(f"Erreur : {e}")

    input("\nAppuie sur Entrée pour fermer...")
