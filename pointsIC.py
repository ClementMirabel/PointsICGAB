import os
import sys
import openpyxl

import roster

# écrit toujours le xlsx à côté de l'exécutable, peu importe le dossier
# "courant" depuis lequel l'app a été lancée (double-clic Finder/Explorer).
# Sur Mac, le binaire réel vit caché dans pointsIC.app/Contents/MacOS/bin ;
# le lanceur du bundle exporte POINTSIC_OUTPUT_DIR vers l'emplacement visible
# de l'app (là où l'utilisateur s'attend à voir apparaître le xlsx).
if getattr(sys, "frozen", False):
    APP_DIR = os.environ.get("POINTSIC_OUTPUT_DIR") or os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))


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
    output_path = os.path.join(APP_DIR, "pointsIC.xlsx")
    wb.save(output_path)
    return len(data), output_path


if __name__ == "__main__":
    try:
        with roster.new_session() as session:
            players = roster.build_players(session)

        results = [compute_IC_points(player) for player in players]
        written, output_path = write_summaries(results)
        print(f"{len(results)} joueurs traités, {written} écrits (Valeur IC >= 24) -> {output_path}")
    except Exception as e:
        print(f"Erreur : {e}")

    # ne bloque que si un humain est vraiment devant un terminal
    # (évite un EOFError quand le script tourne sans surveillance, ex. CI)
    if sys.stdin.isatty():
        input("\nAppuie sur Entrée pour fermer...")
