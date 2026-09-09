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
    tiers = {k: roster.TABLEAUX.index(player["Classement Actuel"][k]) for k in ("Simple", "Double", "Mixte")}
    best_t = min(tiers.values())
    # parmi les disciplines à ce meilleur tableau (tier), celle avec le plus
    # de points sert de référence pour "Points/Rang meilleur tableau"
    candidats = [k for k, val in tiers.items() if val == best_t]
    ref = max(candidats, key=lambda k: player["Points Actuel"][k])

    player["Meilleur tableau"] = roster.TABLEAUX[best_t]
    player["Points meilleur tableau"] = player["Points Actuel"][ref]
    player["Rang meilleur tableau"] = player.get("Rang Actuel", {}).get(ref)
    player["Valeur IC"] = roster.valeur_ic(
        player["Meilleur tableau"],
        player["Sexe"],
        player["Points Actuel"]["Simple"],
        player["Points Actuel"]["Double"],
        player["Points Actuel"]["Mixte"],
    )

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
