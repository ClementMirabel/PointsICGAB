"""
Valeur IC de tous les joueurs du club (classements publics myffbad) ->
pointsIC.xlsx, filtré (Valeur IC >= 24) et trié. Exécuté chaque semaine par
le workflow ; compute_IC_points est aussi utilisé par statsJoueurs.py.

    python pointsIC.py
"""
import os

import openpyxl

import roster

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
    # pas de try/except : une erreur doit faire échouer le job CI (code de
    # sortie non nul) plutôt que produire un export manquant "réussi"
    with roster.new_session() as session:
        players = roster.build_players(session)

    results = [compute_IC_points(player) for player in players]
    written, output_path = write_summaries(results)
    print(f"{len(results)} joueurs traités, {written} écrits (Valeur IC >= 24) -> {output_path}")
