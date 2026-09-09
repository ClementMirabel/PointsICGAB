"""
Parsing des sections de la page joueur qui nécessitent une connexion
(Résultats de la saison, Évolution du classement).

Ces fonctions travaillent sur du HTML déjà chargé (BeautifulSoup), obtenu
via Selenium après connexion (voir statsJoueurs.py). Le format est validé
contre des dumps réels enregistrés dans debug/ (voir tests_results.py).
"""
import re
from datetime import date

MY_CLUB = "GAB38"


def _parse_date_slash(text):
    """'01/09/2026' -> date(2026, 9, 1)."""
    jour, mois, annee = text.split("/")
    return date(int(annee), int(mois), int(jour))


def parse_journal_cote(soup):
    """Cote au 1er septembre (ligne 'Initialisation saison ...' du tableau
    Journal de suivi), pour le tableau (Simple/Double/Mixte) actuellement
    affiché. None si le bouton "Journal de suivi" n'a pas été cliqué (la
    section est encore en mode graphique) ou si la ligne est introuvable.
    """
    section = soup.find(attrs={"data-testid": "player-ranking-elo"})
    if section is None:
        return None
    table = section.find(attrs={"data-testid": "table"})
    if table is None:
        return None
    tbody = table.find("tbody")
    if tbody is None:
        return None
    for row in tbody.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 8:
            continue
        commentaire = cells[5].get_text(strip=True)
        if "Initialisation saison" in commentaire:
            try:
                return float(cells[7].get_text(strip=True))
            except ValueError:
                return None
    return None


def parse_classement_evolution(soup):
    """Panel "Évolution classement" (data-testid='player-classement-
    evolution', page /joueur/<licence>/classement-historique) -> liste de
    {date, Simple: {rang, classement, points}, Double: {...}, Mixte: {...}}
    triée du plus récent au plus ancien (ordre du site). Identique quel que
    soit l'état des boutons Simple/Double/Mixte de la page (vérifié) : les
    3 tableaux sont toujours présents sur chaque ligne."""
    section = soup.find(attrs={"data-testid": "player-classement-evolution"})
    if section is None:
        return []
    table = section.find(attrs={"data-testid": "table"})
    if table is None:
        return []
    tbody = table.find("tbody")
    if tbody is None:
        return []

    entries = []
    for row in tbody.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 4:
            continue
        try:
            entry = {"date": _parse_date_slash(cells[0].get_text(strip=True))}
        except ValueError:
            continue

        for i, tableau in enumerate(("Simple", "Double", "Mixte"), start=1):
            texte = cells[i].get_text(" ", strip=True)
            badge = cells[i].find(attrs={"data-testid": "badge-ranking"})
            rang = None
            premier_mot = texte.split(" ", 1)[0] if texte else ""
            if premier_mot.isdigit():
                rang = int(premier_mot)
            points_match = re.search(r"([\d.]+)\s*points", texte)
            entry[tableau] = {
                "rang": rang,
                "classement": badge.get_text(strip=True) if badge else None,
                "points": float(points_match.group(1)) if points_match else None,
            }
        entries.append(entry)
    return entries


def find_september_1(entries, annee):
    """Ligne du 1er septembre <annee> (départ de saison) dans la liste
    renvoyée par parse_classement_evolution, ou None si absente."""
    cible = date(annee, 9, 1)
    return next((e for e in entries if e["date"] == cible), None)


def _parse_score(score_table, mine_idx):
    """Table de score (data-testid='row-details-score') -> (sets_gagnes,
    sets_perdus) du point de vue de mon côté (mine_idx, 0 ou 1 : PAS
    toujours la rangée du haut, voir _side_entries), ou None si illisible.
    Les scores non numériques (forfait...) sont ignorés set par set plutôt
    que de faire échouer tout le match."""
    if score_table is None:
        return None
    rows = score_table.find_all("tr")
    if len(rows) != 2:
        return None
    sides = [[c.get_text(strip=True) for c in r.find_all("td")] for r in rows]
    mine, theirs = sides[mine_idx], sides[1 - mine_idx]

    sets_gagnes = sets_perdus = 0
    for m, t in zip(mine, theirs):
        try:
            m_val, t_val = int(m), int(t)
        except ValueError:
            continue
        if m_val > t_val:
            sets_gagnes += 1
        elif t_val > m_val:
            sets_perdus += 1
    if sets_gagnes == sets_perdus:
        return None  # score incomplet/illisible : match ignoré
    return sets_gagnes, sets_perdus


def _side_entries(div):
    """Un des deux blocs de row-details-name -> liste de (nom, est_moi,
    licence). est_moi=True si l'élément n'est pas un lien : le site ne
    linke jamais vers son propre profil, seulement vers celui des autres
    joueurs. licence=None pour "moi", ou si le lien ne suit pas le format
    /joueur/<licence>."""
    entries = []
    for el in div.find_all(["a", "span"], recursive=False):
        text = el.get_text(strip=True)
        if not text:
            continue
        licence = None
        if el.name == "a":
            href = el.get("href", "")
            if href.startswith("/joueur/"):
                licence = href.rsplit("/", 1)[-1]
        entries.append((text, el.name != "a", licence))
    return entries


def _side_clubs(div):
    return [a.get_text(strip=True) for a in div.find_all("a")]


def parse_match_row(tr):
    """Une ligne <tr> de row-details (un match) -> dict, ou None si la ligne
    n'a pas de score exploitable (bye, forfait non chiffré...).

    Important : le site liste toujours le côté vainqueur en premier (rangée
    du haut), pas "moi" en premier - "moi" doit donc être identifié par
    l'absence de lien (voir _side_entries), pas par la position.
    """
    cells = tr.find_all("td", recursive=False)
    if len(cells) < 9:
        return None

    tour = cells[1].get_text(strip=True)

    club_div = cells[2].find(attrs={"data-testid": "row-details-club"})
    club_sides = club_div.find_all("div", recursive=False) if club_div else []
    clubs = [_side_clubs(side) for side in club_sides]

    name_div = cells[6].find(attrs={"data-testid": "row-details-name"})
    name_sides = name_div.find_all("div", recursive=False) if name_div else []
    entries = [_side_entries(side) for side in name_sides]

    mine_idx = next((i for i, side in enumerate(entries) if any(is_me for _, is_me, _ in side)), None)
    if mine_idx is None or len(entries) != 2:
        return None  # ne devrait pas arriver : on visite toujours sa propre page
    opp_idx = 1 - mine_idx
    mine_clubs = clubs[mine_idx] if len(clubs) > mine_idx else []

    # partenaire (double uniquement) : l'autre entrée de mon côté, celle qui
    # n'est pas "moi". None en simple (une seule entrée de mon côté).
    mine_side = entries[mine_idx]
    me_pos = next(i for i, (_, is_me, _) in enumerate(mine_side) if is_me)
    partner_pos = next((i for i in range(len(mine_side)) if i != me_pos), None)
    partenaire_nom = mine_side[partner_pos][0] if partner_pos is not None else None
    partenaire_licence = mine_side[partner_pos][2] if partner_pos is not None else None
    partenaire_club = (mine_clubs[partner_pos]
                        if partner_pos is not None and len(mine_clubs) > partner_pos else None)

    score = _parse_score(cells[7].find(attrs={"data-testid": "row-details-score"}), mine_idx)
    if score is None:
        return None
    sets_gagnes, sets_perdus = score

    tournoi_id = None
    action = cells[8].find(attrs={"data-testid": "row-details-action"})
    if action is not None:
        href = action.get("href", "")
        if href.startswith("/tournoi/resultats/"):
            tournoi_id = href.rsplit("/", 1)[-1]

    return {
        "tour": tour,
        "mes_clubs": mine_clubs,
        "clubs_adverses": clubs[opp_idx] if len(clubs) > opp_idx else [],
        "mes_noms": [n for n, _, _ in entries[mine_idx]],
        "partenaire_nom": partenaire_nom,
        "partenaire_licence": partenaire_licence,
        "partenaire_club": partenaire_club,
        "noms_adverses": [n for n, _, _ in entries[opp_idx]],
        "sets_gagnes": sets_gagnes,
        "sets_perdus": sets_perdus,
        "victoire": sets_gagnes > sets_perdus,
        "tournoi_id": tournoi_id,
    }


def parse_results(soup):
    """Section "Résultats" (data-testid='player-results') pour le tableau
    actuellement affiché (Simple/Double/Mixte) -> liste d'événements, chacun
    avec ses matchs imbriqués. Un "événement" est soit une rencontre
    d'Interclubs (evenement commence par "Interclubs"), soit un tournoi
    individuel."""
    section = soup.find(attrs={"data-testid": "player-results"})
    if section is None:
        return []
    table = section.find(attrs={"data-testid": "table"})
    if table is None:
        return []
    tbody = table.find("tbody")
    if tbody is None:
        return []

    events = []
    pending_event = None
    for row in tbody.find_all("tr", recursive=False):
        cells = row.find_all("td", recursive=False)
        if len(cells) >= 5:
            # ligne d'événement (date / évènement / tableau / points / action)
            evenement = cells[1].get_text(strip=True)
            pending_event = {
                "date": cells[0].get_text(strip=True),
                "evenement": evenement,
                "est_interclub": evenement.startswith("Interclubs"),
                "tableau_ou_rencontre": cells[2].get_text(strip=True),
                "points": cells[3].get_text(strip=True),
                "matchs": [],
            }
            events.append(pending_event)
            continue

        # ligne de détail (row-details), rattachée au dernier événement lu
        detail = row.find(attrs={"data-testid": "row-details"})
        if detail is None or pending_event is None:
            continue
        detail_table = detail.find(attrs={"data-testid": "table"})
        if detail_table is None:
            continue
        detail_tbody = detail_table.find("tbody")
        if detail_tbody is None:
            continue
        for match_tr in detail_tbody.find_all("tr", recursive=False):
            match = parse_match_row(match_tr)
            if match is not None:
                pending_event["matchs"].append(match)

    return events
