"""Observation de la saisie des feuilles de rencontre icbad pendant un week-end
d'interclubs : à chaque passage, un instantané horodaté de chaque rencontre du
jour des équipes 1 à 6 (statut, score, matchs déjà saisis) est ajouté à un
fichier JSON Lines. Sert à savoir si les feuilles sont remplies au fil de la
journée (onglet « Live » envisageable) ou en une fois le soir.

    python observation_live.py                       # rencontres d'aujourd'hui -> observation_live.jsonl
    python observation_live.py --date 2026-11-08 --fichier test.jsonl

Pages publiques icbad, une dizaine de requêtes par passage ; rien n'est écrit
s'il n'y a pas de rencontre ce jour-là."""
import argparse
import json
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import icbad
import roster

FICHIER = "observation_live.jsonl"


def rencontres_du_jour(session, jour):
    """[(numéro d'équipe, code, equipe_id, rencontre du calendrier)] pour la date `jour` (AAAA-MM-JJ)."""
    out = []
    for div in icbad.parse_instance(icbad._get(session, f"/instance/{icbad.CLUB}")):
        for ligne in div["classement"]:
            numero = icbad._numero_equipe(ligne["code"]) if ligne["club"] else None
            if numero is None or numero > icbad.EQUIPES_PLANIFICATEUR or not ligne["equipe_id"]:
                continue
            page = icbad.parse_equipe(icbad._get(session, f"/equipe/{ligne['equipe_id']}"))
            out += [(numero, ligne["code"], ligne["equipe_id"], r) for r in page["rencontres"] if r["date"] == jour]
    return out


def instantanes(session, jour, releve):
    """Un enregistrement par rencontre du jour, vue depuis le club."""
    lignes = []
    for numero, code, equipe_id, r in rencontres_du_jour(session, jour):
        detail = None
        if r["id"]:
            try:
                detail = icbad.parse_rencontre(icbad._get(session, f"/rencontre/{r['id']}"))
            except Exception as e:  # une feuille illisible ne doit pas empêcher les autres
                print(f"rencontre {r['id']} illisible : {e}")
        x = icbad._orienter(r, detail, equipe_id, code)
        lignes.append({
            "releve": releve, "jour": jour, "equipe": numero, "rencontre": x["id"], "heure": x["heure"],
            "adversaire": x["adversaire_code"] or x["adversaire"], "statut": x["statut"],
            # score affiché dans le calendrier de l'équipe et score de la feuille (peuvent différer pendant la saisie)
            "score_calendrier": r["score"], "score": [x["score_club"], x["score_adv"]],
            "feuille": detail is not None,
            "matchs_saisis": sum(1 for m in x["matchs"] if m["victoire"] is not None or m["sets"]),
            "matchs_total": len(x["matchs"]),
            "matchs": [{"code": m["code"], "victoire": m["victoire"], "sets": m["sets"],
                        "joueurs": [p.get("nom") for p in m["club"]]} for m in x["matchs"]],
        })
    return lignes


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--date", help="jour observé (AAAA-MM-JJ), aujourd'hui (heure de Paris) par défaut")
    parser.add_argument("--fichier", default=FICHIER)
    args = parser.parse_args()
    jour = args.date or datetime.now(ZoneInfo("Europe/Paris")).date().isoformat()
    releve = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    with roster.new_session() as session:
        lignes = instantanes(session, jour, releve)
    if not lignes:
        print(f"Pas de rencontre le {jour} : rien à relever.")
        return
    with open(args.fichier, "a", encoding="utf-8") as f:
        for ligne in lignes:
            f.write(json.dumps(ligne, ensure_ascii=False) + "\n")
    for l in lignes:
        print(f"équipe {l['equipe']} vs {l['adversaire']} : {l['statut'] or '?'} · "
              f"{l['score'][0]}-{l['score'][1]} · {l['matchs_saisis']}/{l['matchs_total']} matchs saisis")


if __name__ == "__main__":
    main()
