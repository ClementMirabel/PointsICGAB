"""
Interclubs du club depuis icbad.ffbad.org (pages publiques, pas de
connexion) : divisions et classements de poule, joueurs de chaque équipe,
calendrier des rencontres et feuille de match de chaque rencontre jouée
(qui a joué quel match - SH1, SD2, DX1... - et le résultat).

    python icbad.py              # -> icbad.json (test rapide, ~1 min)

Complète les résultats joueurs de myffbad (statsJoueurs.py), qui disent
qu'un match a été joué en interclub et pour quelle équipe, mais pas sur
quel match de la rencontre (SH1 ou SH2...) ni la rencontre complète.

Structure HTML vérifiée sur les pages de la saison 2026-2027 :
- /instance/<CLUB> : un <h2 id="tableau<N>"> par division, suivi du tableau
  de classement de la poule (lignes du club en class="focus") ;
- /equipe/<id> : liste des joueurs (genre en icône, licence, classements,
  capitaine) puis calendrier (tr.clickable-row, lien /rencontre/<id>) ;
- /rencontre/<id> : table.rencontre, une ligne par match (code en <th>,
  joueurs de chaque côté, scores des sets, vainqueur en class
  winner-left-gradient / winner-right-gradient).
"""
import json
import os
import re
import sys
import time
from datetime import date

from bs4 import BeautifulSoup

import roster

BASE_URL = "https://icbad.ffbad.org"
CLUB = "GAB38"
PAUSE = 0.3  # entre deux requêtes - ~90 pages par run, inutile de charger le site
# équipes du club couvertes par le planificateur IC : pour celles-ci, on
# récupère aussi la page de chaque adversaire de la poule (composition type,
# joueurs et cotes) - les équipes 7+ jouent en départemental, hors planificateur
EQUIPES_PLANIFICATEUR = 6

_CODE_EQUIPE = re.compile(r"\(([^()]+-\d+)\)\s*$")       # "Nom du club (38-GAB-1)" -> "38-GAB-1"
_ID = re.compile(r"/(?:equipe|rencontre|joueur)/(\w+)")
_SCORE = re.compile(r"^\s*(\d+)\s*-\s*(\d+)\s*$")
_CODE_MATCH = re.compile(r"^(SH|SD|DH|DD|DX|DM|MX)\s*(\d*)$")
_DATE_CAL = re.compile(r"(\d{1,2})/(\d{1,2})(?:\D+(\d{1,2}):(\d{2}))?")


def _texte(el):
    return el.get_text(" ", strip=True) if el else ""


def _int(texte):
    try:
        return int(texte.strip().replace("−", "-"))
    except (ValueError, AttributeError):
        return None


def _float_pct(texte):
    """"66.67 %" -> 66.67."""
    try:
        return float(str(texte).replace("%", "").replace(",", ".").strip())
    except ValueError:
        return None


def _id(href):
    m = _ID.search(href or "")
    return m.group(1) if m else None


def _code_equipe(nom):
    m = _CODE_EQUIPE.search(nom or "")
    return m.group(1) if m else None


def _numero_equipe(code):
    """"38-GAB-3" -> 3."""
    m = re.search(r"-(\d+)$", code or "")
    return int(m.group(1)) if m else None


def division_courte(libelle):
    """Libellé de la liste des compétitions -> code court de division :
    "National 2 Poule 4" -> "N2", "Pré-National ..." -> "PN", "Régional 1
    ..." -> "R1", "Pré-Régional ..." -> "PR", "Départemental 4 ..." -> "D4"."""
    libelle = (libelle or "").strip()
    for motif, code in ((r"^Pr[ée]-?National", "PN"), (r"^Pr[ée]-?R[ée]gional", "PR"),
                        (r"^National\s*(\d)", "N{}"), (r"^R[ée]gional\s*(\d)", "R{}"),
                        (r"^D[ée]partemental\s*(\d)", "D{}"), (r"^Top\s*12", "TOP12")):
        m = re.match(motif, libelle, re.I)
        if m:
            return code.format(*m.groups())
    return libelle.split(" ")[0] if libelle else None


def _saison_debut(today=None):
    today = today or date.today()
    return today.year if today.month >= 9 else today.year - 1


# ------------------------------------------------------------------ parsing

def parse_instance(html):
    """Page club -> liste de divisions : titre, libellé court, code de
    division, classement complet de la poule (une ligne par équipe, club=True
    pour les équipes du club)."""
    soup = BeautifulSoup(html, "html.parser")
    libelles = {a["href"].lstrip("#"): _texte(a) for a in soup.select("a.scrollTo[href^='#tableau']")}
    divisions = []
    for h2 in soup.find_all("h2", id=re.compile(r"^tableau\d+")):
        table = h2.find_next("table")
        if table is None or not table.find("td", class_="nom-equipe"):
            continue
        classement = []
        for tr in table.find_all("tr"):
            nom_td = tr.find("td", class_="nom-equipe")
            if nom_td is None:
                continue
            tds = tr.find_all("td")
            valeurs = [_int(_texte(td)) for td in tds[tds.index(nom_td) + 1:]]
            valeurs += [None] * (11 - len(valeurs))
            a = nom_td.find("a")
            nom = _texte(nom_td)
            cles = ("j", "g", "n", "p", "f", "bonus", "penalites", "pts", "diff_matchs", "diff_sets", "diff_points")
            classement.append({
                "rang": _int(_texte(tds[0])), "nom": nom, "code": _code_equipe(nom),
                "equipe_id": _id(a["href"]) if a else None, "club": "focus" in (tr.get("class") or []),
                **dict(zip(cles, valeurs)),
            })
        libelle = libelles.get(h2["id"], "")
        divisions.append({"id": h2["id"], "titre": _texte(h2), "libelle": libelle,
                          "division": division_courte(libelle), "classement": classement})
    return divisions


def _est_mute(el):
    """Badge "M" (infobulle "Muté") d'icbad à côté du nom d'un joueur : joueur
    ayant changé de club - au plus 2 mutés alignés par rencontre (ICR AURA
    art. 13, CD38 art. 11-4)."""
    return any((ab.get("uk-tooltip") or "").startswith("Mut") for ab in el.find_all("abbr")) if el else False


def _classements_joueur(td):
    return [{"classement": _texte(div), "cote": _int(div.get("data-cote", ""))}
            for div in td.find_all("div", class_="ic-match-clsmt")]


def parse_equipe(html, saison_debut=None):
    """Page équipe -> joueurs (licence, nom, genre, classements S/D/M,
    capitaine, bilan icbad) et calendrier (journée, date, lieu, équipes,
    score, id de rencontre)."""
    soup = BeautifulSoup(html, "html.parser")
    saison_debut = saison_debut or _saison_debut()
    joueurs, vus = [], set()
    # seulement le tableau "Liste des joueurs" (en-tête "Licence") : des
    # liens /joueur/ apparaissent aussi ailleurs sur la page
    table_joueurs = next((t for t in soup.find_all("table")
                          if any(_texte(th) == "Licence" for th in t.find_all("th"))), None)
    for a in (table_joueurs.find_all("a", href=re.compile(r"/joueur/")) if table_joueurs else []):
        tr = a.find_parent("tr")
        if tr is None or _id(a["href"]) in vus:
            continue
        vus.add(_id(a["href"]))
        tds = tr.find_all("td")
        sexe = "F" if tr.find("i", class_="fa-female") else "H" if tr.find("i", class_="fa-male") else None
        chiffres = [td for td in tds if "uk-text-center" in (td.get("class") or []) and _texte(td)]
        derniers = [_texte(td) for td in chiffres[-3:]]
        joueurs.append({
            "licence": _id(a["href"]), "nom": _texte(a), "sexe": sexe,
            "classements": _classements_joueur(tr),
            "capitaine": any(_texte(ab) == "C" for ab in tr.find_all("abbr")),
            "mute": _est_mute(tr),
            "rencontres": _int(derniers[0]) if len(derniers) == 3 else None,
            "matchs": _int(derniers[1]) if len(derniers) == 3 else None,
            "pct": _float_pct(derniers[2]) if len(derniers) == 3 else None,
        })
    rencontres = []
    for tr in soup.select("tr.clickable-row"):
        lien = tr.find("a", href=re.compile(r"/rencontre/"))
        tds = tr.find_all("td")
        if lien is None or len(tds) < 6:
            continue
        quand = _DATE_CAL.search(_texte(lien))
        d = heure = None
        if quand:
            jour, mois = int(quand.group(1)), int(quand.group(2))
            d = date(saison_debut if mois >= 8 else saison_debut + 1, mois, jour).isoformat()
            heure = f"{quand.group(3)}:{quand.group(4)}" if quand.group(3) else None
        score = _SCORE.match(_texte(tds[4]))
        statut = tr.find("span", class_="rencontre-statut")
        rencontres.append({
            "id": _id(lien["href"]), "journee": _texte(tds[0]), "date": d, "heure": heure,
            "lieu": _texte(tds[2]).rstrip(","), "domicile": _texte(tds[3]), "exterieur": _texte(tds[5]),
            "score": [int(score.group(1)), int(score.group(2))] if score else None,
            "statut": statut.get("title") if statut else None,
        })
    return {"joueurs": joueurs, "rencontres": rencontres, "composition_type": parse_composition_type(soup)}


def parse_composition_type(soup):
    """Tableau "Composition type" d'une page équipe (joueurs habituellement
    alignés sur chaque match, d'après les rencontres jouées) -> [{code,
    joueurs: [{licence, nom, classement}]}]. Pas de cote ici : elle se
    trouve dans la liste des joueurs de la même page (croiser par licence)."""
    titre = soup.find(lambda t: t.name in ("h2", "h3") and "Composition type" in t.get_text())
    table = titre.find_next("table") if titre else None
    out = []
    for tr in (table.find_all("tr") if table else []):
        tds = tr.find_all("td")
        if len(tds) < 2:
            continue
        code = _texte(tds[0]).upper().replace(" ", "")
        if not _CODE_MATCH.match(code):
            continue
        joueurs = []
        for a in tds[1].find_all("a", href=re.compile(r"/joueur/")):
            clsmt = a.find_next_sibling("div", class_="ic-match-clsmt")
            joueurs.append({"licence": _id(a["href"]), "nom": _texte(a), "classement": _texte(clsmt) or None})
        out.append({"code": code, "joueurs": joueurs})
    return out


def _joueurs_cote(td):
    out = []
    for span in td.find_all("span", class_="joueur-nom"):
        a = span.find("a")
        ligne = span.find_parent("tr")
        clsmt = ligne.find("div", class_="ic-match-clsmt") if ligne else None
        attrs = span.find_next_sibling("span", class_="joueur-attrs")
        out.append({"licence": _id(a["href"]) if a else None, "nom": _texte(span), "mute": _est_mute(attrs),
                    "classement": _texte(clsmt) or None, "cote": _int(clsmt.get("data-cote", "")) if clsmt else None})
    return out


def parse_rencontre(html):
    """Feuille de rencontre -> équipes (gauche = qui reçoit, droite), score
    et un élément par match : code (SH1, DX2...), joueurs de chaque côté,
    sets (points gauche, points droite) et vainqueur ("gauche"/"droite",
    None si non joué)."""
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", class_="rencontre")
    if table is None:
        return None
    entete = table.find("tr")
    liens = entete.find_all("a", href=re.compile(r"/equipe/")) if entete else []
    ths = entete.find_all("th") if entete else []
    score = _SCORE.match(_texte(ths[2])) if len(ths) > 2 else None
    matchs = []
    for tr in table.find_all("tr", recursive=False):
        th = tr.find("th", recursive=False)
        tds = tr.find_all("td", recursive=False)
        code = _texte(th).upper().replace(" ", "") if th else ""
        if not _CODE_MATCH.match(code) or len(tds) < 3:
            continue
        sets = []
        for ligne in tds[1].find_all("tr"):
            pts = [_int(_texte(td)) for td in ligne.find_all("td")]
            if len(pts) == 2 and None not in pts:
                sets.append(pts)
        classes_g, classes_d = tds[0].get("class") or [], tds[2].get("class") or []
        vainqueur = ("gauche" if any("winner-left" in c for c in classes_g)
                     else "droite" if any("winner-right" in c for c in classes_d) else None)
        matchs.append({"code": code, "gauche": _joueurs_cote(tds[0]), "droite": _joueurs_cote(tds[2]),
                       "sets": sets, "vainqueur": vainqueur})
    return {
        "gauche": {"nom": _texte(liens[0]), "equipe_id": _id(liens[0]["href"])} if len(liens) > 0 else None,
        "droite": {"nom": _texte(liens[1]), "equipe_id": _id(liens[1]["href"])} if len(liens) > 1 else None,
        "score": [int(score.group(1)), int(score.group(2))] if score else None,
        "matchs": matchs,
    }


# ------------------------------------------------------------------ collecte

def _get(session, path):
    time.sleep(PAUSE)
    r = session.get(BASE_URL + path, timeout=30)
    r.raise_for_status()
    return r.text


def _orienter(rencontre_cal, detail, equipe_id, code):
    """Rencontre vue depuis le club : domicile ou non, adversaire, score
    club/adversaire, et matchs avec joueurs/sets côté club d'abord."""
    domicile = code in (rencontre_cal["domicile"] or "")
    adversaire = rencontre_cal["exterieur"] if domicile else rencontre_cal["domicile"]
    score = rencontre_cal["score"]
    out = {**{k: rencontre_cal[k] for k in ("id", "journee", "date", "heure", "lieu", "statut")},
           "domicile": domicile, "adversaire": adversaire, "adversaire_code": _code_equipe(adversaire),
           "score_club": None, "score_adv": None, "jouee": False, "matchs": []}
    if detail:
        club_a_gauche = (detail["gauche"] or {}).get("equipe_id") == str(equipe_id)
        if detail["score"]:
            score = detail["score"] if club_a_gauche else detail["score"][::-1]
        else:
            score = None
        for m in detail["matchs"]:
            club, adv = (m["gauche"], m["droite"]) if club_a_gauche else (m["droite"], m["gauche"])
            gagnant = m["vainqueur"]
            out["matchs"].append({
                "code": m["code"], "club": club, "adverse": adv,
                "sets": [s if club_a_gauche else s[::-1] for s in m["sets"]],
                "victoire": None if gagnant is None else (gagnant == "gauche") == club_a_gauche,
            })
    elif score and not domicile:
        score = score[::-1]
    if score:
        out["score_club"], out["score_adv"] = score
    out["jouee"] = bool(out["matchs"] and any(m["victoire"] is not None for m in out["matchs"])) or \
        bool(score and sum(score) > 0)
    return out


def collecter(session=None, club=CLUB, log=print):
    """Toutes les données interclubs du club : une entrée par équipe du club
    (numéro, division, classement de poule, joueurs, rencontres avec feuille
    de match pour celles déjà jouées)."""
    session = session or roster.new_session()
    divisions = parse_instance(_get(session, f"/instance/{club}"))
    equipes, adversaires = [], {}
    for div in divisions:
        for ligne in div["classement"]:
            if not ligne["club"] or not ligne["equipe_id"]:
                continue
            # équipes 7+ (départemental) hors périmètre de la page Interclubs
            numero_eq = _numero_equipe(ligne["code"])
            if numero_eq is not None and numero_eq > EQUIPES_PLANIFICATEUR:
                continue
            log(f"  équipe {ligne['code']} ({div['division']})...")
            page = parse_equipe(_get(session, f"/equipe/{ligne['equipe_id']}"))
            rencontres = []
            for r in page["rencontres"]:
                detail = None
                if r["score"] and sum(r["score"]) > 0 and r["id"]:
                    try:
                        detail = parse_rencontre(_get(session, f"/rencontre/{r['id']}"))
                    except Exception as e:  # une feuille illisible ne doit pas faire tomber le reste
                        log(f"    rencontre {r['id']} illisible : {e}")
                rencontres.append(_orienter(r, detail, ligne["equipe_id"], ligne["code"]))
            ids_poule = {c["code"]: c["equipe_id"] for c in div["classement"]}
            for r in rencontres:
                r["adversaire_equipe_id"] = ids_poule.get(r["adversaire_code"])
            numero = _numero_equipe(ligne["code"])
            if numero is not None and numero <= EQUIPES_PLANIFICATEUR:
                for c in div["classement"]:
                    if c["club"] or not c["equipe_id"] or c["equipe_id"] in adversaires:
                        continue
                    try:
                        adv = parse_equipe(_get(session, f"/equipe/{c['equipe_id']}"))
                        adversaires[c["equipe_id"]] = {
                            "nom": c["nom"], "code": c["code"], "joueurs": adv["joueurs"],
                            "composition_type": adv["composition_type"],
                            # leurs rencontres de la poule (scores) : fiche adversaire
                            "rencontres": [{k: r[k] for k in ("journee", "date", "domicile", "exterieur", "score")}
                                           for r in adv["rencontres"]],
                        }
                    except Exception as e:
                        log(f"    adversaire {c['code']} illisible : {e}")
            equipes.append({
                "equipe_id": ligne["equipe_id"], "nom": ligne["nom"], "code": ligne["code"],
                "numero": _numero_equipe(ligne["code"]), "division": div["division"],
                "division_titre": div["titre"], "division_libelle": div["libelle"],
                "rang": ligne["rang"], "nb_equipes": len(div["classement"]),
                "classement": div["classement"], "joueurs": page["joueurs"], "rencontres": rencontres,
            })
    equipes.sort(key=lambda e: (e["numero"] is None, e["numero"] or 0))
    return {"club": club, "url": f"{BASE_URL}/instance/{club}", "equipes": equipes, "adversaires": adversaires,
            "equipes_planificateur": EQUIPES_PLANIFICATEUR}


if __name__ == "__main__":
    data = collecter()
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icbad.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    nb_r = sum(len([r for r in e["rencontres"] if r["jouee"]]) for e in data["equipes"])
    print(f"{len(data['equipes'])} équipe(s), {nb_r} rencontre(s) jouée(s) -> {out}", file=sys.stderr)
