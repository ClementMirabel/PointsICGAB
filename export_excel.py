"""Écriture du classeur Excel à 8 onglets (SH/SD/DH/DD/MX, Tournois, Bilan
joueur, Stats club) à partir des stats produites par stats.py, avec mise en
forme conditionnelle (classements colorés comme sur le site, dégradé de
performance par colonne)."""
from collections import Counter, defaultdict

import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import stats

DISCIPLINE_SHEETS = [
    # (titre, tableau, sexe ou None si le tableau n'est pas séparé par genre)
    ("SH", "Simple", "H"),
    ("SD", "Simple", "F"),
    ("DH", "Double", "H"),
    ("DD", "Double", "F"),
    ("MX", "Mixte", None),
]

# ---------------------------------------------------------------- couleurs


def _uni(color):
    """PatternFill "solid" utilisable en mise en forme conditionnelle :
    Excel affiche bgColor (pas fgColor) pour un remplissage solid dans un
    dxf de mise en forme conditionnelle - contrairement à un remplissage de
    cellule classique. On met les deux pour être sûr que ça s'affiche."""
    return PatternFill(fill_type="solid", start_color=color, end_color=color)


FILL_N = _uni("FFE74C3C")  # rouge
FILL_R = _uni("FF3498DB")  # bleu
FILL_D = _uni("FF27AE60")  # vert
FONT_BLANC = Font(color="FFFFFFFF")
FONT_ROUGE = Font(color="FFE74C3C")

FILL_ENTETE = _uni("FFE8E8E8")  # gris clair
FILL_TITRE_BLOC = _uni("FFD9D9D9")  # gris un peu plus soutenu, pour le titre d'un mini-tableau
FONT_GRAS = Font(bold=True)


def _styliser_ligne(ws, row, last_col, fill=FILL_ENTETE, gras=True):
    """Gras + fond gris sur une ligne (colonnes 1..last_col) - sans ça, une
    ligne d'en-tête ou de titre ne se distingue pas visuellement d'une
    ligne de données, ce qui rend un onglet à plusieurs tableaux empilés
    difficile à lire (pas de repère pour savoir où un bloc commence)."""
    for col in range(1, last_col + 1):
        cell = ws.cell(row=row, column=col)
        if gras:
            cell.font = FONT_GRAS
        cell.fill = fill


def _styliser_mini_tableau(ws, titre_row, header_row, last_row, last_col):
    """Style un mini-tableau empilé (titre + en-tête + données) : titre en
    gras sur fond gris soutenu, en-tête en gras sur fond gris clair,
    bordure fine autour de l'ensemble - pour qu'il se détache visuellement
    des tableaux voisins sur le même onglet (sans ça, plusieurs
    mini-tableaux à la suite se lisent comme un seul mur de chiffres)."""
    _styliser_ligne(ws, titre_row, last_col, fill=FILL_TITRE_BLOC)
    _styliser_ligne(ws, header_row, last_col, fill=FILL_ENTETE)
    _encadrer_bloc(ws, titre_row, last_row, last_col)


def _encadrer_bloc(ws, first_row, last_row, last_col):
    """Bordure fine tout autour d'un bloc (titre + en-tête + données) -
    sépare visuellement les mini-tableaux empilés les uns des autres."""
    for row in range(first_row, last_row + 1):
        for col in range(1, last_col + 1):
            cell = ws.cell(row=row, column=col)
            border = cell.border
            cell.border = Border(
                left=BORDURE_LEGERE if col == 1 else border.left,
                right=BORDURE_LEGERE if col == last_col else border.right,
                top=BORDURE_LEGERE if row == first_row else border.top,
                bottom=BORDURE_LEGERE if row == last_row else border.bottom,
            )

# ordre tableau : une couleur par combinaison (pas juste par discipline en
# tête) - même famille de teinte selon la discipline en tête (vert=Simple,
# bleu=Double, orange=Mixte), une nuance plus soutenue pour la 2e place.
FILL_S_D = _uni("FFD5E8D4")  # Simple > Double > Mixte
FILL_S_M = _uni("FFA9D18E")  # Simple > Mixte > Double
FILL_D_S = _uni("FFDAE8FC")  # Double > Simple > Mixte
FILL_D_M = _uni("FF9DC3E6")  # Double > Mixte > Simple
FILL_M_S = _uni("FFFFE6CC")  # Mixte > Simple > Double
FILL_M_D = _uni("FFF4B183")  # Mixte > Double > Simple

ORDRE_TABLEAU_COULEURS = {
    "Simple > Double > Mixte": FILL_S_D,
    "Simple > Mixte > Double": FILL_S_M,
    "Double > Simple > Mixte": FILL_D_S,
    "Double > Mixte > Simple": FILL_D_M,
    "Mixte > Simple > Double": FILL_M_S,
    "Mixte > Double > Simple": FILL_M_D,
}

BORDURE_GROUPE = Side(style="thick", color="FF999999")
BORDURE_LEGERE = Side(style="thin", color="FFCCCCCC")

# colonnes texte qu'on ne colore pas en dégradé (déjà colorées autrement,
# ou pas un indicateur de performance)
CLASSEMENT_HEADERS = {
    "Classement",
    "Classement sept. S", "Classement sept. D", "Classement sept. M",
    "Classement actuel S", "Classement actuel D", "Classement actuel M",
}
TEXT_HEADERS = {"Nom", "Sexe", "Mois", "Catégorie", "Meilleur partenaire", "Meilleur partenaire (D+M)",
                "Meilleur partenaire au club (si différent)", "Ordre tableau",
                "Meilleure victoire", "Meilleure victoire (score)",
                "Pire défaite", "Pire défaite (score)",
                "Score max infligé S", "Score max infligé D", "Score max infligé M",
                "Score max reçu S", "Score max reçu D", "Score max reçu M",
                "Résultat", "Tournoi", "Joueur A", "Joueur B", "Type", "Joueur",
                "Tableau", "Plus grosse perf", "Plus grosse perf (score)",
                "Plus grosse contre-perf", "Plus grosse contre-perf (score)",
                "Plus gros écart de cote", "Résultat (écart max)", "Score (écart max)"} | CLASSEMENT_HEADERS


def _est_colonne_delta_partenaire(header):
    """Delta % avec/sans partenaire du club : presque toujours proche de 0
    ou positif, une simple barre verte suffit (pas besoin d'un dégradé
    rouge-blanc-vert par-dessus)."""
    return "Delta" in header


def _est_colonne_diff(header):
    """Colonnes pouvant être négatives (diff cote, gain de places/tableau,
    tendance hebdo) : dégradé divergent rouge-blanc-vert centré sur 0,
    plutôt que le dégradé blanc->vert habituel (qui n'aurait pas de sens sur
    une échelle qui traverse 0)."""
    return "Diff" in header or "Gain" in header or "Tendance" in header


def _pct(value):
    return round(value, 1)


def _colorer_classement(ws, col_idx, first_row, last_row):
    """N=rouge, R=bleu, D=vert, comme les badges du site. P10-P12 non colorés."""
    if last_row < first_row:
        return
    col = get_column_letter(col_idx)
    plage = f"{col}{first_row}:{col}{last_row}"
    ref = f"{col}{first_row}"
    for lettre, fill in (("N", FILL_N), ("R", FILL_R), ("D", FILL_D)):
        ws.conditional_formatting.add(
            plage, FormulaRule(formula=[f'LEFT({ref},1)="{lettre}"'], fill=fill, font=FONT_BLANC))


def _colorer_ordre_tableau(ws, col_idx, first_row, last_row):
    """Une couleur différente pour chacune des 6 combinaisons possibles
    (pas seulement 3 selon la seule discipline en tête)."""
    if last_row < first_row:
        return
    col = get_column_letter(col_idx)
    plage = f"{col}{first_row}:{col}{last_row}"
    ref = f"{col}{first_row}"
    for combo, fill in ORDRE_TABLEAU_COULEURS.items():
        ws.conditional_formatting.add(plage, FormulaRule(formula=[f'{ref}="{combo}"'], fill=fill))


def _degrade(ws, col_idx, first_row, last_row):
    """Dégradé blanc (bas) -> vert (haut) sur la colonne."""
    if last_row < first_row:
        return
    col = get_column_letter(col_idx)
    plage = f"{col}{first_row}:{col}{last_row}"
    rule = ColorScaleRule(
        start_type="min", start_color="FFFFFFFF",
        end_type="max", end_color="FF63BE7B",
    )
    ws.conditional_formatting.add(plage, rule)


def _barre_verte(ws, col_idx, first_row, last_row):
    """Barre de données verte (longueur proportionnelle à la valeur) + texte
    rouge quand la valeur est négative - pour le delta % avec/sans
    partenaire du club (un dégradé rouge-blanc-vert par-dessus une barre
    rendait la colonne difficile à lire pour un delta qui reste presque
    toujours proche de 0 ou positif).

    Excel sait nativement faire une barre en "remplissage plein" avec les
    valeurs négatives en rouge et un axe à 0 (noir) au milieu de la cellule,
    mais c'est une fonctionnalité étendue (Excel 2010+, extension x14 du
    format) que openpyxl n'expose pas du tout dans son API d'écriture -
    seule la barre "classique" (une couleur, dégradé, sans axe réglable) est
    accessible ici. On s'en approche avec une barre verte + police rouge sur
    le négatif ; le réglage exact (remplissage plein, axe noir au centre)
    reste à faire à la main dans Excel si besoin - clic droit sur la colonne
    > Mise en forme conditionnelle > Barres de données > Autres règles, une
    fois le classeur ouvert."""
    if last_row < first_row:
        return
    col = get_column_letter(col_idx)
    plage = f"{col}{first_row}:{col}{last_row}"
    ws.conditional_formatting.add(plage, DataBarRule(
        start_type="min", end_type="max", color="FF63BE7B", showValue=True))
    ws.conditional_formatting.add(
        plage, FormulaRule(formula=[f"{col}{first_row}<0"], font=FONT_ROUGE))


def _degrade_diverge(ws, col_idx, first_row, last_row):
    """Rouge (négatif) -> blanc (zéro) -> vert (positif), sans barre de
    données par-dessus (une barre en plus du dégradé rendait la colonne
    difficile à lire) - pour gain de places/tableau, diff cote, qui
    peuvent être négatives."""
    if last_row < first_row:
        return
    col = get_column_letter(col_idx)
    plage = f"{col}{first_row}:{col}{last_row}"
    couleurs = ColorScaleRule(
        start_type="min", start_color="FFE74C3C",
        mid_type="num", mid_value=0, mid_color="FFFFFFFF",
        end_type="max", end_color="FF63BE7B",
    )
    ws.conditional_formatting.add(plage, couleurs)


def _appliquer_mise_en_forme(ws, headers, first_row, last_row):
    """Classements colorés par tableau, dégradé de performance sur les
    autres colonnes numériques (divergent pour les diffs, barre verte pour
    le delta partenaire), % affiché avec un signe pourcentage - une seule
    fois toutes les lignes écrites."""
    for idx, header in enumerate(headers, start=1):
        if header in CLASSEMENT_HEADERS:
            _colorer_classement(ws, idx, first_row, last_row)
        elif header == "Ordre tableau":
            _colorer_ordre_tableau(ws, idx, first_row, last_row)
        elif header in TEXT_HEADERS:
            pass
        elif _est_colonne_delta_partenaire(header):
            _barre_verte(ws, idx, first_row, last_row)
        elif _est_colonne_diff(header):
            _degrade_diverge(ws, idx, first_row, last_row)
        else:
            _degrade(ws, idx, first_row, last_row)

        if "%" in header and last_row >= first_row:
            # les valeurs sont déjà sur une échelle 0-100 (pas 0-1), donc un
            # format pourcent standard afficherait x100 en trop - on ajoute
            # juste le signe "%" à l'affichage sans re-multiplier.
            col = get_column_letter(idx)
            for row in range(first_row, last_row + 1):
                ws[f"{col}{row}"].number_format = '0.0"%"'


def _grouper(headers, *paires):
    """paires : (premier_header, dernier_header) de chaque groupe de
    colonnes à encadrer -> liste de (première_colonne, dernière_colonne),
    1-indexées comme Excel."""
    groupes = []
    for debut, fin in paires:
        i0 = headers.index(debut)
        i1 = headers.index(fin)
        groupes.append((i0 + 1, i1 + 1))
    return groupes


def _encadrer_groupe(ws, first_col, last_col, first_row, last_row):
    """Bordure épaisse autour d'un bloc de colonnes (première ligne =
    en-tête, y compris la/les ligne(s) de sur-en-tête s'il y en a) - aide à
    repérer visuellement les groupes de colonnes apparentées (classement/
    cote, matchs/victoires/%, indices...)."""
    if last_row < first_row:
        return
    for row in range(first_row, last_row + 1):
        for col in range(first_col, last_col + 1):
            cell = ws.cell(row=row, column=col)
            border = cell.border
            cell.border = Border(
                left=BORDURE_GROUPE if col == first_col else border.left,
                right=BORDURE_GROUPE if col == last_col else border.right,
                top=BORDURE_GROUPE if row == first_row else border.top,
                bottom=BORDURE_GROUPE if row == last_row else border.bottom,
            )


def _lignes_legeres(ws, headers, noms_colonnes, first_row, last_row):
    """Trait fin sur le bord gauche de chacune des colonnes nommées, de
    first_row à last_row - divise des sous-groupes à l'intérieur d'un même
    bloc thématique (ex: Classement septembre vs Classement actuel), plus
    discret que la bordure épaisse qui encadre le bloc entier."""
    if last_row < first_row:
        return
    for nom in noms_colonnes:
        col_idx = headers.index(nom) + 1
        for row in range(first_row, last_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            border = cell.border
            cell.border = Border(left=BORDURE_LEGERE, right=border.right,
                                  top=border.top, bottom=border.bottom)


def _ajouter_filtre(ws, nb_colonnes, header_row, last_row):
    """Filtre Auto Excel (Data > Filter) sur la ligne des noms de colonnes
    (pas la ligne de sur-en-tête s'il y en a une), pour trier/filtrer
    facilement sans dérouler un menu Excel à la main."""
    if last_row < header_row:
        return
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(nb_colonnes)}{last_row}"


def _ecrire_sur_entete(ws, headers, niveau1, niveau2=()):
    """Écrit 2 lignes de sur-en-tête AVANT les noms de colonnes (donc à
    appeler avant tout ws.append(headers)) : des cellules fusionnées
    libellées au-dessus de chaque groupe thématique de colonnes -
    complète les bordures (qui délimitent le groupe visuellement) par un
    texte qui dit ce qu'il représente.

    niveau1 : liste de (label, premier_header, dernier_header) - grands
    thèmes, sur la ligne 1. niveau2 : liste de (label, premier_header,
    dernier_header) - sous-thèmes plus fins, sur la ligne 2, seulement pour
    les colonnes qui en ont besoin (les autres : le libellé niveau1 s'étend
    verticalement sur les 2 lignes plutôt que de laisser la ligne 2 vide en
    dessous). Renvoie le numéro de ligne où écrire `headers` ensuite (les
    noms de colonnes)."""
    def _colonnes(debut, fin):
        return headers.index(debut) + 1, headers.index(fin) + 1

    def _a_des_enfants(c0, c1):
        return any(_colonnes(d, f)[0] >= c0 and _colonnes(d, f)[1] <= c1 for _, d, f in niveau2)

    for label, debut, fin in niveau1:
        c0, c1 = _colonnes(debut, fin)
        r1, r2 = (1, 1) if _a_des_enfants(c0, c1) else (1, 2)
        if c1 > c0 or r2 > r1:
            ws.merge_cells(start_row=r1, start_column=c0, end_row=r2, end_column=c1)
        cell = ws.cell(row=1, column=c0, value=label)
        cell.font = Font(bold=True)
        cell.fill = FILL_TITRE_BLOC
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for label, debut, fin in niveau2:
        c0, c1 = _colonnes(debut, fin)
        if c1 > c0:
            ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c1)
        cell = ws.cell(row=2, column=c0, value=label)
        cell.font = Font(italic=True, size=9)
        cell.fill = FILL_ENTETE
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    # ws.append() se base sur un compteur de ligne interne qui n'est mis à
    # jour que par des écritures via ws.cell(), pas par merge_cells() seul -
    # sans ce "toucher" explicite de la ligne 2, un groupe niveau1 fusionné
    # verticalement sur les lignes 1-2 (aucun enfant niveau2, cas des
    # onglets par tableau) laisse le compteur bloqué sur 1, et le prochain
    # ws.append(headers) écrase la ligne 2 au lieu de continuer en ligne 3.
    ws.cell(row=2, column=1)

    return 3


# ------------------------------------------------------- onglets par tableau

def _fmt_meilleur_partenaire(mp):
    if not mp:
        return [None, None, None, None, None, None]
    return [mp["nom"], mp["matchs_joues"], mp["victoires"], _pct(mp["pct_victoire"]),
            round(mp["indice_performance"], 2), mp.get("points_cote_total")]


def _fmt_fait_marquant(fm):
    if not fm:
        return [None, None, None]
    return [fm["adversaire_nom"], round(fm["adversaire_cote"], 1), fm["score"]]


def _fmt_fait_marquant_points(fm):
    if not fm:
        return [None, None, None]
    return [fm["adversaire_nom"], round(fm["points_cote"], 1), fm["score"]]


def _fmt_ecart_cote(fm):
    if not fm:
        return [None, None, None, None]
    return [fm["adversaire_nom"], round(fm["ecart_cote"], 1),
            "Victoire" if fm["victoire"] else "Défaite", fm["score"]]


def _write_discipline_sheet(wb, title, tableau, sexe, all_stats):
    ws = wb.create_sheet(title)
    has_partner = tableau in ("Double", "Mixte")
    inclure_sexe = tableau == "Mixte"  # seul tableau qui mélange les deux genres

    rows = []
    for s in all_stats:
        if sexe is not None and s["Sexe"] != sexe:
            continue
        entry = s["par_tableau"][tableau]
        if entry["matchs_joues"] == 0:
            continue
        rows.append((s, entry))
    rows.sort(key=lambda r: -r[1]["indice_global"])

    mois_tries = sorted({mois for _, entry in rows for mois in entry["evolution_mensuelle_cote"]})

    headers = ["Nom"]
    if inclure_sexe:
        headers += ["Sexe"]
    headers += ["Classement", "Cote", "Diff cote"]
    for mois in mois_tries:
        headers += [f"Diff cote {mois}"]
    headers += ["Matchs joués", "Victoires", "% victoire"]
    headers += ["Indice performance", "Indice niveau", "Indice qualité", "Indice global"]
    headers += [
        "Cote adverse moyenne",
        "Meilleure victoire", "Meilleure victoire (cote adv)", "Meilleure victoire (score)",
        "Pire défaite", "Pire défaite (cote adv)", "Pire défaite (score)",
        "Plus grosse perf", "Plus grosse perf (points)", "Plus grosse perf (score)",
        "Plus grosse contre-perf", "Plus grosse contre-perf (points)", "Plus grosse contre-perf (score)",
        "Plus gros écart de cote", "Diff de cote (écart max)", "Résultat (écart max)", "Score (écart max)",
        "Sets serrés joués", "Sets serrés gagnés", "% clutch", "Diff clutch (%)",
    ]
    if has_partner:
        headers += [
            "Matchs avec partenaire club", "% avec partenaire club",
            "Matchs sans partenaire club", "% sans partenaire club",
            "Delta % (avec - sans)",
            "Meilleur partenaire", "Matchs avec lui/elle", "Victoires avec lui/elle",
            "% victoire avec lui/elle", "Indice perf. avec lui/elle", "Points cote marqués ensemble",
            "Meilleur partenaire au club (si différent)", "Matchs avec lui/elle (club)",
            "Victoires avec lui/elle (club)", "% victoire avec lui/elle (club)",
            "Indice perf. avec lui/elle (club)", "Points cote marqués ensemble (club)",
        ]

    niveau1 = [("Classement & cote", "Classement", "Diff cote")]
    if mois_tries:
        niveau1 += [("Évolution cote", f"Diff cote {mois_tries[0]}", f"Diff cote {mois_tries[-1]}")]
    niveau1 += [
        ("Résultats", "Matchs joués", "% victoire"),
        ("Indices", "Indice performance", "Indice global"),
        ("Faits marquants & clutch", "Cote adverse moyenne", "Diff clutch (%)"),
    ]
    if has_partner:
        niveau1 += [
            ("Partenaire club", "Matchs avec partenaire club", "Delta % (avec - sans)"),
            ("Meilleur partenaire", "Meilleur partenaire", "Points cote marqués ensemble"),
            ("Meilleur partenaire au club", "Meilleur partenaire au club (si différent)",
             "Points cote marqués ensemble (club)"),
        ]
    header_row = _ecrire_sur_entete(ws, headers, niveau1)
    ws.append(headers)
    _styliser_ligne(ws, header_row, len(headers))
    ws.freeze_panes = f"B{header_row + 1}"

    for s, entry in rows:
        row = [s["Nom"]]
        if inclure_sexe:
            row += [s["Sexe"]]
        diff_cote = s["progression"][tableau]["diff_points"]
        row += [entry["classement"], entry["cote"],
                round(diff_cote, 1) if diff_cote is not None else None]
        for mois in mois_tries:
            row += [round(entry["evolution_mensuelle_cote"].get(mois, 0.0), 1)]
        row += [entry["matchs_joues"], entry["victoires"], _pct(entry["pct_victoire"])]
        row += [_pct(entry["indice_performance"]), _pct(entry["indice_niveau"]),
                _pct(entry["indice_qualite"]), _pct(entry["indice_global"])]
        cam = entry["cote_adverse_moyenne"]
        row += [round(cam, 1) if cam is not None else None]
        row += _fmt_fait_marquant(entry["meilleure_victoire"])
        row += _fmt_fait_marquant(entry["pire_defaite"])
        row += _fmt_fait_marquant_points(entry["plus_grosse_perf"])
        row += _fmt_fait_marquant_points(entry["plus_grosse_defaite_points"])
        row += _fmt_ecart_cote(entry["plus_grande_difference_cote"])
        clutch = entry["clutch"]
        row += [
            clutch["sets_serres_joues"], clutch["sets_serres_gagnes"],
            round(clutch["taux_clutch"], 1) if clutch["taux_clutch"] is not None else None,
            round(clutch["delta_clutch"], 1) if clutch["delta_clutch"] is not None else None,
        ]
        if has_partner:
            p = entry["partenaire"]
            avec, sans = p["avec_partenaire_club"], p["sans_partenaire_club"]
            row += [
                avec["matchs_joues"], _pct(avec["pct_victoire"]),
                sans["matchs_joues"], _pct(sans["pct_victoire"]),
                _pct(p["delta_pct"]),
            ]
            row += _fmt_meilleur_partenaire(entry["meilleur_partenaire"])
            row += _fmt_meilleur_partenaire(entry["meilleur_partenaire_club"])
        ws.append(row)

    _appliquer_mise_en_forme(ws, headers, header_row + 1, ws.max_row)

    groupes = [(g[1], g[2]) for g in niveau1]
    for c0, c1 in _grouper(headers, *groupes):
        _encadrer_groupe(ws, c0, c1, 1, ws.max_row)
    _ajouter_filtre(ws, len(headers), header_row, ws.max_row)
    return ws


# ------------------------------------------------------------- Bilan joueur

def _par_tableau(prog, cle):
    return [prog[t][cle] for t in ("Simple", "Double", "Mixte")]


def _par_tableau_arrondi(prog, cle, echelle=1):
    return [round(prog[t][cle] * echelle, 1) if prog[t][cle] is not None else None
            for t in ("Simple", "Double", "Mixte")]


def _cote_saison_valeurs(entries, cle, arrondi=False):
    """entries : (simple, double, mixte) - entrées par_tableau, chacune avec
    un sous-dict "cote_saison" (voir stats._cote_saison_stats)."""
    valeurs = [e["cote_saison"][cle] for e in entries]
    if arrondi:
        return [round(v, 1) if v is not None else None for v in valeurs]
    return valeurs


def _write_bilan_sheet(wb, all_stats):
    ws = wb.create_sheet("Bilan joueur")
    headers = [
        "Nom", "Sexe", "Ordre tableau",
        "Matchs S", "Victoires S", "% S",
        "Matchs D", "Victoires D", "% D",
        "Matchs M", "Victoires M", "% M",
        "Matchs total", "Victoires total", "% total",
        # rappel des indices par tableau (pas seulement le global agrégé,
        # pour qu'on sache de quel tableau ils parlent), puis le global
        "Indice performance S", "Indice niveau S", "Indice qualité S", "Indice global S",
        "Indice performance D", "Indice niveau D", "Indice qualité D", "Indice global D",
        "Indice performance M", "Indice niveau M", "Indice qualité M", "Indice global M",
        "Indice performance (global)", "Indice niveau (global)", "Indice qualité (global)",
        "Indice global (global)",
        "Matchs avec partenaire club (D+M)", "% avec partenaire club",
        "Matchs sans partenaire club (D+M)", "% sans partenaire club",
        "Delta % partenaire (avec - sans)",
        # un seul meilleur partenaire (D+M), pas de distinction club/overall ici
        "Meilleur partenaire (D+M)", "Matchs avec lui/elle", "Victoires avec lui/elle",
        "% victoire avec lui/elle", "Indice perf. avec lui/elle", "Points cote marqués ensemble",
        "Nb tournois individuels", "Nb interclubs (par jour)",
        # regroupé par thème (classement, puis place, puis cote), et DANS
        # chaque thème sept./actuel/diff sont adjacents plutôt qu'éclatés
        # dans des blocs séparés - on lit la progression d'un coup d'oeil.
        "Classement sept. S", "Classement sept. D", "Classement sept. M",
        "Classement actuel S", "Classement actuel D", "Classement actuel M",
        "Gain tableau S", "Gain tableau D", "Gain tableau M",
        "Place sept. S", "Place sept. D", "Place sept. M",
        "Place actuelle S", "Place actuelle D", "Place actuelle M",
        "Gain places S", "Gain places D", "Gain places M",
        "Cote sept. S", "Cote sept. D", "Cote sept. M",
        "Cote actuelle S", "Cote actuelle D", "Cote actuelle M",
        "Diff cote S", "Diff cote D", "Diff cote M", "Diff cote cumulée",
        # pas de version "relative" (%) du diff cote : sur cette échelle, un
        # même diff_points représente un effort comparable quel que soit le
        # niveau de départ - une version relative ferait paraître un joueur
        # bas niveau plus "progressif" qu'un joueur haut niveau pour un
        # progrès équivalent, ça n'a pas de sens (voir stats.diff_classement)
        "Tendance cote S (pts/semaine)", "Tendance cote D (pts/semaine)",
        "Tendance cote M (pts/semaine)", "Tendance cote cumulée (pts/semaine)",
        "Cote min S", "Cote min D", "Cote min M",
        "Cote max S", "Cote max D", "Cote max M",
        "Cote moyenne S", "Cote moyenne D", "Cote moyenne M",
        "Stabilité cote S (%)", "Stabilité cote D (%)", "Stabilité cote M (%)",
    ]

    niveau1 = [
        ("Résultats", "Matchs S", "% total"),
        ("Indices", "Indice performance S", "Indice global (global)"),
        ("Partenaire club", "Matchs avec partenaire club (D+M)", "Delta % partenaire (avec - sans)"),
        ("Meilleur partenaire", "Meilleur partenaire (D+M)", "Points cote marqués ensemble"),
        ("Tournois", "Nb tournois individuels", "Nb interclubs (par jour)"),
        ("Évolution classement", "Classement sept. S", "Gain tableau M"),
        ("Évolution place", "Place sept. S", "Gain places M"),
        ("Évolution cote", "Cote sept. S", "Tendance cote cumulée (pts/semaine)"),
        ("Cote sur la saison", "Cote min S", "Stabilité cote M (%)"),
    ]
    niveau2 = [
        ("Simple", "Matchs S", "% S"), ("Double", "Matchs D", "% D"),
        ("Mixte", "Matchs M", "% M"), ("Total", "Matchs total", "% total"),

        ("Simple", "Indice performance S", "Indice global S"),
        ("Double", "Indice performance D", "Indice global D"),
        ("Mixte", "Indice performance M", "Indice global M"),
        ("Global", "Indice performance (global)", "Indice global (global)"),

        ("Classement septembre", "Classement sept. S", "Classement sept. M"),
        ("Classement actuel", "Classement actuel S", "Classement actuel M"),
        ("Gain tableau", "Gain tableau S", "Gain tableau M"),

        ("Place septembre", "Place sept. S", "Place sept. M"),
        ("Place actuelle", "Place actuelle S", "Place actuelle M"),
        ("Gain places", "Gain places S", "Gain places M"),

        ("Cote septembre", "Cote sept. S", "Cote sept. M"),
        ("Cote actuelle", "Cote actuelle S", "Cote actuelle M"),
        ("Diff cote", "Diff cote S", "Diff cote cumulée"),
        ("Tendance cote", "Tendance cote S (pts/semaine)", "Tendance cote cumulée (pts/semaine)"),

        ("Min", "Cote min S", "Cote min M"),
        ("Max", "Cote max S", "Cote max M"),
        ("Moyenne", "Cote moyenne S", "Cote moyenne M"),
        ("Stabilité", "Stabilité cote S (%)", "Stabilité cote M (%)"),
    ]
    header_row = _ecrire_sur_entete(ws, headers, niveau1, niveau2)
    ws.append(headers)
    _styliser_ligne(ws, header_row, len(headers))
    ws.freeze_panes = f"B{header_row + 1}"

    rows = sorted(all_stats, key=lambda s: -s["global"]["indice_global"])
    for s in rows:
        simple, double, mixte = (s["par_tableau"][t] for t in ("Simple", "Double", "Mixte"))
        g, t, prog = s["global"], s["tournois"], s["progression"]
        p = s["partenaire_double_mixte"]
        avec, sans = p["avec_partenaire_club"], p["sans_partenaire_club"]
        cum = prog["cumule"]

        row = [
            s["Nom"], s["Sexe"], " > ".join(s["ordre_disciplines"]),
            simple["matchs_joues"], simple["victoires"], _pct(simple["pct_victoire"]),
            double["matchs_joues"], double["victoires"], _pct(double["pct_victoire"]),
            mixte["matchs_joues"], mixte["victoires"], _pct(mixte["pct_victoire"]),
            g["matchs_joues"], g["victoires"], _pct(g["pct_victoire"]),
        ]
        for entry in (simple, double, mixte, g):
            row += [_pct(entry["indice_performance"]), _pct(entry["indice_niveau"]),
                    _pct(entry["indice_qualite"]), _pct(entry["indice_global"])]
        row += [
            avec["matchs_joues"], _pct(avec["pct_victoire"]),
            sans["matchs_joues"], _pct(sans["pct_victoire"]),
            _pct(p["delta_pct"]),
        ]
        row += _fmt_meilleur_partenaire(s["meilleur_partenaire_double_mixte"])
        row += [t["nb_tournois"], t["nb_interclubs"]]

        row += _par_tableau(prog, "septembre_classement")
        row += _par_tableau(prog, "actuel_classement")
        row += _par_tableau(prog, "gain_tableau")
        row += _par_tableau(prog, "septembre_rang")
        row += _par_tableau(prog, "actuel_rang")
        row += _par_tableau(prog, "gain_places")
        row += _par_tableau(prog, "septembre_points")
        row += _par_tableau(prog, "actuel_points")
        row += _par_tableau_arrondi(prog, "diff_points")
        row += [round(cum["diff_points"], 1) if cum["diff_points"] is not None else None]
        row += _par_tableau_arrondi(prog, "tendance_hebdo")
        row += [round(cum["tendance_hebdo"], 1) if cum["tendance_hebdo"] is not None else None]
        row += _cote_saison_valeurs((simple, double, mixte), "cote_min")
        row += _cote_saison_valeurs((simple, double, mixte), "cote_max")
        row += _cote_saison_valeurs((simple, double, mixte), "cote_moyenne", arrondi=True)
        row += _cote_saison_valeurs((simple, double, mixte), "stabilite", arrondi=True)
        ws.append(row)

    _appliquer_mise_en_forme(ws, headers, header_row + 1, ws.max_row)

    for c0, c1 in _grouper(headers, *[(g[1], g[2]) for g in niveau1]):
        _encadrer_groupe(ws, c0, c1, 1, ws.max_row)

    _lignes_legeres(ws, headers, ["Classement actuel S"], 2, ws.max_row)
    _lignes_legeres(ws, headers, ["Place actuelle S", "Gain places S"], 2, ws.max_row)
    _lignes_legeres(ws, headers,
                     ["Cote actuelle S", "Diff cote S", "Tendance cote S (pts/semaine)"], 2, ws.max_row)

    _ajouter_filtre(ws, len(headers), header_row, ws.max_row)
    return ws


# ----------------------------------------------------------------- Tournois

def _joueur_par_mois(t):
    """t = s["tournois"] -> {mois: {weekend, soiree, interclub}}."""
    par_mois = defaultdict(lambda: {"weekend": 0, "soiree": 0, "interclub": 0})
    for tournoi in t["tournois"]:
        par_mois[tournoi["date"].strftime("%Y-%m")][tournoi["type"]] += 1
    for d in t["interclubs_dates"]:
        par_mois[d.strftime("%Y-%m")]["interclub"] += 1
    return par_mois


def _write_tournois_sheet(wb, all_stats):
    ws = wb.create_sheet("Tournois")

    mois_tries = sorted({
        mois
        for s in all_stats
        for mois in _joueur_par_mois(s["tournois"])
    })

    header = ["Nom"]
    for mois in mois_tries:
        header += [f"{mois} weekend", f"{mois} soirée", f"{mois} IC"]
    header += ["Total weekend", "Total soirée", "Total IC", "Total total"]
    ws.append(header)

    rows = []
    par_mois_club = defaultdict(lambda: {"weekend": 0, "soiree": 0, "interclub": 0})
    for s in all_stats:
        t = s["tournois"]
        par_mois = _joueur_par_mois(t)
        if not par_mois:
            continue

        row = [s["Nom"]]
        for mois in mois_tries:
            b = par_mois.get(mois, {"weekend": 0, "soiree": 0, "interclub": 0})
            row += [b["weekend"], b["soiree"], b["interclub"]]
        total = t["nb_tournois_weekend"] + t["nb_tournois_soiree"] + t["nb_interclubs"]
        row += [t["nb_tournois_weekend"], t["nb_tournois_soiree"], t["nb_interclubs"], total]
        rows.append(row)

        for mois, b in par_mois.items():
            for cle in ("weekend", "soiree", "interclub"):
                par_mois_club[mois][cle] += b[cle]

    rows.sort(key=lambda r: -r[-1])  # total total décroissant
    for row in rows:
        ws.append(row)
    _appliquer_mise_en_forme(ws, header, 2, 1 + len(rows))
    _ajouter_filtre(ws, len(header), 1, 1 + len(rows))
    _styliser_ligne(ws, 1, len(header))
    ws.freeze_panes = "B2"

    # --- tableau + graphique club par mois ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Par mois (club)"])
    mois_header_row = ws.max_row + 1
    ws.append(["Mois", "Tournois weekend", "Tournois soirée", "Interclubs"])
    mois_first_row = ws.max_row + 1
    for mois in mois_tries:
        b = par_mois_club[mois]
        ws.append([mois, b["weekend"], b["soiree"], b["interclub"]])
    mois_last_row = ws.max_row

    if mois_last_row >= mois_first_row:
        chart = BarChart()
        chart.type = "col"
        chart.title = "Tournois et interclubs par mois"
        chart.y_axis.title = "Nombre"
        chart.x_axis.title = "Mois"
        data = Reference(ws, min_col=2, max_col=4, min_row=mois_header_row, max_row=mois_last_row)
        categories = Reference(ws, min_col=1, min_row=mois_first_row, max_row=mois_last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        ws.add_chart(chart, f"F{mois_header_row}")
    _styliser_mini_tableau(ws, titre_row, mois_header_row, mois_last_row, 4)

    # --- distribution : nombre de joueurs par nombre de tournois individuels ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Joueurs par nombre de tournois individuels"])
    dist_header_row = ws.max_row + 1
    ws.append(["Nb tournois", "Nb joueurs"])
    dist_first_row = ws.max_row + 1
    compte = Counter(s["tournois"]["nb_tournois"] for s in all_stats)
    for nb in sorted(compte):
        ws.append([nb, compte[nb]])
    dist_last_row = ws.max_row

    if dist_last_row >= dist_first_row:
        chart2 = BarChart()
        chart2.type = "col"
        chart2.title = "Nombre de joueurs par nombre de tournois individuels"
        chart2.y_axis.title = "Nombre de joueurs"
        chart2.x_axis.title = "Nombre de tournois"
        data2 = Reference(ws, min_col=2, min_row=dist_header_row, max_row=dist_last_row)
        categories2 = Reference(ws, min_col=1, min_row=dist_first_row, max_row=dist_last_row)
        chart2.add_data(data2, titles_from_data=True)
        chart2.set_categories(categories2)
        ws.add_chart(chart2, f"F{dist_header_row + 18}")
    _styliser_mini_tableau(ws, titre_row, dist_header_row, dist_last_row, 2)

    return ws


# --------------------------------------------------------------- Stats club

# catégories affichées en colonnes, chacune avec 3 sous-colonnes
# (matchs/victoires/%) : SH/SD/DH/DD/MX comme les onglets par tableau, puis
# des vues transverses (tout tableau confondu par genre, et Simple/Double
# unisexe). Les 5 premières alimentent le graphique combiné (barres
# empilées + courbes de %).
CLUB_CATEGORIES = ["SH", "SD", "DH", "DD", "MX",
                    "H (tous tableaux)", "F (tous tableaux)",
                    "S (unisexe)", "D (unisexe)"]

# Dames -> "D" dans SH/SD/DH/DD (pas "F"), pour matcher l'intitulé des onglets
_GENRE_CODE = {"H": "H", "F": "D"}


CATEGORIE_CLUB_TOTAL = "Club (tous tableaux)"


def _write_club_sheet(wb, all_stats):
    ws = wb.create_sheet("Stats club")

    toutes_categories = CLUB_CATEGORIES + [CATEGORIE_CLUB_TOTAL]
    par_mois = defaultdict(lambda: {cat: {"joues": 0, "victoires": 0} for cat in toutes_categories})
    # total vs interclub uniquement, par catégorie (pas par mois) - un match
    # intra-club (adversaire 100% GAB38, en tournoi individuel - un
    # interclub oppose toujours deux clubs différents) apporte TOUJOURS 1
    # victoire + 1 défaite au club, sans rien dire de sa performance face à
    # l'extérieur : on l'exclut de tous les compteurs de cet onglet (mais
    # pas des stats individuelles des joueurs concernés, qui restent des
    # résultats bien réels pour eux).
    par_categorie = defaultdict(lambda: {"total": {"joues": 0, "victoires": 0},
                                          "interclub": {"joues": 0, "victoires": 0}})
    # matchs bruts par catégorie (pas juste les compteurs) pour pouvoir
    # rappeler stats.indice_clutch() club-wide plus bas.
    matchs_par_categorie = defaultdict(list)

    def _add(mois, cat, m):
        b = par_mois[mois][cat]
        b["joues"] += 1
        b["victoires"] += 1 if m["victoire"] else 0
        matchs_par_categorie[cat].append(m)
        bc = par_categorie[cat]
        bc["total"]["joues"] += 1
        bc["total"]["victoires"] += 1 if m["victoire"] else 0
        if m["est_interclub"]:
            bc["interclub"]["joues"] += 1
            bc["interclub"]["victoires"] += 1 if m["victoire"] else 0

    for s in all_stats:
        sexe = s["Sexe"]
        genre_total = "H (tous tableaux)" if sexe == "H" else "F (tous tableaux)"
        for m in s["match_log"]:
            if m["intra_club"]:
                continue
            tableau, mois = m["tableau"], m["mois"]
            _add(mois, CATEGORIE_CLUB_TOTAL, m)
            _add(mois, genre_total, m)
            if tableau == "Mixte":
                _add(mois, "MX", m)
            else:
                _add(mois, tableau[0] + _GENRE_CODE[sexe], m)  # SH/SD/DH/DD
                _add(mois, f"{tableau[0]} (unisexe)", m)  # S/D (unisexe)

    header = ["Mois"]
    for cat in CLUB_CATEGORIES:
        header += [f"{cat} - matchs", f"{cat} - victoires", f"{cat} - %"]
    ws.append(header)

    mois_tries = sorted(par_mois)
    for mois in mois_tries:
        row = [mois]
        for cat in CLUB_CATEGORIES:
            b = par_mois[mois][cat]
            pct = (b["victoires"] / b["joues"] * 100) if b["joues"] else 0.0
            row += [b["joues"], b["victoires"], _pct(pct)]
        ws.append(row)

    header_row = 1
    first_row = 2
    last_row = ws.max_row
    _appliquer_mise_en_forme(ws, header, first_row, last_row)
    _ajouter_filtre(ws, len(header), header_row, last_row)
    _styliser_ligne(ws, header_row, len(header))
    ws.freeze_panes = f"B{first_row}"

    if last_row >= first_row:
        bar = BarChart()
        bar.type = "col"
        bar.grouping = "stacked"
        bar.overlap = 100
        bar.title = "Matchs joués par mois (par tableau) et % de victoire"
        bar.y_axis.title = "Matchs joués"
        bar.x_axis.title = "Mois"

        line = LineChart()
        line.y_axis.axId = 200
        line.y_axis.title = "% victoire"
        line.y_axis.crosses = "max"

        for i in range(5):  # SH, SD, DH, DD, MX
            matchs_col = 2 + i * 3
            pct_col = 4 + i * 3
            bar.add_data(Reference(ws, min_col=matchs_col, min_row=header_row, max_row=last_row),
                         titles_from_data=True)
            line.add_data(Reference(ws, min_col=pct_col, min_row=header_row, max_row=last_row),
                          titles_from_data=True)

        categories = Reference(ws, min_col=1, min_row=first_row, max_row=last_row)
        bar.set_categories(categories)
        line.set_categories(categories)

        bar += line
        ws.add_chart(bar, f"B{last_row + 3}")

    # --- total vs interclub uniquement, par catégorie ---
    # "Diff interclub (%)" = choke (négatif) ou overperform (positif) en
    # interclub par rapport au total (qui inclut déjà l'interclub - une
    # comparaison contre le hors-interclub serait plus "pure", mais total
    # vs interclub est ce qui a été demandé, et reste lisible tant que
    # l'interclub n'est pas l'essentiel du volume de matchs).
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Total vs interclub, par catégorie"])
    ti_header_row = ws.max_row + 1
    ti_headers = ["Catégorie", "Matchs (total)", "Victoires (total)", "% (total)",
                  "Matchs (interclub)", "Victoires (interclub)", "% (interclub)", "Diff interclub (%)"]
    ws.append(ti_headers)
    ti_first_row = ws.max_row + 1

    for cat in toutes_categories:
        b = par_categorie[cat]
        t, ic = b["total"], b["interclub"]
        pct_t = (t["victoires"] / t["joues"] * 100) if t["joues"] else None
        pct_ic = (ic["victoires"] / ic["joues"] * 100) if ic["joues"] else None
        diff = (pct_ic - pct_t) if pct_t is not None and pct_ic is not None else None
        ws.append([
            cat, t["joues"], t["victoires"], round(pct_t, 1) if pct_t is not None else None,
            ic["joues"], ic["victoires"], round(pct_ic, 1) if pct_ic is not None else None,
            round(diff, 1) if diff is not None else None,
        ])
    ti_last_row = ws.max_row
    _appliquer_mise_en_forme(ws, ti_headers, ti_first_row, ti_last_row)
    _styliser_mini_tableau(ws, titre_row, ti_header_row, ti_last_row, len(ti_headers))

    # --- points de cote gagnés/perdus, total vs interclub, par catégorie ---
    # "Diff interclub" (gagnés et perdus) : l'interclub rapporte/coûte-t-il
    # plus de points en moyenne qu'un match hors interclub - barème FFBad
    # différent en compétition par équipe ?
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Points de cote gagnés/perdus, par catégorie"])
    pc_header_row = ws.max_row + 1
    pc_headers = ["Catégorie", "Pts gagnés (total)", "Pts perdus (total)",
                  "Pts gagnés (interclub)", "Pts perdus (interclub)",
                  "Diff gagnés (interclub)", "Diff perdus (interclub)"]
    ws.append(pc_headers)
    pc_first_row = ws.max_row + 1

    for cat in toutes_categories:
        matchs_cat = matchs_par_categorie.get(cat, [])
        matchs_ic = [m for m in matchs_cat if m["est_interclub"]]
        total = stats.points_cote_moyens(matchs_cat)
        interclub = stats.points_cote_moyens(matchs_ic)
        gagne_t, perdu_t = total["moyenne_gagne"], total["moyenne_perdu"]
        gagne_ic, perdu_ic = interclub["moyenne_gagne"], interclub["moyenne_perdu"]
        diff_gagne = (gagne_ic - gagne_t) if gagne_t is not None and gagne_ic is not None else None
        diff_perdu = (perdu_ic - perdu_t) if perdu_t is not None and perdu_ic is not None else None
        ws.append([
            cat,
            round(gagne_t, 1) if gagne_t is not None else None,
            round(perdu_t, 1) if perdu_t is not None else None,
            round(gagne_ic, 1) if gagne_ic is not None else None,
            round(perdu_ic, 1) if perdu_ic is not None else None,
            round(diff_gagne, 1) if diff_gagne is not None else None,
            round(diff_perdu, 1) if diff_perdu is not None else None,
        ])
    pc_last_row = ws.max_row
    _appliquer_mise_en_forme(ws, pc_headers, pc_first_row, pc_last_row)
    _styliser_mini_tableau(ws, titre_row, pc_header_row, pc_last_row, len(pc_headers))

    # --- dynamique de match club-wide, par catégorie ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Dynamique de match, par catégorie"])
    dy_header_row = ws.max_row + 1
    dy_headers = ["Catégorie", "Matchs gagnés", "Taux propre (%)",
                  "1er set gagné", "Taux après 1er set gagné (%)",
                  "1er set perdu", "Taux comeback (%)"]
    ws.append(dy_headers)
    dy_first_row = ws.max_row + 1

    for cat in toutes_categories:
        d = stats.dynamique_match(matchs_par_categorie.get(cat, []))
        ws.append([
            cat, d["nb_gagnes"], round(d["taux_propre"], 1) if d["taux_propre"] is not None else None,
            d["nb_premier_set_gagne"],
            round(d["taux_apres_premier_set"], 1) if d["taux_apres_premier_set"] is not None else None,
            d["nb_premier_set_perdu"],
            round(d["taux_comeback"], 1) if d["taux_comeback"] is not None else None,
        ])
    dy_last_row = ws.max_row
    _appliquer_mise_en_forme(ws, dy_headers, dy_first_row, dy_last_row)
    _styliser_mini_tableau(ws, titre_row, dy_header_row, dy_last_row, len(dy_headers))

    # --- David vs Goliath, par catégorie ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["David vs Goliath, par catégorie (cote adverse vs la sienne)"])
    dvg_header_row = ws.max_row + 1
    dvg_headers = ["Catégorie", "Matchs en outsider", "% victoire en outsider",
                   "Matchs en favori", "% victoire en favori"]
    ws.append(dvg_headers)
    dvg_first_row = ws.max_row + 1

    for cat in toutes_categories:
        dvg = stats.david_vs_goliath(matchs_par_categorie.get(cat, []))
        ws.append([
            cat, dvg["nb_outsider"],
            round(dvg["taux_outsider"], 1) if dvg["taux_outsider"] is not None else None,
            dvg["nb_favori"],
            round(dvg["taux_favori"], 1) if dvg["taux_favori"] is not None else None,
        ])
    dvg_last_row = ws.max_row
    _appliquer_mise_en_forme(ws, dvg_headers, dvg_first_row, dvg_last_row)
    _styliser_mini_tableau(ws, titre_row, dvg_header_row, dvg_last_row, len(dvg_headers))

    # --- clutchness club-wide, par catégorie ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Clutchness club (sets à 2 points d'écart ou moins)"])
    clutch_header_row = ws.max_row + 1
    clutch_headers = ["Catégorie", "Sets serrés joués", "Sets serrés gagnés", "% clutch", "Diff clutch (%)"]
    ws.append(clutch_headers)
    clutch_first_row = ws.max_row + 1

    for cat in toutes_categories:
        c = stats.indice_clutch(matchs_par_categorie.get(cat, []))
        ws.append([
            cat, c["sets_serres_joues"], c["sets_serres_gagnes"],
            round(c["taux_clutch"], 1) if c["taux_clutch"] is not None else None,
            round(c["delta_clutch"], 1) if c["delta_clutch"] is not None else None,
        ])
    clutch_last_row = ws.max_row
    _appliquer_mise_en_forme(ws, clutch_headers, clutch_first_row, clutch_last_row)
    _styliser_mini_tableau(ws, titre_row, clutch_header_row, clutch_last_row, len(clutch_headers))

    return ws


# --------------------------------------------------------------- Stats avancées

# une ligne par métrique de profil de score, répétée pour chaque tableau
# (suffixe S/D/M, comme "Classement sept. S" ailleurs) - voir stats.
# profil_score / stats.sets_extremes.
_METRIQUES_SCORE = ["Score moyen (moi)", "Score moyen (adversaire)", "Score max infligé", "Score max reçu",
                     "Points moyens (victoire)", "Points moyens (défaite)",
                     "Sets extrêmes joués", "Sets extrêmes gagnés",
                     "Taux propre (%)", "Taux après 1er set gagné (%)", "Taux comeback (%)"]
_SUFFIXE_TABLEAU = {"Simple": "S", "Double": "D", "Mixte": "M"}


def _regrouper_dates_consecutives(dates):
    """dates : liste de `date` -> liste de groupes (listes) de dates qui se
    suivent jour par jour (au plus 1 jour d'écart) - un tournoi sur un
    weekend (samedi + dimanche) reste un seul groupe, mais deux occurrences
    du même nom à des dates éloignées (une compétition qui revient chaque
    mois par exemple) restent distinctes."""
    dates_triees = sorted(set(dates))
    groupes = []
    groupe_courant = []
    for d in dates_triees:
        if groupe_courant and (d - groupe_courant[-1]).days > 1:
            groupes.append(groupe_courant)
            groupe_courant = []
        groupe_courant.append(d)
    if groupe_courant:
        groupes.append(groupe_courant)
    return groupes


def _tournois_participation(all_stats, min_joueurs=5):
    """Pour chaque tournoi individuel (hors interclubs et hors matchs
    intra-club - qui ne disent rien de la performance face à l'extérieur),
    les joueurs du club engagés et leur bilan agrégé. Un même tournoi peut
    s'étaler sur plusieurs jours consécutifs (ex: weekend) : les matchs du
    même nom d'événement à des dates qui se suivent (voir
    _regrouper_dates_consecutives) sont fusionnés en une seule occurrence,
    plutôt que comptés comme des tournois séparés jour par jour.

    Renvoie (meilleur, pire) - les deux tournois au % de victoire le plus
    haut/bas, parmi ceux avec au moins `min_joueurs` joueurs du club
    engagés (sinon trop peu pour être significatif). (None, None) si aucun
    tournoi n'atteint ce seuil."""
    par_date = defaultdict(lambda: {"joueurs": set(), "joues": 0, "victoires": 0})
    for s in all_stats:
        for m in s["match_log"]:
            if m["est_interclub"] or m["intra_club"]:
                continue
            b = par_date[(m["evenement"], m["date"])]
            b["joueurs"].add(s["Nom"])
            b["joues"] += 1
            b["victoires"] += 1 if m["victoire"] else 0

    dates_par_nom = defaultdict(list)
    for nom, d in par_date:
        dates_par_nom[nom].append(d)

    candidats = []
    for nom, dates in dates_par_nom.items():
        for groupe in _regrouper_dates_consecutives(dates):
            joueurs, joues, victoires = set(), 0, 0
            for d in groupe:
                b = par_date[(nom, d)]
                joueurs |= b["joueurs"]
                joues += b["joues"]
                victoires += b["victoires"]
            if len(joueurs) < min_joueurs:
                continue
            pct = victoires / joues * 100 if joues else 0.0
            candidats.append({
                "nom": nom, "date": groupe[0], "joueurs": len(joueurs),
                "matchs": joues, "victoires": victoires, "pct": pct,
            })
    if not candidats:
        return None, None
    return max(candidats, key=lambda c: c["pct"]), min(candidats, key=lambda c: c["pct"])


def _rivalites_simple(all_stats):
    """Pour chaque paire de joueurs du club qui se sont affrontés en Simple
    (matchs intra-club), nombre de confrontations et qui mène - les
    rivalités les plus disputées de la saison, triées par nombre de
    confrontations décroissant. Limité au Simple : en double/mixte, une
    "rivalité" opposerait deux PAIRES (4 personnes), pas 2 joueurs - une
    tout autre question."""
    par_paire = defaultdict(lambda: {"total": 0, "victoires": defaultdict(int)})
    for s in all_stats:
        for m in s["match_log"]:
            if m["tableau"] != "Simple" or not m["intra_club"]:
                continue
            adversaire = " / ".join(m["noms_adverses"])
            if not adversaire or s["Nom"] >= adversaire:
                continue  # évite de compter 2 fois la même rencontre (vue des 2 côtés)
            b = par_paire[(s["Nom"], adversaire)]
            b["total"] += 1
            b["victoires"][s["Nom"] if m["victoire"] else adversaire] += 1

    resultat = [
        {
            "joueur_a": a, "joueur_b": b_nom, "total": b["total"],
            "victoires_a": b["victoires"].get(a, 0), "victoires_b": b["victoires"].get(b_nom, 0),
        }
        for (a, b_nom), b in par_paire.items()
    ]
    resultat.sort(key=lambda r: -r["total"])
    return resultat


def _pics_mensuels(all_stats):
    """Pour chaque tableau et chaque genre, le plus gros gain et la plus
    grosse perte de cote sur un seul mois (voir stats.
    evolution_mensuelle_cote) - le joueur, le mois, le delta."""
    resultat = []
    for tableau in ("Simple", "Double", "Mixte"):
        for sexe in ("H", "F"):
            candidats = [
                (s["Nom"], mois, delta)
                for s in all_stats if s["Sexe"] == sexe
                for mois, delta in s["par_tableau"][tableau]["evolution_mensuelle_cote"].items()
            ]
            if not candidats:
                continue
            pic = max(candidats, key=lambda c: c[2])
            chute = min(candidats, key=lambda c: c[2])
            resultat.append({"tableau": tableau, "sexe": sexe, "type": "Pic",
                              "nom": pic[0], "mois": pic[1], "delta": pic[2]})
            resultat.append({"tableau": tableau, "sexe": sexe, "type": "Chute",
                              "nom": chute[0], "mois": chute[1], "delta": chute[2]})
    return resultat


def _write_stats_avancees_sheet(wb, all_stats):
    ws = wb.create_sheet("Stats avancées")

    headers = ["Nom", "Sexe"]
    for tableau in ("Simple", "Double", "Mixte"):
        s = _SUFFIXE_TABLEAU[tableau]
        headers += [f"{m} {s}" for m in _METRIQUES_SCORE]

    niveau1 = [(tableau, f"{_METRIQUES_SCORE[0]} {_SUFFIXE_TABLEAU[tableau]}",
                f"{_METRIQUES_SCORE[-1]} {_SUFFIXE_TABLEAU[tableau]}")
               for tableau in ("Simple", "Double", "Mixte")]
    header_row = _ecrire_sur_entete(ws, headers, niveau1)
    ws.append(headers)
    _styliser_ligne(ws, header_row, len(headers))
    ws.freeze_panes = f"B{header_row + 1}"

    rows = [s for s in all_stats
            if any(s["par_tableau"][t]["matchs_joues"] > 0 for t in ("Simple", "Double", "Mixte"))]
    rows.sort(key=lambda s: -s["global"]["indice_global"])

    for s in rows:
        row = [s["Nom"], s["Sexe"]]
        for tableau in ("Simple", "Double", "Mixte"):
            entry = s["par_tableau"][tableau]
            ps, se, dy = entry["profil_score"], entry["sets_extremes"], entry["dynamique"]
            row += [
                round(ps["score_moyen_mien"], 1) if ps["score_moyen_mien"] is not None else None,
                round(ps["score_moyen_adverse"], 1) if ps["score_moyen_adverse"] is not None else None,
                ps["score_max_inflige"],
                ps["score_max_recu"],
                round(ps["points_moyens_victoire"], 1) if ps["points_moyens_victoire"] is not None else None,
                round(ps["points_moyens_defaite"], 1) if ps["points_moyens_defaite"] is not None else None,
                se["nb_joues"],
                se["nb_gagnes"],
                round(dy["taux_propre"], 1) if dy["taux_propre"] is not None else None,
                round(dy["taux_apres_premier_set"], 1) if dy["taux_apres_premier_set"] is not None else None,
                round(dy["taux_comeback"], 1) if dy["taux_comeback"] is not None else None,
            ]
        ws.append(row)

    _appliquer_mise_en_forme(ws, headers, header_row + 1, ws.max_row)
    for c0, c1 in _grouper(headers, *[(g[1], g[2]) for g in niveau1]):
        _encadrer_groupe(ws, c0, c1, 1, ws.max_row)
    _ajouter_filtre(ws, len(headers), header_row, ws.max_row)

    # --- sets au plafond de prolongation (30-29 / 21-20), résumé club ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Sets au plafond de prolongation (30-29 avant le 1er septembre 2026, 21-20 depuis)"])
    resume_header_row = ws.max_row + 1
    ws.append(["Indicateur", "Valeur"])
    resume_first_row = ws.max_row + 1

    nb_joues = nb_gagnes = joueurs_concernes = 0
    for s in rows:
        nb_joueur = sum(s["par_tableau"][t]["sets_extremes"]["nb_joues"] for t in ("Simple", "Double", "Mixte"))
        if nb_joueur > 0:
            joueurs_concernes += 1
        nb_joues += nb_joueur
        nb_gagnes += sum(s["par_tableau"][t]["sets_extremes"]["nb_gagnes"] for t in ("Simple", "Double", "Mixte"))
    nb_perdus = nb_joues - nb_gagnes
    pct = (nb_gagnes / nb_joues * 100) if nb_joues else None

    ws.append(["Joueurs concernés (au moins 1 set)", joueurs_concernes])
    ws.append(["Total joueurs", len(rows)])
    ws.append(["Sets joués", nb_joues])
    gagnes_row = ws.max_row + 1
    ws.append(["Sets gagnés", nb_gagnes])
    ws.append(["Sets perdus", nb_perdus])
    ws.append(["% victoire dans ces sets", round(pct, 1) if pct is not None else None])
    resume_last_row = ws.max_row
    _styliser_mini_tableau(ws, titre_row, resume_header_row, resume_last_row, 2)

    if nb_joues:
        pie = PieChart()
        pie.title = "Sets au plafond de prolongation : gagnés vs perdus"
        data = Reference(ws, min_col=2, min_row=gagnes_row, max_row=gagnes_row + 1)
        labels = Reference(ws, min_col=1, min_row=gagnes_row, max_row=gagnes_row + 1)
        pie.add_data(data)
        pie.set_categories(labels)
        ws.add_chart(pie, f"D{resume_header_row}")

    # --- tournoi le plus / le moins victorieux (min. 5 joueurs engagés) ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Tournoi le plus / le moins victorieux (au moins 5 joueurs du club engagés)"])
    tournoi_header_row = ws.max_row + 1
    tournoi_headers = ["Résultat", "Tournoi", "Joueurs engagés", "Matchs", "Victoires", "% victoire"]
    ws.append(tournoi_headers)
    tournoi_first_row = ws.max_row + 1
    meilleur, pire = _tournois_participation(all_stats)
    for label, t in (("Meilleur", meilleur), ("Pire", pire)):
        if t is None:
            ws.append([label, "Aucun tournoi n'atteint le seuil de 5 joueurs", None, None, None, None])
        else:
            ws.append([label, t["nom"], t["joueurs"], t["matchs"], t["victoires"], round(t["pct"], 1)])
    tournoi_last_row = ws.max_row
    _appliquer_mise_en_forme(ws, tournoi_headers, tournoi_first_row, tournoi_last_row)
    _styliser_mini_tableau(ws, titre_row, tournoi_header_row, tournoi_last_row, len(tournoi_headers))

    # --- rivalités entre joueurs du club, en Simple ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Rivalités entre joueurs du club (Simple, matchs intra-club)"])
    riv_header_row = ws.max_row + 1
    riv_headers = ["Joueur A", "Joueur B", "Confrontations", "Victoires A", "Victoires B"]
    ws.append(riv_headers)
    riv_first_row = ws.max_row + 1
    rivalites = _rivalites_simple(all_stats)
    for r in rivalites:
        ws.append([r["joueur_a"], r["joueur_b"], r["total"], r["victoires_a"], r["victoires_b"]])
    riv_last_row = ws.max_row
    if riv_last_row >= riv_first_row:
        _appliquer_mise_en_forme(ws, riv_headers, riv_first_row, riv_last_row)
    _styliser_mini_tableau(ws, titre_row, riv_header_row, riv_last_row, len(riv_headers))

    # --- plus gros pics/chutes de cote sur un mois, par tableau et genre ---
    ws.append([])
    titre_row = ws.max_row + 1
    ws.append(["Plus gros pics/chutes de cote sur un mois, par tableau et par genre"])
    pic_header_row = ws.max_row + 1
    pic_headers = ["Tableau", "Sexe", "Type", "Joueur", "Mois", "Delta cote"]
    ws.append(pic_headers)
    pic_first_row = ws.max_row + 1
    for p in _pics_mensuels(all_stats):
        ws.append([p["tableau"], p["sexe"], p["type"], p["nom"], p["mois"], round(p["delta"], 1)])
    pic_last_row = ws.max_row
    if pic_last_row >= pic_first_row:
        _appliquer_mise_en_forme(ws, pic_headers, pic_first_row, pic_last_row)
    _styliser_mini_tableau(ws, titre_row, pic_header_row, pic_last_row, len(pic_headers))

    return ws


def _ajuster_largeurs_colonnes(ws, largeur_min=8, largeur_max=40):
    """Largeur de chaque colonne ajustée au contenu le plus long qu'elle
    contient (en-tête compris) - par défaut openpyxl laisse toutes les
    colonnes à la largeur standard d'Excel, bien trop étroite pour des
    en-têtes comme "Meilleur partenaire au club (si différent)" ou un nom
    de tournoi à rallonge. Plafonnée pour qu'un texte exceptionnellement
    long n'élargisse pas toute la feuille à l'excès."""
    largeurs = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            longueur = len(str(cell.value))
            col = cell.column_letter
            largeurs[col] = max(largeurs.get(col, 0), longueur)
    for col, longueur in largeurs.items():
        ws.column_dimensions[col].width = max(largeur_min, min(largeur_max, longueur + 2))


# couleur d'onglet par feuille - repère rapide parmi les 9 tabs : même
# logique de teinte que "Ordre tableau" (vert=Simple, bleu=Double,
# orange=Mixte) pour les 5 onglets par discipline, une couleur distincte
# pour chacun des 4 onglets de synthèse.
TAB_COLORS = {
    "SH": "FF27AE60", "SD": "FF27AE60",
    "DH": "FF3498DB", "DD": "FF3498DB",
    "MX": "FFE67E22",
    "Tournois": "FF7F8C8D",
    "Bilan joueur": "FF8E44AD",
    "Stats club": "FF16A085",
    "Stats avancées": "FFF1C40F",
}


def write_stats_excel(all_stats, output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # feuille par défaut vide

    for title, tableau, sexe in DISCIPLINE_SHEETS:
        _write_discipline_sheet(wb, title, tableau, sexe, all_stats)

    _write_tournois_sheet(wb, all_stats)
    _write_bilan_sheet(wb, all_stats)
    _write_club_sheet(wb, all_stats)
    _write_stats_avancees_sheet(wb, all_stats)

    for ws in wb.worksheets:
        _ajuster_largeurs_colonnes(ws)
        if ws.title in TAB_COLORS:
            ws.sheet_properties.tabColor = TAB_COLORS[ws.title]

    wb.save(output_path)
    return output_path
