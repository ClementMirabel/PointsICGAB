"""Écriture du classeur Excel à 8 onglets (SH/SD/DH/DD/MX, Tournois, Bilan
joueur, Stats club) à partir des stats produites par stats.py, avec mise en
forme conditionnelle (classements colorés comme sur le site, dégradé de
performance par colonne)."""
from collections import Counter, defaultdict

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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

FILL_SIMPLE = _uni("FFD5E8D4")
FILL_DOUBLE = _uni("FFDAE8FC")
FILL_MIXTE = _uni("FFFFE6CC")

BORDURE_GROUPE = Side(style="medium", color="FF999999")

# colonnes texte qu'on ne colore pas en dégradé (déjà colorées autrement,
# ou pas un indicateur de performance)
CLASSEMENT_HEADERS = {"Classement", "Classement sept. S", "Classement sept. D", "Classement sept. M"}
TEXT_HEADERS = {"Nom", "Sexe", "Mois", "Meilleur partenaire", "Meilleur partenaire (D+M)",
                "Meilleur partenaire au club (si différent)", "Ordre tableau"} | CLASSEMENT_HEADERS


def _est_colonne_delta_partenaire(header):
    """Delta % avec/sans partenaire du club : presque toujours proche de 0
    ou positif, une simple barre verte suffit (pas besoin d'un dégradé
    rouge-blanc-vert par-dessus)."""
    return "Delta" in header


def _est_colonne_diff(header):
    """Colonnes pouvant être négatives (diff cote/relative, gain de places) :
    dégradé divergent rouge-blanc-vert centré sur 0, plutôt que le dégradé
    blanc->vert habituel (qui n'aurait pas de sens sur une échelle qui
    traverse 0)."""
    return "Diff" in header or "Gain places" in header


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
    """Couleur selon la discipline en tête (la plus forte pour le joueur)."""
    if last_row < first_row:
        return
    col = get_column_letter(col_idx)
    plage = f"{col}{first_row}:{col}{last_row}"
    ref = f"{col}{first_row}"
    for prefixe, fill in (("Simple", FILL_SIMPLE), ("Double", FILL_DOUBLE), ("Mixte", FILL_MIXTE)):
        ws.conditional_formatting.add(
            plage, FormulaRule(formula=[f'LEFT({ref},{len(prefixe)})="{prefixe}"'], fill=fill))


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
    """Barre de données verte "classique" (une seule couleur, longueur
    proportionnelle à la valeur) - pour le delta % avec/sans partenaire du
    club : un dégradé rouge-blanc-vert par-dessus une barre bleue rendait la
    colonne difficile à lire pour un delta qui reste presque toujours proche
    de 0 ou positif."""
    if last_row < first_row:
        return
    col = get_column_letter(col_idx)
    plage = f"{col}{first_row}:{col}{last_row}"
    ws.conditional_formatting.add(plage, DataBarRule(
        start_type="min", end_type="max", color="FF63BE7B", showValue=True))


def _degrade_diverge(ws, col_idx, first_row, last_row):
    """Rouge (négatif) -> blanc (zéro) -> vert (positif), sans barre de
    données par-dessus (une barre en plus du dégradé rendait la colonne
    difficile à lire) - pour gain de places / diff cote / diff relative, qui
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
    """Bordure moyenne autour d'un bloc de colonnes (première ligne =
    en-tête) - aide à repérer visuellement les groupes de colonnes
    apparentées (classement/cote, matchs/victoires/%, indices...)."""
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


def _ajouter_filtre(ws, nb_colonnes, last_row):
    """Filtre Auto Excel (Data > Filter) sur l'en-tête, pour trier/filtrer
    facilement sans dérouler un menu Excel à la main."""
    if last_row < 1:
        return
    ws.auto_filter.ref = f"A1:{get_column_letter(nb_colonnes)}{last_row}"


# ------------------------------------------------------- onglets par tableau

def _fmt_meilleur_partenaire(mp):
    if not mp:
        return [None, None, None, None, None, None]
    return [mp["nom"], mp["matchs_joues"], mp["victoires"], _pct(mp["pct_victoire"]),
            round(mp["indice_performance"], 2), mp.get("points_cote_total")]


def _write_discipline_sheet(wb, title, tableau, sexe, all_stats):
    ws = wb.create_sheet(title)
    has_partner = tableau in ("Double", "Mixte")
    inclure_sexe = tableau == "Mixte"  # seul tableau qui mélange les deux genres

    headers = ["Nom"]
    if inclure_sexe:
        headers += ["Sexe"]
    headers += ["Classement", "Cote", "Matchs joués", "Victoires", "% victoire"]
    headers += ["Indice performance", "Indice niveau", "Indice global"]
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
    ws.append(headers)

    rows = []
    for s in all_stats:
        if sexe is not None and s["Sexe"] != sexe:
            continue
        entry = s["par_tableau"][tableau]
        if entry["matchs_joues"] == 0:
            continue
        rows.append((s, entry))
    rows.sort(key=lambda r: -r[1]["indice_global"])

    for s, entry in rows:
        row = [s["Nom"]]
        if inclure_sexe:
            row += [s["Sexe"]]
        row += [entry["classement"], entry["cote"], entry["matchs_joues"], entry["victoires"],
                _pct(entry["pct_victoire"])]
        row += [_pct(entry["indice_performance"]), _pct(entry["indice_niveau"]), _pct(entry["indice_global"])]
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

    _appliquer_mise_en_forme(ws, headers, 2, ws.max_row)

    groupes = [("Classement", "Cote"), ("Matchs joués", "% victoire"),
               ("Indice performance", "Indice global")]
    if has_partner:
        groupes += [
            ("Matchs avec partenaire club", "Delta % (avec - sans)"),
            ("Meilleur partenaire", "Points cote marqués ensemble"),
            ("Meilleur partenaire au club (si différent)", "Points cote marqués ensemble (club)"),
        ]
    for c0, c1 in _grouper(headers, *groupes):
        _encadrer_groupe(ws, c0, c1, 1, ws.max_row)
    _ajouter_filtre(ws, len(headers), ws.max_row)
    return ws


# ------------------------------------------------------------- Bilan joueur

def _par_tableau(prog, cle):
    return [prog[t][cle] for t in ("Simple", "Double", "Mixte")]


def _par_tableau_arrondi(prog, cle, echelle=1):
    return [round(prog[t][cle] * echelle, 1) if prog[t][cle] is not None else None
            for t in ("Simple", "Double", "Mixte")]


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
        "Indice performance S", "Indice niveau S", "Indice global S",
        "Indice performance D", "Indice niveau D", "Indice global D",
        "Indice performance M", "Indice niveau M", "Indice global M",
        "Indice performance (global)", "Indice niveau (global)", "Indice global (global)",
        "Matchs avec partenaire club (D+M)", "% avec partenaire club",
        "Matchs sans partenaire club (D+M)", "% sans partenaire club",
        "Delta % partenaire (avec - sans)",
        # un seul meilleur partenaire (D+M), pas de distinction club/overall ici
        "Meilleur partenaire (D+M)", "Matchs avec lui/elle", "Victoires avec lui/elle",
        "% victoire avec lui/elle", "Indice perf. avec lui/elle", "Points cote marqués ensemble",
        "Nb tournois individuels", "Nb interclubs (par jour)",
        # regroupé par type de métrique (classements, puis places, puis
        # cotes, puis diffs) plutôt que par tableau
        "Classement sept. S", "Classement sept. D", "Classement sept. M",
        "Place sept. S", "Place sept. D", "Place sept. M",
        "Cote sept. S", "Cote sept. D", "Cote sept. M",
        "Place actuelle S", "Place actuelle D", "Place actuelle M",
        "Cote actuelle S", "Cote actuelle D", "Cote actuelle M",
        "Gain places S", "Gain places D", "Gain places M",
        "Diff cote S", "Diff cote D", "Diff cote M", "Diff cote cumulée",
        "Diff relative S (%)", "Diff relative D (%)", "Diff relative M (%)", "Diff relative cumulée (%)",
    ]
    ws.append(headers)

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
            row += [_pct(entry["indice_performance"]), _pct(entry["indice_niveau"]), _pct(entry["indice_global"])]
        row += [
            avec["matchs_joues"], _pct(avec["pct_victoire"]),
            sans["matchs_joues"], _pct(sans["pct_victoire"]),
            _pct(p["delta_pct"]),
        ]
        row += _fmt_meilleur_partenaire(s["meilleur_partenaire_double_mixte"])
        row += [t["nb_tournois"], t["nb_interclubs"]]

        row += _par_tableau(prog, "septembre_classement")
        row += _par_tableau(prog, "septembre_rang")
        row += _par_tableau(prog, "septembre_points")
        row += _par_tableau(prog, "actuel_rang")
        row += _par_tableau(prog, "actuel_points")
        row += _par_tableau(prog, "gain_places")
        row += _par_tableau_arrondi(prog, "diff_points")
        row += [round(cum["diff_points"], 1) if cum["diff_points"] is not None else None]
        row += _par_tableau_arrondi(prog, "diff_relatif", echelle=100)
        row += [round(cum["diff_relatif"] * 100, 1) if cum["diff_relatif"] is not None else None]
        ws.append(row)

    _appliquer_mise_en_forme(ws, headers, 2, ws.max_row)

    groupes = _grouper(
        headers,
        ("Matchs S", "% total"),
        ("Indice performance S", "Indice global (global)"),
        ("Matchs avec partenaire club (D+M)", "Delta % partenaire (avec - sans)"),
        ("Meilleur partenaire (D+M)", "Points cote marqués ensemble"),
        ("Nb tournois individuels", "Nb interclubs (par jour)"),
        ("Classement sept. S", "Classement sept. M"),
        ("Place sept. S", "Cote actuelle M"),
        ("Gain places S", "Diff relative cumulée (%)"),
    )
    for c0, c1 in groupes:
        _encadrer_groupe(ws, c0, c1, 1, ws.max_row)
    _ajouter_filtre(ws, len(headers), ws.max_row)
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
    _ajouter_filtre(ws, len(header), 1 + len(rows))

    # --- tableau + graphique club par mois ---
    ws.append([])
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

    # --- distribution : nombre de joueurs par nombre de tournois individuels ---
    ws.append([])
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


def _write_club_sheet(wb, all_stats):
    ws = wb.create_sheet("Stats club")

    par_mois = defaultdict(lambda: {cat: {"joues": 0, "victoires": 0} for cat in CLUB_CATEGORIES})

    def _add(mois, cat, victoire):
        b = par_mois[mois][cat]
        b["joues"] += 1
        b["victoires"] += 1 if victoire else 0

    for s in all_stats:
        sexe = s["Sexe"]
        genre_total = "H (tous tableaux)" if sexe == "H" else "F (tous tableaux)"
        for m in s["match_log"]:
            tableau, mois, victoire = m["tableau"], m["mois"], m["victoire"]
            _add(mois, genre_total, victoire)
            if tableau == "Mixte":
                _add(mois, "MX", victoire)
            else:
                _add(mois, tableau[0] + _GENRE_CODE[sexe], victoire)  # SH/SD/DH/DD
                _add(mois, f"{tableau[0]} (unisexe)", victoire)  # S/D (unisexe)

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
    _ajouter_filtre(ws, len(header), last_row)

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

    return ws


def write_stats_excel(all_stats, output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # feuille par défaut vide

    for title, tableau, sexe in DISCIPLINE_SHEETS:
        _write_discipline_sheet(wb, title, tableau, sexe, all_stats)

    _write_tournois_sheet(wb, all_stats)
    _write_bilan_sheet(wb, all_stats)
    _write_club_sheet(wb, all_stats)

    wb.save(output_path)
    return output_path
