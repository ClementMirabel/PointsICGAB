"""Écriture du classeur Excel à 8 onglets (SH/SD/DH/DD/MX, Tournois, Bilan
joueur, Stats club) à partir des stats produites par stats.py, avec mise en
forme conditionnelle (classements colorés comme sur le site, dégradé de
performance par colonne)."""
from collections import Counter, defaultdict

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Font, PatternFill
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

# colonnes texte qu'on ne colore pas en dégradé (déjà colorées autrement,
# ou pas un indicateur de performance)
CLASSEMENT_HEADERS = {"Classement", "Classement sept. S", "Classement sept. D", "Classement sept. M"}
TEXT_HEADERS = {"Nom", "Sexe", "Mois", "Meilleur partenaire", "Meilleur partenaire (D+M)",
                "Meilleur partenaire (points)", "Meilleur partenaire (D+M, points)",
                "Ordre tableau"} | CLASSEMENT_HEADERS


def _est_colonne_diff(header):
    """Colonnes pouvant être négatives (delta/diff/gain de places) : dégradé
    divergent rouge-blanc-vert centré sur 0, plutôt que le dégradé blanc->vert
    habituel (qui n'aurait pas de sens sur une échelle qui traverse 0)."""
    return "Delta" in header or "Diff" in header or "Gain places" in header


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


def _degrade_diverge(ws, col_idx, first_row, last_row):
    """Rouge (négatif) -> blanc (zéro) -> vert (positif), avec des barres de
    données par-dessus - pour les colonnes delta/diff qui peuvent être
    négatives (un dégradé blanc->vert n'aurait pas de sens ici : le "bas"
    de l'échelle doit se voir comme mauvais, pas juste "moins vert")."""
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
    barres = DataBarRule(start_type="min", end_type="max", color="638EC6", showValue=True)
    ws.conditional_formatting.add(plage, barres)


def _appliquer_mise_en_forme(ws, headers, first_row, last_row):
    """Classements colorés par tableau, dégradé de performance sur les
    autres colonnes numériques (divergent pour les deltas), % affiché avec
    un signe pourcentage - une seule fois toutes les lignes écrites."""
    for idx, header in enumerate(headers, start=1):
        if header in CLASSEMENT_HEADERS:
            _colorer_classement(ws, idx, first_row, last_row)
        elif header == "Ordre tableau":
            _colorer_ordre_tableau(ws, idx, first_row, last_row)
        elif header in TEXT_HEADERS:
            pass
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


# ------------------------------------------------------- onglets par tableau

def _fmt_meilleur_partenaire(mp):
    if not mp:
        return [None, None, None, None, None, None]
    return [mp["nom"], mp["matchs_joues"], mp["victoires"], _pct(mp["pct_victoire"]),
            round(mp["indice_performance"], 2), mp.get("points_cote_total")]


def _fmt_meilleur_partenaire_points(mp):
    if not mp or mp.get("points_cote_total") is None:
        return [None, None, None]
    return [mp["nom"], mp["matchs_joues"], mp["points_cote_total"]]


def _write_discipline_sheet(wb, title, tableau, sexe, all_stats):
    ws = wb.create_sheet(title)
    has_partner = tableau in ("Double", "Mixte")
    inclure_sexe = tableau == "Mixte"  # seul tableau qui mélange les deux genres

    headers = ["Nom"]
    if inclure_sexe:
        headers += ["Sexe"]
    headers += ["Classement", "Cote", "Matchs joués", "Victoires", "% victoire"]
    if has_partner:
        headers += [
            "Matchs avec partenaire club", "% avec partenaire club",
            "Matchs sans partenaire club", "% sans partenaire club",
            "Delta % (avec - sans)",
            "Meilleur partenaire", "Matchs avec lui/elle", "Victoires avec lui/elle",
            "% victoire avec lui/elle", "Indice perf. avec lui/elle", "Points cote marqués ensemble",
            "Meilleur partenaire (points)", "Matchs avec lui/elle (points)",
            "Points cote cumulés avec lui/elle",
        ]
    headers += ["Indice performance", "Indice niveau", "Indice global"]
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
        if has_partner:
            p = entry["partenaire"]
            avec, sans = p["avec_partenaire_club"], p["sans_partenaire_club"]
            row += [
                avec["matchs_joues"], _pct(avec["pct_victoire"]),
                sans["matchs_joues"], _pct(sans["pct_victoire"]),
                _pct(p["delta_pct"]),
            ]
            row += _fmt_meilleur_partenaire(entry["meilleur_partenaire"])
            row += _fmt_meilleur_partenaire_points(entry["meilleur_partenaire_points"])
        row += [_pct(entry["indice_performance"]), entry["indice_niveau"], _pct(entry["indice_global"])]
        ws.append(row)

    _appliquer_mise_en_forme(ws, headers, 2, ws.max_row)
    return ws


# ------------------------------------------------------------- Bilan joueur

def _fmt_cotes(bloc):
    return [bloc["septembre_rang"], bloc["septembre_points"], bloc["actuel_rang"], bloc["actuel_points"]]


def _fmt_diffs(bloc):
    diff_points, diff_rel = bloc["diff_points"], bloc["diff_relatif"]
    return [
        bloc["gain_places"],
        round(diff_points, 1) if diff_points is not None else None,
        round(diff_rel * 100, 1) if diff_rel is not None else None,
    ]


def _write_bilan_sheet(wb, all_stats):
    ws = wb.create_sheet("Bilan joueur")
    headers = [
        "Nom", "Sexe", "Ordre tableau",
        "Matchs S", "Victoires S", "% S",
        "Matchs D", "Victoires D", "% D",
        "Matchs M", "Victoires M", "% M",
        "Matchs total", "Victoires total", "% total",
        "Matchs avec partenaire club (D+M)", "% avec partenaire club",
        "Matchs sans partenaire club (D+M)", "% sans partenaire club",
        "Delta % partenaire (avec - sans)",
        "Meilleur partenaire (D+M)", "Matchs avec lui/elle", "Victoires avec lui/elle",
        "% victoire avec lui/elle", "Indice perf. avec lui/elle", "Points cote marqués ensemble",
        "Meilleur partenaire (D+M, points)", "Matchs avec lui/elle (points)",
        "Points cote cumulés avec lui/elle",
        "Nb tournois individuels", "Nb interclubs (par jour)",
        # regroupé par type de métrique (classements, puis cotes/places, puis diffs) plutôt que par tableau
        "Classement sept. S", "Classement sept. D", "Classement sept. M",
        "Place sept. S", "Cote sept. S", "Place actuelle S", "Cote actuelle S",
        "Place sept. D", "Cote sept. D", "Place actuelle D", "Cote actuelle D",
        "Place sept. M", "Cote sept. M", "Place actuelle M", "Cote actuelle M",
        "Gain places S", "Diff cote S", "Diff relative S (%)",
        "Gain places D", "Diff cote D", "Diff relative D (%)",
        "Gain places M", "Diff cote M", "Diff relative M (%)",
        "Diff cote cumulée", "Diff relative cumulée (%)",
        "Indice performance", "Indice niveau", "Indice global",
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
            avec["matchs_joues"], _pct(avec["pct_victoire"]),
            sans["matchs_joues"], _pct(sans["pct_victoire"]),
            _pct(p["delta_pct"]),
        ]
        row += _fmt_meilleur_partenaire(s["meilleur_partenaire_double_mixte"])
        row += _fmt_meilleur_partenaire_points(s["meilleur_partenaire_double_mixte_points"])
        row += [t["nb_tournois"], t["nb_interclubs"]]

        row += [prog["Simple"]["septembre_classement"], prog["Double"]["septembre_classement"],
                prog["Mixte"]["septembre_classement"]]
        row += _fmt_cotes(prog["Simple"])
        row += _fmt_cotes(prog["Double"])
        row += _fmt_cotes(prog["Mixte"])
        row += _fmt_diffs(prog["Simple"])
        row += _fmt_diffs(prog["Double"])
        row += _fmt_diffs(prog["Mixte"])
        row += [
            round(cum["diff_points"], 1) if cum["diff_points"] is not None else None,
            round(cum["diff_relatif"] * 100, 1) if cum["diff_relatif"] is not None else None,
        ]
        row += [_pct(g["indice_performance"]), g["indice_niveau"], _pct(g["indice_global"])]
        ws.append(row)

    _appliquer_mise_en_forme(ws, headers, 2, ws.max_row)
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
        ws.add_chart(bar, f"AF{header_row}")

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
