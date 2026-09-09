"""Écriture du classeur Excel à 8 onglets (SH/SD/DH/DD/MX, Tournois, Bilan
joueur, Stats club) à partir des stats produites par stats.py."""
from collections import defaultdict

import openpyxl
from openpyxl.chart import BarChart, Reference

DISCIPLINE_SHEETS = [
    # (titre, tableau, sexe ou None si le tableau n'est pas séparé par genre)
    ("SH", "Simple", "H"),
    ("SD", "Simple", "F"),
    ("DH", "Double", "H"),
    ("DD", "Double", "F"),
    ("MX", "Mixte", None),
]


def _pct(value):
    return round(value, 1)


def _write_discipline_sheet(wb, title, tableau, sexe, all_stats):
    ws = wb.create_sheet(title)
    has_partner = tableau in ("Double", "Mixte")

    headers = ["Nom", "Classement", "Cote", "Matchs joués", "Victoires", "% victoire"]
    if has_partner:
        headers += [
            "Matchs avec partenaire club", "% avec partenaire club",
            "Matchs sans partenaire club", "% sans partenaire club",
            "Delta % (avec - sans)",
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
        rows.append((s["Nom"], entry))
    rows.sort(key=lambda r: -r[1]["indice_global"])

    for nom, entry in rows:
        row = [nom, entry["classement"], entry["cote"], entry["matchs_joues"], entry["victoires"],
               _pct(entry["pct_victoire"])]
        if has_partner:
            p = entry["partenaire"]
            avec, sans = p["avec_partenaire_club"], p["sans_partenaire_club"]
            row += [
                avec["matchs_joues"], _pct(avec["pct_victoire"]),
                sans["matchs_joues"], _pct(sans["pct_victoire"]),
                _pct(p["delta_pct"]),
            ]
        row += [_pct(entry["indice_performance"]), entry["indice_niveau"], _pct(entry["indice_global"])]
        ws.append(row)

    return ws


def _fmt_progression(bloc):
    """classement/place/cote de septembre, place/cote actuelle, gain de
    places, diff de cote et diff relative (%) pour un tableau."""
    diff_points, diff_rel = bloc["diff_points"], bloc["diff_relatif"]
    return [
        bloc["septembre_classement"], bloc["septembre_rang"], bloc["septembre_points"],
        bloc["actuel_rang"], bloc["actuel_points"],
        bloc["gain_places"],
        round(diff_points, 1) if diff_points is not None else None,
        round(diff_rel * 100, 1) if diff_rel is not None else None,
    ]


def _write_bilan_sheet(wb, all_stats):
    ws = wb.create_sheet("Bilan joueur")
    ws.append([
        "Nom", "Sexe", "Ordre tableau",
        "Matchs S", "Victoires S", "% S",
        "Matchs D", "Victoires D", "% D",
        "Matchs M", "Victoires M", "% M",
        "Matchs total", "Victoires total", "% total",
        "Matchs avec partenaire club (D+M)", "% avec partenaire club",
        "Matchs sans partenaire club (D+M)", "% sans partenaire club",
        "Delta % partenaire (avec - sans)",
        "Nb tournois individuels", "Nb interclubs (par jour)",
        "Classement sept. S", "Place sept. S", "Cote sept. S", "Place actuelle S", "Cote actuelle S",
        "Gain places S", "Diff cote S", "Diff relative S (%)",
        "Classement sept. D", "Place sept. D", "Cote sept. D", "Place actuelle D", "Cote actuelle D",
        "Gain places D", "Diff cote D", "Diff relative D (%)",
        "Classement sept. M", "Place sept. M", "Cote sept. M", "Place actuelle M", "Cote actuelle M",
        "Gain places M", "Diff cote M", "Diff relative M (%)",
        "Diff cote cumulée", "Diff relative cumulée (%)",
        "Indice performance", "Indice niveau", "Indice global",
    ])

    rows = sorted(all_stats, key=lambda s: -s["global"]["indice_global"])
    for s in rows:
        simple, double, mixte = (s["par_tableau"][t] for t in ("Simple", "Double", "Mixte"))
        g, t, prog = s["global"], s["tournois"], s["progression"]
        p = s["partenaire_double_mixte"]
        avec, sans = p["avec_partenaire_club"], p["sans_partenaire_club"]
        cum = prog["cumule"]

        row = [
            s["Nom"], s["Sexe"], g["ordre_tableau"],
            simple["matchs_joues"], simple["victoires"], _pct(simple["pct_victoire"]),
            double["matchs_joues"], double["victoires"], _pct(double["pct_victoire"]),
            mixte["matchs_joues"], mixte["victoires"], _pct(mixte["pct_victoire"]),
            g["matchs_joues"], g["victoires"], _pct(g["pct_victoire"]),
            avec["matchs_joues"], _pct(avec["pct_victoire"]),
            sans["matchs_joues"], _pct(sans["pct_victoire"]),
            _pct(p["delta_pct"]),
            t["nb_tournois"], t["nb_interclubs"],
        ]
        row += _fmt_progression(prog["Simple"])
        row += _fmt_progression(prog["Double"])
        row += _fmt_progression(prog["Mixte"])
        row += [
            round(cum["diff_points"], 1) if cum["diff_points"] is not None else None,
            round(cum["diff_relatif"] * 100, 1) if cum["diff_relatif"] is not None else None,
        ]
        row += [_pct(g["indice_performance"]), g["indice_niveau"], _pct(g["indice_global"])]
        ws.append(row)
    return ws


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

    ws.append([])
    ws.append(["Par mois (club)"])
    header_row = ws.max_row + 1
    ws.append(["Mois", "Tournois weekend", "Tournois soirée", "Interclubs"])
    first_data_row = ws.max_row + 1
    for mois in mois_tries:
        b = par_mois_club[mois]
        ws.append([mois, b["weekend"], b["soiree"], b["interclub"]])
    last_data_row = ws.max_row

    if last_data_row >= first_data_row:
        chart = BarChart()
        chart.type = "col"
        chart.title = "Tournois et interclubs par mois"
        chart.y_axis.title = "Nombre"
        chart.x_axis.title = "Mois"
        data = Reference(ws, min_col=2, max_col=4, min_row=header_row, max_row=last_data_row)
        categories = Reference(ws, min_col=1, min_row=first_data_row, max_row=last_data_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        ws.add_chart(chart, f"F{header_row}")

    return ws


# catégories affichées en colonnes, chacune avec 3 sous-colonnes
# (matchs/victoires/%) : SH/SD/DH/DD/MX comme les onglets par tableau, puis
# des vues transverses (tout tableau confondu par genre, et Simple/Double
# unisexe).
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

    for mois in sorted(par_mois):
        row = [mois]
        for cat in CLUB_CATEGORIES:
            b = par_mois[mois][cat]
            pct = (b["victoires"] / b["joues"] * 100) if b["joues"] else 0.0
            row += [b["joues"], b["victoires"], _pct(pct)]
        ws.append(row)

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
