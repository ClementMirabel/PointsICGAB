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

    headers = ["Nom", "Classement", "Matchs joués", "Victoires", "% victoire"]
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
        row = [nom, entry["classement"], entry["matchs_joues"], entry["victoires"],
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


def _write_bilan_sheet(wb, all_stats):
    ws = wb.create_sheet("Bilan joueur")
    ws.append([
        "Nom", "Sexe", "Meilleur tableau",
        "Matchs S", "Victoires S", "% S",
        "Matchs D", "Victoires D", "% D",
        "Matchs M", "Victoires M", "% M",
        "Matchs total", "Victoires total", "% total",
        "Nb tournois individuels", "Nb interclubs (par jour)",
        "Indice performance", "Indice niveau", "Indice global",
    ])

    rows = sorted(all_stats, key=lambda s: -s["global"]["indice_global"])
    for s in rows:
        simple, double, mixte = (s["par_tableau"][t] for t in ("Simple", "Double", "Mixte"))
        g, t = s["global"], s["tournois"]
        ws.append([
            s["Nom"], s["Sexe"], g["meilleur_tableau"],
            simple["matchs_joues"], simple["victoires"], _pct(simple["pct_victoire"]),
            double["matchs_joues"], double["victoires"], _pct(double["pct_victoire"]),
            mixte["matchs_joues"], mixte["victoires"], _pct(mixte["pct_victoire"]),
            g["matchs_joues"], g["victoires"], _pct(g["pct_victoire"]),
            t["nb_tournois"], t["nb_interclubs"],
            _pct(g["indice_performance"]), g["indice_niveau"], _pct(g["indice_global"]),
        ])
    return ws


def _write_tournois_sheet(wb, all_stats):
    ws = wb.create_sheet("Tournois")
    ws.append(["Nom", "Nb tournois individuels", "dont weekend", "dont soirée", "Nb interclubs (par jour)"])

    rows = []
    par_mois = defaultdict(lambda: {"weekend": 0, "soiree": 0, "interclub": 0})
    for s in all_stats:
        t = s["tournois"]
        if t["nb_tournois"] or t["nb_interclubs"]:
            rows.append((s["Nom"], t["nb_tournois"], t["nb_tournois_weekend"], t["nb_tournois_soiree"], t["nb_interclubs"]))
        for tournoi in t["tournois"]:
            par_mois[tournoi["date"].strftime("%Y-%m")][tournoi["type"]] += 1
        for d in t["interclubs_dates"]:
            par_mois[d.strftime("%Y-%m")]["interclub"] += 1
    rows.sort(key=lambda r: -r[1])
    for row in rows:
        ws.append(row)

    ws.append([])
    ws.append(["Par mois (club)"])
    header_row = ws.max_row + 1
    ws.append(["Mois", "Tournois weekend", "Tournois soirée", "Interclubs"])
    first_data_row = ws.max_row + 1
    for mois in sorted(par_mois):
        b = par_mois[mois]
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


def _write_club_sheet(wb, all_stats):
    ws = wb.create_sheet("Stats club")
    ws.append(["Tableau", "Mois", "Genre", "Matchs joués", "Victoires", "% victoire"])

    # (tableau, mois, genre) -> {joues, victoires}
    agg = defaultdict(lambda: {"joues": 0, "victoires": 0})
    for s in all_stats:
        for m in s["match_log"]:
            key = (m["tableau"], m["mois"], s["Sexe"])
            bucket = agg[key]
            bucket["joues"] += 1
            bucket["victoires"] += 1 if m["victoire"] else 0

    for (tableau, mois, genre), b in sorted(agg.items()):
        pct = (b["victoires"] / b["joues"] * 100) if b["joues"] else 0.0
        ws.append([tableau, mois, genre, b["joues"], b["victoires"], _pct(pct)])

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
