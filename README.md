# PointsICGAB — stats et interclubs du GAB38

Statistiques des joueurs du Grenoble Alpes Badminton (GAB38), suivi des interclubs et planificateur de compositions, mis à jour automatiquement chaque semaine à partir de [myffbad.fr](https://myffbad.fr) et [icbad.ffbad.org](https://icbad.ffbad.org/instance/GAB38).

## 🌐 La page web

**[clementmirabel.github.io/PointsICGAB](https://clementmirabel.github.io/PointsICGAB/)** — même adresse chaque semaine, mise à jour dans la nuit de mercredi à jeudi.

| Onglet | Contenu |
|---|---|
| **Club** | Aperçu (avec l'évolution depuis la semaine précédente), classement de forme, compétitions et tournois, stats par catégorie (chaque chiffre ouvre la liste des matchs concernés), profil de jeu, records et rivalités internes, suivi (progression depuis septembre, joueurs proches du classement supérieur, joueurs sans match récent), meilleures paires de double et de mixte |
| **Interclubs** | Équipes 1 à 6 : rencontres, joueurs alignés sur chaque match, bilans par match et par joueur, classement de poule, fiche de chaque adversaire (résultats de poule, composition type, effectif, confrontations avec le GAB), calendriers à ajouter à son agenda ; page **Pilotage** : enjeux de poule (projection, rencontres clés), profondeur d'effectif par équipe, charge et présence des joueurs, classés jamais alignés, bilan par poste et par adversaire |
| **Planificateur** | Affectation des joueurs par week-end d'interclubs, composition des rencontres, contrôles réglementaires, export Excel |
| **Joueurs** | Une fiche par joueur : forme par tableau, évolution de la cote et de l'indice de forme semaine après semaine, faits marquants, confrontations avec les autres joueurs du club, partenaires, matchs, interclubs |

### Le planificateur, en bref

- **Par week-end** : glisser les joueurs vers les équipes (ou choisir l'équipe dans leur ligne). Un joueur ne peut jouer que dans une équipe par week-end. Le bouton « Copier les effectifs du week-end précédent » reprend les équipes de la semaine d'avant.
- **Par rencontre** : choisir les joueurs de chaque match parmi ceux affectés à l'équipe, ou cliquer « Proposer une composition » (placement automatique qui respecte toutes les règles et maximise les écarts de cote, à ajuster ensuite) ; la composition type de l'adversaire (icbad) est préremplie. La hiérarchie utilise les cotes de J-10 dès qu'elles sont publiées.
- **Contrôles** :
  - valeur d'équipe (3 meilleurs joueurs + 3 meilleures joueuses), ICN pour l'équipe 1, ICR pour les autres, et hiérarchie des valeurs entre équipes (règlement ICR AURA art. 5) ;
  - 2 matchs maximum par joueur et pas deux dans la même discipline ; hiérarchie des cotes (SH1 ≥ SH2…) ;
  - joueurs « brûlés » : ICR AURA art. 12 pour les équipes régionales, CD38 art. 8 pour les équipes départementales (bandeau rouge si interdit, jaune si le faire jouer le bloquera ensuite) ;
  - 2 joueurs mutés maximum par rencontre (ICR art. 13, CD38 art. 11-4).
- **Week-ends passés** : affichage de ce qui s'est réellement joué (feuilles icbad), avec les valeurs des joueurs figées à l'export qui a suivi le week-end.
- **Sauvegarde** : le plan est enregistré **dans le navigateur** de chacun ; il est conservé d'une semaine à l'autre mais n'est pas partagé. Pour l'échanger entre capitaines : « Exporter le plan », puis « Importer un plan » de l'autre côté.
- **Export Excel** d'un week-end, au format du classeur « INTERCLUBS » (propositions d'équipe, valeurs, compositions).

Les jeunes (Minibad à Junior, catégorie d'âge myffbad) portent un badge U9 à U19 partout où leur nom apparaît, et un filtre « Jeunes » existe dans le classement et le suivi.

## 📅 Calendriers des interclubs

À ajouter à son agenda (Google Agenda : « Autres agendas » → « À partir de l'URL », pour une mise à jour automatique) : [toutes les équipes](https://clementmirabel.github.io/PointsICGAB/calendrier-interclubs.ics), ou une équipe : `https://clementmirabel.github.io/PointsICGAB/calendrier-equipe-N.ics` (N de 1 à 6).

## ⬇️ Fichiers Excel

Liens fixes, toujours vers la dernière version :

- [**statsJoueurs.xlsx**](https://github.com/ClementMirabel/PointsICGAB/releases/download/weekly-data/statsJoueurs.xlsx) — export complet des stats joueurs et club ;
- [**pointsIC.xlsx**](https://github.com/ClementMirabel/PointsICGAB/releases/download/weekly-data/pointsIC.xlsx) — Valeur IC de tous les joueurs du club (Valeur IC ≥ 24), triée.

---

## Pour la maintenance

### Fonctionnement

Le workflow [`weekly-export.yml`](.github/workflows/weekly-export.yml) a deux modes :

- **export complet** — chaque jeudi à 01:00 UTC, ou lancement manuel (Actions → « Run workflow ») : étapes ci-dessous, ~35 min. Certaines machines GitHub n'arrivent pas à joindre myffbad.fr : un test de connexion (30 s) le détecte et l'export est relancé une fois sur une autre machine ([`export-complet`](.github/actions/export-complet/action.yml), utilisé par les deux tentatives) ;
- **régénération** — à chaque push sur `main` qui touche le code : tests, puis page et Excel recalculés depuis le dump du dernier export complet (`--from-json`), sans accès à myffbad (~2 min). Si aucun dump n'existe encore, le push lance un export complet.

Export complet :

1. `pointsIC.py` : roster public du club → `pointsIC.xlsx`.
2. `statsJoueurs.py` :
   - roster du club (`roster.py`) et interclubs des équipes 1 à 6 (`icbad.py`, pages publiques) ;
   - connexion à myffbad (Selenium) et récupération des résultats de la saison, de la cote et de l'historique de classement des joueurs classés R6 ou mieux, puis résultats seuls des autres joueurs d'interclubs (points de cote) ;
   - calcul des stats (`stats.py`), de l'Excel (`export_excel.py`) et de la page (`export_web.py` + gabarit [`web/statsJoueurs.template.html`](web/statsJoueurs.template.html)) ;
   - valeurs des joueurs figées pour les week-ends d'interclubs terminés (`historique_ic.py`).
3. Publication : la page sur **GitHub Pages**, les fichiers Excel dans la release `weekly-data` (remplacée seulement par un export complet réussi ; une régénération ne publie que la page).

Données conservées entre deux runs :

| Donnée | Où | Rôle |
|---|---|---|
| `historique_valeurs.json` | branche `donnees` du dépôt | valeurs figées par week-end d'interclubs, et instantané de chaque export (cotes de tout l'effectif → cotes J-10 ; indices de forme → tendances) |
| `resume.json` | publié avec la page | chiffres clés, comparés par l'export suivant (Aperçu) |
| `statsJoueurs.json.gz` | branche `donnees` | dump du dernier export complet, utilisé par les régénérations |
| `statsJoueurs_precedent.json` | branche `donnees` | résumé auquel ce dump a été comparé : une régénération garde les mêmes écarts dans l'Aperçu |

### Tests et alertes

- `python -m unittest discover -s tests -t .` : règles de calcul et parcours de la page dans Chrome, sur des données **entièrement fictives** (`tests/donnees_fictives.py`). Lancés à chaque push (workflow `tests.yml`) et avant chaque export ou régénération, qui s'arrête si un test échoue.
- Si l'export échoue (y compris sa seconde tentative), une issue « Export hebdomadaire en échec » est ouverte (notification GitHub par mail), puis fermée automatiquement au premier export réussi.
- La page tient dans un seul fichier HTML (~1,4 Mo mi-novembre, ~200 Ko transférés grâce à la compression) ; un test alerte au-delà de 8 Mo.

### Prérequis du dépôt

- Secrets `MYFFBAD_LICENCE` et `MYFFBAD_PASSWORD` (compte myffbad utilisé pour la connexion).
- Settings → Pages → Source : **GitHub Actions**.
- Issues activées (alertes d'échec de l'export).

### En local

```
pip install -r requirements.txt
python pointsIC.py                       # -> pointsIC.xlsx
python statsJoueurs.py                   # export complet (~35 min, demande licence et mot de passe myffbad)
python statsJoueurs.py 06627061 00123456 # juste ces licences (test rapide)
python statsJoueurs.py --from-json statsJoueurs.json  # régénère Excel + page depuis un dump, sans scraper
python icbad.py                          # interclubs seuls -> icbad.json
```

Pour itérer sur la page : récupérer `statsJoueurs.json.gz` sur la branche `donnees`, le décompresser, puis `--from-json`.
