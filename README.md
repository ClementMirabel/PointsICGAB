# PointsICGAB

Calcule la Valeur IC de tous les joueurs du club GAB38 à partir des classements publics de [myffbad.fr](https://myffbad.fr/recherche/les-tops), et exporte un Excel filtré (Valeur IC ≥ 24) et trié.

## Fichier Excel (mise à jour automatique)

Pas besoin d'installer quoi que ce soit : un export à jour est généré tous les **jeudis vers midi** et disponible ici :

[**⬇️ Télécharger pointsIC.xlsx**](https://github.com/ClementMirabel/PointsICGAB/releases/download/weekly-data/pointsIC.xlsx)

C'est un lien fixe : il pointe toujours vers la dernière version générée, pas besoin de le remettre à jour.

## Page web : stats joueurs, club, interclubs et planificateur

[**🌐 Ouvrir la page**](https://clementmirabel.github.io/PointsICGAB/) (mise à jour tous les jeudis, même adresse chaque semaine).

- **Club** : vue d'ensemble, classement de forme, compétitions, stats par catégorie, profil de jeu, records.
- **Interclubs** : bilan de chaque équipe (rencontres, joueurs alignés, classement de poule), d'après [icbad.ffbad.org](https://icbad.ffbad.org/instance/GAB38).
- **Planificateur** (équipes 1 à 6) : affectation des joueurs par week-end, composition des rencontres, contrôles réglementaires (valeurs ICN/ICR, hiérarchie, joueurs brûlés, mutés).
- **Joueurs** : une fiche par joueur.

Le plan du planificateur est enregistré **dans le navigateur** de chacun : il est conservé d'une semaine à l'autre, mais n'est pas partagé. Pour l'échanger entre capitaines : bouton « Exporter le plan », puis « Importer un plan » de l'autre côté.

L'export Excel complet reste téléchargeable : [**⬇️ statsJoueurs.xlsx**](https://github.com/ClementMirabel/PointsICGAB/releases/download/weekly-data/statsJoueurs.xlsx).

Pour régénérer la page en local sans re-scraper : `python statsJoueurs.py --from-json statsJoueurs.json` (dump produit par chaque run).

## Applications (pour lancer un export à la demande)

Pour générer un export à un autre moment que le jeudi, deux applications autonomes sont aussi disponibles.

### Télécharger

| Système | Lien | Fichier |
|---|---|---|
| **Mac** (Apple Silicon M1/M2/M3...) | [⬇️ Télécharger](https://github.com/ClementMirabel/PointsICGAB/releases/latest/download/pointsIC-macos.app.zip) | `pointsIC-macos.app.zip` |
| **Windows** | [⬇️ Télécharger](https://github.com/ClementMirabel/PointsICGAB/releases/latest/download/pointsIC.exe) | `pointsIC.exe` |

Pas besoin d'installer Python : ce sont des exécutables autonomes.

### Installation — Mac

1. Double-clique sur le fichier téléchargé pour le dézipper (tu obtiens `pointsIC.app`). Fais-le glisser dans ton dossier **Applications** si tu veux.
2. **Débloque l'app (une seule fois)** : l'app n'étant pas signée par un compte développeur Apple (payant), macOS la bloque par défaut. La méthode la plus fiable — le simple clic droit → Ouvrir ne suffit pas toujours ici car l'app contient un programme interne à débloquer aussi :
   - Ouvre **Terminal** (Cmd+Espace, tape "Terminal").
   - Tape `xattr -dr com.apple.quarantine ` (avec l'espace à la fin, sans appuyer sur Entrée).
   - Fais glisser `pointsIC.app` depuis le Finder dans la fenêtre Terminal (ça complète le chemin automatiquement), puis appuie sur Entrée.
3. Double-clique sur `pointsIC.app`. Une fenêtre Terminal s'ouvre toute seule et affiche la progression (macOS peut demander l'autorisation de "contrôler Terminal" la première fois : accepte).

### Installation — Windows

1. Double-clique sur `pointsIC.exe`.
2. Windows SmartScreen va probablement afficher *"Windows a protégé votre ordinateur"* (l'exe n'est pas signé par un éditeur reconnu). Clique sur **Informations complémentaires** → **Exécuter quand même**.
3. Une fenêtre de console s'ouvre et affiche la progression.

### Dans les deux cas

Une fois terminé, le fichier `pointsIC.xlsx` est généré dans le même dossier que l'exécutable.

## Développement

```
pip install -r requirements.txt
python pointsIC.py
```

Un nouveau build Mac + Windows est publié automatiquement (voir [Releases](https://github.com/ClementMirabel/PointsICGAB/releases)) à chaque modification de `pointsIC.py` poussée sur `main`.
