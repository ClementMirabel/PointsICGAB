# PointsICGAB

Calcule la Valeur IC de tous les joueurs du club GAB38 à partir des classements publics de [myffbad.fr](https://myffbad.fr/recherche/les-tops), et exporte un Excel filtré (Valeur IC ≥ 24) et trié.

## Télécharger

| Système | Lien | Fichier |
|---|---|---|
| **Mac** (Apple Silicon M1/M2/M3...) | [⬇️ Télécharger](https://github.com/ClementMirabel/PointsICGAB/releases/latest/download/pointsIC-macos.app.zip) | `pointsIC-macos.app.zip` |
| **Windows** | [⬇️ Télécharger](https://github.com/ClementMirabel/PointsICGAB/releases/latest/download/pointsIC.exe) | `pointsIC.exe` |

Pas besoin d'installer Python : ce sont des exécutables autonomes.

### Installation — Mac

1. Double-clique sur le fichier téléchargé pour le dézipper (tu obtiens `pointsIC.app`).
2. Premier lancement : **clic droit** sur `pointsIC.app` → **Ouvrir** → confirmer dans la boîte de dialogue (macOS bloque par défaut les applications non signées par un compte développeur Apple). Les lancements suivants se font par simple double-clic.
3. Un Terminal s'ouvre et affiche la progression.

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
