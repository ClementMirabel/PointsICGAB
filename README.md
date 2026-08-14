# PointsICGAB

Calcule la Valeur IC de tous les joueurs du club GAB38 à partir des classements publics de [myffbad.fr](https://myffbad.fr/recherche/les-tops), et exporte un Excel filtré (Valeur IC ≥ 24) et trié.

## Télécharger (Mac, Apple Silicon)

[**⬇️ Télécharger pointsIC.app**](https://github.com/ClementMirabel/PointsICGAB/releases/latest/download/pointsIC.app.zip)

### Installation

1. Double-clique sur le fichier téléchargé pour le dézipper.
2. Premier lancement : **clic droit** sur `pointsIC.app` → **Ouvrir** → confirmer dans la boîte de dialogue (macOS bloque par défaut les applications non signées par un compte développeur Apple). Les lancements suivants se font par simple double-clic.
3. Un Terminal s'ouvre et affiche la progression. Une fois terminé, le fichier `pointsIC.xlsx` est généré dans le même dossier que l'application.

## Développement

```
pip install -r requirements.txt
python pointsIC.py
```

Un nouveau build Mac est publié automatiquement (voir [Releases](https://github.com/ClementMirabel/PointsICGAB/releases)) à chaque modification de `pointsIC.py` poussée sur `main`.
