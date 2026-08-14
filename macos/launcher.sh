#!/bin/bash
# Lancé par macOS quand on double-clique sur pointsIC.app (CFBundleExecutable).
# Ouvre une vraie fenêtre Terminal et y exécute le binaire réel, pour que
# les print()/erreurs/la pause finale soient visibles (un .app "console"
# lancé directement par le Finder tourne sans terminal attaché).
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BIN="$DIR/bin/pointsIC"
# dossier où l'utilisateur voit pointsIC.app dans le Finder (Desktop,
# Téléchargements, Applications...) : c'est là que le xlsx doit apparaître,
# pas dans les entrailles du bundle où vit le binaire réel.
OUTPUT_DIR="$(cd "$DIR/../../.." && pwd)"

osascript <<APPLESCRIPT
tell application "Terminal"
    activate
    do script "POINTSIC_OUTPUT_DIR='$OUTPUT_DIR' '$BIN'"
end tell
APPLESCRIPT
