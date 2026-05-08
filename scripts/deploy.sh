#!/bin/bash
# Script de déploiement ALUPS-Mannequin exécuté sur le serveur via SSH
# Étapes : backup BD, pull, install deps si requirements.txt a changé,
# restart du service, healthcheck, rollback automatique en cas d'échec.

set -e

APP_DIR="/var/www/alups-mannequin"
BACKUP_DIR="$APP_DIR/backups"
SERVICE="alups-mannequin"
HEALTH_URL="http://127.0.0.1:5051/"
TIMESTAMP=$(date +%Y%m%d-%H%M%S)
KEEP_BACKUPS=10

cd "$APP_DIR"

echo "[1/6] Sauvegarde de la base de données..."
mkdir -p "$BACKUP_DIR"
if compgen -G "$APP_DIR/data/*.db" > /dev/null; then
    for db in "$APP_DIR"/data/*.db; do
        cp "$db" "$BACKUP_DIR/$(basename "$db").$TIMESTAMP"
    done
    for db in "$APP_DIR"/data/*.db; do
        name=$(basename "$db")
        ls -1t "$BACKUP_DIR/$name".* 2>/dev/null | tail -n +$((KEEP_BACKUPS + 1)) | xargs -r rm --
    done
    echo "  -> Backup OK dans $BACKUP_DIR"
else
    echo "  -> Aucune BD trouvée, on saute"
fi

echo "[2/6] Mémorisation du commit actuel pour rollback..."
PREV_COMMIT=$(git rev-parse HEAD)
echo "  -> Commit avant pull : $PREV_COMMIT"

echo "[3/6] git pull..."
git fetch origin main
git reset --hard origin/main
NEW_COMMIT=$(git rev-parse HEAD)
echo "  -> Nouveau commit : $NEW_COMMIT"

echo "[4/6] Mise à jour des dépendances Python si requirements.txt a changé..."
if [ "$PREV_COMMIT" = "$NEW_COMMIT" ]; then
    echo "  -> Pas de nouveau commit"
elif git diff --name-only "$PREV_COMMIT" "$NEW_COMMIT" | grep -q "^requirements.txt$"; then
    echo "  -> requirements.txt modifié, pip install"
    "$APP_DIR/venv/bin/pip" install -r requirements.txt
else
    echo "  -> requirements.txt inchangé"
fi

echo "[5/6] Permissions et redémarrage du service..."
chown -R www-data:www-data "$APP_DIR"
systemctl restart "$SERVICE"

echo "[6/6] Healthcheck..."
sleep 2
SUCCESS=0
for i in 1 2 3 4 5 6 7 8 9 10; do
    code=$(curl -s -o /dev/null -w "%{http_code}" "$HEALTH_URL" || echo "000")
    if echo "$code" | grep -qE "^(200|301|302|401|403)$"; then
        SUCCESS=1
        echo "  -> Service répond (HTTP $code, tentative $i)"
        break
    fi
    sleep 2
done

if [ "$SUCCESS" -ne 1 ] || ! systemctl is-active --quiet "$SERVICE"; then
    echo "ECHEC du healthcheck. Rollback vers $PREV_COMMIT"
    git reset --hard "$PREV_COMMIT"
    chown -R www-data:www-data "$APP_DIR"
    systemctl restart "$SERVICE"
    echo "Rollback effectué. Le déploiement a échoué."
    exit 1
fi

echo "Déploiement réussi : $(git rev-parse --short HEAD)"
