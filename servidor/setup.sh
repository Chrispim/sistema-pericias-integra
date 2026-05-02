#!/bin/bash
# Setup do servidor Hetzner CX22
# Sistema de Pericias - Integra Contabilidade & Pericias
set -e
GITHUB_USER="Chrispim"
REPO="sistema-pericias-integra"
DOMAIN="pericias.integraconsult.com.br"
APP_DIR="/var/www/pericias"
echo "=== Integra - Setup do Servidor de Pericias ==="
apt-get update -qq && apt-get upgrade -y -qq
apt-get install -y git python3 python3-pip curl ufw fail2ban
apt-get install -y debian-keyring debian-archive-keyring apt-transport-https
curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' | tee /etc/apt/sources.list.d/caddy-stable.list
apt-get update -qq && apt-get install -y caddy
mkdir -p "$APP_DIR"
if [ -d "$APP_DIR/.git" ]; then
  git -C "$APP_DIR" pull
else
  git clone "https://github.com/$GITHUB_USER/$REPO.git" "$APP_DIR"
fi
cat > /etc/caddy/Caddyfile << CADDY
$DOMAIN {
    root * $APP_DIR/painel
    file_server
    handle /dados/* {
        root * $APP_DIR
        file_server
    }
    header /dados/* {
        Access-Control-Allow-Origin "*"
        Access-Control-Allow-Methods "GET"
    }
    log { output file /var/log/caddy/pericias.log }
}
CADDY
mkdir -p /var/log/caddy
systemctl reload caddy
ufw --force enable && ufw allow ssh && ufw allow http && ufw allow https
cat > /usr/local/bin/pericias-deploy << 'DEPLOY'
#!/bin/bash
cd /var/www/pericias
git pull origin main 2>&1
echo "$(date '+%Y-%m-%d %H:%M') Deploy concluido"
DEPLOY
chmod +x /usr/local/bin/pericias-deploy
(crontab -l 2>/dev/null; echo "0 * * * * /usr/local/bin/pericias-deploy >> /var/log/pericias-deploy.log 2>&1") | crontab -
echo "=== Setup concluido! Acesse: https://$DOMAIN ==="
