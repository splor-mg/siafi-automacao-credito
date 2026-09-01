#!/usr/bin/env bash
# Sequencia unica de execucao do robo de CREDITO.
#
# Usado tanto pelo robo_credito.bat (duplo-clique no Windows) quanto pelo bot
# do Telegram, para que os dois caminhos executem exatamente a mesma coisa.
#
# Codigos de saida:
#   0   sucesso (ou cancelado na conferencia manual)
#   1   falha
#   2   interrompido na analise de saldo — nada foi enviado ao SIAFI
#   10  ja existe uma execucao em andamento (lock tomado)
#
# Sem '-e' de proposito: a falha do 'git pull' precisa apenas avisar e seguir
# com a versao local, nao abortar a execucao.
set -uo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CARIMBO="$(date +%Y%m%d-%H%M%S)"

mkdir -p "$REPO/data/logs"

LOG="${ROBO_LOG:-$REPO/data/logs/robo-$CARIMBO.log}"
export RELATO_ARQUIVO="${RELATO_ARQUIVO:-$REPO/data/logs/relato-$CARIMBO.jsonl}"

# LOCK COMPARTILHADO com o robo de cota, de proposito e fora dos dois repos:
# ambos entram no SIAFI com o MESMO usuario, e duas sessoes simultaneas fazem
# o mainframe recusar com 'UNABLE TO ESTABLISH SESSION'.
LOCK="${SIAFI_LOCK:-$HOME/.siafi-robo.lock}"
exec 9>"$LOCK"
if ! flock -n 9; then
    echo "Ja existe uma execucao do robo (cota ou credito) em andamento." >&2
    exit 10
fi

printf '%s\n%s\n' "$LOG" "$RELATO_ARQUIVO" > "$REPO/data/.ultima_execucao"

find "$REPO/data/logs" -maxdepth 1 -type f \
    \( -name 'robo-*.log' -o -name 'relato-*.jsonl' \) \
    -mtime +"${LOG_RETENCAO_DIAS:-30}" -delete 2>/dev/null

{
    echo "=== Robo SIAFI (credito) - $CARIMBO ==="
    cd "$REPO" || exit 1

    echo "Atualizando o robo (git pull na main)..."
    if ! { git checkout main && git pull origin main; }; then
        echo "[aviso] Nao foi possivel atualizar via git pull. Rodando a versao local atual."
    fi

    echo "Iniciando o robo de credito..."
    # shellcheck disable=SC1091
    if ! source venv/bin/activate; then
        echo "[erro] Nao foi possivel ativar o ambiente virtual (venv ausente ou corrompida)."
        echo "       Rode o setup.sh novamente para recriar a venv."
        exit 1
    fi

    PYTHONIOENCODING=utf-8 python siafi_automacao/login.py
} 2>&1 | tee "$LOG"

exit "${PIPESTATUS[0]}"
