#Requires -Version 5.0

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

# ---------------------------------------------------------------------------
# Configuracao do projeto
#
# Este script pressupoe o ambiente ja configurado (WSL + Ubuntu + repositorio
# clonado + venv + .env). Para preparar uma maquina nova, rode antes:
#     wsl -d Ubuntu -- bash -c "bash '/caminho/para/setup_credito.sh'"
# ---------------------------------------------------------------------------

$RepoDir = "~/code/splor-mg/siafi-automacao-credito"

Write-Host ""
Write-Host "=== Robo SIAFI - Remanejamento de Credito ===" -ForegroundColor Cyan
Write-Host ""

# Verificacao minima: ambiente presente
wsl -d Ubuntu -- bash -c "test -d $RepoDir/venv && test -f $RepoDir/.env" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRO: ambiente nao configurado nesta maquina." -ForegroundColor Red
    Write-Host "Repositorio, venv ou .env ausentes em $RepoDir." -ForegroundColor Red
    Write-Host "Rode o setup_credito.sh antes de usar este atalho." -ForegroundColor Red
    exit 1
}

# Atualizar o repositorio
Write-Host "Atualizando o robo (git pull na main)..." -ForegroundColor Cyan
wsl -d Ubuntu -- bash -c "cd $RepoDir && git checkout main && git pull origin main"
if ($LASTEXITCODE -ne 0) {
    Write-Host "[aviso] Nao foi possivel atualizar via git pull. Rodando a versao local atual." -ForegroundColor Yellow
}
Write-Host ""

# O login.py orquestra todo o fluxo:
#   consolida.py -> revisao no Excel -> analise_saldo.py -> solicitacoes
Write-Host "Iniciando o robo SIAFI (credito)..." -ForegroundColor Cyan
wsl -d Ubuntu -- bash -c "cd $RepoDir && source venv/bin/activate && PYTHONIOENCODING=utf-8 python siafi_automacao/login.py"

$roboExit = $LASTEXITCODE

if ($roboExit -eq 2) {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "  Processo INTERROMPIDO na analise de saldo."                  -ForegroundColor Yellow
    Write-Host "  Nenhuma solicitacao foi enviada ao SIAFI."                   -ForegroundColor Yellow
    Write-Host "  Corrija as dotacoes apontadas acima e rode novamente."       -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
}
elseif ($roboExit -ne 0) {
    Write-Host ""
    Write-Host "O robo encerrou com erro (codigo $roboExit)." -ForegroundColor Red
}

exit $roboExit