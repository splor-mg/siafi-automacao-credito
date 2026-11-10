"""Recalculo das formulas do copia.xlsm pelo proprio Excel, via COM.

A aba ROBO nao guarda dados: ela e derivada por formula da aba 'PREENCHER
AQUI' — VLOOKUP contra a tabela qdd, concatenacao do codigo SIAFI completo,
derivacao condicional de elemento e categoria. O openpyxl escreve os dados de
origem mas NAO tem motor de formulas, entao os valores em cache da aba ROBO
ficam vazios ate alguem abrir o arquivo no Excel.

Era exatamente isso que a antiga confirmacao manual fazia: clicar em 'Habilitar
Edicao' (sair do Modo de Exibicao Protegido) e salvar. Aqui o Excel faz o
servico sozinho — mesma engine, mesmos resultados.

Fica fora do login.py para ser importavel: la o modulo executa o fluxo inteiro
no import, e isso deixava esta parte sem teste.
"""

import os
import shutil
import subprocess

# Como servico do systemd o PATH nao traz os diretorios do Windows, entao
# 'powershell.exe' nao e encontrado pelo nome — mesmo com o interop do WSL
# funcionando normalmente. Pelo caminho absoluto funciona.
_PS_PADRAO = '/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
POWERSHELL = (_PS_PADRAO if os.path.exists(_PS_PADRAO)
              else (shutil.which('powershell.exe') or 'powershell.exe'))

# Marca que o PowerShell devolve quando o arquivo abriu como somente leitura.
# So ASCII no texto do comando, de proposito: acento em comando do PowerShell
# 5.1 depende da pagina de codigo do console.
_SOMENTE_LEITURA = 'ABERTO_EM_OUTRO_EXCEL'

_MOTIVO_SOMENTE_LEITURA = (
    'o copia.xlsm já está aberto em outra instância do Excel na máquina do '
    'robô — em geral uma instância escondida que sobrou de uma execução que '
    'falhou. Encerre o Excel pelo Gerenciador de Tarefas e acione de novo')


def _comando(caminho_windows):
    """O script que o PowerShell executa. Devolve 'OK' ou 'ERRO: <motivo>'."""
    caminho = caminho_windows.replace("'", "''")
    return (
        "$ErrorActionPreference='Stop';"
        # Sem isto a resposta vem na pagina de codigo do console (cp850) e o
        # 'ç' de 'Exceção' derrubava o robo com UnicodeDecodeError — justo
        # quando o Excel tinha um erro para contar (10/11/2026, 10:40).
        "[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false);"
        "$xl = $null;"
        "try {"
        "  $xl = New-Object -ComObject Excel.Application;"
        "  $xl.Visible = $false; $xl.DisplayAlerts = $false;"
        "  $xl.AutomationSecurity = 3;"
        f"  $wb = $xl.Workbooks.Open('{caminho}');"
        # Com o arquivo aberto em outra instancia o Excel abre como somente
        # leitura SEM avisar (DisplayAlerts=false), e o Save() adiante nao
        # grava nada nem reclama: o robo so percebia pela aba ROBO vazia.
        f"  if ($wb.ReadOnly) {{ throw '{_SOMENTE_LEITURA}' }};"
        "  $xl.CalculateFullRebuild(); $wb.Save(); $wb.Close($true);"
        "  Write-Output 'OK'"
        "} catch { Write-Output ('ERRO: ' + $_.Exception.Message) }"
        # Fecha o Excel TAMBEM quando deu errado. Sem o finally a instancia
        # escondida ficava viva segurando o copia.xlsm, e toda execucao
        # seguinte caia no 'somente leitura' acima ate alguem matar o processo.
        " finally { if ($xl) { try { $xl.Quit() } catch {} } }"
    )


def _interpretar(stdout, stderr=b''):
    """Traduz a resposta do PowerShell: None se deu certo, senao o motivo.

    Decodifica sem nunca levantar excecao: se o PowerShell responder fora de
    UTF-8 o motivo sai com um caractere trocado, mas sai.
    """
    saida = stdout.decode('utf-8-sig', errors='replace').strip()
    if saida.endswith('OK'):
        return None

    if _SOMENTE_LEITURA in saida:
        return _MOTIVO_SOMENTE_LEITURA

    if not saida:
        saida = stderr.decode('utf-8-sig', errors='replace').strip()
    if saida.startswith('ERRO: '):
        saida = saida[len('ERRO: '):]
    # O Excel responde em varias linhas, com marcadores; vira uma frase so.
    saida = ' '.join(saida.split())
    return saida or 'o PowerShell não devolveu nenhuma resposta'


def recalcular_no_excel(caminho):
    """Faz o proprio Excel abrir, recalcular e salvar a planilha.

    Devolve None quando deu certo, ou o motivo da falha — que precisa chegar
    ao grupo: quem acionou pelo Telegram nao ve o console.
    """
    try:
        win = subprocess.check_output(['wslpath', '-w', caminho]).decode().strip()
    except (OSError, subprocess.CalledProcessError) as e:
        return f'não consegui converter o caminho para o Windows (wslpath): {e}'

    try:
        resposta = subprocess.run(
            [POWERSHELL, '-NoProfile', '-Command', _comando(win)],
            capture_output=True)
    except OSError as e:
        # Sem isto o FileNotFoundError subia como traceback e o grupo recebia
        # 'FALHOU' com rodape vazio, sem nenhuma pista.
        return f'não consegui executar o PowerShell ({POWERSHELL}): {e}'

    return _interpretar(resposta.stdout, resposta.stderr)
