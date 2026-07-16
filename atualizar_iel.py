# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan - Campanha Prospeccao IEL
Le a base (Google Sheets) e atualiza o bloco 'iel' no index.html.

Por enquanto so ha a base de empresas (ainda sem retorno de ligacoes),
entao apenas o card "Empresas na Base" tem dado real. Os demais campos
(tentativas, interessados, status, evolucao) ficam em branco, no mesmo
layout usado pela campanha Retomada da Trilha, ate a base de ligacoes
ser disponibilizada.
"""

import csv
import io
import os
import re
import urllib.request

# ═══════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════

SHEET_ID   = '10epaIyncKbZO9auh8QWnVm8f9VTCzLCY'
GID        = '1369157082'
INDEX_HTML = r'index.html'


# ═══════════════════════════════════════════════════════════
# FUNÇÕES
# ═══════════════════════════════════════════════════════════

def fmt_num(n):
    return f"{n:,}".replace(",", ".")


def baixar_csv(sheet_id, gid=None):
    url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv'
    if gid:
        url += f'&gid={gid}'
    resp = urllib.request.urlopen(url, timeout=60)
    conteudo = resp.read().decode('utf-8-sig')
    return list(csv.reader(io.StringIO(conteudo)))


def calcular():
    print(f'  Baixando base Prospecção IEL (Google Sheets)...')
    linhas = baixar_csv(SHEET_ID, GID)
    dados = linhas[1:]

    empresas = sum(1 for r in dados if any(c.strip() for c in r))

    print(f'  Empresas na base: {empresas}')
    return {'empresas': fmt_num(empresas)}


def gerar_bloco(k):
    return f"""  /* IEL_START */
  iel: {{
    label: '— Prospecção IEL', desc: 'Campanha Prospecção IEL — dados filtrados', periodo: '',
    empresas: '{k['empresas']}', empresasLabel: '🏢 Empresas na Base',
    mediaLabel: '🔁 Média Tentativas/Empresa', mediaSub: 'por empresa',
    tentativas: '-', interessados: '-', conversao: '-',
    decisor: '-', decisorSub: 'Apenas Prospecção IEL', media: '-', trend: '',
    statusLabels: ['Sem dados'], statusData: [0], statusColors:null,
    statusTooltips: [],
    evolucaoLabels: ['--'], tentDia: [0], convDia: [0],
    showWpp: false,
    wppTitle: '', wppDesc: '',
    wppKpiLabels: [], wppListLabels: [], wppPieLabels: [],
    wppEnv:'-', wppResp:'-', wppTaxa:'-', wppSem:'-', wppInfo:'-', wppEmail:'-', wppPie:[0,1]
  }},
  /* IEL_END */"""


def atualizar_html(index_path, blocos):
    with open(index_path, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    for marcador, novo_bloco in blocos.items():
        padrao = f'/\\* {marcador}_START \\*/.*?/\\* {marcador}_END \\*/'
        if not re.search(padrao, conteudo, re.DOTALL):
            raise ValueError(f'[ERRO] Marcadores {marcador} nao encontrados no index.html.')
        conteudo = re.sub(padrao, novo_bloco, conteudo, flags=re.DOTALL)
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(conteudo)


def main():
    print()
    print('=' * 50)
    print('  ATUALIZADOR — Prospecção IEL')
    print('=' * 50)
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        k = calcular()
        bloco = {'IEL': gerar_bloco(k)}
        atualizar_html(INDEX_HTML, bloco)
        print()
        print('=' * 50)
        print('  CONCLUIDO! index.html atualizado.')
        print('  Rode publicar.bat para enviar ao GitHub.')
        print('=' * 50)
        print()
    except Exception as e:
        print(f'\n[ERRO] {e}')
        import traceback; traceback.print_exc()


if __name__ == '__main__':
    main()
