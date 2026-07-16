# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan - Campanha Colonia Inverno 2026
Le a base (Google Sheets) e atualiza o bloco 'colonia_inverno' no index.html.

Por enquanto so ha a base de clientes (ainda sem a base de ligacoes do
discador), entao apenas o card "Cliente na Base" e o grafico de
Distribuicao (aba "Por Unidade", a partir da coluna K) tem dados reais.
Os demais campos (tentativas, interessados, status, evolucao) ficam em
branco ate a base de ligacoes ser disponibilizada.
"""

import csv
import io
import re
import urllib.request
from collections import Counter

# ═══════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════

SHEET_ID   = '1nKJDb-X5YTUEh4IMop30mgQerGk5FeMP'
GID        = '116593147'
INDEX_HTML = r'index.html'

COL_UNIDADE = 10  # K - Unidade de Interesse


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
    print(f'  Baixando base Colonia Inverno 2026 (Google Sheets)...')
    linhas = baixar_csv(SHEET_ID, GID)
    dados = linhas[1:]

    clientes = 0
    uni_counter = Counter()
    for r in dados:
        if not any(c.strip() for c in r):
            continue
        clientes += 1
        uni = r[COL_UNIDADE].strip() if len(r) > COL_UNIDADE else ''
        if uni:
            uni_counter[uni] += 1

    uni_items = uni_counter.most_common()
    print(f'  Clientes na base: {clientes}')
    print(f'  Com unidade de interesse: {sum(uni_counter.values())} | Unidades distintas: {len(uni_items)}')

    return {
        'clientes':   fmt_num(clientes),
        'uniLabels':  [u[0] for u in uni_items],
        'uniData':    [u[1] for u in uni_items],
    }


def js_str(lst):
    return '[' + ','.join(f"'{str(v).replace(chr(39), chr(92)+chr(39))}'" for v in lst) + ']'


def js_num(lst):
    return '[' + ','.join(str(v) for v in lst) + ']'


def gerar_bloco(k):
    return f"""  /* COLONIA_INVERNO_START */
  colonia_inverno: {{
    label: '— Colônia Inverno 2026', desc: 'Campanha Colônia Inverno 2026 — dados filtrados', periodo: '',
    empresas: '{k['clientes']}', empresasLabel: '🧒 Cliente na Base',
    mediaLabel: '🔁 Média Tentativas/Cliente', mediaSub: 'por cliente',
    tentativas: '-', interessados: '-', conversao: '-',
    decisor: '-', decisorSub: 'Apenas Colônia Inverno 2026', media: '-', trend: '',
    statusLabels: ['Sem dados'], statusData: [0], statusColors:null,
    statusTooltips: [],
    evolucaoLabels: ['--'], tentDia: [0], convDia: [0],
    showWpp: false,
    wppTitle: '', wppDesc: '',
    wppKpiLabels: [], wppListLabels: [], wppPieLabels: [],
    wppEnv:'-', wppResp:'-', wppTaxa:'-', wppSem:'-', wppInfo:'-', wppEmail:'-', wppPie:[0,1],
    distToggle: true,
    uniLabels: {js_str(k['uniLabels'])},
    uniData: {js_num(k['uniData'])}
  }},
  /* COLONIA_INVERNO_END */"""


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
    print('  ATUALIZADOR — Colônia Inverno 2026')
    print('=' * 50)
    import os
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        k = calcular()
        bloco = {'COLONIA_INVERNO': gerar_bloco(k)}
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
