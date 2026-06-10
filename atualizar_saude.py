# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan - Modulo Promocao Saude
Le a planilha do Google Sheets (export CSV) e injeta os dados brutos no index.html.
Os filtros (Data, Acao, Especialidade) e KPIs sao recalculados no navegador.

Colunas usadas (0-based):
  A(0)  = Data            -> so conta linha com este campo preenchido
  F(5)  = Empresa         -> eixo X do grafico "Intencoes por Empresa"
  I(8)  = Acao            -> dropdown de filtro
  J(9)  = Especialidade   -> dropdown de filtro
  L(11) = Agendamento     -> "Sim" => Agendada
  O(14) = Cancelamento    -> "Sim" => Cancelada
  S(18) = Pendencia       -> "Sim" => Pendente
"""

import urllib.request
import csv
import io
import re
import os

# ═══════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════

SHEET_ID  = '1bsovvrAr1a5vi9Ny4kujga732cfJp4zpRNTYphpxJos'
GID       = '0'
CSV_URL   = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}'
INDEX_HTML = r'index.html'

COL_DATA  = 0   # A
COL_EMP   = 5   # F
COL_ACAO  = 8   # I
COL_PROF  = 10  # K  Profissional (Nome do Medico)
COL_AGEND = 11  # L
COL_CANC  = 14  # O
COL_PEND  = 18  # S

# ═══════════════════════════════════════════════════════════
# FUNÇÕES
# ═══════════════════════════════════════════════════════════

def baixar_csv(url):
    print(f'  Baixando planilha do Google Sheets...')
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    data = urllib.request.urlopen(req, timeout=40).read().decode('utf-8', errors='replace')
    rows = list(csv.reader(io.StringIO(data)))
    print(f'  {len(rows)} linhas lidas (com cabecalho).')
    return rows


def is_sim(v):
    return str(v).strip().lower() == 'sim'


def parse_data(v):
    """dd/mm/yyyy -> int yyyymmdd. Retorna 0 se nao parsear."""
    s = str(v).strip()
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', s)
    if not m:
        return 0
    d, mo, y = m.groups()
    y = int(y)
    if y < 100:
        y += 2000
    return y * 10000 + int(mo) * 100 + int(d)


def cel(row, idx):
    return row[idx].strip() if len(row) > idx else ''


def base_prof(nome):
    """Remove '(odonto)/(saude)/...' do nome do profissional."""
    return re.sub(r'\s*\(.*?\)', '', nome).strip()


def norm_prof(nome):
    """Chave de agrupamento: sem parenteses, sem acento, minuscula."""
    import unicodedata
    b = base_prof(nome).lower()
    b = unicodedata.normalize('NFKD', b)
    return ''.join(c for c in b if not unicodedata.combining(c))


def processar(rows):
    from collections import Counter
    body = rows[1:]
    emp_list, aco_list, prof_list = [], [], []
    emp_idx, aco_idx, prof_idx = {}, {}, {}

    def get_idx(val, lst, mp):
        if val not in mp:
            mp[val] = len(lst)
            lst.append(val)
        return mp[val]

    # Pre-scan: nome canonico do profissional por chave normalizada (forma mais frequente)
    formas = {}
    for r in body:
        if not cel(r, COL_DATA):
            continue
        p = cel(r, COL_PROF)
        if not p:
            continue
        formas.setdefault(norm_prof(p), Counter())[base_prof(p)] += 1
    canon = {k: c.most_common(1)[0][0] for k, c in formas.items()}

    data_rows = []
    total = 0
    for r in body:
        if not any(c.strip() for c in r):
            continue
        if not cel(r, COL_DATA):   # so conta linha com Data (col A)
            continue
        total += 1
        dt  = parse_data(cel(r, COL_DATA))
        ei  = get_idx(cel(r, COL_EMP) or 'Firjan', emp_list, emp_idx)   # sem empresa = Firjan
        ai  = get_idx(cel(r, COL_ACAO), aco_list, aco_idx)
        praw = cel(r, COL_PROF)
        prof = canon.get(norm_prof(praw), base_prof(praw)) if praw else ''
        si  = get_idx(prof, prof_list, prof_idx)
        ag  = 1 if is_sim(cel(r, COL_AGEND)) else 0
        ca  = 1 if is_sim(cel(r, COL_CANC))  else 0
        pe  = 1 if is_sim(cel(r, COL_PEND))  else 0
        data_rows.append([dt, ei, ai, si, ag, ca, pe])

    # Dropdowns nao mostram valor vazio
    aco_drop  = [a for a in aco_list if a]
    prof_drop = [p for p in prof_list if p]

    print(f'  Interacoes (Data preenchida): {total}')
    print(f'  Agendadas (L=Sim): {sum(x[4] for x in data_rows)}')
    print(f'  Canceladas (O=Sim): {sum(x[5] for x in data_rows)}')
    print(f'  Pendentes (S=Sim): {sum(x[6] for x in data_rows)}')
    print(f'  Empresas: {len(emp_list)} | Acoes: {len(aco_drop)} | Medicos: {len(prof_drop)}')

    return emp_list, aco_list, prof_list, data_rows


def gerar_bloco(emp_list, aco_list, prof_list, data_rows):
    def js_str(lst):
        return '[' + ','.join("'" + str(v).replace('\\', '\\\\').replace("'", "\\'") + "'" for v in lst) + ']'
    def js_rows(rows):
        return '[' + ','.join('[' + ','.join(str(c) for c in r) + ']' for r in rows) + ']'

    return (
        '/* SAUDE_DATA_START */\n'
        f'const SAUDE_EMP={js_str(emp_list)};\n'
        f'const SAUDE_ACO={js_str(aco_list)};\n'
        f'const SAUDE_MED={js_str(prof_list)};\n'
        f'const SAUDE_ROWS={js_rows(data_rows)};\n'
        '/* SAUDE_DATA_END */'
    )


def atualizar_html(index_path, bloco):
    with open(index_path, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    padrao = r'/\* SAUDE_DATA_START \*/.*?/\* SAUDE_DATA_END \*/'
    if not re.search(padrao, conteudo, re.DOTALL):
        raise ValueError('[ERRO] Marcadores SAUDE_DATA nao encontrados no index.html.')
    conteudo = re.sub(padrao, lambda m: bloco, conteudo, flags=re.DOTALL)  # lambda = substituicao literal
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(conteudo)


def carimbar_atualizacao(index_path):
    """Grava a data/hora da atualizacao no cabecalho do dashboard."""
    from datetime import datetime
    agora = datetime.now().strftime('%d/%m/%Y %H:%M')
    with open(index_path, 'r', encoding='utf-8') as f:
        c = f.read()
    c = re.sub(r'<!--LU-->.*?<!--/LU-->', lambda m: f'<!--LU-->{agora}<!--/LU-->', c, flags=re.DOTALL)
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(c)
    return agora


def main():
    print()
    print('=' * 50)
    print('  ATUALIZADOR - Promocao Saude')
    print('=' * 50)
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        print('\n[1/3] Lendo Google Sheets...')
        rows = baixar_csv(CSV_URL)

        print('\n[2/3] Processando dados...')
        emp_list, aco_list, esp_list, data_rows = processar(rows)

        print('\n[3/3] Atualizando index.html...')
        bloco = gerar_bloco(emp_list, aco_list, esp_list, data_rows)
        atualizar_html(INDEX_HTML, bloco)
        ts = carimbar_atualizacao(INDEX_HTML)
        print(f'  Atualizado em: {ts}')

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
    try:
        input('Pressione ENTER para fechar...')
    except EOFError:
        pass
