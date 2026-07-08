# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan - Modulo CSAT
Le 6 Google Sheets (um por canal) e agrega os dados de CSAT.
"""

import io
import csv
import urllib.request
from datetime import datetime, timedelta

# ═══════════════════════════════════════════════════════════
# CONFIGURAÇÃO - Google Sheets por Canal
# ═══════════════════════════════════════════════════════════

CSAT_SOURCES = {
    'VOZ': {
        'sheet_id': '1oEurhulMAqCK7Ar0iKsrnhiLTkxd_4hHP1YIivxhCuc',
        'gid': '1192760984',
        'data_col': 0,        # A: Carimbo de data/hora
        'csat_col': 4,        # E: pergunta de CSAT
        'ces_col': 8,         # I: CES
        'canal_label': 'Telefone',   # nome canonico (bate com o canal da BSales2)
    },
    'WHATSAPP': {
        'sheet_id': '11zjE-QUB0CKCt8zK9-ZK3vBBFXQ-JCmR-b4Dh2-B9KA',
        'gid': '1820980124',
        'data_col': 0,        # A: Carimbo de data/hora
        'csat_col': 3,        # D: pergunta de CSAT
        'ces_col': 7,         # H: CES
        'canal_label': 'Whatsapp',
    },
    'WHATSAPP2': {
        'sheet_id': '1RmPiKTXqgtLvcX1-Hy-sCR01G-vFnRoGRd4oCztHt5E',
        'gid': '1953725700',
        'data_col': 0,        # A: Carimbo de data/hora
        'csat_col': 4,        # E: pergunta de CSAT
        'ces_col': 8,         # I: CES
        'canal_label': 'Whatsapp',
    },
    'E-MAIL': {
        'sheet_id': '1JIGluf8a2lE4f2jjHEbzblGnx29vZBVsFn1PF2wk4fk',
        'gid': None,
        'data_col': 0,        # A: Carimbo de data/hora
        'csat_col': 3,        # D: pergunta de CSAT
        'ces_col': 7,         # H: CES
        'canal_label': 'E-mail',
    },
    'REDES SOCIAIS': {
        'sheet_id': '19mIN_TTzdzdEb-7bjopVsMjFInCa7mhq8K1COUkywR0',
        'gid': '1235116693',
        'data_col': 0,        # A: Carimbo de data/hora
        'csat_col': 3,        # D: pergunta de CSAT
        'ces_col': 7,         # H: CES
        # sem canal_label: nao vira opcao no filtro (o detalhe do canal/rede
        # social - Instagram/Facebook/Messenger - vem da coluna K da aba
        # "Redes Sociais" do SharePoint, nao dessa planilha de CSAT)
        'canal_label': None,
    },
    'CHAT': {
        'sheet_id': '1eoxDEmbrzxk_zK8Y2uVHkHJcwnMlxBUxEYW2Yzi0zxI',
        'gid': None,
        'data_col': 0,        # A: Carimbo de data/hora
        'csat_col': 3,        # D: pergunta de CSAT
        'ces_col': 7,         # H: CES
        'canal_label': 'Chat',
    },
}

CSAT_BONS = {'muito satisfeito', 'satisfeito', 'sim'}


# ═══════════════════════════════════════════════════════════
# FUNÇÕES
# ═══════════════════════════════════════════════════════════

def to_dt(val):
    """Converte valor para datetime."""
    if isinstance(val, datetime):
        return val
    if isinstance(val, (int, float)) and val > 0:
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(val))
        except Exception:
            return None
    # Parse text format: dd/mm/yyyy ou dd/mm/yyyy HH:MM:SS
    if isinstance(val, str):
        val = val.strip()
        # Tira a parte de hora se houver
        data_part = val.split()[0] if ' ' in val else val
        import re as re_mod
        m = re_mod.match(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', data_part)
        if m:
            dd, mo, yy = m.groups()
            yy = int(yy)
            if yy < 100:
                yy += 2000
            try:
                return datetime(yy, int(mo), int(dd))
            except Exception:
                return None
    return None


def txt(v):
    """Converte valor para texto (minúsculo, sem espaço extra)."""
    return str(v).strip().lower() if v is not None else ''


def ler_google_sheets_csv(sheet_id, gid=None):
    """Baixa Google Sheet como CSV e retorna lista de linhas."""
    url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv'
    if gid:
        url += f'&gid={gid}'
    try:
        resp = urllib.request.urlopen(url, timeout=60)
        conteudo = resp.read().decode('utf-8-sig')
        leitor = csv.reader(io.StringIO(conteudo))
        rows = list(leitor)
        return rows
    except Exception as e:
        print(f'  [ERRO] Falha ao baixar Google Sheets {sheet_id}: {e}')
        return []


def cel(row, idx):
    """Acessa célula de forma segura."""
    if len(row) <= idx:
        return None
    v = row[idx]
    return v if v else None


def processar_canal(canal, config):
    """Processa uma Google Sheet de um canal CSAT."""
    csat_rows = []
    try:
        rows = ler_google_sheets_csv(config['sheet_id'], config.get('gid'))
        if not rows:
            print(f'  [AVISO] {canal}: nenhum dado lido')
            return csat_rows

        for i, r in enumerate(rows):
            if i == 0:
                continue  # cabeçalho
            if not any(c for c in r):
                continue  # linha vazia

            # Parse data
            data_raw = cel(r, config['data_col'])
            if not data_raw:
                continue
            c = to_dt(data_raw)
            dt = (c.year * 10000 + c.month * 100 + c.day) if c else 0
            if not dt:
                continue

            # Parse CSAT
            csat_resp = txt(cel(r, config['csat_col']))
            bons = tot = 0
            if csat_resp:
                tot = 1
                if 'satisfeito' in csat_resp or 'sim' in csat_resp:
                    bons = 1

            # Parse CES
            ces_resp = txt(cel(r, config['ces_col']))
            bons_ces = 1 if ces_resp in CSAT_BONS else 0

            # canal_label = nome canonico usado no filtro (bate com BSales2);
            # se None (ex.: Redes Sociais), mantem a chave crua so para o
            # total geral, mas nao vira opcao selecionavel no filtro.
            canal_row = config.get('canal_label') or canal

            # [dt, canal, bons_csat, total_csat, bons_ces]
            csat_rows.append([dt, canal_row, bons, tot, bons_ces])

        print(f'  {canal}: {len(csat_rows)} linhas lidas')
    except Exception as e:
        print(f'  [ERRO] Falha ao processar {canal}: {e}')

    return csat_rows


def main():
    """Processa todos os 6 canais e retorna dados consolidados."""
    print('[CSAT] Atualizando dados de satisfação...')
    todos_csat_rows = []

    for canal, config in CSAT_SOURCES.items():
        rows = processar_canal(canal, config)
        todos_csat_rows.extend(rows)

    print(f'  Total: {len(todos_csat_rows)} respostas agregadas')
    return todos_csat_rows


if __name__ == '__main__':
    csat_data = main()
    print(f'Retornou {len(csat_data)} linhas')
