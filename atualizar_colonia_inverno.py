# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan - Campanha Colonia Inverno 2026
Le a base de clientes (Google Sheets) + a base de ligacoes (Discagem) e
atualiza o bloco 'colonia_inverno' no index.html.

Base de clientes -> card "Cliente na Base" e grafico de Distribuicao,
aba "Por Unidade" (coluna K - Unidade de Interesse).
Base de ligacoes (Discagem) -> Tentativas, Interessados, Decisor,
Conversao, Media, Distribuicao "Por Status" e Evolucao Diaria.
"""

import csv
import glob
import io
import os
import re
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime, timedelta

# ═══════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════

SHEET_ID   = '1nKJDb-X5YTUEh4IMop30mgQerGk5FeMP'
GID        = '116593147'
INDEX_HTML = r'index.html'
DEPARA_PATH = r'Arquivos\bases_apoio\tab_de-para.xlsx'
PASTA_DISCAGEM = r'Arquivos\nao_atualizaveis\Ativo_Colônia de Férias'

COL_UNIDADE = 10  # K - Unidade de Interesse (base de clientes)

STATUS_MAP = {}
LABELS_NAO_DECISOR = {'Telefonia', 'Tentativa', 'Engano', 'Alo'}
LABEL_INTERESSADO  = 'Interessado'

MES_PT = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',
          7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}


# ═══════════════════════════════════════════════════════════
# FUNÇÕES AUXILIARES
# ═══════════════════════════════════════════════════════════

def fmt_num(n):
    return f"{n:,}".replace(",", ".")


def fmt_pct(n):
    return f"{n:.2f}%".replace(".", ",")


def fmt_dec(n):
    return f"{n:.2f}".replace(".", ",")


def norm_tel(t):
    return re.sub(r'\D', '', str(t)) if t else ''


def to_datetime(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, (int, float)) and val > 0:
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(val))
        except Exception:
            return None
    return None


def encontrar_unico_xlsx(pasta):
    arquivos = [a for a in glob.glob(os.path.join(pasta, '*.xlsx')) if not os.path.basename(a).startswith('~$')]
    if not arquivos:
        raise FileNotFoundError(f'[ERRO] Nenhum .xlsx encontrado em {pasta}')
    return sorted(arquivos)[-1]


def ler_depara(caminho):
    mapa = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        ws = wb['status'] if 'status' in wb.sheetnames else wb.active
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            if row[0] and row[1]:
                mapa[str(row[0]).strip()] = str(row[1]).strip()
        wb.close()
        print(f'  De-para: {len(mapa)} entradas')
    except FileNotFoundError:
        print(f'  [AVISO] De-para nao encontrado: {caminho}')
    return mapa


def normalizar_status(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    for k, v in STATUS_MAP.items():
        if s.upper() == k.upper():
            return v
    return s


# ═══════════════════════════════════════════════════════════
# BASE DE CLIENTES (Google Sheets)
# ═══════════════════════════════════════════════════════════

def baixar_csv(sheet_id, gid=None):
    url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv'
    if gid:
        url += f'&gid={gid}'
    resp = urllib.request.urlopen(url, timeout=60)
    conteudo = resp.read().decode('utf-8-sig')
    return list(csv.reader(io.StringIO(conteudo)))


def calcular_base():
    print('  Baixando base de clientes Colônia Inverno 2026 (Google Sheets)...')
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
        'clientes':  clientes,
        'uniLabels': [u[0] for u in uni_items],
        'uniData':   [u[1] for u in uni_items],
    }


# ═══════════════════════════════════════════════════════════
# BASE DE LIGAÇÕES (Discagem)
# ═══════════════════════════════════════════════════════════

def calcular_discagem(caminho):
    """Le o Discagem_Ativo_Colônia de Férias.xlsx.
    Cada linha com DATA (col A) = 1 tentativa de ligacao.
    Status: col L (STATUS_NEGOCIO); se vazia, usa col K (STATUS).
    """
    import openpyxl
    print(f'  Lendo Discagem: {os.path.basename(caminho)}')
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    aba = next((s for s in wb.sheetnames if s.startswith('chamadas_')), wb.sheetnames[0])
    ws = wb[aba]

    data_idx = 0    # A DATA
    orig_idx = 7    # H ORIGEM
    dest_idx = 8    # I DESTINO
    st_idx   = 10   # K STATUS
    sn_idx   = 11   # L STATUS_NEGOCIO

    total_tent = 0
    status_counter = Counter()
    raw_por_label  = defaultdict(set)
    evolucao = {}
    tel_decisor   = set()
    tel_interesse = set()

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        dt = to_datetime(row[data_idx])
        if not dt:
            continue
        total_tent += 1

        tel = norm_tel(row[orig_idx]) or norm_tel(row[dest_idx])

        sn_raw = row[sn_idx] if len(row) > sn_idx else None
        st_raw = row[st_idx] if len(row) > st_idx else None
        raw    = str(sn_raw).strip() if sn_raw else (str(st_raw).strip() if st_raw else '')
        label  = normalizar_status(raw) if raw else None

        if label:
            status_counter[label] += 1
            if raw:
                raw_por_label[label].add(raw)

        if label and label not in LABELS_NAO_DECISOR:
            tel_decisor.add(tel)
        if label == LABEL_INTERESSADO:
            tel_interesse.add(tel)

        dk = f"{dt.day:02d}/{MES_PT[dt.month]}"
        if dk not in evolucao:
            evolucao[dk] = {'date': dt.date(), 'tent': 0, 'conv': 0}
        evolucao[dk]['tent'] += 1
        if label == LABEL_INTERESSADO:
            evolucao[dk]['conv'] += 1

    wb.close()

    dias_ord = sorted(evolucao.items(), key=lambda x: x[1]['date'])
    st_items = status_counter.most_common()
    st_tooltips = [', '.join(sorted(raw_por_label.get(s[0], set()))) for s in st_items]

    print(f'  Total Tentativas: {total_tent}')
    print(f'  Decisor: {len(tel_decisor)} | Interessados: {len(tel_interesse)}')
    print(f'  Status: {dict(st_items[:5])} ...')

    return {
        'tentativas':    total_tent,
        'decisor':       len(tel_decisor),
        'interessados':  len(tel_interesse),
        'statusLabels':  [s[0] for s in st_items],
        'statusData':    [s[1] for s in st_items],
        'statusTooltips': st_tooltips,
        'evoLabels':     [d[0] for d in dias_ord],
        'tentDia':       [d[1]['tent'] for d in dias_ord],
        'convDia':       [d[1]['conv'] for d in dias_ord],
    }


# ═══════════════════════════════════════════════════════════
# GERAÇÃO DO BLOCO JS
# ═══════════════════════════════════════════════════════════

def js_str(lst):
    return '[' + ','.join(f"'{str(v).replace(chr(39), chr(92)+chr(39))}'" for v in lst) + ']'


def js_num(lst):
    return '[' + ','.join(str(v) for v in lst) + ']'


def gerar_bloco(base, disc):
    clientes = base['clientes']
    tent     = disc['tentativas']
    decisor  = disc['decisor']
    interess = disc['interessados']
    taxa     = (interess / decisor * 100) if decisor > 0 else 0
    media    = (tent / clientes) if clientes > 0 else 0

    periodo = f"{disc['evoLabels'][0]} — {disc['evoLabels'][-1]}" if disc['evoLabels'] else ''
    status_labels = disc['statusLabels'] or ['Sem dados']
    status_data   = disc['statusData'] or [0]
    evo_labels    = disc['evoLabels'] or ['--']
    tent_dia      = disc['tentDia'] or [0]
    conv_dia      = disc['convDia'] or [0]

    return f"""  /* COLONIA_INVERNO_START */
  colonia_inverno: {{
    label: '— Colônia Inverno 2026', desc: 'Campanha Colônia Inverno 2026 — dados filtrados', periodo: '{periodo}',
    empresas: '{fmt_num(clientes)}', empresasLabel: '🧒 Cliente na Base',
    mediaLabel: '🔁 Média Tentativas/Cliente', mediaSub: 'por cliente',
    tentativas: '{fmt_num(tent)}', interessados: '{fmt_num(interess)}', conversao: '{fmt_pct(taxa)}',
    decisor: '{fmt_num(decisor)}', decisorSub: 'Apenas Colônia Inverno 2026', media: '{fmt_dec(media)}', trend: '',
    statusLabels: {js_str(status_labels)}, statusData: {js_num(status_data)}, statusColors:null,
    statusTooltips: {js_str(disc['statusTooltips'])},
    evolucaoLabels: {js_str(evo_labels)},
    tentDia: {js_num(tent_dia)}, convDia: {js_num(conv_dia)},
    showWpp: false,
    wppTitle: '', wppDesc: '',
    wppKpiLabels: [], wppListLabels: [], wppPieLabels: [],
    wppEnv:'-', wppResp:'-', wppTaxa:'-', wppSem:'-', wppInfo:'-', wppEmail:'-', wppPie:[0,1],
    distToggle: true,
    uniLabels: {js_str(base['uniLabels'])},
    uniData: {js_num(base['uniData'])}
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


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════

def main():
    print()
    print('=' * 50)
    print('  ATUALIZADOR — Colônia Inverno 2026')
    print('=' * 50)
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        STATUS_MAP.update(ler_depara(DEPARA_PATH))
        base = calcular_base()
        disc = calcular_discagem(encontrar_unico_xlsx(PASTA_DISCAGEM))
        bloco = {'COLONIA_INVERNO': gerar_bloco(base, disc)}
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
