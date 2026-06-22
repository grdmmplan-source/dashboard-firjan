# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan - Modulo SAC
Le a planilha SAC do SharePoint (link publico/anonimo) e injeta os dados brutos
no index.html. Os filtros (Data, Regional, Tipo de Registro, Canal) e os KPIs
sao recalculados no navegador.

Aba: "Sac"
Colunas usadas (0-based):
  A(0)  = Protocolo         -> Total de SACs (linhas com este campo preenchido)
  C(2)  = Data do Atendimento -> filtro Data + base do Tempo Medio
  E(4)  = Canal             -> filtro
  I(8)  = Regional          -> filtro
  J(9)  = Tipo de Registro  -> filtro
  M(12) = Data de Finalizacao -> Tempo Medio = media(M - C)
  N(13) = Nivel de Satisfacao -> Satisfeito / Insatisfeito
"""

import urllib.request
import http.cookiejar
import io
import re
import os
from datetime import datetime, timedelta

# ═══════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════

# Link de compartilhamento do SharePoint (somente leitura/consulta).
# O ?download=1 faz o SharePoint devolver o arquivo .xlsx direto.
SAC_URL = ('https://ddmadvbr-my.sharepoint.com/:x:/g/personal/'
           'fernanda_castro_grupoddm_com_br/'
           'IQCobMqlyAWfT7aBxT5EGURCAaAxahCZQjDD0eCyvVChMv0?download=1')

# Link "limpo" (sem ?download=1) para registrar no arquivo de erros
SAC_LINK = ('https://ddmadvbr-my.sharepoint.com/:x:/g/personal/'
            'fernanda_castro_grupoddm_com_br/'
            'IQCobMqlyAWfT7aBxT5EGURCAaAxahCZQjDD0eCyvVChMv0')

ABA_SAC    = 'Sac'
INDEX_HTML = r'index.html'
ERROS_TXT  = r'erros_sac.txt'

# Tempo de resposta acima deste limite (ou negativo) e tratado como outlier
# e EXCLUIDO do calculo do Tempo Medio (registrado em erros_sac.txt).
MAX_DELTA_DIAS = 30

# --- Colunas fonte 1: SharePoint ---
COL_PROT    = 0   # A  Protocolo
COL_DATA    = 2   # C  Data do Atendimento
COL_CANAL   = 4   # E  Canal
COL_PROD    = 5   # F  Produto/Servico
COL_UNI     = 7   # H  Unidade
COL_ENT     = 8   # I  Entidade
COL_REG     = 9   # J  Regional
COL_TIPO    = 10  # K  Tipo de Registro
COL_FIM     = 13  # N  Data de Finalizacao
COL_SAT     = 14  # O  Nivel de Satisfacao
COL_ASSUNTO = 19  # T  Assunto

# --- Colunas fonte 2: Google Sheets ---
SAC2_URL = ('https://docs.google.com/spreadsheets/d/'
            '1z-4hnVB7JRoqTZZ0yVrZ0KONKR3RlualzrrfM3vw9xs/'
            'export?format=xlsx&gid=1326039238')

C2_PROT    = 0   # A  Protocolo
C2_DATA    = 2   # C  Data do Atendimento
C2_CANAL   = 4   # E  Canal
C2_PROD    = 5   # F  Produto/Servico
C2_TIPO    = 8   # I  Tipo de Registro
C2_FIM     = 11  # L  Data de Finalizacao
C2_SAT     = 12  # M  Nivel de Satisfacao
C2_ASSUNTO = 17  # R  Assunto
C2_UNI     = 21  # V  Unidade
C2_ENT     = 22  # W  Entidade
C2_REG     = 23  # X  Regional

# ═══════════════════════════════════════════════════════════
# FUNÇÕES
# ═══════════════════════════════════════════════════════════

def baixar_xlsx(url):
    """Baixa o .xlsx do SharePoint usando cookie jar (link anonimo)."""
    print('  Baixando planilha SAC do SharePoint...')
    cj = http.cookiejar.CookieJar()
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    op.addheaders = [('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)')]
    resp = op.open(url, timeout=60)
    data = resp.read()
    ctype = resp.headers.get('Content-Type', '')
    if 'spreadsheet' not in ctype and 'octet-stream' not in ctype:
        raise RuntimeError(f'Resposta nao e xlsx (Content-Type: {ctype}). '
                           f'O link pode ter mudado de permissao.')
    print(f'  OK ({len(data):,} bytes)')
    return data


def baixar_gsheets(url):
    """Baixa o .xlsx exportado do Google Sheets (link publico)."""
    print('  Baixando planilha SAC complementar do Google Sheets...')
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    resp = urllib.request.urlopen(req, timeout=60)
    data = resp.read()
    print(f'  OK ({len(data):,} bytes)')
    return data


_HOJE = datetime.now()

def to_dt(val):
    d = None
    if isinstance(val, datetime):
        d = val
    elif isinstance(val, (int, float)) and val > 0:
        try:
            d = datetime(1899, 12, 30) + timedelta(days=float(val))
        except Exception:
            d = None
    elif isinstance(val, str):
        m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', val.strip())
        if m:
            dd, mo, yy = m.groups()
            yy = int(yy); yy = yy + 2000 if yy < 100 else yy
            try:
                d = datetime(yy, int(mo), int(dd))
            except ValueError:
                d = None
    if d is None:
        return None
    # Correcao BR x EUA: se a data caiu no FUTURO, troca dia<->mes
    if d > _HOJE:
        try:
            sw = d.replace(month=d.day, day=d.month)
            if sw <= _HOJE:
                d = sw
        except ValueError:
            pass
    return d


def cel(row, idx):
    if len(row) <= idx:
        return None
    return row[idx]


def sat_code(v):
    """1=Satisfeito, 2=Insatisfeito, 3=Nao Avaliado, 0=outros."""
    s = str(v).strip().lower() if v is not None else ''
    if s == 'satisfeito':
        return 1
    if s == 'insatisfeito':
        return 2
    if s in ('não avaliado', 'nao avaliado'):
        return 3
    return 0


def processar(xlsx_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    if ABA_SAC not in wb.sheetnames:
        raise ValueError(f'Aba "{ABA_SAC}" nao encontrada. Abas: {wb.sheetnames}')
    ws = wb[ABA_SAC]
    rows = list(ws.iter_rows(values_only=True, min_row=2))
    wb.close()

    canal_list, reg_list, tipo_list = [], [], []
    canal_idx, reg_idx, tipo_idx = {}, {}, {}
    ent_list, assunto_list, prod_list, uni_list = [], [], [], []
    ent_idx, assunto_idx, prod_idx, uni_idx = {}, {}, {}, {}

    def get_idx(val, lst, mp):
        v = str(val).strip() if val is not None else ''
        if v not in mp:
            mp[v] = len(lst)
            lst.append(v)
        return mp[v]

    data_rows = []
    erros = []
    outros_set = set()
    total = sat = insat = 0
    soma_delta = 0
    n_delta = 0
    limite_seg = MAX_DELTA_DIAS * 86400

    for i, r in enumerate(rows):
        linha_excel = i + 2  # +2: header (linha 1) + base 0
        prot = cel(r, COL_PROT)
        if prot in (None, ''):
            continue  # Total de SACs = linhas com Protocolo (col A)
        total += 1

        c = to_dt(cel(r, COL_DATA))
        m = to_dt(cel(r, COL_FIM))
        dt = (c.year * 10000 + c.month * 100 + c.day) if c else 0

        ci = get_idx(cel(r, COL_CANAL), canal_list, canal_idx)
        ri = get_idx(cel(r, COL_REG),   reg_list,   reg_idx)
        ti = get_idx(cel(r, COL_TIPO),  tipo_list,  tipo_idx)
        ei = get_idx(cel(r, COL_ENT), ent_list, ent_idx)   # Entidade (col I, direta)
        ui = get_idx(cel(r, COL_UNI), uni_list, uni_idx)   # Unidade (col H) -> ranking
        aci = get_idx(cel(r, COL_ASSUNTO), assunto_list, assunto_idx)
        pi = get_idx(cel(r, COL_PROD),    prod_list,    prod_idx)
        sc = sat_code(cel(r, COL_SAT))
        if sc == 1:
            sat += 1
        elif sc == 2:
            insat += 1
        elif sc == 0:  # "Outros" — guarda os rotulos para a legenda
            raw = cel(r, COL_SAT)
            outros_set.add('(em branco)' if raw in (None, '') else str(raw).strip())

        dl = None
        if c and m:
            d = int((m - c).total_seconds())
            motivo = None
            if d < 0:
                motivo = 'Data de Finalizacao (M) ANTERIOR a Data do Atendimento (C)'
            elif d > limite_seg:
                motivo = f'Tempo de resposta acima do limite de {MAX_DELTA_DIAS} dias'
            if motivo:
                erros.append({
                    'linha': linha_excel,
                    'protocolo': str(prot).strip(),
                    'atendimento': c.strftime('%d/%m/%Y %H:%M'),
                    'finalizacao': m.strftime('%d/%m/%Y %H:%M'),
                    'delta_dias': round(d / 86400, 2),
                    'motivo': motivo,
                })
            else:
                soma_delta += d
                n_delta += 1
                dl = d

        data_rows.append([dt, ci, ri, ti, sc, dl, ei, aci, pi, ui])

    tmr = (soma_delta / n_delta) if n_delta else 0
    print(f'  Total de SACs        : {total}')
    print(f'  Satisfeitos          : {sat}')
    print(f'  Insatisfeitos        : {insat}')
    print(f'  Tempo medio (validos): {tmr/3600:.1f}h = {tmr/86400:.2f} dias '
          f'({n_delta} validos | {len(erros)} outliers ignorados)')
    print(f'  Canais: {len(canal_list)} | Regionais: {len(reg_list)} | Tipos: {len(tipo_list)}')
    outros_labels = sorted(outros_set, key=lambda s: (s != '(em branco)', s.lower()))
    print(f'  "Outros" abrange: {outros_labels}')
    print(f'  Entidades: {len([e for e in ent_list if e])} | Assuntos: {len([a for a in assunto_list if a])} | Produtos: {len([p for p in prod_list if p])}')

    extras = {'ent': ent_list, 'assunto': assunto_list, 'prod': prod_list, 'uni': uni_list}
    return canal_list, reg_list, tipo_list, data_rows, erros, outros_labels, extras


def processar2(xlsx_bytes, canal_list, canal_idx, reg_list, reg_idx, tipo_list, tipo_idx,
               ent_list, ent_idx, assunto_list, assunto_idx, prod_list, prod_idx,
               uni_list, uni_idx):
    """Processa a fonte 2 (Google Sheets) reaproveitando as listas da fonte 1."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True, min_row=2))
    wb.close()

    def get_idx(val, lst, mp):
        v = str(val).strip() if val is not None else ''
        if v not in mp:
            mp[v] = len(lst)
            lst.append(v)
        return mp[v]

    data_rows = []
    erros = []
    outros_set = set()
    total = sat = insat = 0
    soma_delta = 0
    n_delta = 0
    limite_seg = MAX_DELTA_DIAS * 86400

    for i, r in enumerate(rows):
        linha_excel = i + 2
        prot = cel(r, C2_PROT)
        if prot in (None, ''):
            continue
        total += 1

        c = to_dt(cel(r, C2_DATA))
        m = to_dt(cel(r, C2_FIM))
        dt = (c.year * 10000 + c.month * 100 + c.day) if c else 0

        ci  = get_idx(cel(r, C2_CANAL),   canal_list,   canal_idx)
        ri  = get_idx(cel(r, C2_REG),     reg_list,     reg_idx)
        ti  = get_idx(cel(r, C2_TIPO),    tipo_list,    tipo_idx)
        ei  = get_idx(cel(r, C2_ENT),     ent_list,     ent_idx)
        ui  = get_idx(cel(r, C2_UNI),     uni_list,     uni_idx)
        aci = get_idx(cel(r, C2_ASSUNTO), assunto_list, assunto_idx)
        pi  = get_idx(cel(r, C2_PROD),    prod_list,    prod_idx)
        sc  = sat_code(cel(r, C2_SAT))
        if sc == 1:
            sat += 1
        elif sc == 2:
            insat += 1
        elif sc == 0:
            raw = cel(r, C2_SAT)
            outros_set.add('(em branco)' if raw in (None, '') else str(raw).strip())

        dl = None
        if c and m:
            d = int((m - c).total_seconds())
            motivo = None
            if d < 0:
                motivo = 'Data de Finalizacao (L) ANTERIOR a Data do Atendimento (C)'
            elif d > limite_seg:
                motivo = f'Tempo de resposta acima do limite de {MAX_DELTA_DIAS} dias'
            if motivo:
                erros.append({
                    'linha': linha_excel,
                    'protocolo': str(prot).strip(),
                    'atendimento': c.strftime('%d/%m/%Y %H:%M'),
                    'finalizacao': m.strftime('%d/%m/%Y %H:%M'),
                    'delta_dias': round(d / 86400, 2),
                    'motivo': motivo,
                })
            else:
                soma_delta += d
                n_delta += 1
                dl = d

        data_rows.append([dt, ci, ri, ti, sc, dl, ei, aci, pi, ui])

    print(f'  [Fonte 2] Total: {total} | Satisfeitos: {sat} | Insatisfeitos: {insat} | Outliers: {len(erros)}')
    return data_rows, erros


def gerar_bloco(canal_list, reg_list, tipo_list, data_rows, outros_labels=None, extras=None):
    def js_str(lst):
        return '[' + ','.join("'" + str(v).replace('\\', '\\\\').replace("'", "\\'") + "'" for v in lst) + ']'
    def js_rows(rows):
        partes = []
        for r in rows:
            cells = []
            for c in r:
                cells.append('null' if c is None else str(c))
            partes.append('[' + ','.join(cells) + ']')
        return '[' + ','.join(partes) + ']'

    extras = extras or {}
    return (
        '/* SAC_DATA_START */\n'
        f'const SAC_CANAL={js_str(canal_list)};\n'
        f'const SAC_REG={js_str(reg_list)};\n'
        f'const SAC_TIPO={js_str(tipo_list)};\n'
        f'const SAC_ENT={js_str(extras.get("ent", []))};\n'
        f'const SAC_ASSUNTO={js_str(extras.get("assunto", []))};\n'
        f'const SAC_PROD={js_str(extras.get("prod", []))};\n'
        f'const SAC_UNI={js_str(extras.get("uni", []))};\n'
        f'const SAC_OUTROS={js_str(outros_labels or [])};\n'
        f'const SAC_ROWS={js_rows(data_rows)};\n'
        '/* SAC_DATA_END */'
    )


def escrever_erros(caminho, erros, total_sacs):
    """Gera o arquivo de erros com os outliers excluidos do Tempo Medio."""
    from datetime import datetime
    agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    linhas = []
    linhas.append('=' * 78)
    linhas.append('  RELATORIO DE OUTLIERS - TEMPO MEDIO DE RESPOSTA (SAC)')
    linhas.append('=' * 78)
    linhas.append(f'Gerado em      : {agora}')
    linhas.append(f'Planilha (link): {SAC_LINK}')
    linhas.append(f'Aba            : {ABA_SAC}')
    linhas.append(f'Total de SACs  : {total_sacs}')
    linhas.append(f'Outliers       : {len(erros)} (EXCLUIDOS do calculo do Tempo Medio)')
    linhas.append('')
    linhas.append('Regra de validacao:')
    linhas.append('  - Tempo = Data de Finalizacao (col M) - Data do Atendimento (col C)')
    linhas.append('  - INVALIDO se for negativo, ou maior que '
                  f'{MAX_DELTA_DIAS} dias.')
    linhas.append('=' * 78)
    linhas.append('')
    if not erros:
        linhas.append('Nenhum outlier encontrado. Todos os tempos sao validos.')
    else:
        cab = f'{"Linha":>6} | {"Protocolo":<12} | {"Atendimento (C)":<16} | {"Finalizacao (M)":<16} | {"Dias":>8} | Motivo'
        linhas.append(cab)
        linhas.append('-' * len(cab))
        for e in sorted(erros, key=lambda x: x['linha']):
            linhas.append(
                f'{e["linha"]:>6} | {e["protocolo"]:<12} | {e["atendimento"]:<16} | '
                f'{e["finalizacao"]:<16} | {e["delta_dias"]:>8} | {e["motivo"]}'
            )
    with open(caminho, 'w', encoding='utf-8') as f:
        f.write('\n'.join(linhas) + '\n')
    print(f'  Arquivo de erros: {caminho} ({len(erros)} outliers)')


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


def atualizar_html(index_path, bloco):
    with open(index_path, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    padrao = r'/\* SAC_DATA_START \*/.*?/\* SAC_DATA_END \*/'
    if not re.search(padrao, conteudo, re.DOTALL):
        raise ValueError('[ERRO] Marcadores SAC_DATA nao encontrados no index.html.')
    conteudo = re.sub(padrao, lambda m: bloco, conteudo, flags=re.DOTALL)  # lambda = literal
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(conteudo)


def main():
    print()
    print('=' * 50)
    print('  ATUALIZADOR - SAC')
    print('=' * 50)
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        print('\n[1/4] Baixando SharePoint (fonte 1)...')
        xlsx1 = baixar_xlsx(SAC_URL)

        print('\n[2/4] Processando fonte 1...')
        canal, reg, tipo, drows, erros, outros, extras = processar(xlsx1)

        print('\n[3/4] Baixando Google Sheets (fonte 2)...')
        try:
            xlsx2 = baixar_gsheets(SAC2_URL)
            mk = lambda lst: {v: i for i, v in enumerate(lst)}
            drows2, erros2 = processar2(
                xlsx2,
                canal,             mk(canal),
                reg,               mk(reg),
                tipo,              mk(tipo),
                extras['ent'],     mk(extras['ent']),
                extras['assunto'], mk(extras['assunto']),
                extras['prod'],    mk(extras['prod']),
                extras['uni'],     mk(extras['uni']),
            )
            drows += drows2
            erros += erros2
            print(f'  Total combinado: {len(drows)} registros')
        except Exception as e2:
            print(f'  [AVISO] Fonte 2 falhou ({e2}). Continuando apenas com fonte 1.')

        escrever_erros(ERROS_TXT, erros, len(drows))

        print('\n[4/4] Atualizando index.html...')
        bloco = gerar_bloco(canal, reg, tipo, drows, outros, extras)
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
