# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan - Base de Indicadores (Receptivo)
Le Base_Indicadores.xlsx e atualiza o bloco INDICADORES no index.html.

Colunas da planilha:
  Data Inicio | Data fim | TIPO | Nome do indicador | Meta | Entregue | Bloco | Canal

Cada linha representa um indicador valido para um periodo (Data Inicio..Data fim),
associado a um bloco da tela (ex: 'Disponibilidade dos Canais',
'PERFORMANCE - CANAIS ATIVOS', 'Indicadores por Canal — Metas'). Quando 'Canal'
esta preenchido, a linha tambem deve ser filtrada pelo canal selecionado.
"""

import openpyxl
import os
import re
import glob

PASTA      = r'Arquivos\atualizaveis'
PREFIXO    = 'Base_Indicadores'
INDEX_HTML = r'index.html'

COL_INI    = 0
COL_FIM    = 1
COL_TIPO   = 2
COL_NOME   = 3
COL_META   = 4
COL_ENTREG = 5
COL_BLOCO  = 6
COL_CANAL  = 7


def encontrar_arquivo(pasta, prefixo):
    padrao = os.path.join(pasta, f'{prefixo}*.xlsx')
    arquivos = [a for a in glob.glob(padrao) if not os.path.basename(a).startswith('~$')]
    if not arquivos:
        raise FileNotFoundError(f'[ERRO] Arquivo nao encontrado: {prefixo}*.xlsx em {pasta}')
    return sorted(arquivos)[-1]


def data_int(dt):
    if dt is None:
        return None
    return dt.year * 10000 + dt.month * 100 + dt.day


def fmt_valor(v):
    """Formata Meta/Entregue para exibicao. Retorna (texto_exibicao, numero_bruto_ou_None)."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return '—', None
    if isinstance(v, (int, float)):
        if 0 < v <= 1:
            return f'{v * 100:.2f}'.replace('.', ',') + '%', v * 100
        if float(v).is_integer():
            return str(int(v)), float(v)
        return f'{v:.2f}'.replace('.', ','), float(v)
    return str(v).strip(), None


def calcular(caminho):
    print(f'  Lendo: {os.path.basename(caminho)}')
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    linhas = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if not row or row[COL_BLOCO] is None:
            continue
        ini = data_int(row[COL_INI]) if len(row) > COL_INI else None
        fim = data_int(row[COL_FIM]) if len(row) > COL_FIM else None
        tipo = str(row[COL_TIPO]).strip() if row[COL_TIPO] is not None else ''
        nome = str(row[COL_NOME]).strip() if row[COL_NOME] is not None else ''
        meta_disp, meta_num = fmt_valor(row[COL_META] if len(row) > COL_META else None)
        entreg_disp, entreg_num = fmt_valor(row[COL_ENTREG] if len(row) > COL_ENTREG else None)
        bloco = str(row[COL_BLOCO]).strip()
        canal = str(row[COL_CANAL]).strip() if (len(row) > COL_CANAL and row[COL_CANAL]) else ''
        linhas.append({
            'ini': ini, 'fim': fim, 'tipo': tipo, 'nome': nome,
            'metaDisp': meta_disp, 'metaNum': meta_num,
            'entregDisp': entreg_disp, 'entregNum': entreg_num,
            'bloco': bloco, 'canal': canal,
        })

    wb.close()
    print(f'  Linhas lidas: {len(linhas)}')
    return linhas


def js_str(v):
    return "'" + str(v).replace(chr(39), chr(92) + chr(39)) + "'"


def js_num(v):
    return 'null' if v is None else repr(v)


def gerar_bloco(linhas):
    itens = ',\n    '.join(
        '{ini:%s, fim:%s, tipo:%s, nome:%s, metaDisp:%s, metaNum:%s, entregDisp:%s, entregNum:%s, bloco:%s, canal:%s}' % (
            js_num(l['ini']), js_num(l['fim']), js_str(l['tipo']), js_str(l['nome']),
            js_str(l['metaDisp']), js_num(l['metaNum']), js_str(l['entregDisp']), js_num(l['entregNum']),
            js_str(l['bloco']), js_str(l['canal']),
        )
        for l in linhas
    )
    return f"""  /* INDICADORES_START */
  const INDICADORES_BASE = [
    {itens}
  ];
  /* INDICADORES_END */"""


def atualizar_html(index_path, bloco):
    with open(index_path, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    padrao = r'/\* INDICADORES_START \*/.*?/\* INDICADORES_END \*/'
    if not re.search(padrao, conteudo, re.DOTALL):
        raise ValueError('[ERRO] Marcadores INDICADORES nao encontrados no index.html.')
    conteudo = re.sub(padrao, lambda m: bloco.replace('\\', '\\\\'), conteudo, flags=re.DOTALL)
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(conteudo)


def main():
    print()
    print('=' * 50)
    print('  ATUALIZADOR — Base de Indicadores')
    print('=' * 50)
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        caminho = encontrar_arquivo(PASTA, PREFIXO)
        linhas = calcular(caminho)
        bloco = gerar_bloco(linhas)
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
