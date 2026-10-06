# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan — TODAS AS CAMPANHAS
Atualiza o index.html: Propag Nova Iguacu, Propag Tres Rios, IMES, Saude, SAC, Receptivo, URA etc.
As demais campanhas do Ativo estao congeladas (nao rodam aqui).
"""

import os, sys
import atualizar_saude           as asa
import atualizar_sac             as asc
import atualizar_receptivo       as arc
import atualizar_ura             as aura
import atualizar_qualidade       as aq
import atualizar_indicadores     as aind
import atualizar_ocupacao        as aoc
import atualizar_propag_ni       as apni
import atualizar_propag_tres_rios as aptr
import atualizar_imes            as aimes

def main():
    print()
    print('=' * 50)
    print('  ATUALIZADOR — Todas as Campanhas')
    print('=' * 50)

    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    try:
        # Campanhas do Ativo congeladas (Retomada, Smart Factory, Smart Agendamentos, Niteroi,
        # Colonia Inverno, IEL e PotencializEE): nao sao mais atualizadas aqui; o index.html
        # mantem os ultimos dados. Para atualizar uma delas, rode o atualizar_*.py correspondente.

        # 8. Promocao Saude (Google Sheets) — nao quebra se estiver offline
        print('\n[8/11] Promocao Saude (Google Sheets)...')
        try:
            asa.main()
        except Exception as e:
            print(f'  [AVISO] Saude nao atualizada (offline ou erro): {e}')

        # 9. SAC (SharePoint + Google Sheets) — nao quebra se estiver offline
        print('\n[9/11] SAC (SharePoint)...')
        try:
            asc.main()
        except Exception as e:
            print(f'  [AVISO] SAC nao atualizado (offline ou erro): {e}')

        # 10. Receptivo (arquivo local em Arquivos\atualizaveis)
        print('\n[10/11] Receptivo (BSales2)...')
        try:
            arc.main()
        except Exception as e:
            print(f'  [AVISO] Receptivo nao atualizado: {e}')

        print('\n[10b/11] Qualidade (Monitoria)...')
        try:
            aq.main()
        except Exception as e:
            print(f'  [AVISO] Qualidade nao atualizada: {e}')

        print('\n[10c/11] Base de Indicadores...')
        try:
            aind.main()
        except Exception as e:
            print(f'  [AVISO] Base de Indicadores nao atualizada: {e}')

        print('\n[10e/11] Propag Nova Iguaçu...')
        try:
            apni.main()
        except Exception as e:
            print(f'  [AVISO] Propag Nova Iguaçu nao atualizada: {e}')

        print('\n[10f/11] Propag Três Rios...')
        try:
            aptr.main()
        except Exception as e:
            print(f'  [AVISO] Propag Três Rios nao atualizada: {e}')

        print('\n[10g/11] IMES...')
        try:
            aimes.main()
        except Exception as e:
            print(f'  [AVISO] IMES nao atualizado: {e}')

        print('\n[10d/11] Ocupação...')
        try:
            aoc.main()
        except Exception as e:
            print(f'  [AVISO] Ocupação nao atualizada: {e}')

        # 11. URA (arquivo local em Arquivos\atualizaveis)
        print('\n[11/11] URA (BASE URA)...')
        try:
            aura.main()
        except Exception as e:
            print(f'  [AVISO] URA nao atualizada: {e}')

        # 10. Carimbo de data/hora da atualizacao
        ts = asc.carimbar_atualizacao(asc.INDEX_HTML)
        print(f'\n  Dashboard atualizado em: {ts}')

        print()
        print('=' * 50)
        print('  CONCLUIDO! index.html atualizado.')
        if '--no-pause' not in sys.argv:
            print('  Rode publicar.bat para enviar ao GitHub.')
        print('=' * 50)
        print()

    except Exception as e:
        print(f'\n[ERRO] {e}')
        import traceback; traceback.print_exc()

if __name__ == '__main__':
    main()
    if '--no-pause' not in sys.argv:
        input('Pressione ENTER para fechar...')
