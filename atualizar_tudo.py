# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan — TODAS AS CAMPANHAS
Roda Retomada da Trilha + Smart Factory e atualiza o index.html completo.
"""

import os, sys
import atualizar_retomada        as ar
import atualizar_smart           as asm
import atualizar_cursos_niteroi  as an
import atualizar_saude           as asa
import atualizar_sac             as asc
import atualizar_receptivo       as arc
import atualizar_ura             as aura
import atualizar_colonia_inverno as aci
import atualizar_iel             as ail
import atualizar_qualidade       as aq

def main():
    print()
    print('=' * 50)
    print('  ATUALIZADOR — Todas as Campanhas')
    print('=' * 50)

    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    try:
        # 0. De-para compartilhado
        print('\n[0/5] Carregando de-para...')
        depara = ar.ler_depara(ar.DEPARA_PATH)
        ar.STATUS_MAP.update(depara)
        asm.STATUS_MAP.update(depara)
        an.STATUS_MAP.update(depara)

        # 1. Retomada da Trilha
        print('\n[1/5] Calculando Retomada da Trilha...')
        kr = ar.calcular_kpis(ar.PASTA_MAILING)
        print(f'  Emp:{kr["empresas"]}  Tent:{kr["tentativas"]}  Int:{kr["interessados"]}  Dec:{kr["decisor"]}  Taxa:{kr["conversao"]}')

        # 2. Smart Factory
        print('\n[2/5] Calculando Smart Factory...')
        ks = asm.calcular_kpis_smart()
        print(f'  Emp:{ks["empresas"]}  Tent:{ks["tentativas"]}  Int:{ks["interessados"]}  Dec:{ks["decisor"]}  Taxa:{ks["conversao"]}')

        # 3. Cursos Técnicos Niterói
        print('\n[3/5] Calculando Cursos Técnicos Niterói...')
        kn = an.calcular_kpis_niteroi()
        print(f'  Insc:{kn["inscricoes"]}  Tent:{kn["tentativas"]}  Int:{kn["interessados"]}  Dec:{kn["decisor"]}  Taxa:{kn["conversao"]}')

        # 4. Todas (soma)
        print('\n[4/5] Calculando Todas as Campanhas (soma)...')
        kt = ar.somar_campanhas([kr, ks, kn])
        print(f'  Total:{kt["empresas"]}  Tent:{kt["tentativas"]}  Int:{kt["interessados"]}  Dec:{kt["decisor"]}  Taxa:{kt["conversao"]}')

        # 5. Smart Factory - Agendamentos (SharePoint + Discagem local) — nao quebra se offline
        print('\n[5/6] Calculando Smart Factory - Agendamentos...')
        bloco_smart_agend = None
        try:
            ka = asm.calcular_kpis_smart_agend()
            print(f'  Emp:{ka["empresas"]}  Tent:{ka["tentativas"]}  Dec:{ka["decisor"]}  Agend:{ka["agendamentos"]}')
            bloco_smart_agend = asm.gerar_bloco_smart_agend(ka)
        except Exception as e:
            print(f'  [AVISO] Smart Factory - Agendamentos nao atualizado: {e}')

        # 6. Gerar blocos JS
        print('\n[6/6] Gerando blocos JavaScript...')
        blocos = {
            'TODAS':    ar.gerar_bloco_todas(kt),
            'RETOMADA': ar.gerar_bloco_retomada(kr),
            'SMART':    asm.gerar_bloco_smart(ks),
            'NITEROI':  an.gerar_bloco_niteroi(kn),
        }
        if bloco_smart_agend:
            blocos['SMART_AGEND'] = bloco_smart_agend

        # 7. Atualizar HTML (campanhas Ativo)
        print('\n[7/7] Atualizando index.html (Ativo)...')
        ar.atualizar_html(ar.INDEX_HTML, blocos)

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

        # 11. URA (arquivo local em Arquivos\atualizaveis)
        print('\n[11/11] URA (BASE URA)...')
        try:
            aura.main()
        except Exception as e:
            print(f'  [AVISO] URA nao atualizada: {e}')

        # 12. Colônia Inverno 2026 (Google Sheets + Discagem local) — nao quebra se offline
        print('\n[12/13] Colônia Inverno 2026...')
        try:
            aci.main()
        except Exception as e:
            print(f'  [AVISO] Colônia Inverno 2026 nao atualizada: {e}')

        # 13. Prospecção IEL (Google Sheets + Discagem local) — nao quebra se offline
        print('\n[13/13] Prospecção IEL...')
        try:
            ail.main()
        except Exception as e:
            print(f'  [AVISO] Prospecção IEL nao atualizada: {e}')

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
