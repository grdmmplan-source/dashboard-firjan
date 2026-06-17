# -*- coding: utf-8 -*-
"""
Atualizador Dashboard Firjan — TODAS AS CAMPANHAS
Roda Retomada da Trilha + Smart Factory e atualiza o index.html completo.
"""

import os, sys
import atualizar_retomada       as ar
import atualizar_smart          as asm
import atualizar_cursos_niteroi as an
import atualizar_saude          as asa
import atualizar_sac            as asc
import atualizar_receptivo      as arc
import atualizar_ura            as aura

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

        # 5. Gerar blocos JS
        print('\n[5/5] Gerando blocos JavaScript...')
        blocos = {
            'TODAS':    ar.gerar_bloco_todas(kt),
            'RETOMADA': ar.gerar_bloco_retomada(kr),
            'SMART':    asm.gerar_bloco_smart(ks),
            'NITEROI':  an.gerar_bloco_niteroi(kn),
        }

        # 6. Atualizar HTML (campanhas Ativo)
        print('\n[6/6] Atualizando index.html (Ativo)...')
        ar.atualizar_html(ar.INDEX_HTML, blocos)

        # 6. Promocao Saude (Google Sheets) — nao quebra se estiver offline
        print('\n[6/7] Promocao Saude (Google Sheets)...')
        try:
            rows = asa.baixar_csv(asa.CSV_URL)
            aux_cap = asa.baixar_aux(asa.AUX_URL) if asa.AUX_GID else {}
            emp, aco, esp, drows, saude_aux = asa.processar(rows, aux_cap)
            bloco_saude = asa.gerar_bloco(emp, aco, esp, drows, saude_aux)
            asa.atualizar_html(asa.INDEX_HTML, bloco_saude)
            print('  Saude atualizada.')
        except Exception as e:
            print(f'  [AVISO] Saude nao atualizada (offline ou erro): {e}')

        # 7. SAC (SharePoint) — nao quebra se estiver offline
        print('\n[7/7] SAC (SharePoint)...')
        try:
            xlsx = asc.baixar_xlsx(asc.SAC_URL)
            canal, reg, tipo, srows, erros, outros, extras = asc.processar(xlsx)
            asc.escrever_erros(asc.ERROS_TXT, erros, len(srows))
            bloco_sac = asc.gerar_bloco(canal, reg, tipo, srows, outros, extras)
            asc.atualizar_html(asc.INDEX_HTML, bloco_sac)
            print('  SAC atualizado.')
        except Exception as e:
            print(f'  [AVISO] SAC nao atualizado (offline ou erro): {e}')

        # 8. Receptivo (arquivo local em Arquivos\atualizaveis)
        print('\n[8/8] Receptivo (BSales2)...')
        try:
            caminho = arc.encontrar_arquivo(arc.PASTA, arc.PREFIXO)
            canal, ent, uni, rrows, disc, disc2, csat, extra, assunto, reg, iecT, iecD, seg, pes = arc.processar(caminho)
            bloco_rec = arc.gerar_bloco(canal, ent, uni, rrows, disc, disc2, csat, extra, assunto, reg, iecT, iecD, seg, pes)
            arc.atualizar_html(arc.INDEX_HTML, bloco_rec)
            print('  Receptivo atualizado.')
        except Exception as e:
            print(f'  [AVISO] Receptivo nao atualizado: {e}')

        # 9. URA (arquivo local em Arquivos\atualizaveis)
        print('\n[9/9] URA (BASE URA)...')
        try:
            caminho = aura.encontrar_arquivo(aura.PASTA, aura.PREFIXO)
            tipos, motivos, urows = aura.processar(caminho)
            bloco_ura = aura.gerar_bloco(tipos, motivos, urows)
            aura.atualizar_html(aura.INDEX_HTML, bloco_ura)
            print('  URA atualizada.')
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
