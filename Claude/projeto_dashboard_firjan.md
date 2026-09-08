---
name: projeto-dashboard-firjan
description: "Estado atual do Dashboard Firjan — campanhas, indicadores, ocupação, infraestrutura"
metadata: 
  node_type: memory
  type: project
  originSessionId: dc7e6f4f-9b60-4182-8676-8b44b2e62b45
  modified: 2026-09-03T13:36:41.960Z
---

## Resumo Executivo

Dashboard Firjan é um sistema de monitoramento de campanhas e receptivo hospedado em GitHub Pages. Processa dados de múltiplos canais (Telefone, WhatsApp, Chat, E-mail, Redes Sociais) com atualização diária via Agendador de Tarefas do Windows.

**Estrutura**: Python scripts (`atualizar_*.py`) → Excel bases (`Arquivos/atualizaveis/`) → JavaScript arrays injetados → `index.html` renderiza.

---

## Campanhas Implementadas (6 abas)

Cada campanha segue padrão: `calcular_*() → combinar_*() → gerar_bloco_*()`

| Campanha | Status | Arch. Python | Sucesso Statuses | Tentativa Label | Notas |
|----------|--------|-------------|------------------|-----------------|-------|
| **Retomada da Trilha** | ✅ Completo | `atualizar_retomada.py` | INTERESSADO(A), Interessado, JÀ INFORMADO, JÁ INFORMADO, NÃO INTERESSADO(A), Não Interessado, Retornar (7) | Agrupa: Falhou, Fora de Área/Cx Msgs, Ligação Muda, Não Atendeu, Ocupado, Tel Não Atende/Ocupado | Row-level counter (não set), dataMin/dataMax p/ período |
| **Smart Factory** | ✅ Completo | `atualizar_smart.py` | Enviar E-mail, Formulário Preenchido, INFORMADO, Não tem interesse, Retornar (5) | Mesmas 6 de Retomada | Dois blocos: principal + Agendamentos |
| **Cursos Técnicos Niterói** | ✅ Completo | `atualizar_cursos_niteroi.py` | INFORMADO, JA INSCRITO, NAO INTERESSADO, Não tem interesse, INTERESSADO (5) | Falhou, Ligação Muda, Não Atendeu, Ocupado, Atendido | Row-level, corrigido NameError `dias_ord` |
| **Colônia Inverno 2026** | ✅ Completo | `atualizar_colonia_inverno.py` | JÁ INSCRITO, MATRÍCULA ONLINE, Não tem interesse, RETORNAR, INTERESSADO (5) | Atendido, FORA DE ÁREA/CX MENSAGENS, Ligação Muda, Não Atendeu, Ocupado, TEL NÃO ATENDE/OCUPADO | Grid 2D (por Status, por Unidade); corrigido NameError `dias_ord` |
| **Prospecção IEL** | ✅ Completo | `atualizar_iel.py` | INFORMAÇÕES POR E-MAIL, NAO INTERESSADO, Retornar, INTERESSADO (4) | Falhou, Fora de Area/Cx de Mensagens, Ligação Muda, Não Atendeu, Ocupado, Tel Não Atende/Ocupado, Atendido (7) | Row-level, dataMin/dataMax p/ período |
| **Smart Factory - Agendamentos** | ✅ Completo | `atualizar_smart.py` (bloco 2) | Agendamento Realizado, NAO INTERESSADO, Retornar, INTERESSADO (4) | Mesmas 7 de Retomada | Compartilha código com Smart Factory principal |

**Padrão de Sucesso**: Cada campanha define estatuses que contam como "Contatos de Sucesso" → soma esses para o card, resto vai para gráfico "Tentativas de Contato Sem Sucesso" (bar chart, evoBar=true).

---

## Receptivo (aba nova)

| Componente | Status | Arquivo | Detalhes |
|-----------|--------|---------|----------|
| **CSAT Geral** | ✅ | `atualizar_receptivo.py` + JS | Lê BASE_CSAT.xlsx, mostra % de boas respostas; filtro por Data/Canal; sem Regional/Entidade/Unidade |
| **Respostas CSAT** | ✅ | `atualizar_receptivo.py` + JS | Conta linhas da BASE_CSAT no período filtrado |
| **Ocupação** | ✅ Novo | `atualizar_ocupacao.py` + JS | Calcula ocupação (Voz vs Digital separado, combinado em "Todos") |

**Ocupação — Fórmulas**:
- **Voz (Tempos Callflex)**: `tempo_em_atendimento (T) / (tempo_logado (G) − pausas (N))` em segundos
- **Digital (Tempos Sales)**: `soma(F-D onde tipo=atendimento) / (soma(F todas) − soma(F pausas))` em segundos
- **Regra**: Sem filtro → soma combinada Voz+Digital; Com Canal → só fonte correspondente; Regional/Entidade/Unidade filtrados → "--"

---

## Indicadores por Canal — Metas (Base Indicadores)

Novo sistema para que IEC, IAL, TME, ICT venham da planilha `Base_Indicadores.xlsx` em vez de valores fixos no JS.

| Indicador | Voz | WhatsApp/Chat/E-mail | Digital (Instagram/Facebook/Messenger) | Autonomia e Renda | Implementação |
|-----------|-----|--------|----------|----------|----------|
| **IEC** | ✅ Lê planilha | ✅ Lê planilha | ✅ Lê planilha | — | `getIECFixed(src)` em JS, lê INDICADORES_BASE |
| **IAL** | ✅ Calcula descador | ✅ Lê planilha | ✅ Lê planilha | — | `getIALFixed(src)`, fallback p/ card.fix se não tiver linha |
| **TME** | ✅ Calcula discador | ✅ Lê planilha (horário→seg) | ✅ Lê planilha | — | `getTMEFixed(src)`, converte `HH:MM:SS`/`--` em segundos |
| **ICT** | ✅ Calcula discador | ✅ Lê planilha | ✅ Lê planilha | — | `getICTFixed(src)`, fallback se não tiver linha |

**Importância**: TME e ICT dos digitais agora vêm de dados reais (planilha preenchida pela operação) em vez de "0" hardcoded. Quando usuário adiciona linhas em Base_Indicadores com Data Início/Fim/Canal/Tipo/Entregue, roda `atualizar_indicadores.py` e JS puxa o valor automaticamente.

---

## Infraestrutura & Pipeline

### Arquivos-chave
- **index.html**: ~1900 linhas, renderiza todas as abas; marcadores `/* NAME_START/END */` para injeção JS
- **atualizar_tudo.py**: Orquestra 11+ módulos de atualização em sequência
- **Base_Indicadores.xlsx**: Centraliza IEC/IAL/TME/ICT por período/canal
- **Base Ocupação.xlsx**: Tempos Callflex (Voz) + Tempos Sales (Digital)

### Pipeline Diário
1. **Agendador de Tarefas** roda `atualizar_e_publicar_agendado.bat` (com retry automático + stdin não-interativo)
2. `atualizar_tudo.py` → 11 scripts Python em série
3. `git add/commit/pull/push origin main` com 3 tentativas cada
4. GitHub Actions → GitHub Pages publica em ~1min
5. **URL**: https://grdmmplan-source.github.io/dashboard-firjan/

### Fixes Recentes (Sept 2026)
- **Retry automático**: `atualizar_e_publicar_agendado.bat` usa `< NUL` para bloquear prompts interativos; 3 tentativas git pull/push com 5s de intervalo — resolve travamento que bloqueava execuções seguintes (10h → 11h/12h ficavam presas)
- **Motivo case-sensitivity**: "Incompatibilidade de Agenda" vs lowercase = chaves diferentes em dicts → normaliza p/ lowercase em agregação, preserva first-seen em display
- **Card vs gráfico mismatch**: Mudou de set-based (unique telefones) p/ row-level counting (cada linha = 1 contador) → card agora bate com somas

---

## Fórmulas de Cálculo Padrão

### Contatos de Sucesso (Card)
```
decisor_count = 0
FOR each row IN discagem:
  IF row.status normalized IN SUCESSO_MAP:
    decisor_count += 1
```

### Tentativas de Contato Sem Sucesso (Gráfico bar)
```
FOR each row IN discagem:
  IF row.status normalized NOT IN SUCESSO_MAP:
    normalized_label = SEM_OPERADOR_NORM.get(row.status) || 'Tentativa'
    raw_por_label_ns[raw_status] = raw_status  # p/ tooltip
    counter[normalized_label] += 1
```

### Período Dinâmico
```
dataMin = min(row.date para todas linhas)
dataMax = max(row.date para todas linhas)
periodo = f"{dataMin.day:02d}/{MÊS_PT[dataMin.month]} — {dataMax.day:02d}/{MÊS_PT[dataMax.month]}"
```

---

## Problemas Conhecidos & Melhorias Futuras

| Problema | Status | Notas |
|----------|--------|-------|
| **Ocupação Digital não é por-canal** | 🟡 By design | Base Tempos Sales só tem "Digital" genérico, sem distinção WhatsApp/Chat/E-mail → qualquer canal digital mostra mesmo %. Seria preciso nova fonte de dados |
| **Regional/Entidade/Unidade granularidade** | ❌ Não tem | Ocupação, CSAT e Indicadores por Canal não têm esses níveis (bases não têm essa info) → filtrar por eles mostra "--" |
| **URA (aba 11)** | ✅ Implementado | Roda via `atualizar_ura.py`, independente das campanhas |
| **Potencializ EE (aba PotencializEE)** | ✅ Implementado | Tem gráfico de Distribuição + "Tentativas" bar chart igual ao padrão das campanhas |

---

## Tarefas Concluídas (Sessão Atual)

1. ✅ **Campanhas Sucesso/Tentativa**: Padronizou todos 6 campanhas p/ usar card "Contatos de Sucesso" + bar chart "Tentativas"
2. ✅ **Indicadores Dinâmicos**: IEC/IAL/TME/ICT agora consultam Base_Indicadores.xlsx (com fallback p/ valores fixos)
3. ✅ **TME Horário→Segundos**: Converter `HH:MM:SS`/`--` em segundos na função `fmt_valor()` do `atualizar_indicadores.py`
4. ✅ **Ocupação Card**: Novo card "🎯 Ocupação" em Receptivo; calcula Voz (Tempos Callflex) vs Digital (Tempos Sales) separado e combinado
5. ✅ **Retry Automático**: Ajustou `atualizar_e_publicar_agendado.bat` p/ bloquear prompts interativos e fazer retry automático → resolve travamento de execuções sequenciais

---

## Próximas Etapas Esperadas

- Usuário preenche Base_Indicadores.xlsx com linhas de IAL/TME/ICT por período/canal → roda `atualizar_indicadores.py` → valores aparecem nos cards
- Usuário preenche Base Ocupação com novos dados (Tempos Callflex/Sales) → roda `atualizar_ocupacao.py` → card Ocupação atualiza
- Observar execuções agendadas (10h, 11h, 12h...) para confirmar que retry automático previne travamentos

---

## Links & Referências

- **Repo GitHub**: https://github.com/grdmmplan-source/dashboard-firjan
- **Dashboard Live**: https://grdmmplan-source.github.io/dashboard-firjan/
- **Bases Excel**: `\\192.168.1.20\ddmrj-dados\Planejamento\FIRJAN\Dashboard\dashboard-firjan\Arquivos\atualizaveis\`
- **Scripts Python**: `\\192.168.1.20\ddmrj-dados\Planejamento\FIRJAN\Dashboard\dashboard-firjan\atualizar_*.py`
- **Agendador**: Task Scheduler → "Atualizar Dashboard Firjan" (múltiplas vezes ao dia)
