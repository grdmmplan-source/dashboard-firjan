# Arquitetura Técnica — Dashboard Firjan

Documentação técnica para desenvolvedores que precisam manter, estender ou depurar o projeto.

---

## 🏗️ Arquitetura Geral

```
┌─────────────────────┐
│   Fontes de Dados   │
├─────────────────────┤
│ • Excel (local)     │
│ • Google Sheets     │
│ • SharePoint        │
└──────────┬──────────┘
           │
           ↓
┌─────────────────────────────────┐
│   Scripts Python (atualizar_*.py)│
├─────────────────────────────────┤
│ • Extração de dados             │
│ • Processamento/transformação   │
│ • Geração de blocos JS          │
│ • Injeção em HTML               │
└──────────────┬──────────────────┘
               │
               ↓
        ┌──────────────┐
        │  index.html  │
        │ (JS puro)    │
        ├──────────────┤
        │ • Filtros    │
        │ • Gráficos   │
        │ • Tabelas    │
        └──────────────┘
               │
               ↓
       ┌──────────────┐
       │ GitHub Pages │
       │ (publicado)  │
       └──────────────┘
```

---

## 📐 Pipeline de Dados

### 1. **Extração** (`atualizar_*.py`)

Cada módulo tem seu próprio script:

```python
# atualizar_sac.py
def processar(xlsx_bytes):
    # 1. Parse do XLSX
    # 2. Filtragem de linhas (col A = Protocolo)
    # 3. Mapeamento de valores (de-para)
    # 4. Agregação por dimensões (canal, regional, etc.)
    # 5. Cálculos (Tempo Médio, CSAT, outliers)
    # 6. Retorno: (canal_list, reg_list, tipo_list, data_rows, ...)
    return canal, reg, tipo, data_rows, erros, outros, extras
```

### 2. **Geração de Bloco JS**

```python
def gerar_bloco(canal_list, reg_list, tipo_list, data_rows, outros_labels, extras):
    return f"""
    /* SAC_DATA_START */
    const SAC_CANAL={js_str(canal_list)};
    const SAC_REG={js_str(reg_list)};
    const SAC_ROWS={js_rows(data_rows)};
    /* SAC_DATA_END */
    """
```

Blocos são injetados entre marcadores: `/* SAC_DATA_START */` ... `/* SAC_DATA_END */`

### 3. **Injeção em HTML**

```python
def atualizar_html(index_path, bloco):
    with open(index_path, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    conteudo = re.sub(r'/\* SAC_DATA_START \*/.*?/\* SAC_DATA_END \*/', 
                      bloco, conteudo, flags=re.DOTALL)
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(conteudo)
```

### 4. **Renderização no JS** (`renderSac()`)

```javascript
function renderSac(){
    // 1. Coleta filtros
    const di=giv('flt-sac-ini'), df=giv('flt-sac-fim');
    const diN=di?parseInt(di.replace(/-/g,''),10):0;
    
    // 2. Loop sobre SAC_ROWS com filtros
    for(const r of SAC_ROWS){
        const dt=r[0], sc=r[4];
        if(dt&&(dt<diN||dt>dfN))continue;  // filtro data
        if(fr!==''&&SAC_REG[r[2]]!==fr)continue;  // filtro regional
        // ... mais filtros
        // contadores += 1
    }
    
    // 3. Atualiza cards
    setTxt('sc-total', total.toLocaleString('pt-BR'));
    
    // 4. Recarrega gráficos
    mkChart('sv-sac-line', 'line', LINE);
}
```

---

## 🎯 Campanhas Ativo — Objeto CAMPAIGNS

Cada campanha é um objeto JS injetado no `index.html` dentro do objeto `CAMPAIGNS`:

```javascript
const CAMPAIGNS = {
    /* TODAS_START */
    'todas': { label:'...', empresas:..., tentativas:..., ... },
    /* TODAS_END */
    /* RETOMADA_START */
    'retomada': { ... },
    /* RETOMADA_END */
    /* SMART_START */
    'smart': { ... },
    /* SMART_END */
    /* NITEROI_START */
    'niteroi': { ... },
    /* NITEROI_END */
};
```

### Campos obrigatórios de cada campanha

| Campo | Tipo | Descrição |
|-------|------|-----------|
| `label` | string | Nome completo exibido no dropdown |
| `empresas` | int | Total de empresas (ou inscrições) |
| `tentativas` | int | Total de tentativas |
| `interessados` | int | Total de interessados |
| `decisor` | int | Total de contatos c/ decisor |
| `conversao` | string | Taxa de conversão formatada |
| `media` | string | Média de tentativas formatada |
| `evoLabels` | array | Datas da evolução diária |
| `evoData` | array | Valores da evolução diária |
| `statusLabels` | array | Labels do gráfico de status |
| `statusData` | array | Valores do gráfico de status |

### Campos de personalização (dinâmicos via `applyFilter()`)

| Campo | Padrão | Niterói |
|-------|--------|---------|
| `empresasLabel` | `'🏢 Empresas na Base'` | `'📋 CPFs Inscritos'` |
| `mediaLabel` | `'🔁 Média Tentativas/Empresa'` | `'🔁 Média Tentativas/Inscrição'` |
| `mediaSub` | `'por empresa'` | `'por inscrição'` |
| `periodo` | `'01/05 — 15/06'` | idem |
| `wppTitle` | `'WhatsApp — Smart Factory'` | `'WhatsApp — Cursos Técnicos Niterói'` |
| `wppDesc` | descrição | descrição |
| `wppKpiLabels` | `['📤 Enviados','✅ Respostas',...]` | `['📤 Total Disparado','✅ Entregues',...]` |
| `wppListLabels` | `['Enviados','Lidos',...]` | `['Total Disparado','Lidos',...]` |
| `wppPieLabels` | `['Respostas','Sem Resposta']` | `['Entregues','Com Erro']` |

> **`'todas'`** tem `periodo: ''` — o campo de período fica oculto no dashboard.

### Soma das campanhas (`somar_campanhas`)

```python
# atualizar_retomada.py
def somar_campanhas(lista):
    total_emp = sum(k['_empresas'] for k in lista)
    # _empresas é a chave interna; para Niterói = _inscricoes
    ...
```

Niterói define `_empresas = _inscricoes` internamente para compatibilidade com a soma.

---

## 🗺️ De-para (Mapeamentos)

Arquivo: `Arquivos\bases_apoio\tab_de-para.xlsx`

Abas:
- **status:** raw_status → normalized_label
- **entidade:** unidade_raw → entidade_group

### Entradas de status relevantes

| Raw | Normalizado |
|-----|-------------|
| `Tentativa de contato` | `Tentativa` |
| `Contato com decisor` | `Contato c/ Decisor` |
| `JA INSCRITO` | `Contato c/ Decisor` |
| `NAO INTERESSADO` | `Contato c/ Decisor` |
| `INTERESSADO` | `Interessado` |
| `Nao localizado` | `Não Localizado` |

> **Atenção:** entradas sem acentos (`JA INSCRITO`, `NAO INTERESSADO`) devem constar explicitamente — o de-para não normaliza acentos automaticamente.

**Uso em Python:**
```python
def ler_depara(caminho):
    wb = openpyxl.load_workbook(caminho, read_only=True)
    ws = wb['status']
    mapa = {}
    for row in ws.iter_rows(values_only=True):
        if row[0] and row[1]:
            mapa[str(row[0]).strip()] = str(row[1]).strip()
    return mapa
```

O de-para é carregado uma vez em `atualizar_tudo.py` e distribuído a todos os scripts via `STATUS_MAP.update(depara)`.

---

## 📊 Scripts Ativo em Detalhe

### `atualizar_retomada.py`

- **Mailing:** lê `Mailing_*.xlsx` na pasta, conta empresas por CNPJ único
- **Retorno:** lê `Retorno_*.xlsx`, conta tentativas/status/evolução por linha
- Exporta: `calcular_kpis()`, `gerar_bloco_retomada()`, `gerar_bloco_todas()`, `somar_campanhas()`, `ler_depara()`, `atualizar_html()`

### `atualizar_smart.py`

- **Mailing:** conta CNPJs únicos
- **Retorno:** conta tentativas/status/evolução
- **WhatsApp:** lê `Whatsapp_*.xlsx` aba `Planilha1`, colunas Enviados/Respondidos/Lidos
- Exporta: `calcular_kpis_smart()`, `gerar_bloco_smart()`

### `atualizar_cursos_niteroi.py`

- **Mailing:** conta apenas o total de linhas (1 linha = 1 CPF inscrito). **Não lê status** — evita "LIGAR" no gráfico de status.
- **Retorno:** aba `chamadas_*` — conta tentativas, status e evolução
- **WhatsApp:** lê `Whatsapp_*.xlsx` aba `Planilha1`, coluna `Status` com valores: `Erro`, `Enviado`, `Entregue`, `Lido`
  - KPIs: Total Disparado | Entregues (Entregue+Lido) | Taxa de Entrega | Com Erro
- Exporta: `calcular_kpis_niteroi()`, `gerar_bloco_niteroi()`

---

## 🖥️ Modo Painel (Kiosk)

O Modo Painel roda em loop automático em TVs/monitores. Estrutura:

```javascript
const KP = ['kp-ativo','kp-receptivo','kp-saude','kp-sac','kp-ura','kp-campanha'];
const KN = ['Ativo','Receptivo','Promoção Saúde','SAC','URA','Campanha'];
```

Cada página é um `<div id="kp-*">` com cards KPI específicos. O loop avança entre páginas via `nextKiosk()` / `prevKiosk()`.

### Páginas do Kiosk

| ID | Título | Dados |
|----|--------|-------|
| `kp-ativo` | Ativo | Campanhas ativas (tentativas, interessados, conversão) |
| `kp-receptivo` | Receptivo | IAL, ICT, TMA, CSAT |
| `kp-saude` | Promoção Saúde | Intenções, Agendadas, Pendentes |
| `kp-sac` | SAC | Total, Satisfeitos, TMR |
| `kp-ura` | URA | Retidos, Resolvidos, Abandonos, TMA |
| `kp-campanha` | Campanha | Total, CSAT Respostas, CSAT Geral |

### Inicialização

```javascript
function openKiosk(){
    renderAtivo();
    renderReceptivo();
    renderSaude();
    renderSac();
    renderUra();
    renderCamp();   // <-- popula kp-campanha
    initStaticCharts();
    startKiosk();
}
```

---

## 🔗 Fluxo de Cascata em Filtros (SAC)

```
Regional
   ↓ (depende de Data)
Entidade
   ↓ (depende de Data + Regional)
Canal
   ↓ (depende de Data + Regional + Entidade)
Tipo → Produto → Assunto
```

**Implementação:**
```javascript
const SAC_FILTERS=[
    {id:'flt-sac-reg',   get:()=>SAC_REG,   col:2, ph:'Todas'},
    {id:'flt-sac-ent',   get:()=>SAC_ENT,   col:6, ph:'Todas'},
    {id:'flt-sac-canal', get:()=>SAC_CANAL, col:1, ph:'Todos'},
    // ...
];

function refreshSacOptions(){
    for(let k=0; k<SAC_FILTERS.length; k++){
        const f=SAC_FILTERS[k];
        const vals=new Set();
        for(const r of SAC_ROWS){
            if(!passaData(r)) continue;
            let ok=true;
            for(let j=0; j<k; j++){
                const g=SAC_FILTERS[j];
                if(g.val!==''&&g.get()[r[g.col]]!==g.val){ok=false; break;}
            }
            if(!ok) continue;
            const v=f.get()[r[f.col]];
            if(v) vals.add(v);
        }
        repopSel(el, sortPT(vals), f.ph);
    }
}
```

---

## 📊 Cálculos Especiais

### **CSAT (Satisfação)**
```python
# SAC: Satisfeitos / (Satisfeitos + Insatisfeitos)
# Ignora: Não Avaliado, Em branco
csat_pct = (satisfeitos / (satisfeitos + insatisfeitos)) * 100
```

### **Tempo Médio de Resposta (SAC)**
```python
# Exclui outliers: delta < 0 ou > 30 dias
tma_segundos = soma_deltas / quantidade_valida
```

### **IAR (Atendimentos Registrados, Receptivo)**
```python
# Linhas com Classificação do Ticket / Total de linhas BSales2
iar_pct = (com_classificacao / total_bsales2) * 100
```

### **Taxa de Entrega (WhatsApp Niterói)**
```python
# Entregues = Entregue + Lido (já chegou ao destino)
entregues = contagem['Entregue'] + contagem['Lido']
taxa_pct  = entregues / total_disparado * 100
```

### **Taxa de Ocupação (Saúde)**
```python
# Para cada médico presente na aba AUX (col A=nome, col B=capacidade diária):
#   dias_trabalhados = datas distintas na col M (Data Agendamento) com esse médico
#   agendamentos     = linhas com esse médico onde col M está preenchida
#   capacidade_total = dias_trabalhados × capacidade_diária
# Taxa global = sum(agendamentos) / sum(capacidade_total) × 100
```
- Calculada dinamicamente no JS respeitando filtros de data e médico
- Requer `AUX_GID` preenchido em `atualizar_saude.py` (GID da aba AUX no Google Sheets)
- Injetada como `SAUDE_AUX = [[idx_medico, capacidade], ...]` e `r[7]` em `SAUDE_ROWS` = `dt_agend`

---

## ➕ Adicionar Nova Campanha Ativo

Para adicionar uma nova campanha Ativo ao dashboard:

### Passo 1 — Criar o script `atualizar_{id}.py`

Copie `atualizar_cursos_niteroi.py` como template e ajuste:

```python
# Constantes de caminho
PASTA    = r'Arquivos\nao_atualizaveis\Ativo_{NomeDaPasta}'
MARKER   = 'NOVA_CAMP'   # deve ser único

# Funções a implementar:
# calcular_mailing_*()  — contar empresas/inscrições
# calcular_retorno_*()  — ler tentativas/status/evolução
# calcular_wpp_*()      — opcional, se tiver WhatsApp
# combinar_*()          — juntar os dados
# gerar_bloco_*()       — gerar JS (inclui todos os campos de personalização)
```

Campos obrigatórios no `gerar_bloco_*()`:

```python
{
    'label':         'Nome da Campanha',
    'empresasLabel': '🏢 Empresas na Base',   # ou '📋 CPFs Inscritos'
    'mediaLabel':    '🔁 Média Tentativas/Empresa',
    'mediaSub':      'por empresa',
    'periodo':       f"{evoLabels[0]} — {evoLabels[-1]}",
    'wppTitle':      'WhatsApp — Nome',
    'wppDesc':       '...',
    'wppKpiLabels':  ['📤 Enviados','✅ Respostas','📊 Taxa de Resposta','❌ Sem Resposta'],
    'wppListLabels': ['Enviados','Lidos','Sem Resposta','Com Erro'],
    'wppPieLabels':  ['Respostas','Sem Resposta'],
    '_empresas':     total,   # chave interna para somar_campanhas()
}
```

### Passo 2 — Adicionar marcadores no `index.html`

Em `<script>`, dentro do objeto `CAMPAIGNS`, adicione após a última campanha:

```html
/* NOVA_CAMP_START */
/* NOVA_CAMP_END */
```

E no `<select>` de filtro de campanha:
```html
<option value="nova_camp">Nome da Campanha</option>
```

### Passo 3 — Atualizar `atualizar_tudo.py`

```python
import atualizar_nova_camp as anc

# Na main():
kn2 = anc.calcular_kpis_nova_camp()
kt  = ar.somar_campanhas([kr, ks, kn, kn2])  # adicionar à lista

blocos['NOVA_CAMP'] = anc.gerar_bloco_nova_camp(kn2)
anc.STATUS_MAP.update(depara)
```

### Passo 4 — Atualizar `gerar_bloco_todas()` em `atualizar_retomada.py`

Atualize a descrição (`desc`) do bloco "Todas" para incluir a nova campanha:

```python
'desc': 'Retomada + Smart Factory + Niterói + Nova Campanha',
```

### Passo 5 — Atualizar `applyFilter()` no `index.html`

Adicione o badge da nova campanha:

```javascript
// Na lógica de badge color (buscar "badge.style.background")
currentCamp==='nova_camp' ? '#cor_hex' : ...
```

---

## 🎨 Estrutura de Cards

### Card simples (KPI)
```html
<div class="kpi">
    <div class="kpi-label" id="kpi-label-empresas">🏢 Empresas na Base</div>
    <div class="kpi-val lg" id="kv-empresas">0</div>
</div>
```

Os IDs com prefixo `kpi-label-*` são atualizados dinamicamente por `applyFilter()` ao trocar de campanha.

### Card com barra (indicador %)
```javascript
setQ('rc-csat', 'rc-csat-bar', 86.7, true);
// setQ(idVal, idBar, percentual, visível?)
```

### Ranking (Top 5)
```javascript
function rankHTML(counter, topN){
    const top = Object.entries(counter)
        .sort((a,b) => b[1].n - a[1].n)
        .slice(0, topN);
    return top.map((p,i) => `<div class="rank-row">...</div>`).join('');
}
```

---

## 🔍 Debug & Logs

### Python
```python
print(f'  Total de SACs: {total}')
escrever_erros('erros_sac.txt', erros, total)
```

### JavaScript (console do navegador)
```javascript
// Inspecionar dados injetados
console.log(CAMPAIGNS)             // todos os objetos de campanha
console.log(SAC_ROWS.slice(0,5))   // primeiras 5 linhas SAC
console.log(SAC_CANAL.filter(v=>v))// canais distintos
```

---

## 🧪 Testes Manuais

### Verificar injeção JS
```javascript
// No console do navegador:
console.log(CAMPAIGNS['niteroi'].empresas)  // deve ser > 0
console.log(CAMPAIGNS['todas'].tentativas)   // soma das campanhas
```

### Limpar + recarregar
```bash
limpar.bat          # apaga dados
atualizar.bat       # recarrega
# abrir index.html → deve aparecer os dados
```

### Checar status no gráfico
Se aparecer "LIGAR" no gráfico de distribuição por status:
1. Verifique se `calcular_mailing_*()` está lendo a coluna Status (não deve)
2. Verifique se `combinar_*()` usa apenas dados do Retorno para status/tentativas

---

## 📚 Padrões de Código

### Python
- **Encoding:** `# -*- coding: utf-8 -*-` no topo
- **Paths:** Relativos: `os.chdir(os.path.dirname(os.path.abspath(__file__)))`
- **Excel:** `openpyxl.load_workbook(..., read_only=True, data_only=True)`
- **Strings seguras:** `txt(val) = str(val).strip() if val else ''`
- **Datas:** `yyyymmdd int` (ex: 20260608) para comparação rápida
- **Pausas:** `if '--no-pause' not in sys.argv: input('...')`

### JavaScript
- **IDs:** `flt-{modulo}-{campo}` (ex: `flt-sac-reg`)
- **Constantes globais:** `SAC_ROWS`, `SAC_CANAL` (maiúsculas)
- **Função de get:** `giv(id) = document.getElementById(id)?.value || ''`
- **Datas:** Parse `yyyy-mm-dd` (input) → `yyyymmdd int` (comparação)

---

## 🐛 Problemas Conhecidos & Workarounds

| Problema | Causa | Solução |
|----------|-------|---------|
| Datas com mês > 12 | Excel lê 07/12 como "July 12" (US) | `to_dt()` swaps dia↔mês se > hoje |
| SAC_ROWS vazio | SharePoint offline | Script continua com [AVISO], SAC fica com dados antigos |
| Filtros não repopulam | `refreshSacOptions()` não chamada | Verificar `addEventListener` em `initSacFilters()` |
| Gráfico fica preto | Muitos dados/browser lento | Limitar a 1 ano (`filterByYear()`) |
| "LIGAR" no status chart | `calcular_mailing_*()` lendo coluna Status | Simplificar para contar apenas linhas |
| "JA INSCRITO" não agrupa | Raw sem acento ausente no de-para | Adicionar entrada exata em `tab_de-para.xlsx` aba `status` |
| Taxa de Ocupação mostra `--` | `AUX_GID` vazio ou médicos da AUX sem correspondência em `SAUDE_MED` | Preencher `AUX_GID` em `atualizar_saude.py` e checar ortografia dos nomes |
| Colunas Receptivo deslocadas | Inserção de novas colunas no BSales2 | Atualizar índices `COL_*` em `atualizar_receptivo.py` |
| Filtro Unidade/Regional/Entidade mistura fontes | BSales2 e Redes Sociais compartilham as mesmas listas | Esperado — cada fonte contribui com seus valores; verificar `RS_UNI`, `RS_ENT`, `RS_REGIONAL` |

---

## 📖 Referências

- **Chart.js:** https://www.chartjs.org/ (gráficos)
- **openpyxl:** https://openpyxl.readthedocs.io/ (Excel em Python)
- **Python datetime:** https://docs.python.org/3/library/datetime.html

---

---

## 📋 Estrutura de `RECEP_ROWS` (Receptivo / Campanha)

Cada linha: `[dt, canal, entidade, unidade, hc, assunto, regional, src, segmento, pesquisa]`

| Índice | Campo | Fonte BSales2 | Fonte Redes | Fonte Autonomia |
|--------|-------|--------------|-------------|-----------------|
| 0 | `dt` | Col D (Data abertura) | Col J (Resolvido) | Col DATA INICIAL |
| 1 | `canal` | Col E (Origem do caso) | Col K (Observação) | fixo "Autonomia e Renda" |
| 2 | `entidade` | Col X (Entidade2) | Col G (Entidade) | vazio |
| 3 | `unidade` | Col W (Unidade2) | Col F (Unidade) | vazio |
| 4 | `hc` | Col L (Classificação ≠ vazio) | sempre 0 | sempre 0 |
| 5 | `assunto` | Col L (Classificação) | Col E (Assunto) | Col CATEGORIA |
| 6 | `regional` | Col Y (Regional2) | Col H (Regional) | vazio |
| 7 | `src` | 0 | 1 | 2 |
| 8 | `segmento` | Col P | -1 | -1 |
| 9 | `pesquisa` | Col Q | -1 | -1 |

> A aba **Campanha** filtra apenas linhas BSales2 (`src===0`) cujo assunto começa com `"Sub"`.

### Mapeamento de colunas — Redes Sociais (SharePoint)

| Col | Letra | Campo no filtro |
|-----|-------|----------------|
| 4 | E | Assunto (top serviços) |
| 5 | F | Unidade |
| 6 | G | Entidade |
| 7 | H | Regional |
| 9 | J | Data (Resolvido) |
| 10 | K | Canal (Observação) |

---

**Versão:** 2.2  
**Última atualização:** 17/06/2026
