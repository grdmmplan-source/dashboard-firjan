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

**Padrão de linha** (SAC):
```python
data_rows.append([
    dt,           # 0: Data (int yyyymmdd)
    ci,           # 1: Canal (idx em SAC_CANAL[])
    ri,           # 2: Regional (idx em SAC_REG[])
    ti,           # 3: Tipo (idx em SAC_TIPO[])
    sc,           # 4: Satisfação (0=outro, 1=sim, 2=não, 3=nao-avaliado)
    dl,           # 5: Delta tempo (segundos, ou None se outlier)
    ei,           # 6: Entidade (idx em SAC_ENT[])
    aci,          # 7: Assunto (idx em SAC_ASSUNTO[])
    pi,           # 8: Produto (idx em SAC_PROD[])
    ui            # 9: Unidade (idx em SAC_UNI[])
])
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

## 🔗 Fluxo de Cascata em Filtros

Exemplo: SAC com 6 filtros em cascata

```
Regional
   ↓ (depende de Data)
Entidade
   ↓ (depende de Data + Regional)
Canal
   ↓ (depende de Data + Regional + Entidade)
Tipo
   ↓ ...
Produto
   ↓ ...
Assunto
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
            // verifica todos os filtros ANTES deste (k-1, k-2, ...)
            for(let j=0; j<k; j++){
                const g=SAC_FILTERS[j];
                if(g.val!==''&&g.get()[r[g.col]]!==g.val){
                    ok=false; break;
                }
            }
            if(!ok) continue;
            const v=f.get()[r[f.col]];
            if(v) vals.add(v);
        }
        // repopula o select
        repopSel(el, sortPT(vals), f.ph);
    }
}
```

---

## 🗺️ De-para (Mapeamentos)

Arquivo: `Arquivos\bases_apoio\tab_de-para.xlsx`

Abas:
- **status:** raw_status → normalized_label
  - Ex: "Tentativa de contato" → "Tentativa"
- **entidade:** unidade_raw → entidade_group
  - Ex: "SENAI Benfica" → "SENAI"

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

# statusMap['Tentativa de contato'] = 'Tentativa'
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
# Soma de (Data Finalização - Data Atendimento) / quantidade válida
# Exclui outliers: delta < 0 ou > 30 dias
tma_segundos = soma_deltas / quantidade_valida
tma_horas = tma_segundos / 3600
tma_dias = tma_segundos / 86400
```

### **IAR (Atendimentos Registrados, Receptivo)**
```python
# Linhas com Classificação do Ticket (col L) / Total de linhas BSales2
iar_pct = (com_classificacao / total_bsales2) * 100
```

### **Corrição de Data BR × EUA**
```python
# Se data > hoje (impossível), tenta trocar dia↔mês
if data_parsed > datetime.now():
    try:
        data_swapped = data.replace(month=data.day, day=data.month)
        if data_swapped <= datetime.now():
            data = data_swapped
    except ValueError:
        pass
```

---

## 🎨 Estrutura de Cards

### Card simples (KPI)
```html
<div class="kpi">
    <div class="kpi-label">Satisfeitos</div>
    <div class="kpi-val lg" id="sc-sat">0</div>
</div>
```
```javascript
setTxt('sc-sat', '1.245');  // atualiza textContent
```

### Card com barra (indicador %)
```javascript
setQ('rc-csat', 'rc-csat-bar', 86.7, true);
// função setQ: id_valor, id_barra, percentual, visível?

function setQ(idVal, idBar, val, ok){
    const e=document.getElementById(idVal);
    const b=document.getElementById(idBar);
    if(e) e.textContent = ok ? val.toFixed(2).replace('.',',')+'%' : '-- ';
    if(b) b.style.width = (ok ? val : 0) + '%';
}
```

### Ranking (Top 5)
```javascript
function rankHTML(counter, topN){
    // counter = {label: {n, sat, insat}}
    const top = Object.entries(counter)
        .sort((a,b) => b[1].n - a[1].n)
        .slice(0, topN);
    return top.map((p,i) => 
        `<div class="rank-row">...${p[0]}...${p[1].n}...</div>`
    ).join('');
}
```

---

## 🔍 Debug & Logs

### Python
```python
# Sempre printe status no console
print(f'  Total de SACs: {total}')
print(f'  Satisfeitos: {sat}')

# Logs em arquivo (SAC)
escrever_erros('erros_sac.txt', erros, total)
```

### JavaScript
```javascript
console.log('Filtros:', fsel);
console.log('Total após filtro:', total);

// Inspecione SAC_ROWS direto no console
console.log(SAC_ROWS.slice(0,5));  // primeiras 5 linhas
```

---

## 🧪 Testes Manuais

### Verificar injeção JS
```javascript
// No console do navegador:
console.log(SAC_ROWS.length);  // deve ser > 0
console.log(SAC_CANAL.filter(v=>v).length);  // canais distintos
```

### Limpar + recarregar
```bash
limpar.bat          # apaga dados
atualizar.bat       # recarrega
# abrir index.html → deve aparecer os dados
```

### Checar outliers de SAC
```bash
# Arquivo erros_sac.txt tem relatório de outliers excluídos
cat erros_sac.txt
```

---

## 📚 Padrões de Código

### Python
- **Encoding:** `# -*- coding: utf-8 -*-` no topo
- **Paths:** Relativos: `os.chdir(os.path.dirname(os.path.abspath(__file__)))`
- **Excel:** `openpyxl.load_workbook(..., read_only=True, data_only=True)`
- **Strings seguras:** `txt(val) = str(val).strip() if val else ''`
- **Datas:** `yyyymmdd int` (ex: 20260608) pra comparação rápida

### JavaScript
- **IDs:** `flt-{modulo}-{campo}` (ex: `flt-sac-reg`)
- **Constantes globais:** `SAC_ROWS`, `SAC_CANAL` (maiúsculas)
- **Função de get:** `giv(id) = document.getElementById(id)?.value || ''`
- **Datas:** Parse `yyyy-mm-dd` (input) → `yyyymmdd int` (comparação)

---

## 🚀 Adicionar Novo Módulo

Exemplo: adicionar módulo "Cobrança"

1. **Criar script:** `atualizar_cobranca.py`
   ```python
   import openpyxl
   
   def processar(xlsx_path):
       wb = openpyxl.load_workbook(xlsx_path, ...)
       # ... extração + cálculos
       return categoria_list, status_list, data_rows, extras
   
   def gerar_bloco(cat_list, status_list, data_rows, extras):
       return f"/* COB_DATA_START */\nconst COB_CAT={js_str(cat_list)};\n..."
   
   def atualizar_html(index_path, bloco):
       # mesmo padrão
   ```

2. **Integrar em `atualizar_tudo.py`:**
   ```python
   import atualizar_cobranca as acob
   
   # ... na main():
   canal, status, drows, extras = acob.processar(caminho)
   bloco_cob = acob.gerar_bloco(canal, status, drows, extras)
   acob.atualizar_html(INDEX_HTML, bloco_cob)
   ```

3. **Adicionar HTML:**
   ```html
   <div id="page-cobranca" class="page" style="display:none">
       <h2>Cobrança</h2>
       <div class="saude-filters">
           <!-- filtros -->
       </div>
       <!-- gráficos e cards -->
   </div>
   ```

4. **Implementar `renderCobranca()`:**
   ```javascript
   function renderCobranca(){
       // mesmo padrão que renderSac() etc
   }
   
   function initCobFilters(){
       // listeners para filtros
   }
   ```

---

## 🐛 Problemas Conhecidos & Workarounds

| Problema | Causa | Solução |
|----------|-------|---------|
| Datas com mês > 12 | Excel lê 07/12 como "July 12" (US) | `to_dt()` swaps dia↔mês se > hoje |
| SAC_ROWS vazio | SharePoint offline | Script continua com [AVISO], SAC fica com dados antigos |
| Filtros não repopulam | `refreshSacOptions()` não chamada | Verificar `addEventListener` em `initSacFilters()` |
| Gráfico fica preto | Muitos dados/browser lento | Limitar a 1 ano (`filterByYear()`) |

---

## 📖 Referências

- **Chart.js:** https://www.chartjs.org/ (gráficos)
- **openpyxl:** https://openpyxl.readthedocs.io/ (Excel em Python)
- **Python datetime:** https://docs.python.org/3/library/datetime.html

---

**Versão:** 1.0  
**Última atualização:** 08/06/2026
