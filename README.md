# Dashboard Firjan Unificado

Dashboard de monitoramento em tempo real das campanhas de atendimento do Grupo Firjan, consolidando dados de múltiplas fontes (Ativo, Saúde, SAC, Receptivo).

🌐 **Link:** https://grdmmplan-source.github.io/dashboard-firjan/

---

## 📊 Módulos

| Módulo | Descrição | Dados | Filtros |
|--------|-----------|-------|---------|
| **Ativo** | Retomada da Trilha + Smart Factory | Excel local | Data, Campanha |
| **Saúde** | Promoção de Saúde | Google Sheets | Data, Ação, Médico |
| **SAC** | Serviço de Atendimento ao Cliente | SharePoint | Data, Regional, Entidade, Canal, Tipo, Produto, Assunto |
| **Receptivo** | Receptivo (telefone, chat, email, redes) | Excel local + SharePoint | Data, Canal, Entidade, Unidade |

---

## 🚀 Início Rápido

### 1. **Clonar o repositório**
```bash
git clone https://github.com/grdmmplan-source/dashboard-firjan.git
cd dashboard-firjan
```

### 2. **Instalar dependências**
```bash
pip install openpyxl
# Git já deve estar instalado
```

### 3. **Atualizar dados locais**
```bash
# Atualiza tudo (Retomada, Smart, Saúde, SAC, Receptivo)
atualizar.bat

# Ou versão com auto-publicação no GitHub
atualizar_e_publicar.bat
```

### 4. **Publicar no GitHub Pages**
```bash
publicar.bat
```

Site atualiza em ~1 minuto em https://grdmmplan-source.github.io/dashboard-firjan/

---

## 📁 Estrutura de Arquivos

```
dashboard-firjan/
├── index.html                      # Dashboard (HTML/JS puro)
├── atualizar_tudo.py              # Orquestrador principal
├── atualizar_retomada.py          # Ativo: Retomada da Trilha
├── atualizar_smart.py             # Ativo: Smart Factory
├── atualizar_saude.py             # Saúde (Google Sheets)
├── atualizar_sac.py               # SAC (SharePoint)
├── atualizar_receptivo.py         # Receptivo (Excel + SharePoint)
├── atualizar.bat                  # Batch: rodar atualizar_tudo.py
├── publicar.bat                   # Batch: git push (publica no GitHub)
├── limpar.bat                     # Batch: remove todos os dados do HTML
├── erros_sac.txt                  # Relatório de outliers de SAC
├── Arquivos/
│   ├── atualizaveis/
│   │   ├── Retorno_RECEPTIVO.xlsx         # BSales2, BASE DISCADOR 1/2, BASE CSAT
│   │   └── Retorno_Ativo_B+P_*            # Retomada da Trilha (Mailing + Retorno)
│   └── bases_apoio/
│       └── tab_de-para.xlsx               # Mapeamentos (status, entidade)
└── README.md                       # Esta documentação
```

---

## 🔄 Fluxo de Atualização

```
atualizar_tudo.py
  ├→ atualizar_retomada.py   → extrai Retomada da Trilha
  ├→ atualizar_smart.py      → extrai Smart Factory
  ├→ atualizar_saude.py      → baixa Google Sheets
  ├→ atualizar_sac.py        → baixa SharePoint + analisa SAC
  └→ atualizar_receptivo.py  → lê Excel local + SharePoint (Redes/Autonomia)
        ↓
  Gera blocos JS → injeita em index.html → carimba data/hora
```

**Frequência recomendada:** Rodas o script diariamente (ou conforme necessário).

---

## 📊 Fontes de Dados

### **Ativo (Retomada + Smart Factory)**
- **Mailing:** `C:\Users\...\Arquivos\atualizaveis\Mailing_Ativo_*.xlsx`
- **Retorno:** `C:\Users\...\Arquivos\atualizaveis\Retorno_Ativo_*.xlsx`
- **Tipo:** Excel local
- **Atualização:** Manual (copiar arquivo)

### **Saúde**
- **Fonte:** Google Sheets (anonimamente via CSV export)
- **URL:** `https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv`
- **Tipo:** Público (sem credenciais)

### **SAC**
- **Fonte:** SharePoint (anonimamente via link público)
- **Tipo:** Público (sem credenciais)
- **Dados:**
  - Data de atendimento / Finalização → Tempo Médio
  - Satisfação → KPI de satisfação
  - Regional, Canal, Tipo, Produto, Assunto → Filtros + Rankings

### **Receptivo**
- **BSales2:** Excel local (`Retorno_RECEPTIVO.xlsx`)
- **Discador 1 & 2:** Mesmo arquivo Excel, abas diferentes
- **BASE CSAT:** Mesmo arquivo Excel
- **Redes Sociais:** SharePoint (anonimamente)
- **Autonomia e Renda:** SharePoint (anonimamente)

---

## 🔐 Credenciais & Segurança

### ✅ **Sem credenciais hardcoded**
- SharePoint: Link público com `?download=1`
- Google Sheets: CSV export público
- Arquivos locais: Leitura do disco
- GitHub: Credenciais do Git em cache local (Windows Credential Manager)

### 🔑 **Para publicar no GitHub**
1. Adicione sua conta como **Collaborator** no repo (Settings → Collaborators)
2. No primeiro `publicar.bat`, Git pede login → entra no Credential Manager
3. Daí em diante, push automático

---

## 📈 Indicadores Principais por Módulo

### **Ativo**
- Tentativas (ligações)
- Interessados
- Contatos c/ Decisor
- Taxa de conversão
- Gráfico de evolução diária
- Distribuição por status

### **Saúde**
- Total de intenções
- Agendadas
- Canceladas
- Pendentes
- Gráfico: Intenções por empresa

### **SAC**
- Total de SACs
- Satisfeitos / Insatisfeitos
- Tempo médio de resposta (horas/dias)
- Rankings: Unidade, Canal, Produtos
- Comparativo anual (qtd × TMR por mês)
- CSAT por linha
- Distribuição por satisfação

### **Receptivo**
- Total de atendimentos
- IAL (Índice de Abandono)
- ICT (Atendidas em 30s)
- TMA (Tempo Médio Atendimento)
- TME (Tempo Médio Espera)
- IAR (Atendimentos Registrados)
- CSAT Geral
- Respostas CSAT
- Top 5 Serviços Mais Procurados

---

## 🎨 Modos de Visualização

- **Página (padrão):** Layout desktop responsivo
- **Painel (Kiosk):** Cards grandes para TV/monitor
- **Videowall (2×2):** 4 módulos em grid

---

## 🛠️ Desenvolvimento

### Adicionar um novo filtro em cascata (ex: SAC)
1. Adicione a coluna no `.py` (índice + array)
2. Atualize `SAC_FILTERS` no JS com `{id, get, col, ph}`
3. O `refreshSacOptions()` vai automaticamente repopular em cascata

### Adicionar um novo módulo (ex: Atendimento)
1. Crie `atualizar_atendimento.py` seguindo o padrão
2. Importe em `atualizar_tudo.py`
3. Adicione o bloco JS injeção
4. Crie a seção HTML com `<div id="page-atendimento">` e filtros
5. Implemente `renderAtendimento()`

---

## 🐛 Troubleshooting

### "Arquivo não encontrado: Retorno_RECEPTIVO.xlsx"
→ Coloque o arquivo em `Arquivos\atualizaveis\`

### "Conexão recusada ao SharePoint"
→ Verifique internet e se o link é público

### "Git: credenciais inválidas"
→ Rode `git config --global credential.helper wincred` e tente de novo

### "Dashboard não atualiza após rodar o script"
→ Limpe o cache do navegador (Ctrl+F5)

---

## 📞 Contato & Suporte

- **Repositório:** https://github.com/grdmmplan-source/dashboard-firjan
- **Issues:** Abra uma issue no GitHub
- **Contato:** planejamento@grupoddm.com.br

---

## 📄 Licença

Interno — Grupo DDM

---

**Última atualização:** 08/06/2026
