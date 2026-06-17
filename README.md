# Dashboard Firjan Unificado

Dashboard de monitoramento em tempo real das campanhas de atendimento do Grupo Firjan, consolidando dados de múltiplas fontes (Ativo, Saúde, SAC, Receptivo, URA e Campanha).

🌐 **Link:** https://grdmmplan-source.github.io/dashboard-firjan/

---

## 📊 Módulos

| Módulo | Descrição | Dados | Filtros |
|--------|-----------|-------|---------|
| **Ativo** | Retomada da Trilha + Smart Factory + Cursos Técnicos Niterói | Excel local | Data, Campanha |
| **Saúde** | Promoção de Saúde | Google Sheets | Data, Ação, Médico |
| **SAC** | Serviço de Atendimento ao Cliente | SharePoint | Data, Regional, Entidade, Unidade, Canal, Tipo, Produto, Assunto |
| **Receptivo** | Receptivo (telefone, chat, email, redes sociais) | Excel local + SharePoint | Data, Canal, Regional, Entidade, Unidade |
| **URA** | Unidade de Resposta Audível | Excel local | Data, Tipo |
| **Campanha** | Atendimentos BSales2 com classificação "Sub" | Excel local + SharePoint | Data, Canal, Regional, Entidade, Unidade |

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
# Atualiza tudo (Retomada, Smart, Niterói, Saúde, SAC, Receptivo, URA)
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
├── atualizar_tudo.py               # Orquestrador principal
├── atualizar_retomada.py           # Ativo: Retomada da Trilha
├── atualizar_smart.py              # Ativo: Smart Factory
├── atualizar_cursos_niteroi.py     # Ativo: Cursos Técnicos Niterói
├── atualizar_saude.py              # Saúde (Google Sheets)
├── atualizar_sac.py                # SAC (SharePoint)
├── atualizar_receptivo.py          # Receptivo (Excel + SharePoint)
├── atualizar_ura.py                # URA (Excel local)
├── atualizar.bat                   # Clique 2x pra atualizar
├── publicar.bat                    # Clique 2x pra publicar no GitHub
├── limpar.bat                      # Clique 2x pra limpar dados
├── erros_sac.txt                   # Relatório de outliers de SAC
├── Arquivos/
│   ├── atualizaveis/
│   │   ├── Retorno_RECEPTIVO.xlsx        # BSales2, BASE DISCADOR 1/2, BASE CSAT
│   │   └── BASE_URA_*.xlsx               # Dados de URA
│   └── bases_apoio/
│       └── tab_de-para.xlsx              # Mapeamentos (status, entidade)
└── Arquivos/nao_atualizaveis/
    ├── Ativo_Retomada_da_Trilha/         # Mailing + Retorno Retomada
    ├── Ativo_Smart_Factory/              # Mailing + Retorno + WPP Smart Factory
    └── Ativo_Cursos_Técnicos_Unidade_Niterói/  # Mailing + Retorno + WPP Niterói
```

---

## 🔄 Fluxo de Atualização

```
atualizar_tudo.py
  ├→ atualizar_retomada.py          → extrai Retomada da Trilha
  ├→ atualizar_smart.py             → extrai Smart Factory
  ├→ atualizar_cursos_niteroi.py    → extrai Cursos Técnicos Niterói
  ├→ atualizar_saude.py             → baixa Google Sheets
  ├→ atualizar_sac.py               → baixa SharePoint + analisa SAC
  ├→ atualizar_receptivo.py         → lê Excel local + SharePoint (Redes/Autonomia)
  └→ atualizar_ura.py               → lê Excel local (BASE URA)
        ↓
  Gera blocos JS → injeita em index.html → carimba data/hora
```

**Frequência recomendada:** Rodar o script diariamente (ou conforme necessário).

---

## 📊 Fontes de Dados

### **Ativo (Retomada + Smart Factory + Cursos Técnicos Niterói)**

| Campanha | Mailing | Retorno | WhatsApp |
|----------|---------|---------|----------|
| Retomada da Trilha | `nao_atualizaveis/Ativo_Retomada_da_Trilha/Mailing_*.xlsx` | `Retorno_*.xlsx` | — |
| Smart Factory | `nao_atualizaveis/Ativo_Smart_Factory/Mailing_*.xlsx` | `Retorno_*.xlsx` | `Whatsapp_*.xlsx` |
| Cursos Técnicos Niterói | `nao_atualizaveis/Ativo_Cursos_Técnicos_Unidade_Niterói/Mailing_*.xlsx` | `Retorno_*.xlsx` | `Whatsapp_*.xlsx` |

> **Niterói:** o Mailing conta apenas inscrições (1 linha = 1 CPF). O card "Empresas na Base" muda para "CPFs Inscritos" ao filtrar esta campanha.

### **Saúde**
- **Fonte:** Google Sheets (anonimamente via CSV export)
- **Tipo:** Público (sem credenciais)

### **SAC**
- **Fonte:** SharePoint (anonimamente via link público)
- **Dados:** Data, Satisfação, Regional, Canal, Tipo, Produto, Assunto

### **Receptivo**
- **BSales2 + Discador 1 & 2 + BASE CSAT:** Excel local (`Retorno_RECEPTIVO.xlsx`)
- **Redes Sociais:** SharePoint (Fernanda Castro) — aba "Redes Sociais"; colunas: E=Assunto, F=Unidade, G=Entidade, H=Regional, J=Data, K=Canal
- **Autonomia e Renda:** SharePoint (Fernanda Castro) — múltiplas abas/meses

### **URA**
- **Fonte:** Excel local (`BASE_URA_*.xlsx`)
- **Dados:** Data, Tipo (Digital/Voz), Classificação (Retido/Resolvido/Abandono/Transferência), TMA, Motivo

### **Campanha**
- **Fonte:** mesmo `Retorno_RECEPTIVO.xlsx` (aba BSales2) + Redes Sociais SharePoint
- Subconjunto do Receptivo: linhas BSales2 cuja classificação começa com "Sub"
- Colunas adicionais: P = Segmento, Q = Pesquisa (Como conheceram)

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

| Indicador | Retomada | Smart Factory | Niterói |
|-----------|----------|---------------|---------|
| Empresas / Inscrições | 🏢 Empresas na Base | 🏢 Empresas na Base | 📋 CPFs Inscritos |
| Tentativas | Média por empresa | Média por empresa | Média por inscrição |
| Interessados | ✅ | ✅ | ✅ |
| Contato c/ Decisor | ✅ | ✅ | ✅ |
| Taxa de conversão | ✅ | ✅ | ✅ |
| Evolução diária | ✅ | ✅ | ✅ |
| Distribuição por status | ✅ | ✅ | ✅ |
| WhatsApp | — | Enviados / Respostas | Total Disparado / Entregues / Com Erro |
| Período trabalhado | ✅ | ✅ | ✅ |

> Ao filtrar "Todas", o card mostra "🏢 Empresas / Inscrições" (soma consolidada).

### **Saúde**
- Total de intenções, Agendadas, Canceladas, Pendentes
- **🏥 Taxa de Ocupação** — agendamentos reais / (dias trabalhados × capacidade diária por médico), baseado na aba AUX da planilha e coluna M (Data de Agendamento)
- Gráfico: Intenções por empresa

### **SAC**
- Total de SACs, Satisfeitos / Insatisfeitos, Tempo médio de resposta
- Rankings: Unidade, Canal, Produtos
- Comparativo anual (qtd × TMR por mês), CSAT por linha

### **Receptivo**
- Total, IAL, ICT, TMA, TME, IAR, CSAT Geral
- Top 5 Serviços Mais Procurados
- Filtros em cascata: Data, Canal, Regional, Entidade, Unidade
- Fontes do filtro de Unidade: BSales2 col W (Unidade2) + Redes Sociais col F
- Fontes do filtro de Regional: BSales2 col Y (Regional2) + Redes Sociais col H
- Fontes do filtro de Entidade: BSales2 col X (Entidade2) + Redes Sociais col G

### **URA**
- Total de acionamentos, Retidos, Resolvidos, Abandonos, Transferências para Humano
- TMA por classificação
- Motivos de transferência para humano

### **Campanha**
- Total de Atendimentos
- Gráfico Distribuição: Por Canal | Por Unidade | Por Segmento (col P) | Por Faixa Etária
- Gráfico "Como os participantes conheceram a Colônia de Férias" (col Q — Pesquisa)
- Indicadores de Qualidade: TMA, TME, ICT, IAR, IAL, CSAT

---

## 🎨 Modos de Visualização

- **Página (padrão):** Layout desktop responsivo
- **Painel (Kiosk):** Cards grandes para TV/monitor — inclui abas Ativo, Receptivo, Promoção Saúde, SAC, URA e Campanha
- **Videowall (2×2):** 4 módulos em grid

---

## 🛠️ Desenvolvimento

### Adicionar um novo filtro em cascata (ex: SAC)
1. Adicione a coluna no `.py` (índice + array)
2. Atualize `SAC_FILTERS` no JS com `{id, get, col, ph}`
3. O `refreshSacOptions()` vai automaticamente repopular em cascata

### Adicionar uma nova campanha Ativo
Ver [ARQUITETURA.md](ARQUITETURA.md) — seção "Adicionar Nova Campanha Ativo".

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

### "LIGAR aparece no gráfico de status"
→ Verifique se o `calcular_mailing_*()` está só contando linhas, sem ler a coluna Status. Apenas o Retorno alimenta status/tentativas.

---

## 📞 Contato & Suporte

- **Repositório:** https://github.com/grdmmplan-source/dashboard-firjan
- **Issues:** Abra uma issue no GitHub
- **Contato:** planejamento@grupoddm.com.br

---

## 📄 Licença

Interno — Grupo DDM

---

**Última atualização:** 17/06/2026 (rev 3)
