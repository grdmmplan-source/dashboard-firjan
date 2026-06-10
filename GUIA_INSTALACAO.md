# Guia de Instalação e Uso — Dashboard Firjan

Para **qualquer pessoa** (não precisa ser programador) atualizar o dashboard.

---

## 🖥️ Pré-requisitos

Você precisa ter instalado (use as versões mais recentes):

1. **Python** (3.9+) → https://www.python.org/downloads/
   - ✅ Marque "Add Python to PATH" na instalação
   
2. **Git** → https://git-scm.com/download/win
   - ✅ Use as opções padrão

3. **GitHub Desktop** (opcional, mas fácil) → https://desktop.github.com/
   - Alternativa: use o Git via terminal

---

## 📥 Passo 1: Clonar o Repositório

### Opção A: Git Bash (terminal)
```bash
# Abra o Git Bash (clique direito em qualquer pasta → Git Bash Here)
git clone https://github.com/grdmmplan-source/dashboard-firjan.git
cd dashboard-firjan
```

### Opção B: GitHub Desktop
1. Abra GitHub Desktop
2. File → Clone Repository
3. Cole: `https://github.com/grdmmplan-source/dashboard-firjan.git`
4. Escolha a pasta de destino (Desktop, Documents, etc.)

---

## 🐍 Passo 2: Instalar Dependências Python

Abra o **Prompt de Comando** (cmd) ou **PowerShell**, navegue até a pasta do projeto e rode:

```bash
pip install openpyxl
```

(Demora ~30 segundos)

---

## 📂 Passo 3: Preparar os Arquivos de Dados

### Ativo (Retomada + Smart Factory)
1. Coloque os arquivos Excel em: `Arquivos\atualizaveis\`
   - `Mailing_Ativo_*.xlsx`
   - `Retorno_Ativo_*.xlsx`

### Receptivo
1. Coloque em: `Arquivos\atualizaveis\`
   - `Retorno_RECEPTIVO.xlsx` (com abas: BSales2, BASE DISCADOR 1/2, BASE CSAT)

### Saúde & SAC
- ✅ Já vêm da internet (Google Sheets + SharePoint públicos)

---

## 🚀 Passo 4: Rodar a Atualização

Na pasta do projeto, **clique 2x** em:

```
atualizar.bat
```

Ou via terminal:
```bash
python atualizar_tudo.py
```

**Você verá:**
```
==================================================
  ATUALIZADOR — Todas as Campanhas
==================================================

[0/8] Carregando de-para...
[1/8] Calculando Retomada da Trilha...
[2/8] Calculando Smart Factory...
...
[8/8] Concluido! index.html atualizado.
```

⏱️ Demora: ~30–60 segundos (depende da internet)

---

## 🌐 Passo 5: Ver o Dashboard Localmente

Abra o arquivo **`index.html`** no navegador (clique 2x).

Se quiser ver as **últimas mudanças**, faça um refresh:
- **Chrome/Firefox/Edge:** `Ctrl + F5` (limpa cache)

---

## 📤 Passo 6: Publicar no GitHub (Opcional)

Se quiser que as mudanças fiquem no link público:

### Na primeira vez:
1. **Peça acesso ao repositório**
   - Um admin adiciona sua conta como Collaborator no GitHub
   
2. Clique 2x em:
   ```
   publicar.bat
   ```
   
3. O Git vai pedir seu **login do GitHub**
   - Usuário: seu username do GitHub
   - Senha: seu Personal Access Token (se usar 2FA) ou senha
   - ✅ Depois fica salvo na máquina

### Próximas vezes:
Só clicar 2x em `publicar.bat` → automático!

---

## ⏰ Automação (Opcional)

### Agenda diária (Windows)
Para rodar o script **automaticamente todo dia às 9h:**

1. Abra **Agendador de Tarefas** (Pesquisa Windows → "agendar")
2. **Criar Tarefa Básica**
3. Nome: `Atualizar Dashboard Firjan`
4. Gatilho: Diário, às 9:00 AM
5. Ação:
   - Programa: `cmd.exe`
   - Argumentos: `/c cd "C:\caminho\dashboard-firjan" && atualizar.bat`

---

## 🧹 Limpar Dados (Teste)

Para ver como fica o dashboard **vazio** (útil pra testar):

```bash
limpar.bat
```

Isso remove todos os dados do `index.html` (não apaga nada no disco). Você pode rodar `atualizar.bat` de novo pra recarregar.

---

## 🐛 Erros Comuns

| Erro | Solução |
|------|---------|
| `ModuleNotFoundError: No module named 'openpyxl'` | Rode `pip install openpyxl` |
| `FileNotFoundError: Retorno_RECEPTIVO.xlsx` | Coloque o arquivo em `Arquivos\atualizaveis\` |
| `[AVISO] SAC nao atualizado (offline ou erro)` | Verifique a internet; o SAC usa SharePoint (que precisa de conexão) |
| Git pede senha a cada `publicar.bat` | Configure: `git config --global credential.helper wincred` |
| Dashboard não atualiza no navegador | Pressione `Ctrl + F5` pra limpar cache |

---

## 📊 Estrutura do Projeto (para referência)

```
dashboard-firjan/
├── index.html                    ← Arquivo do dashboard (abra isto!)
├── README.md                     ← Documentação completa
├── GUIA_INSTALACAO.md            ← Você está aqui
│
├── atualizar_tudo.py             ← Script principal (chama tudo)
├── atualizar_retomada.py         ← Ativo: Retomada
├── atualizar_smart.py            ← Ativo: Smart Factory
├── atualizar_saude.py            ← Saúde (Google Sheets)
├── atualizar_sac.py              ← SAC (SharePoint)
├── atualizar_receptivo.py        ← Receptivo (Excel + SharePoint)
│
├── atualizar.bat                 ← Clique 2x pra atualizar
├── publicar.bat                  ← Clique 2x pra publicar no GitHub
├── limpar.bat                    ← Clique 2x pra limpar dados
│
├── Arquivos/
│   ├── atualizaveis/             ← Coloque seus Excel aqui
│   │   ├── Retorno_RECEPTIVO.xlsx
│   │   └── Mailing_Ativo_*.xlsx
│   └── bases_apoio/
│       └── tab_de-para.xlsx      ← Mapeamentos (não mexer)
│
└── erros_sac.txt                 ← Relatório (gerado automaticamente)
```

---

## ✅ Checklist de Primeiro Uso

- [ ] Python instalado (verificar: `python --version` no cmd)
- [ ] Git instalado (verificar: `git --version` no cmd)
- [ ] Repositório clonado (`git clone ...` ou GitHub Desktop)
- [ ] openpyxl instalado (`pip install openpyxl`)
- [ ] Arquivos Excel em `Arquivos\atualizaveis\`
- [ ] Rodei `atualizar.bat` com sucesso
- [ ] Abri `index.html` no navegador ✅ Dashboard aparece!
- [ ] (Opcional) Pedi acesso ao GitHub e rodei `publicar.bat`

---

## 🔗 Links Úteis

- **Dashboard público:** https://grdmmplan-source.github.io/dashboard-firjan/
- **Repositório:** https://github.com/grdmmplan-source/dashboard-firjan
- **Python downloads:** https://www.python.org/downloads/
- **Git downloads:** https://git-scm.com/download/win
- **GitHub Desktop:** https://desktop.github.com/

---

## 📞 Precisa de Ajuda?

1. Verifique se todos os **Pré-requisitos** estão instalados
2. Leia o **Troubleshooting** acima
3. Abra uma **Issue** no GitHub (com screenshot do erro)
4. Contacte: **planejamento@grupoddm.com.br**

---

**Boa sorte! 🚀**
