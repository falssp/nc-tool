# [UL] Dashboard_Taxonomias

Dashboard operacional de taxonomias Unilever BR — Grasp x StormX.

**🔗 Acesso:** https://falssp.github.io/ul-dashboard-taxonomias/

Sincroniza tickets do Jira (projeto UL) automaticamente e os organiza em um painel visual hospedado no GitHub Pages.

---

## Como funciona

```
Jira (projeto UL)
      ↓  busca paginada encadeada (sem timeout)
Google Sheets (planilha pessoal)
      ↓  sync completo na primeira vez, incremental a cada 1h
Apps Script Web App (API de dados)
      ↓  responde requisições JSONP com os dados do PAINEL
GitHub Pages — index.html
      ↓  carrega dados via JSONP e renderiza o painel no browser
```

O sync é encadeado em triggers de 4.5 min cada — escala para qualquer volume de issues sem risco de timeout.

---

## Arquivos

| Arquivo | Onde vai | Descrição |
|---|---|---|
| `01_Config.gs` | Apps Script | Menu, credenciais, acionadores, popups de progresso |
| `02_Sync.gs` | Apps Script | Busca paginada no Jira, buffer, gravação RAW, etapas encadeadas |
| `03_Abas.gs` | Apps Script | Gravação e formatação de todas as abas da planilha |
| `04_WebApp.gs` | Apps Script | doGet (JSONP/JSON), doPost (exportar), getDados |
| `Painel.html` | Apps Script | Redirect para o GitHub Pages |
| `index.html` | GitHub Pages (raiz) | Painel completo — carrega dados via JSONP |

---

## Abas da planilha

| Aba | Descrição |
|---|---|
| 🏠 INICIO | Manual gerado automaticamente a cada sync |
| PAINEL | Fonte de dados do painel (pai × filho × veículo) |
| 📋 TABELA | Cópia formatada do PAINEL para consulta direta |
| RAW_PAI | Todas as Tarefas do Jira (campanhas) |
| RAW_FILHO | Todas as Subtarefas do Jira (tickets de trabalho) |
| DE_PARA | Dicionário de normalização — edite para mapear veículos e marcas |
| ⚠️ INCORRETOS | Issues com tipo inválido — corrija no Jira (aparece só quando há erros) |

---

## Contagem de tickets

O Jira conta todas as issues. O dashboard conta só o **trabalho executável**:

| Tipo | Jira | Dashboard |
|---|---|---|
| Subtarefas (com pai) | ✅ | ✅ conta |
| Solos (tarefa sem subtarefas) | ✅ | ✅ conta |
| Pais (tarefa com subtarefas) | ✅ | ❌ não conta — são agrupadores |
| Incorretos (Epic, Bug etc) | ✅ | ❌ não conta — aparecem em banner vermelho |

```
Jira total     = Pais + Subtarefas + Solos + Incorretos
Dashboard total = Subtarefas + Solos
```

A diferença entre os dois números é sempre `Pais + Incorretos` — e muda conforme novos tickets entram no Jira.

---

## Setup

1. Planilha Google Sheets → Extensões → Apps Script
2. Crie 4 arquivos `.gs` e 1 arquivo HTML com os nomes exatos acima
3. Cole o conteúdo de cada arquivo correspondente
4. Deploy → Nova implantação → Web App → Executar como: Eu → Qualquer pessoa
5. Menu 📋 Dashboard UL → 1. Configurar credenciais
6. Informe e-mail Atlassian + API Token do Jira + data de validade
7. O sync completo inicia automaticamente (~10–15 min para concluir)

### GitHub Pages

1. Suba o `index.html` na raiz do repositório
2. Settings → Pages → Source: main / root
3. URL ativa em ~1 min: `https://falssp.github.io/ul-dashboard-taxonomias/`

---

## Sync

- **Completo:** roda na primeira configuração ou quando forçado pelo menu (Menu → 3. Sincronizar). Busca todas as issues do projeto UL. Encadeado em 4 triggers independentes para evitar timeout.
- **Incremental:** roda automaticamente a cada 1h. Busca só issues atualizadas na última hora e faz merge com os dados existentes.

---

## DE_PARA — como mapear

Abra a aba `DE_PARA` na planilha e adicione uma linha:

| TIPO | TOKEN_JIRA | NOME_OFICIAL |
|---|---|---|
| MARCA | Dove Deos | Dove |
| VEICULO | YT | YouTube |

Se aparecer "⚠️ Não mapeado" no painel, o termo do título da issue não está no dicionário — adicione aqui.

---

## Token do Jira

O sistema envia e-mail automático quando faltam 10 dias para o token vencer.

Para renovar: [id.atlassian.com](https://id.atlassian.com) → Segurança → Tokens de API → Depois: Menu → 1. Configurar credenciais.

---

**Responsável técnico:** StormX Data & Tech  
**Projeto:** Unilever BR · Grasp x StormX
