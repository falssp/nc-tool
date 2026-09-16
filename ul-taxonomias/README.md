# Dicionário de Taxonomias UL

**Projeto:** NC Tool — Unilever BR × StormX Data & Tech  
**Atualização:** Automática todo dia 1 do mês

---

## Acesso

🔗 **[https://falssp.github.io/nc-tool/ul-taxonomias/taxonomias-ul.html](https://falssp.github.io/nc-tool/ul-taxonomias/taxonomias-ul.html)**

---

## Arquivos

| Arquivo | Descrição |
|---|---|
| `taxonomias-ul.html` | Dicionário interativo (HTML standalone) |
| `atualizar-taxonomias.gs` | Script GAS — lê planilha UL e commita o HTML |

---

## O que é

Dicionário interativo de taxonomias da Unilever BR para uso interno da StormX.  
Busca, filtra e consulta todos os parâmetros de naming (Galileo e Freetext) com siglas, descrições e plataformas.

**566 opções · 22 campos · Dark mode · Responsivo**

---

## Fontes de dados

| Item | Valor |
|---|---|
| Planilha original UL | `1qIJIAz8UnYxHsRk1I5eRl1S9oPbewJhi7l7PjDnvgs0` |
| Aba | `Galielo e Freetext` |
| Planilha corp StormX | `1tYSRIxPXjsJNiLZ0f1bwbh_1MP74yXFc7kGrZOYL38w` |
| Planilha pessoal | `1hlLAFEkiU8bR67vCg337f9kcRGxhDeKg6XKwocE8J2Y` |

---

## Automação

| Script | Trigger | Função |
|---|---|---|
| `atualizar-taxonomias.gs` (corp) | Todo dia 1 às 08h | `atualizarTaxonomias` |

### Forçar atualização manual
1. Abrir planilha corp ou pessoal → Extensões → Apps Script
2. Selecionar `atualizarTaxonomias` → Executar
