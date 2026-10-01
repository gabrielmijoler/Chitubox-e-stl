# Shopee Upload — Contexto e Plano de Ação

Documento único de contexto do fluxo de upload em massa da Shopee (miniaturas RPG),
gerado a partir da `New Calculadora 3D - Copia.xlsx`.

**Estado atual**: 47 anúncios individuais já subidos e ativos; em curso a migração
para **anúncios agrupados com variações** (9 anúncios × até 10 variações).
Sem interface gráfica — tudo feito por código (`preencher_shopee.py`).

---

## 1. Ficheiros e paths

| Ficheiro | Path | Papel |
|---|---|---|
| Calculadora (fonte) | `D:\Arquivos Impressão 3D\Shopee\New Calculadora 3D - Copia.xlsx` | Folha `Pecas`, linhas **339–385** (47 peças da migração); col A=nome, K=Preço Final Shopee, L=Valor Ideal Shopee |
| Template Shopee | `D:\Downloads\Shopee_mass_upload_2026-09-30_basic_template.xlsx` | Template "basic" de criação em massa; folha `Modelo`; **nunca alterado** |
| Ficheiro gerado (47 individuais) | `D:\Downloads\Shopee_mass_upload_2026-09-30_preenchido.xlsx` | Já subido pelo utilizador (47 anúncios separados) |
| Ficheiro gerado (agrupado) | `D:\Downloads\Shopee_mass_upload_2026-09-30_agrupado.xlsx` | **Destino** — 9 anúncios com variações (a gerar) |
| Exemplo do utilizador | `D:\Downloads\shopee_upload_preenchido.xlsx` | Rascunho **nunca subido**; mostra 1 anúncio com 5 variações (linhas 7–11) |
| Script | `C:\Users\Gustavo\Documents\Repositorios\RoboChitu\preencher_shopee.py` | Fases 1–4 (ver §5) |
| Migração | `MIGRACAO.md` | Histórico da migração das 47 linhas |
| Ambiente | `.venv` (Python 3.14), pandas/openpyxl | Executar: `.\.venv\Scripts\python.exe preencher_shopee.py` |

---

## 2. Estrutura do template (folha `Modelo`)

- **Linhas 1–6 = metadados** (obrigatório/opcional, descrição, limites) — **nunca tocar**
- **Dados começam na linha 7**; 52 colunas **A..AZ**
- Template tem **5005 DVs** (5004 list + 1 whole na col N) — validação exige diff=0
- **Bug openpyxl**: XML tem `activePane="bottom_left"` → patch em memória para `bottomLeft` antes de abrir
- PowerShell 5.1 stripa aspas duplas em `-c` → usar here-strings `@'...'@`; para `"` em Python usar `chr(34)`

### Colunas-chave

| Col | Campo | Uso neste fluxo |
|---|---|---|
| A | Categoria | `101385` (Hobbies/Coleções/Figuras de Ação) |
| B | Nome do produto | Título genérico do **grupo** (corrigível à mão) |
| C | Descrição | Descrição-modelo com rótulo do lote |
| D | SKU principal | `MINI.LOTE{década}.{n}` (igual em todo o grupo) |
| **E** | Número de integração de variação | **Chave do agrupamento** — mesmo E = mesmo anúncio; obrigatório p/ produtos com variação |
| F | Nome da variação 1 | `Modelo` |
| G | Opção da variação | Nome da peça (limpo, **≤30 chars**, único no grupo) |
| H | Imagem por variação | **Manual** (fica vazio no script) |
| K | Preço (R$) | Preço da peça (max(K,L) da calculadora), 2 casas |
| L | Estoque | `99999` |
| M | SKU da variação | `MINI.`+slug do nome (≤100 chars, único no grupo) |
| P | GTIN | `00` = item sem GTIN (texto, const `GTIN_SEM`); checkbox "Item sem GTIN" no upload **a confirmar** |
| R..Z | Capa + galeria | **Manual** (fica vazio no script) |
| AA..AD | Peso/dims | 0.5 kg / 6 / 16 / 8 (cm) |
| AE | Canal 90016 | `Ligado` (DV `Ligado,Desativado`, obrigatório) |
| AF | Canal 90023 | **Vazia** ("Please do not edit this column") |
| **AG** | Prazo de postagem encomenda | **`3`** → ativa "Sob encomenda = Sim, prazo 3 dias" (faixa cat 101385 = "3 - 15"; sem DV, numérico) |
| AV | Item agrupável | `No` (única DV `allowBlank=false`, lista `Yes,No`) |
| AZ | Coluna de saída do sistema | Nunca preencher |

Regras do template (linha 6): título 2–120 chars; descrição 10–5000; preço 1.00–100000.00;
G/M ≤30/≤100 chars; **razão preço mais caro/mais barato dentro do anúncio ≤ 4.00**.

---

## 3. Regras de agrupamento (coluna E) — DECIDIDO

**Regra por valor de preço em décadas** (aprovada pelo utilizador):

1. Preço final = `max(K, L)` arredondado a 2 casas (ROUND_HALF_UP)
2. Grupo = mesma **década** de preço: `int(preço // 10) * 10` → 10–19,99 | 20–29,99 | 30–39,99 | ...
3. Ordenar por preço crescente dentro da década
4. **Máx. 10 variações por grupo**; década que exceder divide em lotes (Lote 1, Lote 2, ...)
5. E = 1..N sequencial; linhas do mesmo grupo gravadas **contíguas**
6. Exemplo do utilizador bate: Dark Wraiths 22,62–24,85 → grupo 20–29; dark riders 30,00–32,61 → grupo 30–39

### Grupos resultantes das 47 peças

| E | Grupo | Preços | Var. |
|---|---|---|---|
| 1 | R$10–19,99 Lote 1 | 17,15–19,40 | 10 |
| 2 | R$10–19,99 Lote 2 | 19,52–19,81 | 6 |
| 3 | R$20–29,99 Lote 1 | 20,27–23,20 | 10 |
| 4 | R$20–29,99 Lote 2 | 23,63–26,04 | 10 |
| 5 | R$20–29,99 Lote 3 | 29,33 | 1 |
| 6 | R$30–39,99 | 30,00–35,21 | 7 |
| 7 | R$40–49 | 42,31 (Archangel Azazel) | 1 |
| 8 | R$100–109 | 100,04 (Ooze Dragon) | 1 |
| 9 | R$120–129 | 127,11 (Witch King Mounted) | 1 |

Razão de preço por grupo ≤ 1,44 (limite Shopee 4,00) ✔

### Campos por grupo vs por linha

- **Iguais em todo o grupo**: A, B, C, D, E, F, AA, AB, AC, AD, AE, AF (vazia), AG=3, AV
- **Por linha (variação)**: G (nome da peça), K (preço), L (estoque), M (SKU)
- **Vazios (manuais)**: H, R..Z — **P** = `00` (texto, sem GTIN)

### Título e descrição genéricos (corrigíveis à mão)

- Título: `Miniaturas RPG - Peças {d} a {d+9} Reais[ Lote n] - D&D Pathfinder Tormenta 20`
- Descrição: `DESCRICAO_MODELO` com `@NOME@` → rótulo do lote (ex.: `Peças 10 a 19 Reais Lote 1`)

---

## 4. Decisões confirmadas (Q&A)

| Tema | Decisão |
|---|---|
| Interface | **Sem GUI** — "quero que seja feito por código" |
| Entrada de dados | Só coluna E (agrupamento) por **env ou const** → const `AGRUPAR_E = True` no script |
| Regra de E | Por **valores** (preço, décadas), máx. 10 variações |
| Nome/descrição do grupo | **Genérico** p/ o utilizador corrigir antes de subir |
| Âmbito | **Substituir os 47 atuais** já subidos (apagar antes de subir o novo) |
| F/G/M | **Derivar automático**: F=`Modelo`, G=nome da peça, M=`MINI.`+slug |
| Fotos H e R..Z | **Manuais** (script deixa vazio) |
| GTIN | P=`00` (texto) escrito pelo script; checkbox "Item sem GTIN" no upload **a confirmar** no 1º upload |
| Sob encomenda | **Sim, 3 dias** → AG=3 |
| Preço | `max(K,L)`, 2 casas (min 17.15 / médio 27.35 / máx 127.11) |
| Peso/dims | 0.5 kg; caixa 6×16×8 cm |
| Estoque | 99999 por variação |
| Canais | AE=`Ligado`; AF vazia; fiscais AH–AT vazios |

### Discrepâncias do exemplo do utilizador (rascunho, nunca subido) — usar valores provados

| Campo | Exemplo | Template provado (47 subidos) |
|---|---|---|
| AE/AF | `Ativado` | `Ligado` / vazio |
| AV | `Não` | `No` |
| P (GTIN) | `'00'` | vazio nos 47 subidos; agora **`'00'`** pelo script (decisão do utilizador) |
| E/F/G/H | preenchidos ✔ | referência da estrutura de variações |

---

## 5. Pipeline do script (`preencher_shopee.py`)

1. **Fase 1 — preços**: Excel COM read-only (`CalculateFullRebuild`), lê K/L das linhas 339–385; ficheiro fonte intacto
2. **Fase 2 — montar linhas**: nomes via openpyxl (read-only); preço `max(K,L)`; agrupamento por décadas (novo); campos por grupo/linha (novo)
3. **Fase 3 — gravar**: patch `activePane` em memória → openpyxl → validar valores contra DVs da planilha → salvar saída
4. **Fase 4 — validar pós-gravação**:
   - DVs no XML = (5005, 5004, 1); diff DVs template vs saída = 0
   - Linhas 1–6 intactas
   - Revalidação das listas no ficheiro final
   - Valores esperados por coluna (**incl. P=GTIN `00`**, texto)
   - **Novo**: mesmo E ⇒ mesmo B/C/D/F; ≤10 linhas por E; G≤30 único no grupo; M≤100 único; razão preço ≤4,00 por grupo; AG=3
   - **Aviso** (não falha) se H ou R..Z vazios → "fotos pendentes p/ preencher à mão"
5. **Teste Excel COM**: abrir ficheiro gerado read-only e confirmar que não pede reparação

---

## 6. Plano de ação

- [x] Contexto e plano de ação neste `.md`
- [x] Implementar modo agrupado em `preencher_shopee.py` (const `AGRUPAR_E`; décadas; E/F/G/M/D; AG=3; H/R..Z vazios; saída `_agrupado.xlsx`)
- [x] Alargar Fase 4 (validações de grupo + aviso de fotos)
- [x] Executar script → `Shopee_mass_upload_2026-09-30_agrupado.xlsx` + teste COM
      (Fases 1–4 OK; 9 anúncios × ≤10 var.; razão preço ≤4,0; Excel abriu sem reparação;
      470 células de foto vazias — pendente manual)
- [x] GTIN: escrever `00` (texto) na coluna P das 47 linhas (const `GTIN_SEM`) — Fases 1–4 OK
- [ ] **Utilizador**: apagar os 47 anúncios individuais da Shopee
- [ ] **Utilizador**: preencher fotos (H por variação; R..Z galeria) e corrigir títulos/descrições dos 9 grupos
- [ ] **Utilizador**: subir em massa; confirmar se o checkbox "Item sem GTIN" ainda é necessário
      (P já vem com `00`) e confirmar "Sob encomenda: Sim, 3 dias" após upload

### Pendências herdadas (menos relevantes com o ficheiro agrupado)

- Editar em Massa p/ corrigir AG dos 47 atuais → **desnecessário** se forem apagados e substituídos
- Checkbox GTIN manual → **a confirmar**: P já traz `00`; pode já não ser preciso
