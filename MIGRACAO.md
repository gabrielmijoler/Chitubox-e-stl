# Migração de dados → New Calculadora 3D

Registo da migração feita a **29/09/2026** e como a repetir se voltar a acontecer.

## Ficheiros envolvidos

| Papel | Caminho |
|---|---|
| Origem (dados do robô) | `C:\Users\Gustavo\OneDrive\Documentos\resultados_do_robo.xlsx` (folha `Planilha1`) |
| Destino (planilha de cálculo) | `D:\Arquivos Impressão 3D\Shopee\New Calculadora 3D - Copia.xlsx` (folha **`Pecas`**) |
| CSVs por rodada do robô | `<projeto>\resultados\resultados_AAAA-MM-DD_HH-MM-SS.csv` |
| Script de importação (permanente) | `importar_excel.py` |

## Estrutura da folha "Pecas"

- Cabeçalho: `A=Nome do arquivo`, `B=Altura (mm)`, `C=Largura (mm)`, `D=Peso (g)`, `E=Tempo (h)`, `F..N`=fórmulas de custo/preço
- Tabela `TabelaPecas` = `A1:N2206` (linhas vazias dentro da tabela já têm as fórmulas prontas)
- **Só escrever em A, D e E** — B/C são preenchidas à mão e F–N são fórmulas
- O script expande a tabela automaticamente se passar da linha 2206

## Regras de mapeamento

| Origem | Destino | Regra |
|---|---|---|
| `ficheiro` (`resultado_Aguara.png`) | `A Nome do arquivo` | remove prefixo `resultado_` e extensão `.png` → `Aguara` |
| `weight_g` (`10.039 g`) | `D Peso (g)` | remove unidade, `,`→`.`, número float (10.039) |
| `time` (`3h2m34s`) | `E Tempo (h)` | horas decimais com 2 casas → **3.04** |

- Cabeçalhos comparados em minúsculas sem espaços extra (aceita `ficheiro`/`file`, `weight_g`/`peso`, `time`/`tempo`)
- **Dedup pela chave (Nome + Peso + Tempo)** — se só o tempo diferir, o registo entra
- ⚠️ Aproveitar: dados introduzidos manualmente na folha usam às vezes a convenção `3.26` = 3h26 (não é decimal verdadeiro); os do robô são decimais (`3.44` = 3h26m). Pode gerar "duplicados" com tempos diferentes.

## Como correr

```powershell
cd C:\Users\Gustavo\Documents\Repositorios\RoboChitu
.\.venv\Scripts\python importar_excel.py          # CSV mais recente de resultados/ → Excel
.\.venv\Scripts\python importar_excel.py csv xlsx  # caminhos à escolha
```

- O ficheiro de destino (e a origem) **tem de estar fechado no Excel** antes de correr
- Antes de gravar é criado sempre `{nome}_backup_{AAAA-MM-DD_HH-MM-SS}.xlsx` na mesma pasta
- Escreve apenas em linhas **realmente vazias** (nunca sobrescreve, mesmo com buracos no meio)
- Excel recalcula as fórmulas ao abrir (`fullCalcOnLoad`)

## Migração executada a 29/09/2026 (resultado)

```
Origem: 54 linhas → 7 duplicados internos → 47 únicos
Destino: 337 peças → 384 peças (linhas 339..385 escritas)
Dedup vs destino: 0 (Sunathaer Caex entrou outra vez — 3.44 ≠ 3.26 existente)
Validação: 0 células alteradas nas linhas 1..338; fórmulas F–N intactas
Backup: New Calculadora 3D - Copia_backup_2026-09-29_19-28-19.xlsx
```

## Próxima migração (xlsx de origem)

O script one-off `migrar_resultados.py` foi usado e **apagado**. Para repetir, recriá-lo
importando as funções já testadas de `importar_excel.py`:

```python
# migrar_resultados.py — mínimo viável
import re, sys, shutil, time
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from importar_excel import (
    COL_NOME, COL_PESO, COL_TEMPO, FOLHA_DESTINO,
    mapear_colunas, extrair_registos, limpar,
    primeira_linha_vazia, proxima_linha_livre, expandir_tabela,
    converter_peso, converter_tempo,
)

ORIGEM  = r"C:\Users\Gustavo\OneDrive\Documentos\resultados_do_robo.xlsx"
DESTINO = r"D:\Arquivos Impressão 3D\Shopee\New Calculadora 3D - Copia.xlsx"
FOLHA_ORIGEM = "Planilha1"

def limpar_nome(v):
    n = re.sub(r"\s+", " ", limpar(v)).strip()
    n = re.sub(r"^resultado_", "", n, flags=re.I)
    return re.sub(r"\.png$", "", n, flags=re.I)

def main():
    df = pd.read_excel(ORIGEM, sheet_name=FOLHA_ORIGEM, dtype=object).dropna(how="all")
    df.columns = [re.sub(r"\s+", " ", str(c)).strip().lower() for c in df.columns]
    colunas = mapear_colunas(list(df.columns))
    registos, _ = extrair_registos(df, colunas)
    registos = list(dict.fromkeys((limpar_nome(n), p, t) for n, p, t in registos))  # dedup interno

    wb = load_workbook(DESTINO)
    ws = wb[FOLHA_DESTINO]
    existentes = {
        (str(ws.cell(row=r, column=COL_NOME).value).strip(),
         converter_peso(ws.cell(row=r, column=COL_PESO).value),
         converter_tempo(ws.cell(row=r, column=COL_TEMPO).value))
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=COL_NOME).value not in (None, "")
    }

    linha = primeira_linha_vazia(ws)
    inseridos = 0
    for chave in registos:
        if chave in existentes:
            print(f"  ⏭️ {chave[0]}: já existe")
            continue
        linha = proxima_linha_livre(ws, linha)
        ws.cell(row=linha, column=COL_NOME, value=chave[0])
        ws.cell(row=linha, column=COL_PESO, value=chave[1])
        ws.cell(row=linha, column=COL_TEMPO, value=chave[2])
        print(f"  ✅ {chave[0]} → linha {linha}")
        existentes.add(chave); linha += 1; inseridos += 1

    if not inseridos:
        print("Nada a gravar."); return 0

    expandir_tabela(ws, linha - 1)
    wb.calculation.fullCalcOnLoad = True
    backup = DESTINO.replace(".xlsx", f"_backup_{time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")
    shutil.copy2(DESTINO, backup)
    print("💾 Backup:", backup)
    wb.save(DESTINO)
    print(f"📊 Inseridos: {inseridos}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
```

**Checklist:** 1) guardar o bloco acima como `migrar_resultados.py` **na pasta do projeto** (assim
encontra `importar_excel.py`) → 2) fechar origem e destino no Excel → 3) correr primeiro numa
**cópia** do destino → 4) validar (contagem de peças, linhas novas, fórmulas em F) → 5) correr no
destino real (faz backup) → 6) apagar o script one-off se for uso único.
