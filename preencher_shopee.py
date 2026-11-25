#!/usr/bin/env python3
"""
Gera a planilha de criação em massa da Shopee (template "basic") com as peças
fatiadas no run atual.

Cada peça do CSV de resultados vira UMA variação
(coluna G = nome do modelo, K = preço). Cada execução cria um ficheiro novo em
pasta shopee/ do projeto (Shopee_mass_upload_<data-hora>.xlsx); os anteriores não são tocados.

As variações são agrupadas em tantos anúncios quantos forem precisos para respeitar a regra
da Shopee: preço mais caro <= 4 x preço mais barato dentro de cada anúncio.

Fases:
  1. Preços  - localizar as peças na calculadora e ler K/L (Excel COM, read-only)
  2. Linhas  - montar em memória (variações, preços, canais)
  3. Gravar  - patch activePane -> openpyxl -> validar listas -> salvar em ficheiro novo
  4. Validar  - reabrir ficheiro final (DV counts, diff, valores, aviso de preço)

Uso: .\\.venv\\Scripts\\python.exe preencher_shopee.py [caminho do CSV]
     (sem argumento: usa o CSV mais recente de resultados/)
"""

from __future__ import annotations

import io
import re
import subprocess
import sys
import time
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import importar_excel as ie
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if isinstance(sys.stderr, io.TextIOWrapper):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# --- CONFIGURAÇÃO ----------------------------------------------------------
CALCULADORA = Path(r"D:\Arquivos Impressão 3D\Shopee\New Calculadora 3D.xlsx")
TEMPLATE = Path(r"D:\Downloads\Shopee_mass_upload_2026-09-30_basic_template.xlsx")
PASTA_SAIDA = Path(__file__).resolve().parent / "shopee"
PASTA_RESULTADOS = Path(__file__).resolve().parent / "resultados"

FOLHA_CALC = "Pecas"
FOLHA_TEMPLATE = "Modelo"
LINHA_DADOS_TEMPLATE = 7                  # primeira linha de dados do template

# --- Agrupamento em anúncios (coluna E = 1, 2, ... ; razão preço <= RAZAO_PRECO_MAX) ---
MAX_VARIACOES = 50                        # acima disto só avisa (limite real a confirmar na Shopee)
PRE_ENCOMENDA = 3                         # AG: dias de postagem p/ encomenda (faixa cat 101385 = 3..15)
NOME_VAR = "Modelo"                       # F: nome da variação 1
NOME_TITULO_RUN = "Kit de Miniaturas"     # título genérico do anúncio (corrigível à mão)

CATEGORIA = "101385"                      # Hobbies e Coleções/Itens Colecionáveis/Figuras de Ação
ESTOQUE = 99999
TERMINACAO_PRECO = Decimal("0.90")
          # preço final = parte inteira + 0,90 (ex.: 21,15 -> 21,90)
PESO_KG = 0.5                             # peso fixo do pacote (kg)
COMPRIMENTO, LARGURA, ALTURA = 6, 16, 8   # caixa 6x16x8 cm, ordem escrita pelo utilizador
CANAL_AE = "Ligado"                       # Shopee Xpress CPF - DV "Ligado,Desativado"
ITEM_AGRUPAVEL = "No"                     # DV "Yes,No" com allowBlank=false
GTIN_SEM = "00"                           # P (GTIN): "00" = item sem GTIN (texto)
# AF (Retirada pelo Comprador) fica VAZIA - template: "Please do not edit this column"

PREFIXO_TITULO = "Miniaturas RPG - "
SUFIXO_TITULO = " - D&D Pathfinder Tormenta 20"
LIMITE_TITULO = 120
LIMITE_G = 30                             # coluna G: 1 a 30 caracteres
LIMITE_M = 100                            # coluna M: menos de 100 caracteres
RAZAO_PRECO_MAX = 4.00                    # Shopee: preço mais caro <= 4 x preço mais barato

DESCRICAO_MODELO = (
    "ATENÇÃO\n"
    "\n"
    "Caso não encontre uma miniatura que necessite, ou deseje personalizar com uma "
    "pintura customizada, entre em contato pelo chat, que estaremos felizes em te auxiliar.\n"
    "\n"
    "  \n"
    "\n"
    "Miniaturas na escala 32mm:\n"
    "@NOME@\n"
    "\n"
    "IMPORTANTE: Produto enviado com a base e peças avulsas separadas para maior "
    "segurança no envio. A impressão recebe tratamento em luz UV e já vai com primer "
    "preto aplicado.\n"
    "Somos uma marca preocupada em entregar o melhor produto para todos os compradores. "
    "Nosso foco está em miniaturas de RPG, criação de personagens customizados, cenários "
    "e terrenos para enriquecer a experiência de suas mesas, seja em campos de batalhas, "
    "praças das cidades, esconderijos de ladrões, hordas de zumbis, monstros aterrorizantes "
    "e mercadores em apuros."
)

# Colunas do template (Modelo)
COL = {
    "categoria": 1, "nome": 2, "descricao": 3, "sku_principal": 4,
    "e_agrup": 5, "f_var": 6, "g_opcao": 7, "preco": 11, "estoque": 12,
    "m_sku": 13, "gtin": 16, "peso": 27, "comprimento": 28, "largura": 29,
    "altura": 30, "canal_ae": 31, "ag": 33, "agrupavel": 48,
}

# --- FASE 1: preços via Excel COM ------------------------------------------
PS_TEMPLATE = """
$ErrorActionPreference = 'Stop'
$planilha = '@@PLANILHA@@'
$linhas = @(@@LINHAS@@)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
try {
    $wb = $excel.Workbooks.Open($planilha, 0, $true)
    $excel.CalculateFullRebuild()
    $ws = $wb.Worksheets.Item('Pecas')
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    foreach ($r in $linhas) {
        $k = $ws.Cells.Item($r, 11).Value2
        $l = $ws.Cells.Item($r, 12).Value2
        $o = $ws.Cells.Item($r, 15).Value2
        $ks = if ($null -eq $k -or $k -is [string]) { '' } else { ([double]$k).ToString('R', $inv) }
        $ls = if ($null -eq $l -or $l -is [string]) { '' } else { ([double]$l).ToString('R', $inv) }
        $os = if ($null -eq $o -or $o -is [string]) { '' } else { ([double]$o).ToString('R', $inv) }
        Write-Output ('{0}|{1}|{2}|{3}' -f $r, $ks, $ls, $os)
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
"""


def obter_precos_com(linhas: list[int]) -> dict[int, tuple[float | None, float | None, float | None]]:
    """Abre a New Calculadora read-only, recalcula e devolve {linha: (K, L)} para as linhas pedidas."""
    ps = (
        PS_TEMPLATE
        .replace("@@PLANILHA@@", str(CALCULADORA))
        .replace("@@LINHAS@@", ",".join(str(r) for r in linhas))
    )
    print(f"  Excel COM: abrir read-only, CalculateFullRebuild, ler K/L de {len(linhas)} linha(s)...", flush=True)
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            "Excel COM falhou (verifica se a Excel está licenciada e a planilha fechada):\n"
            + (proc.stderr or proc.stdout)
        )

    dados: dict[int, tuple[float | None, float | None, float | None]] = {}
    for linha in proc.stdout.splitlines():
        m = re.match(r"^(\d+)\|([^|]*)\|([^|]*)\|(.*)$", linha.strip())
        if not m:
            continue
        r = int(m.group(1))
        k = float(m.group(2)) if m.group(2).strip() else None
        l = float(m.group(3)) if m.group(3).strip() else None
        o = float(m.group(4)) if m.group(4).strip() else None
        dados[r] = (k, l, o)

    faltam = [r for r in linhas if r not in dados]
    if faltam:
        raise RuntimeError(f"COM não devolveu as linhas {faltam}. stdout:\n{proc.stdout}")
    return dados


def localizar_linhas(nomes: list[str]) -> tuple[dict[str, int], list[str]]:
    """Procura cada nome na coluna A da folha Pecas. Se houver várias ocorrências, vale a última (mais recente)."""
    wb = load_workbook(CALCULADORA, read_only=True, data_only=True)
    ws = wb[FOLHA_CALC]
    procurados = set(nomes)
    encontrados: dict[str, int] = {}
    for linha in ws.iter_rows(min_col=1, max_col=1):
        cel = linha[0]
        if cel.value is None:
            continue
        nome = str(cel.value).strip()
        if nome in procurados and cel.row is not None:
            encontrados[nome] = cel.row
    wb.close()
    faltam = [n for n in nomes if n not in encontrados]
    return encontrados, faltam


# --- LEITURA DO RUN -----------------------------------------------------------
def resolver_csv(caminho: str | Path | None) -> Path:
    """O CSV indicado, ou o mais recente de resultados/."""
    if caminho:
        p = Path(caminho)
        if not p.is_file():
            raise FileNotFoundError(f"CSV não encontrado: {p}")
        return p
    if not PASTA_RESULTADOS.is_dir():
        raise FileNotFoundError(f"Pasta de CSVs não encontrada: {PASTA_RESULTADOS}")
    candidatos = sorted(PASTA_RESULTADOS.glob("resultados_*.csv"), key=lambda p: p.stat().st_mtime)
    if not candidatos:
        raise FileNotFoundError(f"Nenhum CSV 'resultados_*.csv' encontrado em {PASTA_RESULTADOS}")
    return candidatos[-1]


def ler_run(caminho: Path) -> tuple[list[str], list[str]]:
    """Nomes das peças do run com peso e tempo válidos (os mesmos que o importar_excel aceita).
    Devolve (nomes, ignoradas)."""
    df = ie.ler_csv(caminho)
    colunas = ie.mapear_colunas(list(df.columns))
    nomes: list[str] = []
    ignoradas: list[str] = []
    for _, linha in df.iterrows():
        nome = ie.limpar(linha[colunas["nome"]])
        if not nome:
            continue
        peso = ie.converter_peso(linha[colunas["peso"]])
        tempo = ie.converter_tempo(linha[colunas["tempo"]])
        if peso is None or tempo is None:
            ignoradas.append(nome)
            continue
        if nome not in nomes:
            nomes.append(nome)
    return nomes, ignoradas


# --- FASE 2: montar linhas ---------------------------------------------------
def limpar_nome_titulo(nome: str) -> str:
    """Nome do modelo para G/M/descrição: tira o sufixo das pastas do kit (_32mm_...),
    separa CamelCase e troca '_' por espaço."""
    n = re.sub(r"_?32mm_.*$", "", str(nome), flags=re.IGNORECASE)
    n = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", n)
    n = re.sub(r"\s+", " ", n.replace("_", " ")).strip()
    return n.strip("| ").strip()


def construir_titulo_run(n: int = 1, total: int = 1) -> str:
    nome = NOME_TITULO_RUN if total == 1 else f"{NOME_TITULO_RUN} {n}"
    titulo = PREFIXO_TITULO + nome + SUFIXO_TITULO
    if not 2 <= len(titulo) <= LIMITE_TITULO:
        raise ValueError(f"título inválido ({len(titulo)} car.): {titulo}")
    return titulo


def construir_descricao(nomes: list[str]) -> str:
    lista = "\n".join(f"- {limpar_nome_titulo(n)}" for n in nomes)
    desc = DESCRICAO_MODELO.replace("@NOME@", lista)
    if not 10 <= len(desc) <= 5000:
        raise ValueError(f"descrição inválida ({len(desc)} car.)")
    return desc


def preco_shopee(k: float | None, l: float | None, r: int) -> float:
    candidatos = [v for v in (k, l) if v is not None]
    if not candidatos:
        raise ValueError(f"linha {r}: K e L vazios - sem preço")
    p = Decimal(str(max(candidatos))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    p = Decimal(int(p)) + TERMINACAO_PRECO      # sempre ,90: 21,15 -> 21,90 | 19,86 -> 19,90
    if not Decimal("1.00") <= p <= Decimal("100000.00"):
        raise ValueError(f"linha {r}: preço {p} fora de 1.00..100000.00")
    return float(p)


def agrupar_por_razao(pecas: list[dict]) -> list[list[dict]]:
    """Agrupa as peças em anúncios com preço mais caro <= RAZAO_PRECO_MAX x mais barato.
    Guloso: cada anúncio começa na peça mais barata que ainda sobra; a peça que passa do limite abre o seguinte."""
    limite = Decimal(str(RAZAO_PRECO_MAX))
    grupos: list[list[dict]] = []
    for p in sorted(pecas, key=lambda p: (p["preco"], p["nome"])):
        preco = Decimal(str(p["preco"]))
        if grupos and preco <= limite * Decimal(str(grupos[-1][0]["preco"])):
            grupos[-1].append(p)
        else:
            grupos.append([p])
    return grupos


def ler_taxas_config() -> dict:
    """Lê da aba Config da calculadora as faixas de comissão da Shopee e o MEI/peça."""
    wb = load_workbook(CALCULADORA, read_only=True, data_only=True)
    ws = wb["Config"]

    def b(linha: int) -> Decimal:
        v = ws.cell(row=linha, column=2).value
        if not isinstance(v, (int, float)):
            raise ValueError(f"Config!B{linha} vazio ou não numérico ({v!r}) — guarde a calculadora no Excel")
        return Decimal(str(v))

    cfg = {
        "mei": b(42),
        # (limite superior, comissão %, taxa fixa); a última faixa não tem limite
        "faixas": [(b(46), b(47), b(48)), (b(49), b(50), b(51)), (b(52), b(53), b(54)), (None, b(55), b(56))],
    }
    wb.close()
    return cfg


def preco_kit_bruto(soma: Decimal, cfg: dict) -> Decimal:
    """Mesma lógica da coluna K: primeira faixa cujo preço calculado não ultrapassa o limite dela."""
    for limite, comissao, fixa in cfg["faixas"]:
        p = (soma + cfg["mei"] + fixa) / (1 - comissao / 100)
        if limite is None or p <= limite:
            return p
    raise AssertionError("inalcançável: a última faixa não tem limite")


def calcular_kit(precos_sem_taxas: list[float]) -> dict:
    """Preço do kit com todas as peças numa só venda, pela tabela de faixas da aba Config:
    soma os preços sem as taxas (coluna O), acrescenta MEI e taxa fixa da faixa e divide por
    (1 - comissão da faixa). Arredonda como as peças: parte inteira + ,90."""
    cfg = ler_taxas_config()
    soma = sum(Decimal(str(p)) for p in precos_sem_taxas)
    bruto = preco_kit_bruto(soma, cfg)
    # parte inteira + ,90 nunca passa de um limite terminado em ,99, logo continua na faixa calculada
    preco = Decimal(int(bruto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))) + TERMINACAO_PRECO
    if not Decimal("1.00") <= preco <= Decimal("100000.00"):
        raise ValueError(f"preço do kit {preco} fora de 1.00..100000.00")
    return {"soma": float(soma), "kit": float(preco)}


def construir_titulo_kit(n: int) -> str:
    titulo = PREFIXO_TITULO + f"Kit Completo {n} Miniaturas" + SUFIXO_TITULO
    if not 2 <= len(titulo) <= LIMITE_TITULO:
        raise ValueError(f"título do kit inválido ({len(titulo)} car.): {titulo}")
    return titulo


def slug_sku(nome: str, prefixo: str = "MINI.") -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", ".", limpar_nome_titulo(nome)).strip(".").upper()
    sku = prefixo + slug
    if not 1 <= len(sku) <= LIMITE_M:
        raise ValueError(f"SKU de variação inválido ({len(sku)} car.): {sku}")
    return sku


def opcao_variacao(nome: str, usadas: set[str]) -> str:
    """G = nome do modelo limpo, ≤30 chars, único dentro do anúncio."""
    g = limpar_nome_titulo(nome)[:LIMITE_G].rstrip()
    base, n = g, 2
    while g in usadas:
        sufixo = f" {n}"
        g = base[:LIMITE_G - len(sufixo)].rstrip() + sufixo
        n += 1
    if not 1 <= len(g) <= LIMITE_G:
        raise ValueError(f"opção de variação inválida: {g!r}")
    return g


# --- FASE 3: template + validação das listas ---------------------------------
def carregar_template() -> Workbook:
    """Lê o ZIP do template, corrige activePane em memória e abre com openpyxl."""
    zin = zipfile.ZipFile(TEMPLATE)
    buf = io.BytesIO()
    corrigidos = 0
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/"):
                txt = dados.decode("utf-8")
                if 'activePane="bottom_left"' in txt:
                    txt = txt.replace('activePane="bottom_left"', 'activePane="bottomLeft"')
                    corrigidos += 1
                    dados = txt.encode("utf-8")
            zout.writestr(item, dados)
    zin.close()
    buf.seek(0)
    wb = load_workbook(buf)
    print(f"  Template aberto (activePane corrigido em {corrigidos} folhas)")
    return wb


def resolver_formula1(wb, formula1: str) -> list[str] | None:
    f = (formula1 or "").strip()
    if len(f) >= 2 and f.startswith('"') and f.endswith('"'):
        return [v.strip() for v in f[1:-1].split(",")]
    m = re.match(r"^'?([^'!]+)'?!(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$", f)
    if m:
        folha, ref1, ref2 = m.groups()
        folha = folha.replace("$", "")
        if folha not in wb.sheetnames:
            return None
        ws2 = wb[folha]
        intervalo = f"{ref1.replace('$', '')}:{ref2.replace('$', '')}"
        valores = []
        for linha in ws2[intervalo]:
            for cel in linha:
                if cel.value not in (None, ""):
                    valores.append(str(cel.value))
        return valores
    return None


def validar_contra_listas_da_planilha(ws, linha_final: int, linha_inicial: int = 7) -> list[str]:
    """Confere valores das colunas com DV type='list'. Devolve problemas."""
    wb = ws.parent
    problemas: list[str] = []
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list":
            continue
        permitidos = resolver_formula1(wb, dv.formula1)
        if permitidos is None:
            problemas.append(f"DV list não interpretável: sqref={dv.sqref} formula1={dv.formula1}")
            continue
        set_perm = set(permitidos)
        for token in str(dv.sqref).split():
            min_c, min_r, max_c, max_r = range_boundaries(token)
            if min_c is None or min_r is None or max_c is None or max_r is None:
                continue  # faixa sem limites (não acontece nas validações do template)
            for r in range(max(min_r, linha_inicial), min(max_r, linha_final) + 1):
                for c in range(min_c, max_c + 1):
                    v = ws.cell(row=r, column=c).value
                    if v in (None, ""):
                        continue
                    if str(v).strip() not in set_perm:
                        problemas.append(
                            f"linha {r}, coluna {get_column_letter(c)}: valor {v!r} "
                            f"fora da lista (ex.: {sorted(set_perm)[:3]})"
                        )
    return problemas


def dv_set(wb) -> set:
    s = set()
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            s.add((ws.title, str(dv.sqref), str(dv.formula1), str(dv.type), bool(dv.allowBlank)))
    return s


def contar_dv_xml(caminho: Path) -> tuple[int, int, int]:
    z = zipfile.ZipFile(caminho)
    total = nlist = nwhole = 0
    for nome in z.namelist():
        if not nome.startswith("xl/worksheets/"):
            continue
        xml = z.read(nome).decode("utf-8", "replace")
        for tag in re.findall(r"<dataValidation\b[^>]*?>", xml):
            total += 1
            if 'type="list"' in tag:
                nlist += 1
            elif 'type="whole"' in tag:
                nwhole += 1
    z.close()
    return total, nlist, nwhole


def caminho_saida_novo() -> Path:
    """Ficheiro novo a cada run: Shopee_mass_upload_<data-hora>.xlsx (nunca sobrescreve)."""
    base = f"Shopee_mass_upload_{time.strftime('%Y-%m-%d_%H-%M-%S')}"
    caminho = PASTA_SAIDA / f"{base}.xlsx"
    n = 2
    while caminho.exists():
        caminho = PASTA_SAIDA / f"{base}_{n}.xlsx"
        n += 1
    return caminho


# --- PRINCIPAL ----------------------------------------------------------------
def executar(caminho_csv: str | Path | None = None) -> int:
    print("=== RUN - ler o CSV de resultados ===")
    try:
        csv_path = resolver_csv(caminho_csv)
        nomes, ignoradas = ler_run(csv_path)
    except (FileNotFoundError, ValueError) as erro:
        print(f"❌ {erro}")
        return 1
    print(f"📄 CSV: {csv_path} | peças válidas no run: {len(nomes)}")
    for nome in ignoradas:
        print(f"  ⚠️ {nome}: sem peso/tempo válidos — fora da planilha")
    if not nomes:
        print("❌ Nenhuma peça válida neste run.")
        return 1

    if ie.calculadora_aberta_no_excel(CALCULADORA):
        print("❌ A calculadora está aberta no Excel. Guarde (Ctrl+S) e feche-a, e corra de novo.")
        print(f"   (Se o Excel já estiver fechado, apague o ficheiro '~$...' na pasta: {CALCULADORA.parent})")
        return 1

    print("\n=== FASE 1 - localizar na calculadora e ler preços (Excel COM) ===")
    encontrados, faltam = localizar_linhas(nomes)
    for nome in faltam:
        print(f"  ⚠️ '{nome}' não está na folha Pecas da calculadora — fora da planilha (correr importar_excel?)")
    nomes_ok = [n for n in nomes if n in encontrados]
    if not nomes_ok:
        print("❌ Nenhuma peça do run foi encontrada na calculadora.")
        return 1
    try:
        dados = obter_precos_com([encontrados[n] for n in nomes_ok])
    except Exception as erro:  # noqa: BLE001
        print(f"❌ {erro}")
        return 1

    print("\n=== FASE 2 - montar o anúncio único ===")
    pecas: list[dict] = []
    try:
        for nome in nomes_ok:
            r = encontrados[nome]
            if dados[r][2] is None:
                raise ValueError(f"linha {r} ({nome}): coluna O (Preço sem as taxas) vazia "
                                 f"— guarde e feche a calculadora no Excel e corra de novo")
            pecas.append({"row": r, "nome": nome, "preco": preco_shopee(dados[r][0], dados[r][1], r),
                          "sem_taxas": dados[r][2]})
    except ValueError as erro:
        print(f"❌ {erro}")
        return 1
    pecas.sort(key=lambda p: (p["preco"], p["nome"]))

    precos = [p["preco"] for p in pecas]
    print(f"  {len(pecas)} variações | preço min {min(precos):.2f} "
          f"| médio {sum(precos)/len(precos):.2f} | max {max(precos):.2f}")
    grupos = agrupar_por_razao(pecas)
    print(f"  {len(pecas)} variações -> {len(grupos)} anúncio(s) (razão preço <= {RAZAO_PRECO_MAX:.2f}x)")
    for e, grupo in enumerate(grupos, start=1):
        gp = [p["preco"] for p in grupo]
        print(f"     E={e} | {len(grupo)} var. | {min(gp):.2f} a {max(gp):.2f} | razão {max(gp)/min(gp):.2f}")
    if any(len(g) > MAX_VARIACOES for g in grupos):
        print(f"  ⚠️ algum anúncio acima de {MAX_VARIACOES} variações (confirmar o limite da Shopee)")

    try:
        stamp = time.strftime("%Y%m%d%H%M%S")
        linhas: list[dict] = []
        usadas_m: set[str] = set()
        for e, grupo in enumerate(grupos, start=1):
            titulo = construir_titulo_run(e, len(grupos))
            descricao = construir_descricao([p["nome"] for p in grupo])
            sku_principal = f"MINI.RUN{stamp}.{e}"
            usadas_g: set[str] = set()
            for p in grupo:
                g_opcao = opcao_variacao(p["nome"], usadas_g)
                usadas_g.add(g_opcao)
                m_sku = slug_sku(p["nome"])
                if m_sku in usadas_m:
                    raise ValueError(f"SKU duplicado: {m_sku} ({p['nome']})")
                usadas_m.add(m_sku)
                linhas.append({
                    "row": p["row"], "nome": p["nome"], "preco": p["preco"],
                    "titulo": titulo, "descricao": descricao, "sku_principal": sku_principal,
                    "e": str(e), "f": NOME_VAR, "g": g_opcao, "m": m_sku,
                })
        if len(pecas) >= 2:
            kit = calcular_kit([p["sem_taxas"] for p in pecas])
            soma_publicada = sum(p["preco"] for p in pecas)
            print(f"  Kit completo ({len(pecas)} peças): soma sem taxas {kit['soma']:.2f} "
                  f"| preço do kit {kit['kit']:.2f} | soma dos preços publicados {soma_publicada:.2f}")
            linhas.append({
                "row": None, "nome": "Kit Completo", "preco": kit["kit"],
                "titulo": construir_titulo_kit(len(pecas)),
                "descricao": construir_descricao([p["nome"] for p in pecas]),
                "sku_principal": f"MINI.KIT{stamp}", "e": None, "f": None, "g": None, "m": None,
            })
    except ValueError as erro:
        print(f"❌ {erro}")
        return 1

    print("\n=== FASE 3 - preencher template + validar listas ===")
    try:
        wb = carregar_template()
    except Exception as erro:  # noqa: BLE001
        print(f"❌Erro ao abrir template: {erro}")
        return 1
    ws = wb[FOLHA_TEMPLATE]

    for i, item in enumerate(linhas):
        r = LINHA_DADOS_TEMPLATE + i
        ws.cell(row=r, column=COL["categoria"], value=CATEGORIA)
        ws.cell(row=r, column=COL["nome"], value=item["titulo"])
        ws.cell(row=r, column=COL["descricao"], value=item["descricao"])
        ws.cell(row=r, column=COL["preco"], value=item["preco"])
        ws.cell(row=r, column=COL["estoque"], value=ESTOQUE)
        ws.cell(row=r, column=COL["peso"], value=PESO_KG)
        ws.cell(row=r, column=COL["comprimento"], value=COMPRIMENTO)
        ws.cell(row=r, column=COL["largura"], value=LARGURA)
        ws.cell(row=r, column=COL["altura"], value=ALTURA)
        ws.cell(row=r, column=COL["canal_ae"], value=CANAL_AE)
        ws.cell(row=r, column=COL["agrupavel"], value=ITEM_AGRUPAVEL)
        c_gtin = ws.cell(row=r, column=COL["gtin"], value=GTIN_SEM)
        c_gtin.number_format = "@"               # P: texto ("00" não vira número)
        # AF (32) fica vazia - "Please do not edit this column"
        ws.cell(row=r, column=COL["sku_principal"], value=item["sku_principal"])
        ws.cell(row=r, column=COL["e_agrup"], value=item["e"])
        ws.cell(row=r, column=COL["f_var"], value=item["f"])
        ws.cell(row=r, column=COL["g_opcao"], value=item["g"])
        ws.cell(row=r, column=COL["m_sku"], value=item["m"])
        ws.cell(row=r, column=COL["ag"], value=PRE_ENCOMENDA)
        # H (8) e R..Z: fotos - ficam vazias (preenchimento manual)

    linha_final = LINHA_DADOS_TEMPLATE + len(linhas) - 1
    problemas = validar_contra_listas_da_planilha(ws, linha_final=linha_final)
    if problemas:
        print(f"❌ VALIDAÇÃO FALHOU ({len(problemas)}) - nada gravado:")
        for p in problemas[:40]:
            print("   ", p)
        return 1
    print(f"  Validação das listas: OK (linhas {LINHA_DADOS_TEMPLATE}..{linha_final})")

    saida = caminho_saida_novo()
    saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(saida)
    wb.close()
    print(f"  Gravado: {saida}")

    print("\n=== FASE 4 - validação pós-gravação ===")
    contagem = contar_dv_xml(saida)
    esperado = (5005, 5004, 1)
    ok = True
    if contagem != esperado:
        print(f"  ❌ DVs no ficheiro gravado: {contagem} (esperado {esperado})")
        ok = False
    else:
        print(f"  DVs no XML gravado: {contagem} (total/list/whole) ✔")

    wb_t = carregar_template()
    wb_s = load_workbook(saida)
    s_t, s_s = dv_set(wb_t), dv_set(wb_s)
    diff = s_t ^ s_s
    if diff:
        ok = False
        print(f"  ❌ diff de DVs ({len(diff)}):")
        for d in sorted(diff)[:20]:
            print("   ", d)
    else:
        print("  DVs (sqref+formula1+type+allowBlank): diff = 0 ✔")

    ws_o, ws_f = wb_t[FOLHA_TEMPLATE], wb_s[FOLHA_TEMPLATE]
    alteradas = []
    for r in range(1, 7):
        for c in range(1, 53):
            a = ws_o.cell(row=r, column=c).value
            b = ws_f.cell(row=r, column=c).value
            # openpyxl normaliza '' -> None na gravação; ambos contam como vazio
            if (a if a not in (None, "") else None) != (b if b not in (None, "") else None):
                alteradas.append(f"{get_column_letter(c)}{r}")
    if alteradas:
        ok = False
        print(f"  ❌ linhas 1-6 alteradas: {alteradas}")
    else:
        print("  Linhas 1-6 intactas ✔")

    problemas_f = validar_contra_listas_da_planilha(ws_f, linha_final=linha_final)
    if problemas_f:
        ok = False
        print(f"  ❌ revalidação do ficheiro final: {len(problemas_f)} problemas")
        for p in problemas_f[:20]:
            print("   ", p)
    else:
        print("  Revalidação listas no ficheiro final: OK ✔")

    erros_valores = []
    for i, item in enumerate(linhas):
        r = LINHA_DADOS_TEMPLATE + i
        esperados = {
            COL["categoria"]: CATEGORIA, COL["nome"]: item["titulo"],
            COL["descricao"]: item["descricao"], COL["preco"]: item["preco"],
            COL["estoque"]: ESTOQUE, COL["peso"]: PESO_KG,
            COL["comprimento"]: COMPRIMENTO, COL["largura"]: LARGURA,
            COL["altura"]: ALTURA, COL["canal_ae"]: CANAL_AE,
            COL["agrupavel"]: ITEM_AGRUPAVEL, 32: None,
            COL["gtin"]: GTIN_SEM,
            COL["sku_principal"]: item["sku_principal"], COL["e_agrup"]: item["e"],
            COL["f_var"]: item["f"], COL["g_opcao"]: item["g"],
            COL["m_sku"]: item["m"], COL["ag"]: PRE_ENCOMENDA,
            34: None,                   # AH: sem NCM
        }
        for c, esp in esperados.items():
            v = ws_f.cell(row=r, column=c).value
            if v != esp:
                erros_valores.append(f"{get_column_letter(c)}{r}: {v!r} != {esp!r}")
    if ws_f.cell(row=linha_final + 1, column=COL["nome"]).value not in (None, ""):
        erros_valores.append(f"B{linha_final + 1} devia estar vazia")
    if erros_valores:
        ok = False
        print(f"  ❌ valores divergentes ({len(erros_valores)}):")
        for e in erros_valores[:20]:
            print("   ", e)
    else:
        print(f"  Valores das {len(linhas)} linhas (variações + kit) conferem ✔ "
              "(incl. AE=Ligado, AF=vazia, AV=No, P=00)")

    # --- validação dos anúncios (razão 4x, G único por anúncio, M único, linhas contíguas) ---
    erros_anuncio = []
    por_e: dict[str, list[dict]] = {}
    for it in linhas:
        if it["e"] is not None:          # o kit (sem E) fica fora das regras por anúncio
            por_e.setdefault(it["e"], []).append(it)
    for e, its in por_e.items():
        ps = [it["preco"] for it in its]
        if max(ps) > RAZAO_PRECO_MAX * min(ps) + 1e-9:
            erros_anuncio.append(f"anúncio {e}: razão {max(ps) / min(ps):.2f} > {RAZAO_PRECO_MAX:.2f}")
        if len({it["g"] for it in its}) != len(its):
            erros_anuncio.append(f"anúncio {e}: opção G duplicada")
        idxs = [i for i, it in enumerate(linhas) if it["e"] == e]
        if idxs != list(range(idxs[0], idxs[0] + len(idxs))):
            erros_anuncio.append(f"anúncio {e}: linhas não contíguas no template")
    com_m = [it["m"] for it in linhas if it["m"] is not None]
    if len(set(com_m)) != len(com_m):
        erros_anuncio.append("SKU M duplicado no ficheiro")
    if erros_anuncio:
        ok = False
        print(f"  ❌ validação dos anúncios ({len(erros_anuncio)}):")
        for g in erros_anuncio:
            print("   ", g)
    else:
        print(f"  Anúncios: {len(por_e)} com razão <= {RAZAO_PRECO_MAX:.2f} ✔ (G/M únicos, linhas contíguas)")

    # fotos: H (por variação) e R..Z (galeria) ficam vazias - aviso, não falha
    cols_foto = [8] + list(range(18, 27))
    fotos_pendentes = [
        f"{get_column_letter(c)}{LINHA_DADOS_TEMPLATE + i}"
        for i in range(len(linhas)) for c in cols_foto
        if ws_f.cell(row=LINHA_DADOS_TEMPLATE + i, column=c).value in (None, "")
        and not (linhas[i]["e"] is None and c == 8)
    ]
    if fotos_pendentes:
        print(f"  ⚠ FOTOS PENDENTES: {len(fotos_pendentes)} células vazias "
              "(H por variação + R..Z capa/galeria) - preencher manualmente antes do upload")

    wb_t.close()
    wb_s.close()

    print("\n=== AMOSTRA (linha template | preço | G | título) ===")
    for i in list(range(min(3, len(linhas)))) + list(range(max(3, len(linhas) - 2), len(linhas))):
        item = linhas[i]
        print(f"  {LINHA_DADOS_TEMPLATE + i} | {item['preco']:.2f} | {item['g']} | {item['titulo']}")

    print()
    if ok:
        print(f"✅ TUDO OK - ficheiro pronto para upload: {saida}")
        print(f"   {len([it for it in linhas if it['e'] is not None])} variações em {len({it["e"] for it in linhas if it["e"] is not None})} anúncio(s){" + kit" if any(it["e"] is None for it in linhas) else ""} | linhas "
              f"{LINHA_DADOS_TEMPLATE}..{linha_final} do template")
        print("   Pendente manual: fotos (H + R..Z), corrigir o título genérico")
        print("   e confirmar o checkbox 'Item sem GTIN' no upload")
        return 0
    print("❌ FASE 4 FALHOU - rever erros acima")
    return 1


def main() -> int:
    return executar(sys.argv[1] if len(sys.argv) > 1 else None)


if __name__ == "__main__":
    sys.exit(main())
