import os
import re
import csv
import time
import subprocess
import pyautogui
from PIL import Image

# --- CONFIGURAÇÕES E CAMINHOS ---
CAMINHO_CHITUBOX = r"C:\Program Files\CHITUBOX\CHITUBOX.exe"
PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_ORIGEM_STL = os.path.join(PASTA_BASE, "Arquivos para fatiar")
PASTA_CAPTURAS = os.path.join(PASTA_BASE, "imagens")
os.makedirs(PASTA_ORIGEM_STL, exist_ok=True)
os.makedirs(PASTA_CAPTURAS, exist_ok=True)

# Extensões aceites pelo CHITUBOX
EXTENSOES_VALIDAS = ('.stl', '.chitubox')

# --- PÓS-PROCESSAMENTO DAS CAPTURAS (CORTE + OCR -> CSV) ---
PASTA_RESULTADOS = os.path.join(PASTA_BASE, "resultados")
os.makedirs(PASTA_RESULTADOS, exist_ok=True)
CAMPOS_CSV = ["ficheiro", "weight_g", "time"]
_engine_ocr = None

# Espera antes da captura consoante a quantidade de ficheiros fatiados
TABELA_TEMPOS = [   # (máximo de ficheiros, segundos)
    (2, 30), (4, 45), (6, 60), (8, 80), (10, 120),
    (14, 150), (18, 180),
]
TEMPO_PADRAO = 210   # acima de 18 ficheiros
# Depois da espera mínima, lê o ecrã repetidamente até o Weight/Time aparecer (sem 'Calculating...')
INTERVALO_LEITURA = 10          # segundos entre leituras
FATOR_LIMITE_ESPERA = 4         # limite = tempo mínimo × este fator...
LIMITE_MINIMO_ESPERA = 600      # ...mas nunca menos que isto (segundos)


def tem_extensao_valida(nome_ficheiro):
    return nome_ficheiro.lower().endswith(EXTENSOES_VALIDAS)

def contar_ficheiros_para_fatiar(caminho):
    """Conta os ficheiros válidos de uma pasta; ficheiro único conta como 1."""
    if os.path.isdir(caminho):
        return sum(
            1 for f in os.listdir(caminho)
            if os.path.isfile(os.path.join(caminho, f)) and tem_extensao_valida(f)
        )
    return 1

def tempo_espera_por_quantidade(quantidade):
    for maximo, segundos in TABELA_TEMPOS:
        if quantidade <= maximo:
            return segundos
    return TEMPO_PADRAO

# Imagens de referência
IMG_BOTAO_VOLTAR = os.path.join(PASTA_BASE, "botao_voltar.png")
IMG_SLICE_CURRENT = os.path.join(PASTA_BASE, "botao_current_platform.png")  # dropdown fechado no painel Slice
IMG_SLICE_ALL = os.path.join(PASTA_BASE, "botao_all_platform.png")          # opção 'All Platform' da lista
IMG_COUNT = os.path.join(PASTA_BASE, "count.png")  # caixa 'Count' (valor 5) do painel Add Platform

# True: abre com o 1.º ficheiro e importa os restantes da pasta (Ctrl+O) antes do fluxo de All Platform
IMPORTAR_RESTANTES = True

pyautogui.FAILSAFE = True

def aplicar_comando_lo1():
    """Executa a sequência L, O, 1 antes de iniciar o fatiamento"""
    print("A executar comando pré-slice (L -> O -> 1)...")
    pyautogui.press('l')
    time.sleep(0.5)
    pyautogui.press('o')
    time.sleep(1) # Pausa para o CHITUBOX processar o menu/ação
    pyautogui.press('1')
    time.sleep(2) # Pausa para a ação estabilizar antes de fatiar

def executar_fatiamento_teclado():
    print("A enviar o comando de fatiamento via teclado (S L)...")
    pyautogui.press('s')
    time.sleep(0.5) 
    pyautogui.press('l')
    time.sleep(3) 
    print("A confirmar o fatiamento (Enter)...")
    pyautogui.press('enter')

def localizar_na_tela(imagem, tentativas=5, intervalo=1, confidence=0.8):
    """Devolve o centro de 'imagem' no ecrã, ou None."""
    for _ in range(tentativas):
        try:
            centro = pyautogui.locateCenterOnScreen(imagem, confidence=confidence)
            if centro:
                return centro
        except pyautogui.ImageNotFoundException:
            pass
        time.sleep(intervalo)
    return None

def adicionar_plataforma_por_atalho():
    """Adiciona a 2.ª plataforma: A, P abre o painel Add Platform; clica no 5 da caixa Count,
    escreve 1 e Enter (Apply)."""
    print("A adicionar a 2.ª plataforma (A -> P -> Count 1 -> Enter)...")
    trazer_chitubox_para_frente()
    pyautogui.press('a')
    time.sleep(0.5)
    pyautogui.press('p')
    time.sleep(2)

    centro = localizar_na_tela(IMG_COUNT, tentativas=5, confidence=0.8)
    if not centro:
        caminho_debug = os.path.join(PASTA_BASE, "debug_add_platform.png")
        pyautogui.screenshot().save(caminho_debug)
        print(f"❌ Painel 'Count' não encontrado. Captura de debug: {caminho_debug}")
        pyautogui.press('esc')
        return False

    # count.png é só o rótulo 'Count' (50x30); o '5' fica 44 px à direita do centro do rótulo
    pyautogui.click(centro.x + 44, centro.y)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')   # apaga o 5 selecionado antes de escrever
    pyautogui.write('1', interval=0.05)
    time.sleep(0.5)
    pyautogui.press('enter')        # Apply
    pyautogui.press('enter')        # Apply
    time.sleep(2)

    # afasta o rato da zona dos botões para nenhum tooltip tapar o ecrã
    largura, altura = pyautogui.size()
    pyautogui.moveTo(largura * 0.45, altura * 0.25)
    print("✅ 2.ª plataforma adicionada (Count = 1).")
    return True

def configurar_slice_todas_plataformas():
    """S, L abre o painel Slice; escolhe Current -> All Platform; volta ao Prepare (Esc).
    Corre em cada item, depois de a mesa ter sido limpa e as peças importadas."""
    print("A abrir o painel Slice (S L) para escolher 'All Platform'...")
    pyautogui.press('s')
    time.sleep(0.5)
    pyautogui.press('l')
    time.sleep(3)

    dropdown = localizar_na_tela(IMG_SLICE_CURRENT)
    if dropdown:
        pyautogui.click(dropdown)
        time.sleep(1)
        opcao = localizar_na_tela(IMG_SLICE_ALL)
        if opcao:
            pyautogui.click(opcao)
            time.sleep(1)
            print("✅ Modo de slice: All Platform.")
        else:
            caminho_debug = os.path.join(PASTA_BASE, "debug_slice_mode.png")
            pyautogui.screenshot().save(caminho_debug)
            print(f"❌ Opção 'All Platform' não encontrada. Captura de debug: {caminho_debug}")
            pyautogui.press('esc')
    else:
        print("ℹ️ Dropdown 'Current Platform' não encontrado — a seguir com o modo atual.")

    # Voltar ao Prepare: Esc fecha o painel Slice
    pyautogui.press('esc')
    time.sleep(3)
    print("A voltar ao Prepare.")

def trazer_chitubox_para_frente():
    """Best-effort: ativa a janela do CHITUBOX antes de capturar o ecrã."""
    try:
        import pygetwindow as gw
        for padrao in ("CHITUBOX", ".chitubox"):
            for janela in gw.getWindowsWithTitle(padrao):
                try:
                    if janela.isMinimized:
                        janela.restore()
                    janela.activate()
                    time.sleep(1)
                    return True
                except Exception:
                    continue
    except Exception:
        pass
    return False

def aguardar_e_salvar_print(nome_arquivo, quantidade=None):
    if quantidade is None:
        quantidade = contar_ficheiros_para_fatiar(os.path.join(PASTA_ORIGEM_STL, nome_arquivo))
    tempo_espera = tempo_espera_por_quantidade(quantidade)
    print(f"A fatiar {quantidade} ficheiro(s)... A aguardar {tempo_espera} segundos.")
    time.sleep(tempo_espera)
    
    nome_limpo = os.path.splitext(nome_arquivo)[0]
    caminho_print = os.path.join(PASTA_CAPTURAS, f"resultado_{nome_limpo}.png")

    limite = max(LIMITE_MINIMO_ESPERA, tempo_espera * FATOR_LIMITE_ESPERA)
    inicio = time.time()
    largura, altura = pyautogui.size()
    campos = {}
    while True:
        trazer_chitubox_para_frente()
        # rato para uma zona vazia: o tooltip do painel Slice não pode tapar os números
        pyautogui.moveTo(largura * 0.45, altura * 0.25)
        pyautogui.screenshot().save(caminho_print)
        linhas = transcrever_imagem(recortar_esquerda(caminho_print))
        campos = extrair_campos(linhas)
        if campos.get("weight_g") and campos.get("time"):
            break
        decorrido = int(tempo_espera + (time.time() - inicio))
        if decorrido >= limite:
            break
        a_calcular = any("calculat" in linha.lower() for linha in linhas)
        motivo = "ainda a calcular" if a_calcular else "dados ainda não visíveis"
        print(f"⏳ {motivo}... nova leitura em {INTERVALO_LEITURA}s ({decorrido}/{limite}s)")
        time.sleep(INTERVALO_LEITURA)

    print(f"📸 Captura guardada: resultado_{nome_limpo}.png")
    if campos.get("weight_g"):
        print(f"✅ Dados lidos: {campos.get('weight_g')} | {campos.get('time') or '?'}")
    else:
        print(f"⚠️ Aviso: não foi possível ler Weight na captura após {limite}s.")

def voltar_para_mesa(tentativas=5, intervalo=3):
    print("À procura da imagem 'botao_voltar.png' no ecrã...")
    for tentativa in range(1, tentativas + 1):
        try:
            centro = pyautogui.locateCenterOnScreen(IMG_BOTAO_VOLTAR, confidence=0.8)
            if centro:
                print("✅ Botão de voltar encontrado. A regressar à mesa...")
                pyautogui.click(centro)
                time.sleep(3)
                return True
        except pyautogui.ImageNotFoundException:
            pass
        print(f"⏳ Botão não encontrado (tentativa {tentativa}/{tentativas})...")
        time.sleep(intervalo)

    caminho_debug = os.path.join(PASTA_BASE, "debug_voltar.png")
    pyautogui.screenshot().save(caminho_debug)
    print(f"❌ Botão 'voltar' não encontrado após {tentativas} tentativas.")
    print(f"   Captura de debug guardada: {caminho_debug}")
    print("   Compara-a com 'botao_voltar.png' — se o aspeto do ecrã mudou, recorta o botão de novo.")
    return False

def limpar_mesa():
    """Garante que a mesa é limpa após cada slice"""
    print("🧹 A focar na mesa 3D e a limpar a peça atual...")
    
    # Clica no centro da tela para garantir foco na mesa 3D antes de apagar
    largura_tela, altura_tela = pyautogui.size()
    pyautogui.click(largura_tela / 2, altura_tela / 2)
    time.sleep(1)
    
    print("A executar Ctrl + A e Delete...")
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.press('delete')
    time.sleep(1)

def importar_arquivo_teclado(caminho_completo):
    print("A abrir janela de importação (Ctrl + O)...")
    pyautogui.hotkey('ctrl', 'o')
    time.sleep(2) 
    
    print("A digitar o caminho do ficheiro...")
    pyautogui.write(caminho_completo, interval=0.02)
    time.sleep(1)

    print("A confirmar injeção (Enter)...")
    pyautogui.press('enter')

    print("A aguardar 10 segundos para a peça carregar na mesa...")
    time.sleep(10)

def modo_misto(filtro=None):
    """Lida automaticamente com ficheiros soltos e pastas (Kits/Personagens).
    'filtro' (lista de nomes) limita o processamento a esses itens."""
    todos_itens = os.listdir(PASTA_ORIGEM_STL)

    itens_validos = []
    for item in todos_itens:
        caminho_item = os.path.join(PASTA_ORIGEM_STL, item)
        if os.path.isdir(caminho_item) or (os.path.isfile(caminho_item) and tem_extensao_valida(item)):
            itens_validos.append(item)

    if filtro:
        escolhidos = [i for i in itens_validos if i.lower() in filtro]
        nao_encontrados = [n for n in filtro if n not in [i.lower() for i in itens_validos]]
        for nome in nao_encontrados:
            print(f"⚠️ Não encontrado em '{PASTA_ORIGEM_STL}': {nome}")
        itens_validos = escolhidos

    if not itens_validos:
        print(f"❌ Erro: Nenhum ficheiro {', '.join(EXTENSOES_VALIDAS)} ou pasta válida encontrada em:\n{PASTA_ORIGEM_STL}")
        return

    print(f"\n🚀 A iniciar fatiamento MISTO de {len(itens_validos)} itens (Arquivos Soltos e Pastas/Kits).")
    
    software_aberto = False

    for index, item in enumerate(itens_validos):
        caminho_completo = os.path.join(PASTA_ORIGEM_STL, item)
        
        # CENÁRIO 1: É UMA PASTA
        if os.path.isdir(caminho_completo):
            ficheiros_na_pasta = [f for f in os.listdir(caminho_completo) if tem_extensao_valida(f)]

            if not ficheiros_na_pasta:
                print(f"⚠️ A pasta '{item}' está vazia ou não tem ficheiros válidos. A saltar...")
                continue

            print(f"\n[Item {index+1}/{len(itens_validos)}] 📁 A processar PASTA: {item} ({len(ficheiros_na_pasta)} ficheiros)")

            for i, ficheiro in enumerate(ficheiros_na_pasta):
                caminho_ficheiro = os.path.join(caminho_completo, ficheiro)

                if not software_aberto:
                    print(f"A iniciar o CHITUBOX já com a 1ª peça do kit: {ficheiro}")
                    subprocess.Popen([CAMINHO_CHITUBOX, caminho_ficheiro])
                    print("A aguardar 15 segundos para o software iniciar e carregar a peça...")
                    time.sleep(15)
                    software_aberto = True
                elif IMPORTAR_RESTANTES:
                    print(f"A injetar a peça {i+1}/{len(ficheiros_na_pasta)} na mesma mesa: {ficheiro}")
                    importar_arquivo_teclado(caminho_ficheiro)
                else:
                    print(f"Peça {i+1}/{len(ficheiros_na_pasta)} ({ficheiro}) — já colocada manualmente, não importada.")

            # Todas as peças do item estão na mesa: 2.ª plataforma e depois o All Platform (em cada item)
            adicionar_plataforma_por_atalho()
            configurar_slice_todas_plataformas()
            print(f"A processar e fatiar o lote completo da pasta '{item}'...")
            aplicar_comando_lo1() # <--- NOVO COMANDO AQUI
            executar_fatiamento_teclado()
            aguardar_e_salvar_print(item) 
            voltar_para_mesa()
            limpar_mesa()

        # CENÁRIO 2: É UM FICHEIRO SOLTO (.STL / .CHITUBOX)
        elif os.path.isfile(caminho_completo):
            print(f"\n[Item {index+1}/{len(itens_validos)}] 📄 A processar FICHEIRO SOLTO: {item}")
            
            if not software_aberto:
                print(f"A iniciar o CHITUBOX já com a peça: {item}")
                subprocess.Popen([CAMINHO_CHITUBOX, caminho_completo])
                print("A aguardar 15 segundos para o software iniciar e carregar a peça...")
                time.sleep(15)
                software_aberto = True
            elif IMPORTAR_RESTANTES:
                print(f"A injetar o ficheiro na mesa vazia: {item}")
                importar_arquivo_teclado(caminho_completo)
            else:
                print(f"Ficheiro {item} — já colocado manualmente, não importado.")

            # Ficheiro na mesa: 2.ª plataforma e depois o All Platform (em cada item)
            adicionar_plataforma_por_atalho()
            configurar_slice_todas_plataformas()
            aplicar_comando_lo1() # <--- NOVO COMANDO AQUI
            executar_fatiamento_teclado()
            aguardar_e_salvar_print(item)
            voltar_para_mesa()
            limpar_mesa()
            
    print("\n🏁 Fim da fila de ficheiros e pastas.")
    processar_capturas()

def modo_tudo_junto(arquivos):
    print(f"\n🚀 Modo 2 selecionado: A colocar todos os {len(arquivos)} ficheiros soltos na mesma mesa.")

    lista_comandos = [CAMINHO_CHITUBOX]
    for arquivo in arquivos:
        lista_comandos.append(os.path.join(PASTA_ORIGEM_STL, arquivo))

    subprocess.Popen(lista_comandos)
    time.sleep(20)
    adicionar_plataforma_por_atalho()
    configurar_slice_todas_plataformas()

    aplicar_comando_lo1() # <--- NOVO COMANDO AQUI
    executar_fatiamento_teclado()
    aguardar_e_salvar_print("LOTE_COMPLETO_JUNTO.stl", quantidade=len(arquivos))
    voltar_para_mesa()
    limpar_mesa()
    processar_capturas()

# --- PÓS-PROCESSAMENTO: CORTE + OCR + CSV ---

def recortar_esquerda(caminho):
    """Corta a captura ao meio e devolve o caminho do lado esquerdo guardado."""
    imagem = Image.open(caminho)
    meio = imagem.width // 2
    caminho_recorte = os.path.splitext(caminho)[0] + "_esq.png"
    imagem.crop((0, 0, meio, imagem.height)).save(caminho_recorte)
    return caminho_recorte

def transcrever_imagem(caminho):
    """OCR da imagem -> lista de linhas de texto ordenadas de cima para baixo."""
    global _engine_ocr
    if _engine_ocr is None:
        from rapidocr import RapidOCR
        _engine_ocr = RapidOCR()

    resultado = _engine_ocr(caminho)
    if not resultado or not resultado.txts:
        return []

    itens = []
    for texto, caixa in zip(resultado.txts, resultado.boxes):
        texto = texto.strip()
        if not texto:
            continue
        y = sum(p[1] for p in caixa) / len(caixa)
        x = sum(p[0] for p in caixa) / len(caixa)
        itens.append((y, x, texto))
    itens.sort()

    linhas = []
    for y, x, texto in itens:
        if linhas and abs(y - linhas[-1][0]) <= 15:
            linhas[-1][1].append((x, texto))
        else:
            linhas.append([y, [(x, texto)]])

    return [" ".join(t for _, t in sorted(pecas)) for _, pecas in linhas]

def extrair_campos(linhas):
    """Procura os campos da tela de slice nas linhas do OCR."""
    campos = {campo: "" for campo in CAMPOS_CSV}
    padroes = {
        "weight_g": r"\bWeight\s+([\d.,]+\s*g)\b",
        "time": r"\bTime\s+(\d+h\d+m\d+s)",
    }

    for linha in linhas:
        for campo, padrao in padroes.items():
            if not campos[campo]:
                encontrado = re.search(padrao, linha, re.IGNORECASE)
                if encontrado:
                    campos[campo] = encontrado.group(1).strip()

    return campos

def caminho_csv_novo():
    """Devolve um caminho de CSV novo (um por rodada) na pasta 'resultados'."""
    base = f"resultados_{time.strftime('%Y-%m-%d_%H-%M-%S')}"
    caminho = os.path.join(PASTA_RESULTADOS, f"{base}.csv")
    contador = 2
    while os.path.exists(caminho):
        caminho = os.path.join(PASTA_RESULTADOS, f"{base}_{contador}.csv")
        contador += 1
    return caminho

def processar_capturas():
    """Corta todas as capturas ao meio, transcreve o lado esquerdo e gera o CSV."""
    capturas = sorted(
        f for f in os.listdir(PASTA_CAPTURAS)
        if f.lower().startswith("resultado_") and f.lower().endswith(".png")
        and not f.lower().endswith("_esq.png")
    )

    if not capturas:
        print(f"⚠️ Nenhuma captura 'resultado_*.png' encontrada em:\n{PASTA_CAPTURAS}")
        return

    print(f"\n🔎 A processar {len(capturas)} captura(s): corte ao meio + OCR -> CSV")

    registos = []
    for nome in capturas:
        caminho = os.path.join(PASTA_CAPTURAS, nome)
        caminho_recorte = recortar_esquerda(caminho)
        campos = extrair_campos(transcrever_imagem(caminho_recorte))
        campos["ficheiro"] = os.path.splitext(nome)[0].removeprefix("resultado_")
        registos.append(campos)
        if campos["weight_g"]:
            print(f"  ✅ {campos['ficheiro']}: {campos['weight_g']} | {campos['time'] or '?'}")
        else:
            print(f"  ⚠️ {campos['ficheiro']}: sem dados (CHITUBOX não visível ou ecrã ainda a calcular)")

    caminho_csv = caminho_csv_novo()
    with open(caminho_csv, "w", newline="", encoding="utf-8-sig") as f:
        escritor = csv.DictWriter(f, fieldnames=CAMPOS_CSV)
        escritor.writeheader()
        escritor.writerows(registos)

    print(f"📄 CSV guardado: {caminho_csv}")

def iniciar_robo():
    print("\n==============================================")
    print("🤖 ROBÔ CHITUBOX OTIMIZADO (PASTAS E FICHEIROS)")
    print("==============================================")
    print("Escolha o modo de operação:")
    print(" [ 1 ] Fatiamento Misto (Processa pastas/kits e ficheiros soltos na ordem)")
    print(" [ 2 ] Fatiar TUDO JUNTO (Ignora pastas, joga todos os ficheiros soltos juntos)")
    print(" [ 3 ] Processar capturas existentes (corte ao meio + OCR -> CSV)")
    print(" [ 4 ] Reprocessar item(s) específico(s) (digite o nome da pasta/ficheiro)")
    print("==============================================")

    opcao = input("Digite a opção desejada (1, 2, 3 ou 4): ").strip()

    if opcao == "1":
        modo_misto()
    elif opcao == "4":
        nomes = input("Nome(s) da pasta/ficheiro, separados por vírgula: ").strip()
        filtro = [n.strip().lower() for n in nomes.split(",") if n.strip()]
        if filtro:
            modo_misto(filtro=filtro)
        else:
            print("❌ Nenhum nome digitado.")
    elif opcao == "2":
        arquivos = [f for f in os.listdir(PASTA_ORIGEM_STL) if tem_extensao_valida(f) and os.path.isfile(os.path.join(PASTA_ORIGEM_STL, f))]
        if arquivos:
            modo_tudo_junto(arquivos)
        else:
            print(f"❌ Erro: Nenhum ficheiro {', '.join(EXTENSOES_VALIDAS)} solto encontrado para o Modo 2.")
    elif opcao == "3":
        processar_capturas()
    else:
        print("❌ Opção inválida.")

if __name__ == "__main__":
    iniciar_robo()