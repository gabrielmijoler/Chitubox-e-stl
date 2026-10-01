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

    campos = {}
    for tentativa in range(1, 4):
        trazer_chitubox_para_frente()
        pyautogui.screenshot().save(caminho_print)
        campos = extrair_campos(transcrever_imagem(recortar_esquerda(caminho_print)))
        if campos.get("weight_g"):
            break
        if tentativa < 3:
            print(f"⏳ Dados ainda não prontos ('Calculating...' ou CHITUBOX não visível). Nova tentativa em 15s ({tentativa}/3)...")
            time.sleep(15)

    print(f"📸 Captura guardada: resultado_{nome_limpo}.png")
    if not campos.get("weight_g"):
        print("⚠️ Aviso: não foi possível ler Weight/Time na captura.")

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
    pyautogui.write(caminho_completo)
    time.sleep(1)
    
    print("A confirmar injeção (Enter)...")
    pyautogui.press('enter')
    
    print("A aguardar 10 segundos para a peça carregar na mesa...")
    time.sleep(10)

def modo_misto():
    """Lida automaticamente com ficheiros soltos e pastas (Kits/Personagens)"""
    todos_itens = os.listdir(PASTA_ORIGEM_STL)
    
    itens_validos = []
    for item in todos_itens:
        caminho_item = os.path.join(PASTA_ORIGEM_STL, item)
        if os.path.isdir(caminho_item) or (os.path.isfile(caminho_item) and tem_extensao_valida(item)):
            itens_validos.append(item)

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
                else:
                    print(f"A injetar a peça {i+1}/{len(ficheiros_na_pasta)} na mesma mesa: {ficheiro}")
                    importar_arquivo_teclado(caminho_ficheiro)
            
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
            else:
                print(f"A injetar o ficheiro na mesa vazia: {item}")
                importar_arquivo_teclado(caminho_completo)
            
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
    print("==============================================")

    opcao = input("Digite a opção desejada (1, 2 ou 3): ").strip()

    if opcao == "1":
        modo_misto()
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