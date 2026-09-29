import os
import time
import subprocess
import pyautogui

# --- CONFIGURAÇÕES E CAMINHOS ---
CAMINHO_CHITUBOX = r"C:\Program Files\CHITUBOX\CHITUBOX.exe"
PASTA_BASE = r"C:\Users\Gustavo\Desktop\RoboChitu"
PASTA_ORIGEM_STL = os.path.join(PASTA_BASE, "Arquivos para fatiar")

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

def aguardar_e_salvar_print(nome_arquivo):
    tempo_espera = 60 # Tempo fixo para o fatiamento
    print(f"A fatiar a peça... A aguardar {tempo_espera} segundos.")
    time.sleep(tempo_espera)
    
    nome_limpo = os.path.splitext(nome_arquivo)[0]
    caminho_print = os.path.join(PASTA_ORIGEM_STL, f"resultado_{nome_limpo}.png")
    
    pyautogui.screenshot().save(caminho_print)
    print(f"📸 Captura guardada: resultado_{nome_limpo}.png")

def voltar_para_mesa():
    print("À procura da imagem 'botao_voltar.png' no ecrã...")
    try:
        centro = pyautogui.locateCenterOnScreen(IMG_BOTAO_VOLTAR, confidence=0.8)
        if centro:
            print("✅ Botão de voltar encontrado. A regressar à mesa...")
            pyautogui.click(centro)
            time.sleep(3) 
        else:
            print("❌ Aviso: Botão detetado, mas sem centro. Vou tentar avançar.")
    except pyautogui.ImageNotFoundException:
        print("❌ Aviso: A imagem do botão não foi encontrada. Se o ecrã não mudou, a limpeza pode falhar.")

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
        if os.path.isdir(caminho_item) or (os.path.isfile(caminho_item) and item.lower().endswith('.stl')):
            itens_validos.append(item)

    if not itens_validos:
        print(f"❌ Erro: Nenhum ficheiro .stl ou pasta válida encontrada em:\n{PASTA_ORIGEM_STL}")
        return

    print(f"\n🚀 A iniciar fatiamento MISTO de {len(itens_validos)} itens (Arquivos Soltos e Pastas/Kits).")
    
    software_aberto = False

    for index, item in enumerate(itens_validos):
        caminho_completo = os.path.join(PASTA_ORIGEM_STL, item)
        
        # CENÁRIO 1: É UMA PASTA
        if os.path.isdir(caminho_completo):
            stls_na_pasta = [f for f in os.listdir(caminho_completo) if f.lower().endswith('.stl')]
            
            if not stls_na_pasta:
                print(f"⚠️ A pasta '{item}' está vazia ou não tem STLs. A saltar...")
                continue
                
            print(f"\n[Item {index+1}/{len(itens_validos)}] 📁 A processar PASTA: {item} ({len(stls_na_pasta)} ficheiros)")
            
            for i, stl in enumerate(stls_na_pasta):
                caminho_stl = os.path.join(caminho_completo, stl)
                
                if not software_aberto:
                    print(f"A iniciar o CHITUBOX já com a 1ª peça do kit: {stl}")
                    subprocess.Popen([CAMINHO_CHITUBOX, caminho_stl])
                    print("A aguardar 15 segundos para o software iniciar e carregar a peça...")
                    time.sleep(15)
                    software_aberto = True
                else:
                    print(f"A injetar a peça {i+1}/{len(stls_na_pasta)} na mesma mesa: {stl}")
                    importar_arquivo_teclado(caminho_stl)
            
            print(f"A processar e fatiar o lote completo da pasta '{item}'...")
            aplicar_comando_lo1() # <--- NOVO COMANDO AQUI
            executar_fatiamento_teclado()
            aguardar_e_salvar_print(item) 
            voltar_para_mesa()
            limpar_mesa()

        # CENÁRIO 2: É UM ARQUIVO .STL SOLTO
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

def modo_tudo_junto(arquivos_stl):
    print(f"\n🚀 Modo 2 selecionado: A colocar todos os {len(arquivos_stl)} ficheiros soltos na mesma mesa.")
    
    lista_comandos = [CAMINHO_CHITUBOX]
    for arquivo in arquivos_stl:
        lista_comandos.append(os.path.join(PASTA_ORIGEM_STL, arquivo))
        
    subprocess.Popen(lista_comandos)
    time.sleep(20) 
    
    aplicar_comando_lo1() # <--- NOVO COMANDO AQUI
    executar_fatiamento_teclado()
    aguardar_e_salvar_print("LOTE_COMPLETO_JUNTO.stl")
    voltar_para_mesa()
    limpar_mesa()

def iniciar_robo():
    print("\n==============================================")
    print("🤖 ROBÔ CHITUBOX OTIMIZADO (PASTAS E FICHEIROS)")
    print("==============================================")
    print("Escolha o modo de operação:")
    print(" [ 1 ] Fatiamento Misto (Processa pastas/kits e ficheiros soltos na ordem)")
    print(" [ 2 ] Fatiar TUDO JUNTO (Ignora pastas, joga todos os STLs soltos juntos)")
    print("==============================================")
    
    opcao = input("Digite a opção desejada (1 ou 2): ").strip()
    
    if opcao == "1":
        modo_misto()
    elif opcao == "2":
        arquivos_stl = [f for f in os.listdir(PASTA_ORIGEM_STL) if f.lower().endswith('.stl') and os.path.isfile(os.path.join(PASTA_ORIGEM_STL, f))]
        if arquivos_stl:
            modo_tudo_junto(arquivos_stl)
        else:
            print("❌ Erro: Nenhum ficheiro .stl solto encontrado para o Modo 2.")
    else:
        print("❌ Opção inválida.")

if __name__ == "__main__":
    iniciar_robo()