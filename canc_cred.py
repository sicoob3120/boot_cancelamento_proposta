import pyautogui
import time
import os
from datetime import datetime, timedelta
import pytesseract
import cv2


class CancelCred:
    def __init__(self):
        hoje = datetime.today()
        self.d30_dias = hoje-timedelta(days=30)
        self.d90_dias_d30 = self.d30_dias-timedelta(days=90)

        self.caminho_imagens = r'C:\Users\sicoob\Sicoob Central Crediminas\3120 - Business Intelligence (B.I) - Geral\Automacoes\Imagens_RPA\cancelamentocredito'

        self.img_user = os.path.join(self.caminho_imagens, 'user.png')
        self.img_senha = os.path.join(self.caminho_imagens, 'senha.png')
        self.img_btn_pesquisar = os.path.join(self.caminho_imagens, 'btn_pesquisar.png')
        self.img_credito = os.path.join(self.caminho_imagens, 'credito.png')
        self.img_emprestimo = os.path.join(self.caminho_imagens, 'emprestimo.png')
        self.img_consignado = os.path.join(self.caminho_imagens, 'consignado.png')
        self.img_crural = os.path.join(self.caminho_imagens, 'crural.png')
        self.img_climites = os.path.join(self.caminho_imagens, 'climites.png')
        self.img_pesquisar = os.path.join(self.caminho_imagens, 'pesquisar.png')
        self.img_emprestimo_teste = os.path.join(self.caminho_imagens, 'emprestimo_teste.png')
        self.img_consignado_teste = os.path.join(self.caminho_imagens, 'consignado_teste.png')
        self.img_crural_teste = os.path.join(self.caminho_imagens, 'crural_teste.png')
        self.img_concessao_teste = os.path.join(self.caminho_imagens, 'concessao_teste.png')
        self.img_menu_aplicativos = os.path.join(self.caminho_imagens, 'menu_aplicativos.png')
        self.img_operacoes_credito = os.path.join(self.caminho_imagens, 'operacoes_credito.png')
        self.img_operacao_credito = os.path.join(self.caminho_imagens, 'operacao_credito.png')
        self.img_limpar = os.path.join(self.caminho_imagens, 'limpar.png')
        self.img_antecipacao = os.path.join(self.caminho_imagens, 'antecipacao.png')
        self.img_mesa_operacoes = os.path.join(self.caminho_imagens, 'mesa_de_operacoes.png')
        self.img_data_prop_inicio = os.path.join(self.caminho_imagens, 'data_prop_inicio.png')
        self.img_data_prop_final = os.path.join(self.caminho_imagens, 'data_prop_final.png')
        self.img_procurar = os.path.join(self.caminho_imagens, 'procurar.png')
        self.img_abrir_fase = os.path.join(self.caminho_imagens, 'abrir_fase.png')
        self.img_proposta = os.path.join(self.caminho_imagens, 'proposta.png')
        self.img_documentacao = os.path.join(self.caminho_imagens, 'documentacao.png')
        self.img_garantia = os.path.join(self.caminho_imagens, 'garantia.png')
        self.img_estudo = os.path.join(self.caminho_imagens, 'estudo.png')
        self.img_btn_ok = os.path.join(self.caminho_imagens, 'btn_ok.png')
        self.img_fase_proposta = os.path.join(self.caminho_imagens, 'fase_proposta.png')

    def open_cred(self):
        try:
            pyautogui.PAUSE = 4

            # # Acessar SISBR
            # self.login_sisbr()
            
            # Acessar PLATAFORMA DE CRÉDITO
            self.acessar_credito()
 
            # Loop para acessar todas as platafomas (emprestimo, consignado, credito rural e concessão de limites)
            funcoes = [self.plt_emprestimo, self.plt_consignado, self.plt_crural, self.plt_concessao]
            for func in funcoes:
                self.locate_image(self.img_menu_aplicativos)
                time.sleep(2)

                self.locate_image(self.img_limpar)
                time.sleep(2)
                    
                func()
                time.sleep(3)
            
            self.plt_emprestimo()

        except Exception as e:
            print("Erro ao tentar realizar a automação, reveja os processos", e)
                 
    def login_sisbr(self):
            # Abrir SisBR
            pyautogui.press('winleft')
            pyautogui.write('sis')
            pyautogui.press('enter')
            time.sleep(10)
                
            # Posicionar mouse e digitar usuário            
            self.locate_image(self.img_user)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write('ricardor3120_00')

            # Posicionar mouse e digitar senha
            self.locate_image(self.img_senha)
            pyautogui.write('Dezembro@2024') 
            pyautogui.press('enter')
            time.sleep(20)
    
    def acessar_credito(self):
            # Posicionar mouse, digitar "PLATAFORMA DE CREDITO" e clicar
            self.locate_image(self.img_btn_pesquisar)
            time.sleep(1)
            pyautogui.write('PLATAFORMA DE CREDITO')
            time.sleep(1)
            self.locate_image(self.img_credito)
            time.sleep(3)

    def locate_image(self, img, grayscale=False, max_attemps=10, delay=1, existe=True):
        try:
            encontrado = False
            for tentativa in range(max_attemps):
                # Localizar a imagem na tela
                img_location = pyautogui.locateOnScreen(img, confidence=0.9, grayscale=grayscale)
                

                if img_location is not None:
                    # Encontrar o centro da imagem
                    img_center = pyautogui.center(img_location)
                    
                    # Mover o mouse para o centro da imagem e clicar
                    pyautogui.click(img_center)
                    # print(f'Imagem {img} encontrada na tentativa {attempts + 1}.')
                    encontrado = True
                    break
                    
                else:
                    tentativa += 1
                    print(f"Imagem {img} não encontrada na tela.")
                    time.sleep(delay)
            
            if encontrado and existe:
                return True
            elif not encontrado and not existe:
                return True

            return False
        
        except Exception as e:
            print(f"Erro ao encontrar a imagem {img}!", e)   

    def plt_emprestimo(self):
        try:
            # Pesquisar emprestimo (teste)
            self.locate_image(self.img_pesquisar)
            pyautogui.write('empr')
            time.sleep(2)
            self.locate_image(self.img_emprestimo_teste)
            time.sleep(2)

            # Clicar menu "OPERAÇÕES DE CRÉDITO"
            self.locate_image(self.img_operacoes_credito)
            time.sleep(1)

            # Acessar sub-menu Mesa de Operações
            self.locate_image(self.img_mesa_operacoes)
            time.sleep(1)

            # Data proposta inicio e final
            self.locate_image(self.img_data_prop_inicio)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1)
            pyautogui.write(self.d90_dias_d30.strftime('%d'))
            pyautogui.write(self.d90_dias_d30.strftime('%m'))
            pyautogui.write(self.d90_dias_d30.strftime('%Y'))
            time.sleep(1)

            self.locate_image(self.img_data_prop_final)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1)
            pyautogui.write(self.d30_dias.strftime('%d'))
            pyautogui.write(self.d30_dias.strftime('%m'))
            pyautogui.write(self.d30_dias.strftime('%Y'))
            time.sleep(1)

            fases = [self.img_proposta, self.img_documentacao, self.img_garantia, self.img_estudo]
            for fase in fases:
                self.locate_image(self.img_abrir_fase)
                self.locate_image(fase)
                self.locate_image(self.img_procurar)

                while True:
                    screen_path = os.path.join(self.caminho_imagens, 'imagem_referencia.png')
                    self.capture_region(439, 472, 1040, 354, screen_path)  # Ajustar a resolução conforme necessário

                    # Extrair texto da imagem da tela
                    screen_text_proposta = self.extract_text_from_image(screen_path, config="--psm 4 -l por --dpi 2500 -c tessedit_char_whitelist= 0123456789.,-abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")

                    # Verificar se o texto extraído contém a palavra-chave da proposta
                    if "Proposta" not in screen_text_proposta:
                        # alert_proposta = pyautogui.alert('Nenhuma proposta para analisar.')
                        # alert_proposta  
                        print("Nenhuma proposta para analisar, fechando o sistema...")
                        # Fechar sistema, caso não encontre proposta                  
                        for _ in range(2):  # Supondo que há 2 abas a serem fechadas
                            pyautogui.hotkey('alt', 'f4') # comando para fechar a aba
                            time.sleep(4)  # Pequena pausa entre os fechamentos
                        self.locate_image(self.img_btn_fechar)                    
                        break
                    else:
                        print("Proposta encontrada, processando...")

                self.locate_image(self.img_btn_ok)
                time.sleep(2)


        except Exception as e:
            print("Erro ao tentar realizar a automação, reveja os processos", e)

    def plt_consignado(self):
        try:
            # Pesquisar emprestimo (teste)
            self.locate_image(self.img_pesquisar)
            pyautogui.write('cons')
            time.sleep(2)
            self.locate_image(self.img_consignado_teste)
            time.sleep(4)

            # Clicar menu "OPERAÇÕES DE CRÉDITO"
            self.locate_image(self.img_operacao_credito)
            time.sleep(1)


        except Exception as e:
            print("Erro ao tentar realizar a automação, reveja os processos", e)

    def plt_crural(self):
        try:
            # Pesquisar emprestimo (teste)
            # self.locate_image(self.pesquisar)
            # pyautogui.write('rural')
            pyautogui.moveTo(x=238, y=356, duration=0)
            time.sleep(2)
            pyautogui.scroll(-100000)
            time.sleep(2)
            self.locate_image(self.img_crural)
            time.sleep(5)

            # Clicar menu "OPERAÇÕES DE CRÉDITO"
            # self.locate_image(self)
            time.sleep(1)


        except Exception as e:
            print("Erro ao tentar realizar a automação, reveja os processos", e)
        
    def plt_concessao(self):
        try:
            # Pesquisar emprestimo (teste)
            self.locate_image(self.img_pesquisar)
            pyautogui.write('conc')
            time.sleep(2)
            self.locate_image(self.img_concessao_teste)
            time.sleep(4)

            # Clicar menu "ANTECIPAÇÃO DE RECEBÍVES"
            self.locate_image(self.img_antecipacao)
            time.sleep(1)


        except Exception as e:
            print("Erro ao tentar realizar a automação, reveja os processos", e)
    
    def abrir_fases(self):
        fases = [self.img_proposta, self.img_documentacao, self.img_garantia, self.img_estudo]
        for fase in fases:
            self.locate_image(self.img_abrir_fase)
            self.locate_image(fase)
            self.locate_image(self.img_procurar)
            while True:
                screen_path = os.path.join(self.caminho_imagens, 'imagem_referencia.png')
                self.capture_region(439, 472, 1040, 354, screen_path)  # Ajustar a resolução conforme necessário

                # Extrair texto da imagem da tela
                screen_text_proposta = self.extract_text_from_image(screen_path, config="--psm 4 -l por --dpi 2500 -c tessedit_char_whitelist= 0123456789.,-abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")

                # Verificar se o texto extraído contém a palavra-chave da proposta
                if "Proposta" not in screen_text_proposta:
                    # alert_proposta = pyautogui.alert('Nenhuma proposta para analisar.')
                    # alert_proposta  
                    print("Nenhuma proposta para analisar, fechando o sistema...")
                    # Fechar sistema, caso não encontre proposta                  
                    for _ in range(2):  # Supondo que há 2 abas a serem fechadas
                        pyautogui.hotkey('alt', 'f4') # comando para fechar a aba
                        time.sleep(4)  # Pequena pausa entre os fechamentos
                    self.locate_image(self.img_btn_fechar)                    
                    break
                else:
                    print("Proposta encontrada, processando...")

            self.locate_image(self.img_btn_ok)
            time.sleep(2)
    
    def capture_region(self, x, y, width, height, output_path):
        # Captura a região especificada da tela
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        # Salva a imagem capturada
        screenshot.save(output_path)
        print(f"Screenshot salvo como {output_path}")

    def extract_text_from_image(self, image_path, config=None):
        # image = Image.open(image_path)
        image = self.preprocess_image(image_path)
        text = pytesseract.image_to_string(image, config=config)
        return text

    def preprocess_image(self, image_path):
        image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        
        # Aumentar a resolução
        scale = 2.0
        width = int(image.shape[1] * scale)
        height = int(image.shape[0] * scale)
        image = cv2.resize(image, (width, height), interpolation=cv2.INTER_CUBIC)
        
        # Aplicar binarização
        _, image = cv2.threshold(image, 128, 255, cv2.THRESH_BINARY)
        
        # Salvar a imagem processada (opcional)
        cv2.imwrite('imagem_processada.png', image)
        return image
         
if __name__=='__main__':
    cred = CancelCred()
    cred.open_cred()
