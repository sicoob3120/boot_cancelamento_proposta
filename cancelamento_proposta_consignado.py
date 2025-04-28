import pyautogui
import time
import os
from datetime import timedelta 
from datetime import datetime as dt
import pytesseract
import pandas as pd
import cv2
import re


class CancelCred:
    def __init__(self):
        pyautogui.PAUSE = 2
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        self.caminho_imagens = r'C:\Users\sicoob\Sicoob Central Crediminas\3120 - Business Intelligence (B.I) - Geral\Automacoes\Imagens_RPA\cancelamentocredito'
        self.caminho_dir = r'C:\Users\sicoob\Sicoob Central Crediminas\3120 - Business Intelligence (B.I) - Geral\Automacoes\CancelamentoCredito'

        self.img_user = os.path.join(self.caminho_imagens, 'user.png')
        self.img_senha = os.path.join(self.caminho_imagens, 'senha.png')
        self.img_btn_pesquisar = os.path.join(self.caminho_imagens, 'btn_pesquisar.png')
        self.img_credito = os.path.join(self.caminho_imagens, 'plt_credito.png')
        self.img_consignado = os.path.join(self.caminho_imagens, 'consignado.png')
        self.img_pesquisar = os.path.join(self.caminho_imagens, 'pesquisar.png')
        self.img_consignado_teste = os.path.join(self.caminho_imagens, 'consignado_teste.png')
        self.img_menu_aplicativos = os.path.join(self.caminho_imagens, 'menu_aplicativos.png')
        self.img_operacoes_credito = os.path.join(self.caminho_imagens, 'operacoes_credito.png')
        self.img_operacao_credito = os.path.join(self.caminho_imagens, 'operacao_credito.png')
        self.img_operacao_rural = os.path.join(self.caminho_imagens, 'operacao_rural.png')
        self.img_antecipacao = os.path.join(self.caminho_imagens, 'antecipacao.png')
        self.img_cheque = os.path.join(self.caminho_imagens, 'cheque_especial.png')
        self.img_guardachuva = os.path.join(self.caminho_imagens,'guarda_chuva.png')
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
        self.img_btn_ok_2 = os.path.join(self.caminho_imagens, 'btn_ok_proposta.png')
        self.img_fechar_sistema = os.path.join(self.caminho_imagens, 'btn_fechar.png')
        self.img_fase_estudo = os.path.join(self.caminho_imagens, 'fase_estudo.png')
        self.img_fase_garantia = os.path.join(self.caminho_imagens, 'fase_garantia.png')
        self.img_fase_proposta = os.path.join(self.caminho_imagens, 'fase_proposta.png')
        self.img_fase_documentacao = os.path.join(self.caminho_imagens, 'fase_documentacao.png')
        self.img_abrir_proposta = os.path.join(self.caminho_imagens,'btn_abrir.png')
        self.img_emprestimo_app = os.path.join(self.caminho_imagens, 'emprestimo_teste.png')
        self.img_crural_app = os.path.join(self.caminho_imagens, 'crural.png')
        self.img_concessao_app = os.path.join(self.caminho_imagens, 'concessao_teste.png')
        self.img_guardachuva = os.path.join(self.caminho_imagens,'guarda_chuva.png')
        self.img_cancelar = os.path.join(self.caminho_imagens, 'btn_cancelado.png')
        self.img_justificativa = os.path.join(self.caminho_imagens, 'justificativa.png')
        self.img_motivo = os.path.join(self.caminho_imagens, 'motivo.png')
        self.img_limpar = os.path.join(self.caminho_imagens, 'limpar.png')
        self.btn_sim = os.path.join(self.caminho_imagens, 'btn_sim.png')
    
    def open_consignado(self):
            try:
                pyautogui.PAUSE = 1
                # Pesquisar consignado (teste)
                # self.locate_image(self.img_pesquisar)
                # pyautogui.write('consignado')
                # time.sleep(2)
                self.locate_image(self.img_consignado)
                time.sleep(4)

                # Clicar menu "OPERAÇÕES DE CRÉDITO"
                self.locate_image(self.img_operacao_credito)
                time.sleep(2)

                # Acessar sub-menu Mesa de Operações
                self.locate_image(self.img_mesa_operacoes)
                time.sleep(2)
                self.escolher_datas(90,30)
                self.escolher_fase()                

            except Exception as e:
                print(f"Erro ao realizar a automação!", e) 

    def cancelar(self):
        try:
            pyautogui.PAUSE = 2
            # Posicionar mouse, clicar no botão "CANCELAR"
            self.locate_image(self.img_cancelar)
            time.sleep(1)

            self.locate_image(self.img_motivo)
            pyautogui.write('outros')
            time.sleep(1)

            self.locate_image(self.img_justificativa)
            pyautogui.write('PROPOSTA CANCELADA, MAIS DE 30 DIAS NA MESA DE OPERACOES')
            time.sleep(1)
        
        except Exception as e:
            print(f"Erro ao realizar a automação!", e) 

    def acessar_plt_cred(self):
        try:
            pyautogui.PAUSE = 2
            # # Posicionar mouse, digitar "PLATAFORMA DE CREDITO" e clicar
            # self.locate_image(self.img_btn_pesquisar)
            # time.sleep(1)
            # pyautogui.write('PLATAFORMA DE CREDITO')
            # time.sleep(1)
            self.locate_image(self.img_credito)
            time.sleep(3)

        except Exception as e:
            print(f"Erro ao acessar a plataforma!", e) 

    def login_sisbr(self, usuario, senha):
            pyautogui.PAUSE = 2
            # Abrir SisBR
            pyautogui.press('winleft')
            pyautogui.write('sis')
            pyautogui.press('enter')
            time.sleep(10)
                
            # Posicionar mouse e digitar usuário            
            self.locate_image(self.img_user)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(usuario)

            # Posicionar mouse e digitar senha
            self.locate_image(self.img_senha)
            pyautogui.write(senha) 
            pyautogui.press('enter')
            time.sleep(10)

    def escolher_datas(self, dia_inicial, dia_final):
        time.sleep(2)
        hoje = dt.today()
        self.d30_dias = hoje-timedelta(days=dia_final)
        self.d90_dias_d30 = self.d30_dias-timedelta(days=dia_inicial)

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

    # Função para selecionar as fases, verificar se há proposta e abrir, após realizar o cancelamento
    def escolher_fase(self):
        pyautogui.PAUSE = 2

        fases = [self.img_proposta, self.img_documentacao, self.img_garantia, self.img_estudo]
        for fase in fases:
            self.locate_image(self.img_abrir_fase)
            self.locate_image(fase)
            self.locate_image(self.img_procurar)
            time.sleep(2)
            self.locate_image(self.img_btn_ok_2)
            print(f'Está na fase {fase}')
            
            while True:
                # Capturar a tela para verificar a presença de propostas
                screen_path = os.path.join(self.caminho_dir,'imagem_referencia.png')
                self.capture_region(440, 470, 1032, 339, screen_path)  # Ajustar a resolução conforme necessário
                
                # Extrair texto da imagem da tela
                screen_text_proposta = self.extract_text_from_image(screen_path, config="--psm 6 -l por --dpi 2500 -c tessedit_char_whitelist= 0123456789.,-abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")
                
                # Verificar se o texto extraído contém a palavra-chave da proposta
                if "Proposta" in screen_text_proposta :                            
                    print("Proposta encontrada na fase de Proposta, processando...")               
                    # Posicionar mouse e abrir a proposta
                    self.locate_image(self.img_fase_proposta)  
                    self.locate_image(self.img_abrir_proposta) 
                    time.sleep(5) 

                    screen_path_proposta = os.path.join(self.caminho_dir,'proposta_documentacao.png')
                    self.capture_region(241, 159, 1324, 757, screen_path_proposta)  

                    self.cancelar()  
                    self.locate_image(self.img_btn_ok)
                    time.sleep(2) 
                    self.locate_image(self.btn_sim)
                    time.sleep(4) 
                    self.locate_image(self.img_btn_ok_2)
                    time.sleep(2)
                    self.locate_image(self.img_fechar_sistema) 
                    pyautogui.press('esc')
                    # break                

                elif "Documentação" in screen_text_proposta or "Documentagao" in screen_text_proposta:              
                    start_time = time.time()
                    start_time_str = dt.now().strftime("%d/%m/%Y %H:%M:%S")

                    print("Proposta encontrada na fase de Documentação, processando...") 
                    self.locate_image(self.img_fase_documentacao)
                    self.locate_image(self.img_abrir_proposta)
                    time.sleep(5)                     
                    
                    screen_path_proposta = os.path.join(self.caminho_dir,'proposta_documentacao.png')
                    self.capture_region(241, 159, 1324, 757, screen_path_proposta)  
                    
                    self.cancelar()  
                    self.locate_image(self.img_btn_ok)
                    time.sleep(2) 
                    self.locate_image(self.btn_sim)
                    time.sleep(4) 
                    self.locate_image(self.img_btn_ok_2)
                    time.sleep(2)
                    self.locate_image(self.img_fechar_sistema) 
                    pyautogui.press('esc')

                    # Capturar o horário de fim do processamento da proposta
                    end_time = time.time()
                    end_time_str = dt.now().strftime("%d/%m/%Y %H:%M:%S")

                    # Calcular o tempo gasto no processamento da proposta
                    time_spent = end_time - start_time

                    # Converter o tempo gasto para o formato HH:MM:SS
                    time_spent_str = time.strftime("%H:%M:%S", time.gmtime(time_spent))

                    # Capturar textos e salvar
                    self.processar_imagem_e_extrair(screen_path_proposta, os.path.join(self.caminho_dir,'dados_extraidos.xlsx'), 'Documentação', start_time_str, end_time_str, time_spent_str)
                    time.sleep(7)
                    # break 

                elif "Garantia" in screen_text_proposta:
                    start_time = time.time()
                    start_time_str = dt.now().strftime("%d/%m/%Y %H:%M:%S")
                    
                    print("Proposta encontrada na fase de Garantia, processando...") 
                    self.locate_image(self.img_fase_garantia)
                    self.locate_image(self.img_abrir_proposta)
                    time.sleep(5)
                    screen_path_proposta = os.path.join(self.caminho_dir,'proposta_garantia.png')
                    self.capture_region(241, 159, 1324, 757, screen_path_proposta)
                    time.sleep(3)
                    
                    self.cancelar()   
                    self.locate_image(self.img_btn_ok) 
                    time.sleep(2)
                    self.locate_image(self.btn_sim)
                    time.sleep(4) 
                    self.locate_image(self.img_btn_ok_2)
                    time.sleep(2)
                    self.locate_image(self.img_fechar_sistema)
                    pyautogui.press('esc')
                    
                    # Capturar o horário de fim do processamento da proposta
                    end_time = time.time()
                    end_time_str = dt.now().strftime("%d/%m/%Y %H:%M:%S")

                    # Calcular o tempo gasto no processamento da proposta
                    time_spent = end_time - start_time

                    # Converter o tempo gasto para o formato HH:MM:SS
                    time_spent_str = time.strftime("%H:%M:%S", time.gmtime(time_spent))

                    self.processar_imagem_e_extrair(screen_path_proposta, os.path.join(self.caminho_dir,'dados_extraidos.xlsx'), 'Garantia', start_time_str, end_time_str, time_spent_str) 
                    time.sleep(7)
                    # break 
                
                elif "Estudo" in screen_text_proposta:
                    start_time = time.time()
                    start_time_str = dt.now().strftime("%d/%m/%Y %H:%M:%S")
                    
                    print("Proposta encontrada na fase de Estudo, processando...") 
                    self.locate_image(self.img_fase_estudo)
                    self.locate_image(self.img_abrir_proposta)
                    time.sleep(5) 
                    screen_path_proposta = os.path.join(self.caminho_dir,'proposta_estudo.png')
                    self.capture_region(241, 159, 1324, 757, screen_path_proposta) 
                    time.sleep(3)
                    
                    self.cancelar()
                    time.sleep(3)
                    self.locate_image(self.img_btn_ok)
                    time.sleep(2) 
                    self.locate_image(self.btn_sim)
                    time.sleep(4) 
                    self.locate_image(self.img_btn_ok_2)
                    time.sleep(2)
                    self.locate_image(self.img_fechar_sistema)
                    pyautogui.press('esc')
                    
                    # Capturar o horário de fim do processamento da proposta
                    end_time = time.time()
                    end_time_str = dt.now().strftime("%d/%m/%Y %H:%M:%S")

                    # Calcular o tempo gasto no processamento da proposta
                    time_spent = end_time - start_time

                    # Converter o tempo gasto para o formato HH:MM:SS
                    time_spent_str = time.strftime("%H:%M:%S", time.gmtime(time_spent))

                    self.processar_imagem_e_extrair(screen_path_proposta, os.path.join(self.caminho_dir,'dados_extraidos.xlsx'), 'Estudo', start_time_str, end_time_str, time_spent_str) 
                    time.sleep(7) 
                    # break 

                else:
                    print("Nenhuma proposta para analisar, indo pra próxima fase...")
                    break 

    def processar_imagem_e_extrair(self, image_path, output_path, fase = None, start_time = None, end_time = None, time_elapsed = None):
        plt = "Consignado"
        
        # Extrair texto da imagem
        text = self.extract_text_from_image(image_path)

        # Extrair valores específicos do texto
        cpf, valor_proposta, pa, numero_proposta = self.extrair_textos_especificos(text)

        # Salvar os dados extraídos
        self.salvar_extracao_dados(output_path, cpf, valor_proposta,  pa, numero_proposta, plt, fase, start_time, end_time, time_elapsed)   

    def extrair_textos_especificos(self, text):
        # Normalizar texto
        text = text.replace('\n', ' ').replace('\r', ' ').strip()
        
        # Inicializar variáveis
        cpf = valor_contrato =  pa = numero_proposta = None

        # Extrair CPF/CNPJ
        cpf_match = re.search(r'\d{3}\.\d{3}\.\d{3}-\d{2}', text)
        if cpf_match:
            cpf = cpf_match.group(0)

        # Extrair Valor Contratado
        valor_match = re.search(r'R\$ ?\d{1,3}(\.\d{3})*,\d{2}', text)
        if valor_match:
            valor_contrato = valor_match.group(0)

        
        # Extrair PA (busca mais específica)
        pa_match = re.search(r'\bPA[:\s]*(\d{2,})\b|(?<=\s)(\d{2})(?=\s83388)', text, re.IGNORECASE)
        if pa_match:
            pa = pa_match.group(1) or pa_match.group(2)  # Captura o número correto
            pa = pa.strip()


        # Extrair Nº Contrato
        contrato_match = re.search(r'\b\d{6,}\b', text)
        if contrato_match:
            numero_proposta = contrato_match.group(0).strip()        
        
        print(f'CPF:{cpf}; Valor:{valor_contrato}; PA:{pa}; Contrato:{numero_proposta}')

        return cpf, valor_contrato,  pa, numero_proposta

    def salvar_extracao_dados(self, output_path, cpf, valor_contrato,  pa, numero_proposta, plt, fase, start_time, end_time, time_elapsed):
        today_date = dt.now().strftime("%d/%m/%Y %H:%M:%S")
        results = [(today_date, cpf, valor_contrato,  pa, numero_proposta, plt, fase, start_time, end_time, time_elapsed)]

        # Criar DataFrame
        df = pd.DataFrame(results, columns=[
            'Data Extração', 'CPF/CNPJ', 'Valor Proposta',
            'PA', 'Nº Proposta', 'Plataforma', 'Fase', 
            'Tempo Início', 'Tempo Fim', 'Tempo Gasto'
        ])

        try:
            # Verificar se o arquivo já existe para adicionar as novas entradas
            existing_df = pd.read_excel(output_path)
            df = pd.concat([existing_df, df], ignore_index=True)
        except FileNotFoundError:
            print(f'O arquivo {output_path} não foi encontrado. Criando um novo arquivo.')

        try:
            # Salvar o DataFrame atualizado (ou novo) no arquivo Excel
            df.to_excel(output_path, index=False)
            print(f'Dados salvos com sucesso em {output_path}')
        except Exception as e:
            print(f'Erro ao salvar os dados: {e}')

    # Função para fechar janela do sistema   
    def fechar_sistemas(self, qtd):
        for _ in range(qtd):  # Supondo que há 2 abas a serem fechadas
            pyautogui.hotkey('alt', 'f4') # comando para fechar a aba
            time.sleep(4)  # Pequena pausa entre os fechamentos
        self.locate_image(self.img_fechar_sistema) 


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
        cv2.imwrite(os.path.join(self.caminho_dir,'imagem_processada.png'), image)
        return image
    
    def capture_region(self, x, y, width, height, output_path):
        # Captura a região especificada da tela
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        # Salva a imagem capturada
        screenshot.save(output_path)
        print(f"Screenshot salvo como {output_path}")


if __name__=='__main__':
    cred = CancelCred()
    cred.login_sisbr('usrcanpropost3120_00', 'Vostro%3120')
    cred.acessar_plt_cred()
    cred.open_consignado()
    cred.fechar_sistemas(2)