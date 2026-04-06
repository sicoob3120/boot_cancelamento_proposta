import pyautogui
import time
import os
from datetime import timedelta
from datetime import datetime as dt
import pytesseract
import pandas as pd
import cv2
import re
import logging
import socket
import getpass
from dotenv import load_dotenv

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from gerenciamento_senha.senha_sisbr import TokenEquipamentos

class CancelCred:
    def __init__(self):

        self.usuario_logado = getpass.getuser()

        self.dotenv_path = os.path.join(
            os.environ['USERPROFILE'], 'OneDrive - Sicoob Central Crediminas', '3120 - Business Intelligence (B.I) - Geral', 'Automacoes', '.env'
        )

        load_dotenv(dotenv_path=self.dotenv_path)

        self.usuario = os.getenv("USUARIO_CANPROPOST")
        self.senha = os.getenv("SENHA_CANPROPOST")

        pyautogui.PAUSE = 2
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        self.caminho_imagens = rf'C:\Users\{self.usuario_logado}\OneDrive - Sicoob Central Crediminas\3120 - Business Intelligence (B.I) - Geral\Automacoes\Imagens_RPA\cancelamentocredito'
        self.caminho_dir = rf'C:\Users\{self.usuario_logado}\OneDrive - Sicoob Central Crediminas\3120 - Business Intelligence (B.I) - Geral\Automacoes\__arquivos\Cancelamento_Propostas'
        self.caminho_log = os.path.join(os.environ['USERPROFILE'], 'OneDrive - Sicoob Central Crediminas','3120 - Business Intelligence (B.I) - Geral', 'Automacoes','Logs')

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
        self.img_recebiveis = os.path.join(self.caminho_imagens, 'antecipacao_recebiveis.png')
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
        self.img_fechar_sistema = os.path.join(self.caminho_imagens, 'btn_fechar.png')
        self.img_fase_estudo = os.path.join(self.caminho_imagens, 'fase_estudo.png')
        self.img_fase_garantia = os.path.join(self.caminho_imagens, 'fase_garantia.png')
        self.img_fase_proposta = os.path.join(self.caminho_imagens, 'fase_proposta.png')
        self.img_fase_documentacao = os.path.join(self.caminho_imagens, 'fase_documentacao.png')
        self.img_abrir_proposta = os.path.join(self.caminho_imagens,'btn_abrir.png')
        self.img_emprestimo_app = os.path.join(self.caminho_imagens, 'emprestimo_teste.png')
        self.img_crural_app = os.path.join(self.caminho_imagens, 'crural.png')
        self.img_concessao_app = os.path.join(self.caminho_imagens, 'climites.png')
        self.img_guardachuva = os.path.join(self.caminho_imagens,'guarda_chuva.png')
        self.img_cancelar = os.path.join(self.caminho_imagens, 'btn_cancelado.png')
        self.img_justificativa = os.path.join(self.caminho_imagens, 'justificativa.png')
        self.img_motivo = os.path.join(self.caminho_imagens, 'motivo.png')
        self.img_limpar = os.path.join(self.caminho_imagens, 'limpar.png')
        self.img_btn_ok_2 = os.path.join(self.caminho_imagens, 'btn_ok_proposta.png')
        self.btn_sim = os.path.join(self.caminho_imagens, 'btn_sim.png')

        self.troca_senha = TokenEquipamentos(
            usuario_key="USUARIO_CANPROPOST",
            senha_key="SENHA_CANPROPOST",
            dotenv_path=self.dotenv_path,
            modo_teste=False
        )

        logging.basicConfig(
            filename=os.path.join(self.caminho_log,'cancelamento_propostas.log'),
            filemode='a',
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%d/%m/%Y %H:%M:%S',
            force=True
        )

        logging.info("Logger inicializado com sucesso.")


    def get_local_ip(self):
        try:
            hostname = socket.gethostname()
            ip_address = socket.gethostbyname(hostname)
            return ip_address
        except Exception as e:
            logging.error("Erro ao obter IP da máquina", exc_info=True)
            return "IP não identificado"

    def open_limites(self):
            try:
                pyautogui.PAUSE = 1
                # Pesquisar emprestimo (teste)
                # self.locate_image(self.img_pesquisar)
                # pyautogui.write('conce')
                # time.sleep(2)
                self.locate_image(self.img_concessao_app)
                time.sleep(4)

                # Clicar menu "ANTECIPAÇÃO"
                self.locate_image(self.img_antecipacao)
                time.sleep(2)

                # Clicar menu "ANTECIPAÇÃO DE RECEBÍVEIS"
                self.locate_image(self.img_recebiveis)
                time.sleep(2)

                # Acessar sub-menu Mesa de Operações
                self.locate_image(self.img_mesa_operacoes)
                time.sleep(2)
                self.escolher_datas(90,30)
                self.escolher_fase()

                # # Clicar menu "CHEQUE ESPECIAL"
                # self.locate_image(self.img_cheque)
                # time.sleep(2)

                # # Acessar sub-menu Mesa de Operações
                # self.locate_image(self.img_mesa_operacoes)
                # time.sleep(2)
                # self.escolher_datas(90,10)
                # self.escolher_fase()

                # # Clicar menu "GUARDA CHUVA"
                # self.locate_image(self.img_guardachuva)
                # time.sleep(2)

                # # Acessar sub-menu Mesa de Operações
                # self.locate_image(self.img_mesa_operacoes)
                # time.sleep(2)
                # self.escolher_datas(90,10)
                # self.escolher_fase()

            except Exception as e:
                logging.error(f"Erro ao realizar a automação!", e)

    def cancelar(self):
        try:
            pyautogui.PAUSE = 2
            # Posicionar mouse, clicar no botão "CANCELAR"
            self.locate_image(self.img_cancelar)
            time.sleep(1)

            self.locate_image(self.img_motivo)
            pyautogui.write('REALIZAÇÃO')
            time.sleep(1)

            self.locate_image(self.img_justificativa)
            pyautogui.write('PROPOSTA CANCELA, MAIS DE 30 DIA NA MESA DE OPERACOES')
            time.sleep(1)

        except Exception as e:
            logging.error(f"Erro ao realizar a automação!", e)

    def acessar_plt_cred(self):
        try:
            pyautogui.PAUSE = 2
            # Posicionar mouse, digitar "PLATAFORMA DE CREDITO" e clicar
            # self.locate_image(self.img_btn_pesquisar)
            # time.sleep(1)
            # pyautogui.write('PLATAFORMA DE CREDITO')
            # time.sleep(1)
            self.locate_image(self.img_credito)
            time.sleep(2)

        except Exception as e:
            logging.error(f"Erro ao acessar a plataforma!", e)

    def login_sisbr(self):
            pyautogui.PAUSE = 2
            # Abrir SisBR
            pyautogui.press('winleft')
            pyautogui.write('sis')
            pyautogui.press('enter')
            time.sleep(10)

            # Posicionar mouse e digitar usuário
            self.locate_image(self.img_user)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(self.usuario)
            time.sleep(2)

            # Posicionar mouse e digitar senha
            self.locate_image(self.img_senha)
            pyautogui.write(self.senha)
            pyautogui.press('enter')
            time.sleep(7)

            # Aguarda aparecer mensagem de senha expirada
            end = time.time() + 20
            expirada = False
            while time.time() < end:
                try:
                    achou_expirada = pyautogui.locateOnScreen(self.troca_senha.img_senha_expirada, confidence=0.8)
                except pyautogui.ImageNotFoundException:
                    achou_expirada = None

                if achou_expirada is not None:
                    expirada = True
                    break
                time.sleep(0.5)

            # Se detectou expiração, chama trocar_senha e tenta logar novamente com a nova senha
            if expirada:
                logging.info("Tela senha expirada detectada. Realizando a troca...")
                self.troca_senha.trocar_senha()

                load_dotenv(dotenv_path=self.dotenv_path, override=True)
                self.senha = os.getenv("SENHA_CANPROPOST")

                self.locate_image(self.img_user)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.write(self.usuario)
                time.sleep(1)

                self.locate_image(self.img_senha)
                pyautogui.write(self.senha)
                pyautogui.press('enter')
                time.sleep(10)

            # Iniciar arquivo de LOG
            ip = self.get_local_ip()
            hora_login = dt.now().strftime("%d/%m/%Y %H:%M:%S")
            logging.info("")  # <-- adiciona linha em branco no arquivo de log
            logging.info(f"Login realizado - Usuário: {self.usuario} | IP: {ip} | Horário: {hora_login}")

    def escolher_datas(self, dia_inicial, dia_final):
        hoje = dt.today()
        self.d30_dias = hoje-timedelta(days=dia_final)
        self.d90_dias_d30 = self.d30_dias-timedelta(days=dia_inicial)

        # Data proposta inicio e final
        time.sleep(2)
        self.locate_image(self.img_data_prop_inicio)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(2)
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
            logging.info(f'Está na fase {fase}')

            while True:
                # Capturar a tela para verificar a presença de propostas
                screen_path = os.path.join(self.caminho_dir,'imagem_referencia.png')
                self.capture_region(440, 470, 1032, 339, screen_path)  # Ajustar a resolução conforme necessário

                # Extrair texto da imagem da tela
                screen_text_proposta = self.extract_text_from_image(screen_path, config="--psm 6 -l por --dpi 2500 -c tessedit_char_whitelist= 0123456789.,-abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")

                # Verificar se o texto extraído contém a palavra-chave da proposta
                if "Proposta" in screen_text_proposta :
                    logging.info("Proposta encontrada na fase de Proposta, processando...")
                    # Posicionar mouse e abrir a proposta
                    self.locate_image(self.img_fase_proposta)
                    self.locate_image(self.img_abrir_proposta)
                    time.sleep(2)

                    screen_path_proposta = os.path.join(self.caminho_dir,'proposta_documentacao_.png')
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
                    time.sleep(2)
                    self.locate_image(self.img_btn_ok_2)
                    # break

                elif "Documentação" in screen_text_proposta or "Documentagao" in screen_text_proposta:
                    start_time = time.time()
                    start_time_str = dt.now().strftime("%d/%m/%Y %H:%M:%S")

                    logging.info("Proposta encontrada na fase de Documentação, processando...")
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
                    time.sleep(2)
                    self.locate_image(self.img_btn_ok_2)

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

                    logging.info("Proposta encontrada na fase de Garantia, processando...")
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
                    time.sleep(2)
                    self.locate_image(self.img_btn_ok_2)

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

                    logging.info("Proposta encontrada na fase de Estudo, processando...")
                    self.locate_image(self.img_fase_estudo)
                    self.locate_image(self.img_abrir_proposta)
                    time.sleep(5)
                    screen_path_proposta = os.path.join(self.caminho_dir,'proposta_estudo.png')
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
                    time.sleep(2)
                    self.locate_image(self.img_btn_ok_2)

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
                    logging.info("Nenhuma proposta para analisar, indo pra próxima fase...")
                    break

    def processar_imagem_e_extrair(self, image_path, output_path, fase = None, start_time = None, end_time = None, time_elapsed = None):
        plt = "Concessão de Limites"

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

        # Extrair PA
        pa_match = re.search(r'\bPA[:\s]*([0-9]{2})\b', text, re.IGNORECASE)
        if pa_match:
            pa = pa_match.group(1).strip()

        # Extrair Nº Contrato
        contrato_match = re.search(r'\b\d{8,}\b', text)
        if contrato_match:
            numero_proposta = contrato_match.group(0).strip()

        logging.info(f'CPF:{cpf}; Valor:{valor_contrato}; PA:{pa}; Contrato:{numero_proposta}')

        return cpf, valor_contrato,  pa, numero_proposta

    def salvar_extracao_dados(self, output_path, cpf, valor_contrato,  pa, numero_proposta, plt, fase,start_time, end_time, time_elapsed):
        today_date = dt.now().strftime("%d/%m/%Y %H:%M:%S")
        results = [(today_date, cpf, valor_contrato,  pa, numero_proposta, plt, fase ,start_time, end_time, time_elapsed)]

        # Criar DataFrame
        df = pd.DataFrame(results, columns=[
            'Data Extração', 'CPF/CNPJ', 'Valor Proposta',
            'PA', 'Nº Proposta', 'Plataforma', 'Fase' ,'Tempo Início', 'Tempo Fim', 'Tempo Gasto'
        ])

        try:
            # Verificar se o arquivo já existe para adicionar as novas entradas
            existing_df = pd.read_excel(output_path)
            df = pd.concat([existing_df, df], ignore_index=True)
        except FileNotFoundError:
            logging.error(f'O arquivo {output_path} não foi encontrado. Criando um novo arquivo.')

        try:
            # Salvar o DataFrame atualizado (ou novo) no arquivo Excel
            df.to_excel(output_path, index=False)
            logging.info(f'Dados salvos com sucesso em {output_path}')
        except Exception as e:
            logging.error(f'Erro ao salvar os dados: {e}')

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
                try:
                    img_location = pyautogui.locateOnScreen(img, confidence=0.9, grayscale=grayscale)
                except pyautogui.ImageNotFoundException:
                    img_location = None

                if img_location is not None:
                    # Encontrar o centro da imagem
                    img_center = pyautogui.center(img_location)
                    # Mover o mouse para o centro da imagem e clicar
                    pyautogui.click(img_center)
                    encontrado = True
                    break
                else:
                    tentativa += 1
                    logging.error(f"Imagem {img} não encontrada na tela.")
                    time.sleep(delay)

            if encontrado and existe:
                return True
            elif not encontrado and not existe:
                return True

            return False

        except Exception as e:
            logging.error(f"Erro ao encontrar a imagem {img}!", e)

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
    cred.login_sisbr()
    cred.acessar_plt_cred()
    cred.open_limites()
    cred.fechar_sistemas(2)