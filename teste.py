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
    def testando(self):
        try:
            # image = Image.open(os.path.join(self.caminho_dir,'proposta.png' ))
            # boxes = pytesseract.image_to_osd(image)

            # for box in boxes.splitlines():
            #     print(box)
            output_path_img_proposta = os.path.join(r'C:\Users\sicoob\Sicoob Central Crediminas\3120 - Business Intelligence (B.I) - Geral\Automacoes\CancelamentoCredito','imagem_referencia.png' )
            extracted_text = self.extract_text_from_image(output_path_img_proposta, config=" -l eng+equ+por --psm 6 -c tessedit_char_whitelist= 0123456789.,")
            print(extracted_text)

            self.extrair_textos_especificos(extracted_text)
            
            # self.processar_imagem_e_extrair(output_path_img_proposta, 'dados_extraidos.xlsx')

            
        except Exception as e:
            print("Erro ao tentar realizar a automação, reveja os processos", e) 

    def processar_imagem_e_extrair(self, image_path, output_path):
        # Início do tempo de processamento
        start_time = dt.now()

        # Extrair texto da imagem
        text = self.extract_text_from_image(image_path)

        # Extrair valores específicos do texto
        cpf, valor_proposta, pa, numero_proposta = self.extrair_textos_especificos(text)

        # Fim do tempo de processamento
        end_time = dt.now()
        time_elapsed = (end_time - start_time).total_seconds()

        # Salvar os dados extraídos
        self.salvar_extracao_dados(output_path, cpf, valor_proposta,  pa, numero_proposta, start_time, end_time, time_elapsed)

    def extrair_textos_especificos(self, text):
        # Normalizar texto
        text = text.replace('\n', ' ').replace('\r', ' ').replace('oo', '00').strip()
                
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

        
       # Extrair PA (corrigido para capturar "00" e "08")
        pa_match = re.search(r'\bPA[\s]*([0-9]{2,})\b', text, re.IGNORECASE)
        if pa_match:
            pa = pa_match.group(1).zfill(2)


        # Extrair Nº Contrato
        contrato_match = re.search(r'\b\d{8,}\b', text)
        if contrato_match:
            numero_proposta = contrato_match.group(0).strip()     
        
        print(f'CPF:{cpf}; Valor:{valor_contrato}; PA:{pa}; Contrato:{numero_proposta}')

        return cpf, valor_contrato,  pa, numero_proposta

    def salvar_extracao_dados(self, output_path, cpf, valor_contrato,  pa, numero_proposta, start_time, end_time, time_elapsed):
        today_date = dt.now().strftime("%d/%m/%Y %H:%M:%S")
        results = [(today_date, cpf, valor_contrato,  pa, numero_proposta, start_time, end_time, time_elapsed)]

        # Criar DataFrame
        df = pd.DataFrame(results, columns=[
            'Data Extração', 'CPF/CNPJ', 'Valor Proposta',
            'PA', 'Nº Proposta', 'Tempo Início', 'Tempo Fim', 'Tempo Gasto'
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
    cred.testando()