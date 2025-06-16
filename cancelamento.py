import subprocess
import getpass
import os
import logging

scripts = ['cancelamento_proposta_consignado.py', 'cancelamento_proposta_emprestimo.py', 'cancelamento_proposta_limites.py','cancelamento_proposta_rural.py'] #'',cancelamento_proposta_emprestimo.py 
usuario_logado = getpass.getuser()
caminho_log = rf'C:\Users\{usuario_logado}\Sicoob Central Crediminas\3120 - Business Intelligence (B.I) - Geral\Automacoes\Logs'

logging.basicConfig(
            filename=os.path.join(caminho_log,'cancelamento_propostas.log'),
            filemode='a',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%d/%m/%Y %H:%M:%S')

for script in scripts:
    print(f'Executando {script}...')
    subprocess.run(["python", script])