import subprocess

scripts = ['cancelamento_proposta_consignado.py', 'cancelamento_proposta_emprestimo.py', 'cancelamento_proposta_limites.py','cancelamento_proposta_rural.py'] #'',cancelamento_proposta_emprestimo.py 

for script in scripts:
    print(f'Executando {script}...')
    subprocess.run(["python", script])