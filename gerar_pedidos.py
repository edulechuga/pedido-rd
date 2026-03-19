import pandas as pd
import os

# Nomes dos arquivos de entrada
pedidos_filename = 'Pedidos Aberturas - RD Saúde - 24.09 - CUSTOM.xlsx'
rqcm_template_filename = 'RQCM-001 - Pedido de venda - V05.xlsx'
output_directory = 'pedidos_gerados'

try:
    # Crie o diretório de saída se ele não existir
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Leia o arquivo de pedidos
    pedidos_file = pd.read_csv(pedidos_filename)

    # Leia o conteúdo do template RQCM
    with open(rqcm_template_filename, 'r', encoding='utf-8') as f:
        rqcm_template_lines = f.readlines()

    # Extraia os valores da 55ª coluna (índice 54)
    observacoes = pedidos_file.iloc[:, 54].tolist()

    # Gere um novo arquivo para cada observação
    for i, obs in enumerate(observacoes):
        new_rqcm_lines = list(rqcm_template_lines)

        # Encontre e substitua a linha de observação
        for j, line in enumerate(new_rqcm_lines):
            if line.strip().startswith('OBS.:'):
                new_rqcm_lines[j] = f'OBS.:,"{obs}"\n'
                break
        
        # Defina o nome do novo arquivo
        output_filename = os.path.join(output_directory, f'RQCM-001_Pedido_{i+1}.csv')
        
        # Escreva o novo arquivo
        with open(output_filename, 'w', encoding='utf-8') as f:
            f.writelines(new_rqcm_lines)

    print(f"Sucesso! Foram gerados {len(observacoes)} arquivos na pasta '{output_directory}'.")

except FileNotFoundError as e:
    print(f"Erro: Arquivo não encontrado. Verifique se os arquivos '{pedidos_filename}' e '{rqcm_template_filename}' estão na mesma pasta que o script.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")