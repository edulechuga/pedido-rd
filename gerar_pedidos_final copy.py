import pandas as pd
import openpyxl
import os
import re

# --- Configuração ---
# Nomes dos arquivos de entrada
pedidos_filename = 'Pedidos Aberturas - CUSTOM - 05.12.xlsx'
rqcm_template_filename = 'RQCM-001 - Pedido de venda - V05.xlsx'

# Nome da pasta onde os arquivos gerados serão salvos
output_directory = 'pedidos_gerados_excel'

# Célula exata a ser preenchida no template
target_cell = 'B17'
# --------------------

def sanitize_filename(filename):
    """
    Remove caracteres inválidos de uma string para que ela possa ser usada como nome de arquivo.
    """
    # Substitui caracteres inválidos por um espaço
    return re.sub(r'[\\/*?:"<>|]', " ", str(filename))

try:
    # 1. Cria o diretório de saída se ele não existir
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # 2. Lê a planilha de pedidos
    print("Lendo o arquivo de pedidos...")
    pedidos_df = pd.read_excel(pedidos_filename)

    # 3. Extrai os valores da 58ª coluna (índice 57, que corresponde à coluna BF)
    observacoes = pedidos_df.iloc[:, 57].fillna('').tolist()
    print(f"Encontradas {len(observacoes)} linhas para processar.")

    # 4. Gera um novo arquivo para cada observação
    for i, obs in enumerate(observacoes):
        # Carrega o arquivo de modelo do Excel
        workbook = openpyxl.load_workbook(rqcm_template_filename)
        sheet = workbook.active

        # Escreve a observação diretamente na célula B17
        sheet[target_cell] = obs
        
        # Limpa a observação para usar como nome de arquivo
        safe_obs_name = sanitize_filename(obs)
        
        # Define o nome do novo arquivo de saída, incluindo a observação
        # Limita o comprimento para evitar nomes de arquivo excessivamente longos
        output_filename = os.path.join(output_directory, f'RQCM-001_Pedido_{safe_obs_name[:50]}_{i+1}.xlsx')
        
        # Salva o novo arquivo Excel
        workbook.save(output_filename)
        print(f"Arquivo {i+1} de {len(observacoes)} gerado: {output_filename}")

    print("\n--- Processo Concluído ---")
    print(f"Sucesso! Foram gerados {len(observacoes)} arquivos na pasta '{output_directory}'.")

except FileNotFoundError:
    print(f"\nERRO: Arquivo não encontrado.")
    print(f"Verifique se os nomes dos arquivos estão corretos e se eles estão na mesma pasta que o script.")
    print(f"Nomes esperados: '{pedidos_filename}' e '{rqcm_template_filename}'")
except Exception as e:
    print(f"\nOcorreu um erro inesperado: {e}")