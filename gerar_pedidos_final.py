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

    # Limpando espaços em branco do início e fim do nome das colunas para garantir que a busca não falhe
    pedidos_df.columns = pedidos_df.columns.str.strip()

    # 3. Extrai os valores das colunas necessárias usando busca dinâmica (pelo nome do cabeçalho)
    try:
        observacoes = pedidos_df['Observações NF'].fillna('').tolist()
        val_q_list = pedidos_df['Qtde'].fillna('').tolist()            
        val_ab_list = pedidos_df['VALOR RD - PDF'].fillna('').tolist() 
        val_ax_list = pedidos_df['Custo UNT HW + Patrimonio'].fillna('').tolist()
        val_ar_list = pedidos_df['PTAX'].fillna('').tolist()           
    except KeyError as e:
        print(f"\nERRO: Alguma coluna solicitada não existe na planilha base: {e}")
        # Encerra caso a coluna não exista e retorna o erro sem quebrar
        raise e
    
    print(f"Encontradas {len(observacoes)} linhas para processar.")

    # 4. Gera um novo arquivo para cada linha
    for i in range(len(observacoes)):
        obs = observacoes[i]
        val_q = val_q_list[i]
        val_ab = val_ab_list[i]
        val_ax = val_ax_list[i]
        val_ar = val_ar_list[i]

        # Carrega o arquivo de modelo do Excel
        workbook = openpyxl.load_workbook(rqcm_template_filename)
        sheet = workbook.active

        # Escreve a observação diretamente na célula B17
        sheet[target_cell] = obs
        
        # Preenche as novas células conforme solicitado
        sheet['A27'] = val_q
        sheet['A42'] = val_q
        sheet['D27'] = val_ab
        sheet['D42'] = val_ax
        sheet['B47'] = val_ar
        
        # Insere as fórmulas para cálculo
        sheet['E27'] = '=A27*D27'
        sheet['E42'] = '=A42*D42'
        
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