# -*- coding: utf-8 -*-

import pandas as pd
import os

# Altera o diretório de trabalho para o diretório onde o script está
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Função para processar e gerar a nova planilha
def processar_planilha(input_file, output_file):
    # Ler a planilha de entrada
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return
    
    # Verificar se as colunas relevantes existem
    colunas_necessarias = ['Contato', 'Identificador', 'Protocolo', 'Data da mensagem', 'Agente', 'Mensagem']
    
    if not all(coluna in df.columns for coluna in colunas_necessarias):
        print("A planilha de entrada não contém todas as colunas necessárias.")
        return
    
    # Filtrar as colunas relevantes
    df_relevante = df[colunas_necessarias]
    
    # Gerar a nova planilha de saída
    try:
        df_relevante.to_excel(output_file, index=False)
        print(f"Nova planilha gerada com sucesso: {output_file}")
    except Exception as e:
        print(f"Erro ao salvar a planilha: {e}")

# Caminho para a planilha de entrada e saída
input_file = '20250326130107233696_1.xlsx'  # Altere para o caminho da sua planilha de entrada
output_file = 'atendimento_processado.xlsx'  # Altere para o caminho da planilha de saída

# Chamar a função para processar
processar_planilha(input_file, output_file)
