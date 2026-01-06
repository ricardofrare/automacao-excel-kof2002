import pandas as pd
import os

# 1. Configuração do caminho de entrada
caminho_entrada = r"C:\Users\ricar\Downloads\Kof 2k2 um 2019 Feb japanese ratio tier list.xlsx"
df_kof = pd.read_excel(caminho_entrada, header=1)

# 2. Lógica para salvar na mesma pasta do script
# os.path.dirname(__file__) pega o caminho da pasta onde este arquivo .py está salvo
diretorio_atual = os.path.dirname(os.path.abspath(__file__))
caminho_saida = os.path.join(diretorio_atual, "Kof_2k2_Rankeado.xlsx")

# --- Processamento ---

# Criando a coluna de desempate
df_kof['total_score'] = df_kof['point'] + df_kof['mid'] + df_kof['anchor']

# Ordenação (Maior anchor primeiro, depois maior total_score)
df_sorted = df_kof.sort_values(
    by=['anchor', 'total_score'], 
    ascending=[False, False]
).copy()

# Lista para armazenar as categorias
classificacoes = []

# Loop de Iteração e Lógica Condicional (Desafio 1)
for index, row in df_sorted.iterrows():
    val = row['anchor']
    
    if val > 7:
        rank = "Rank S"
    elif val == 7:
        rank = "Rank A"
    elif val == 6:
        rank = "Rank B"
    elif val == 5:
        rank = "Rank C"
    else:
        rank = "Rank D"
    
    classificacoes.append(rank)

# Adicionando a nova coluna
df_sorted['Classificacao'] = classificacoes

# Removendo coluna auxiliar e salvando
df_final = df_sorted.drop(columns=['total_score'])
df_final.to_excel(caminho_saida, index=False)

print(f"Sucesso! Arquivo salvo em: {caminho_saida}")