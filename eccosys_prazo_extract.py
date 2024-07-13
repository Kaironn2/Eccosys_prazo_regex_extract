import pandas as pd
import re

# Extrai o número após "média"
def extrair_numero(texto):
    if isinstance(texto, str):
        padrao = re.compile(r'média\s+(\d+)', re.IGNORECASE)
        match = padrao.search(texto)
        if match:
            return int(match.group(1))
    return None

# Caminho Planilha Eccosys
caminho_arquivo = r'C:\Users\jonat\OneDrive\Área de Trabalho\regex_prazo_eccosys\planilhas\teste.xlsx'

df = pd.read_excel(caminho_arquivo)

print(df.columns)

# Remove espaços em branco dos nomes das colunas
df.columns = df.columns.str.strip()

# Verifica se a OBSERVAÇÃO INTERNA existe e, se existir, aplica a função de extração
if 'OBSERVAÇÃO INTERNA' in df.columns:
    df['Numero_apos_media'] = df['OBSERVAÇÃO INTERNA'].apply(extrair_numero)
    # Exibe o DataFrame com a nova coluna
    print(df[['OBSERVAÇÃO INTERNA', 'Numero_apos_media']])
else:
    print("A coluna 'OBSERVAÇÃO INTERNA' não foi encontrada no arquivo Excel.")

# Caminho novo arquivo
df.to_excel('arquivo_atualizado.xlsx', index=False)
