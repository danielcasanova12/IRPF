import pandas as pd
import re

def calcular_preco_medio(arquivo):
    df = pd.read_excel(arquivo)
    df.loc[df['Entrada/Saída'] == 'Debito', 'Quantidade'] *= -1
    df['Preço unitário'] = df['Preço unitário'].replace('-', '0').astype(float)
    df['Valor da Operação'] = df['Quantidade'] * df['Preço unitário']
    
    # df = df[df.groupby('Produto')['Quantidade'].transform('sum') >= 0]
    df = df[df.groupby('Produto')['Quantidade'].transform('sum') > 0]

    # Calcular o preço médio para cada ativo
    df['Preço Médio'] = df.groupby('Produto')['Valor da Operação'].transform('sum') / df.groupby('Produto')['Quantidade'].transform('sum')
    
    df['Quantidade_total'] = df.groupby('Produto')['Quantidade'].transform('sum')
    # Remover linhas duplicadas
    df = df.drop_duplicates(subset='Produto')
    df.drop('Entrada/Saída', axis=1, inplace=True)
    df.drop('Data', axis=1, inplace=True)
    df.drop('Movimentação', axis=1, inplace=True)
    df.drop('Instituição', axis=1, inplace=True)
    df.drop('Preço unitário', axis=1, inplace=True)
    df.drop('Quantidade', axis=1, inplace=True)
    df.drop('Valor da Operação', axis=1, inplace=True)
    df['Produto'] = df['Produto'].str.split(' ').str[0]

    return df

def formatar_mensagem(row):
    return f"{row['Produto']} quantidade {row['Quantidade_total']} e com preço médio de R$ {row['Preço Médio']:.2f}"



def completar_tabela(arquivo, resultado):  
    df = pd.read_excel(arquivo)
    df['Produto'] = df['Produto'].str.split(' ').str[0]
    # df['CNPJ da Empresa'] = df['CNPJ da Empresa']
    
    df['CNPJ da Empresa'] = df['CNPJ da Empresa']

    # Merge dos dois DataFrames
    # resultado_final = pd.merge(resultado, df[['Produto', 'CNPJ da Empresa']], on='Produto', how='left')
    resultado_final = pd.merge(resultado, df[['Produto', 'CNPJ da Empresa', 'Valor Atualizado']], on='Produto', how='left')
    
    return resultado_final

arquivo = 'movimentacao-2024-03-29-12-19-54.xlsx'
resultado = calcular_preco_medio(arquivo)

arquivo2 = 'relatorio-consolidado-anual-2023.xlsx'
resultado_final = completar_tabela(arquivo2, resultado)

resultado_final['Mensagem'] = resultado_final.apply(formatar_mensagem, axis=1)

resultado_final.to_excel('resultado3.xlsx', index=False)