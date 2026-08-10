'''
Dados só do Ensino Médio (todos os 3 anos);
Considerar somente ensino regular;
Recortes temporais de 2º trimestre (Abril, Maio e Junho) para 2025 e 2026;
Considerar somente 14 componentes da BNCC:
- "Arte",
- "Biologia",
- "Educação Física",
- "Ensino Religioso",
- "Filosofia",
- "Física",
- "Geografia",
- "História",
- "Língua Espanhola",
- "Língua Inglesa",
- "Língua Portuguesa",
- "Matemática",
- "Química",
- "Sociologia"


Indicadores de Fluxo e Mobilização:
- Média de frequência dos estudantes;
- % de estudantes com frequência abaixo de 75%;
- % das frequências preenchidas (quantidade de frequências lançadas);
- % estudantes em abandono escolar;
- risco de abandono (conforme métrica do próprio estado);
- % escolas com conselho de classe realizada;
- % recuperações paralelas aplicadas;
- % de estudantes que “passaram” na Recuperação;
- % estudantes com ao menos 1 componente abaixo da média;
- % estudantes com ao menos 1 componente abaixo da média (aplicando RPP do estado);
- % de lançamento de notas no sistema.


'''

# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import openpyxl
import re


############################################ Média de frequência dos estudantes ####################################################
'''
Caminho para acesso aos arquivos de frequência dos estudantes:
SigEduc > Diário de Classe > Relatórios > Frequência > Relatório Acompanhamento de Frequências Mensais Aluno > Ano: 2025 ou 2026 > Mês: Abril, Maio e Junho > Gerar Relatório XLSX Calculado > 1 download por DIREC
'''
######## 2025 ########:
# RELATÓRIO DE FREQUÊNCIA - 2º TRIMESTRE (ABRIL, MAIO E JUNHO) - 2025
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Frequencia - Estudantes\2tri 2025"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 4 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=4)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_freq_25 = pd.concat(dfs, ignore_index=True)


# DATAFRAMES RESERVAS
df_freq_25_reserva = df_freq_25.copy(deep=True)


# Manter somente as séries do Ensino Médio (1ª, 2ª e 3ª séries) e ensino regular
df_freq_25 = df_freq_25[df_freq_25['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])]


# Manter somente as etapas do Ensino Regular do Ensino Médio
df_freq_25 = df_freq_25[df_freq_25['ETAPA DE ENSINO'].isin([
    'ENSINO MÉDIO POTIGUAR EM TEMPO INTEGRAL',
    'ENSINO MÉDIO POTIGUAR',
    ])]


# Se tiver valores duplicados de 'MATRÍCULA' no dataframe, somar os valores de aulas dadas e previstas para ter o percentual final
# Garantir que as colunas numéricas estejam como número
df_freq_25['AULAS DADAS'] = pd.to_numeric(df_freq_25['AULAS DADAS'], errors='coerce')
df_freq_25['PRESENÇAS'] = pd.to_numeric(df_freq_25['PRESENÇAS'], errors='coerce')

# Agrupar por matrícula e somar as aulas e presenças
df_freq_25_final = (
    df_freq_25
    .groupby('MATRÍCULA', as_index=False)
    .agg({
        'DIREC': 'first',
        'MUNICÍPIO': 'first',
        'INEP ESCOLA': 'first',
        'ESCOLA': 'first',
        'ALUNO': 'first',
        'CPF': 'first',
        'ETAPA DE ENSINO': 'first',
        'SÉRIE': 'first',
        'AULAS PREVISTAS': 'sum',
        'AULAS DADAS': 'sum',
        'PRESENÇAS': 'sum'
    })
)

# Recalcular o percentual final de frequência
df_freq_25_final['% FREQUÊNCIA'] = (
    df_freq_25_final['PRESENÇAS'] / df_freq_25_final['AULAS DADAS']
) * 100


# Média de frequência dos estudantes:
media_freq_25 = df_freq_25_final['% FREQUÊNCIA'].mean()

print(f"Média de frequência dos estudantes em 2025: {media_freq_25:.2f}%")



######## 2026 ########:
# RELATÓRIO DE FREQUÊNCIA - 2º TRIMESTRE (ABRIL, MAIO E JUNHO) - 2026
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Frequencia - Estudantes\2tri 2026"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 4 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=4)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_freq_26 = pd.concat(dfs, ignore_index=True)


# DATAFRAMES RESERVAS
df_freq_26_reserva = df_freq_26.copy(deep=True)


# Manter somente as séries do Ensino Médio (1ª, 2ª e 3ª séries) e ensino regular
df_freq_26 = df_freq_26[df_freq_26['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])]


# Manter somente as etapas do Ensino Regular do Ensino Médio
df_freq_26 = df_freq_26[df_freq_26['ETAPA DE ENSINO'].isin([
    'ENSINO MÉDIO POTIGUAR EM TEMPO INTEGRAL',
    'ENSINO MÉDIO POTIGUAR',
    ])]


# Se tiver valores duplicados de 'MATRÍCULA' no dataframe, somar os valores de aulas dadas e previstas para ter o percentual final
# Garantir que as colunas numéricas estejam como número
df_freq_26['AULAS DADAS'] = pd.to_numeric(df_freq_26['AULAS DADAS'], errors='coerce')
df_freq_26['PRESENÇAS'] = pd.to_numeric(df_freq_26['PRESENÇAS'], errors='coerce')

# Agrupar por matrícula e somar as aulas e presenças
df_freq_26_final = (
    df_freq_26
    .groupby('MATRÍCULA', as_index=False)
    .agg({
        'DIREC': 'first',
        'MUNICÍPIO': 'first',
        'INEP ESCOLA': 'first',
        'ESCOLA': 'first',
        'ALUNO': 'first',
        'CPF': 'first',
        'ETAPA DE ENSINO': 'first',
        'SÉRIE': 'first',
        'AULAS PREVISTAS': 'sum',
        'AULAS DADAS': 'sum',
        'PRESENÇAS': 'sum'
    })
)

# Recalcular o percentual final de frequência
df_freq_26_final['% FREQUÊNCIA'] = (
    df_freq_26_final['PRESENÇAS'] / df_freq_26_final['AULAS DADAS']
) * 100


# Média de frequência dos estudantes:
media_freq_26 = df_freq_26_final['% FREQUÊNCIA'].mean()

print(f"Média de frequência dos estudantes em 2026: {media_freq_26:.2f}%")


############################################ % de estudantes com frequência abaixo de 75% ####################################################
######## 2025 ########:
# Percentual de linhas com frequência < 75
percentual_freq_25_75 = (df_freq_25_final['% FREQUÊNCIA'] < 75).mean() * 100

print(f'{percentual_freq_25_75:.2f}%')


######## 2026 ########:
# Percentual de linhas com frequência < 75
percentual_freq_26_75 = (df_freq_26_final['% FREQUÊNCIA'] < 75).mean() * 100

print(f'{percentual_freq_26_75:.2f}%')



############################################ % das frequências preenchidas (quantidade de frequências lançadas) ####################################################
'''
Lançamento de frequência pelos professores, através do Relatório de Frequência Mensal de Professsores.

Caminho para acesso aos arquivos de notas dos estudantes:
SigEduc > Diário de Classe > Relatórios > Frequência > Relatório de Acompanhamento de Frequência Mensal Professor > Ano: 2025 ou 2026 > Mês: Abril, Maio e Junho > Gerar Relatório XLSX > 1 download por DIREC
'''

######## 2025 ########:
# RELATÓRIO DE FREQUÊNCIA de PROFESSORES - 2º TRIMESTRE (ABRIL, MAIO E JUNHO) - 2025
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Frequencia - Professores\2tri 2025"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 4 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=4)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_freq_prof_25 = pd.concat(dfs, ignore_index=True)

# DATAFRAMES RESERVAS
df_freq_prof_25_reserva = df_freq_prof_25.copy(deep=True)


# Se tiver valores duplicados de 'MATRÍCULA' no dataframe, somar os valores de aulas dadas e previstas para ter o percentual final
# Garantir que as colunas numéricas estejam como número
df_freq_prof_25['AULAS DADAS'] = pd.to_numeric(df_freq_prof_25['AULAS DADAS'], errors='coerce')
df_freq_prof_25['AULAS PREVISTAS'] = pd.to_numeric(df_freq_prof_25['AULAS PREVISTAS'], errors='coerce')

# Agrupar por matrícula e somar as aulas dadas e aulas previstas
df_freq_prof_25_final = (
    df_freq_prof_25
    .groupby('MATRÍCULA', as_index=False)
    .agg({
        'DIREC': 'first',
        'MUNICÍPIO': 'first',
        'INEP ESCOLA': 'first',
        'ESCOLA': 'first',
        'PROFESSOR': 'first',
        'CPF': 'first',
        'AULAS PREVISTAS': 'sum',
        'AULAS DADAS': 'sum'
    })
)

# Recalcular o percentual final de frequência
df_freq_prof_25_final['% FREQUÊNCIA'] = (
    df_freq_prof_25_final['AULAS DADAS'] / df_freq_prof_25_final['AULAS PREVISTAS']
) * 100


# % de frequências preenchidas (quantidade de frequências lançadas):
perc_freq_prof_25 = (
    df_freq_prof_25_final['AULAS DADAS'].sum() /
    df_freq_prof_25_final['AULAS PREVISTAS'].sum()
) * 100

print(f"Percentual de frequências preenchidas em 2025: {perc_freq_prof_25:.2f}%")



######## 2026 ########:
# RELATÓRIO DE FREQUÊNCIA de PROFESSORES- 2º TRIMESTRE (ABRIL, MAIO E JUNHO) - 2026
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Frequencia - Professores\2tri 2026"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 4 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=4)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_freq_prof_26 = pd.concat(dfs, ignore_index=True)


# DATAFRAMES RESERVAS
df_freq_prof_26_reserva = df_freq_prof_26.copy(deep=True)


# Se tiver valores duplicados de 'MATRÍCULA' no dataframe, somar os valores de aulas dadas e previstas para ter o percentual final
# Garantir que as colunas numéricas estejam como número
df_freq_prof_26['AULAS DADAS'] = pd.to_numeric(df_freq_prof_26['AULAS DADAS'], errors='coerce')
df_freq_prof_26['AULAS PREVISTAS'] = pd.to_numeric(df_freq_prof_26['AULAS PREVISTAS'], errors='coerce')

# Agrupar por matrícula e somar as aulas dadas e aulas previstas
df_freq_prof_26_final = (
    df_freq_prof_26
    .groupby('MATRÍCULA', as_index=False)
    .agg({
        'DIREC': 'first',
        'MUNICÍPIO': 'first',
        'INEP ESCOLA': 'first',
        'ESCOLA': 'first',
        'PROFESSOR': 'first',
        'CPF': 'first',
        'AULAS PREVISTAS': 'sum',
        'AULAS DADAS': 'sum'
    })
)

# Recalcular o percentual final de frequência
df_freq_prof_26_final['% FREQUÊNCIA'] = (
    df_freq_prof_26_final['AULAS DADAS'] / df_freq_prof_26_final['AULAS PREVISTAS']
) * 100


# % de frequências preenchidas (quantidade de frequências lançadas):
perc_freq_prof_26 = (
    df_freq_prof_26_final['AULAS DADAS'].sum() /
    df_freq_prof_26_final['AULAS PREVISTAS'].sum()
) * 100

print(f"Percentual de frequências preenchidas em 2026: {perc_freq_prof_26:.2f}%")



############################################ % estudantes em abandono escolar ####################################################
######## 2025 ########:
'''
Pelo relatório geral de matrículas no SigEduc só teria a informação do recorte de final de 2025, sem conseguir ver informação por trimestre.
Sendo assim, vou considerar a informação oficial do Inep para abandono para Ensino Médio da rede estadual: 3,6%
'''


######## 2026 ########:
# Ler os dados do relatório geral de matrículas do SigEduc para 2026:
df_geral_26 = pd.read_excel(r"C:\Users\hugob\Downloads\20260810_2026_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)


# Manter somente as séries do Ensino Médio (1ª, 2ª e 3ª séries) e ensino regular
df_geral_26 = df_geral_26[df_geral_26['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])]


# Manter somente as etapas do Ensino Regular do Ensino Médio
df_geral_26 = df_geral_26[df_geral_26['ETAPA DE ENSINO'].isin([
    'ENSINO MÉDIO POTIGUAR EM TEMPO INTEGRAL',
    'ENSINO MÉDIO POTIGUAR',
    ])]


# Manter os valores de Matrícula com data de operação mais recente, caso haja duplicidade de matrícula
# Converter 'DATA DA OPERAÇÃO' para datetime
df_geral_26['DATA DA OPERAÇÃO'] = pd.to_datetime(
    df_geral_26['DATA DA OPERAÇÃO'],
    dayfirst=True,
    errors='coerce'
)

# Ordenar pela data mais recente e exclui duplicatas de matrículas
df_geral_26 = (
    df_geral_26
    .sort_values('DATA DA OPERAÇÃO', ascending=False)
    .drop_duplicates(subset='MATRÍCULA', keep='first'))


# Percentual de estudantes com status de abandono:
# Estudantes com Situação = 'DEIXOU DE FREQUENTAR' ou 'CANCELADO'
percentual_abandono_26 = (df_geral_26['SITUAÇÃO'].isin(['DEIXOU DE FREQUENTAR', 'CANCELADO'])).mean() * 100

print(f"Percentual de estudantes com status de abandono em 2026: {percentual_abandono_26:.2f}%")



############################################ risco de abandono (conforme métrica do próprio estado) ####################################################
'''
De acordo com métrica de termômetro de frequência que a Frente de Fluxo Escolar usa.
Risco de Abandono = frequência do estudante de 0% a 50%
'''
######## 2025 ########:
# Percentual de linhas com frequência <= 50
percentual_freq_25_50 = (df_freq_25_final['% FREQUÊNCIA'] <= 50).mean() * 100

print(f'{percentual_freq_25_50:.2f}%')


######## 2026 ########:
# Percentual de linhas com frequência <= 50
percentual_freq_26_50 = (df_freq_26_final['% FREQUÊNCIA'] <= 50).mean() * 100

print(f'{percentual_freq_26_50:.2f}%')


############################################ % escolas com conselho de classe realizada ####################################################
'''
Só tem informação ao final do ano, não tem informação por trimestre.

2025: 100%
2026: só saberemos no final do ano.
'''


############################################ % recuperações paralelas aplicadas ####################################################
'''
Não há acompanhamento de recuperação paralela, então não se sabe se as escolas realizam e quantitativos.

Isso vale para todos os indicadores relacionados à recuperação paralela.
'''


############################################ % estudantes com ao menos 1 componente abaixo da média ####################################################
'''
Caminho para acesso aos arquivos de notas dos estudantes:
SigEduc > Diário de Classe > Relatórios > Resultados Finais > Relatório Acompanhamento de Notas dos Alunos > Ano: 2025 ou 2026 > Gerar Relatório XLSX> 1 download por DIREC
'''

######## 2025 ########:
# RELATÓRIO DE NOTAS de ESTUDANTES - 2º TRIMESTRE (ABRIL, MAIO E JUNHO) - 2025
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Notas\2tri 2025"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_notas_25 = pd.concat(dfs, ignore_index=True)

# DATAFRAMES RESERVAS
df_notas_25_reserva = df_notas_25.copy(deep=True)












######## 2026 ########:
# RELATÓRIO DE NOTAS de ESTUDANTES- 2º TRIMESTRE (ABRIL, MAIO E JUNHO) - 2026
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Notas\2tri 2026"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_notas_26 = pd.concat(dfs, ignore_index=True)


# DATAFRAMES RESERVAS
df_notas_26_reserva = df_notas_26.copy(deep=True)






############################################ % estudantes com ao menos 1 componente abaixo da média (aplicando RPP do estado) ####################################################
######## 2025 ########:









######## 2026 ########:








############################################ % de lançamento de notas no sistema #####################################################
######## 2025 ########:
















######## 2026 ########:














############################################ % de estudantes com 100% das notas lançadas #####################################################
######## 2025 ########:
















######## 2026 ########:























