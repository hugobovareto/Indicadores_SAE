'''
Dados só do Ensino Médio (todos os 3 anos);
Considerar somente ensino regular;
Recortes temporais de 2º trimestre (Abril, Maio e Junho) para 2025 e 2026;
Considerar somente 13 componentes da Formação Geral Básica da BNCC para Ensino Médio (sem Ensino Religioso, pois não reprova):
- "Arte",
- "Biologia",
- "Educação Física",
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

Para as notas, como quer informação do 2º trimestre, estou considerando as notas do 1º e 2º bimestres (para 2025 e 2026).
Preencher essa informação em julho, que é o período que finaliza o 2º bimestre.
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


# Manter somente as séries do Ensino Médio (1ª, 2ª e 3ª séries) e ensino regular
df_notas_25 = df_notas_25[df_notas_25['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])]


# Manter somente as etapas do Ensino Regular do Ensino Médio
df_notas_25 = df_notas_25[df_notas_25['ETAPA ENSINO'].isin([
    'ENSINO MÉDIO POTIGUAR EM TEMPO INTEGRAL',
    'ENSINO MÉDIO POTIGUAR',
    ])]


# Manter somente os componentes curriculares da BNCC da Formação Geral Básica para Ensino Médio (sem Ensino Religioso, pois não reprova)
componentes_bnbc = ['Arte',
                    'Biologia',
                    'Educação Física',
                    'Filosofia',
                    'Física',
                    'Geografia',
                    'História',
                    'Língua Inglesa',
                    'Língua Portuguesa',
                    'Língua Espanhola',
                    'Matemática',
                    'Química',
                    'Sociologia']

df_notas_25 = df_notas_25[df_notas_25['COMPONENTE CURRICULAR'].isin(componentes_bnbc)]


# Substituir vírgula por ponto para reconhecimento das notas como números:
colunas_para_converter = [
    "NOTA 1º BIMESTRE",
    "NOTA 2º BIMESTRE",
    "NOTA 3º BIMESTRE",
    "NOTA 4º BIMESTRE",
    "MÉDIA ANUAL",
    "EXAME FINAL",
    "AVALIAÇÃO ESPECIAL",
    "MÉDIA FINAL"
] 

for col in colunas_para_converter:
    if col in df_notas_25.columns:  # só executa se a coluna estiver no DataFrame
        # Substitui vírgula por ponto
        df_notas_25[col] = df_notas_25[col].str.replace(",", ".", regex=False)
        # Converte para float, erros viram NaN
        df_notas_25[col] = pd.to_numeric(df_notas_25[col], errors="coerce")


# Fazer nota do 1º semestre, fazendo média de 'NOTA 1º BIMESTRE' e 'NOTA 2º BIMESTRE'
df_notas_25['MEDIA_NOTAS'] = df_notas_25[['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE']].mean(axis=1, skipna=True)

# Quantidade de reprovações e aprovações por estudante:
# 1. Cria colunas temporárias booleanas (True se a condição for atendida, False se não)
df_notas_25["aprovado_temp"] = df_notas_25["MEDIA_NOTAS"] >= 6
df_notas_25["reprovado_temp"] = df_notas_25["MEDIA_NOTAS"] < 6


# Agrupa por 'MATRÍCULA ESTUDANTE' e soma as colunas temporárias
df_resumo_notas_25 = (
    df_notas_25.groupby("MATRÍCULA ESTUDANTE")[["aprovado_temp", "reprovado_temp"]]
    .sum()
    .reset_index()
)

# Renomeia as colunas
df_resumo_notas_25.rename(
    columns={
        "aprovado_temp": "Componentes_aprovados",
        "reprovado_temp": "Componentes_reprovados",
    },
    inplace=True,
)

# % de estudantes com ao menos 1 componente abaixo da média:
percentual_estudantes_1_reprovacao_25 = (df_resumo_notas_25['Componentes_reprovados'] >= 1).sum() / len(df_resumo_notas_25) * 100

print(f"Percentual de estudantes com ao menos 1 componente abaixo da média (2025): {percentual_estudantes_1_reprovacao_25:.2f}%")



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


# Manter somente as séries do Ensino Médio (1ª, 2ª e 3ª séries) e ensino regular
df_notas_26 = df_notas_26[df_notas_26['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])]


# Manter somente as etapas do Ensino Regular do Ensino Médio
df_notas_26 = df_notas_26[df_notas_26['ETAPA ENSINO'].isin([
    'ENSINO MÉDIO POTIGUAR EM TEMPO INTEGRAL',
    'ENSINO MÉDIO POTIGUAR',
    ])]


# Manter somente os componentes curriculares da BNCC da Formação Geral Básica para Ensino Médio (sem Ensino Religioso, pois não reprova)
componentes_bnbc = ['Arte',
                    'Biologia',
                    'Educação Física',
                    'Filosofia',
                    'Física',
                    'Geografia',
                    'História',
                    'Língua Inglesa',
                    'Língua Portuguesa',
                    'Língua Espanhola',
                    'Matemática',
                    'Química',
                    'Sociologia']

df_notas_26 = df_notas_26[df_notas_26['COMPONENTE CURRICULAR'].isin(componentes_bnbc)]


# Substituir vírgula por ponto para reconhecimento das notas como números:
colunas_para_converter = [
    "NOTA 1º BIMESTRE",
    "NOTA 2º BIMESTRE",
    "NOTA 3º BIMESTRE",
    "NOTA 4º BIMESTRE",
    "MÉDIA ANUAL",
    "EXAME FINAL",
    "AVALIAÇÃO ESPECIAL",
    "MÉDIA FINAL"
] 

for col in colunas_para_converter:
    if col in df_notas_26.columns:  # só executa se a coluna estiver no DataFrame
        # Substitui vírgula por ponto
        df_notas_26[col] = df_notas_26[col].str.replace(",", ".", regex=False)
        # Converte para float, erros viram NaN
        df_notas_26[col] = pd.to_numeric(df_notas_26[col], errors="coerce")


# Fazer nota do 1º semestre, fazendo média de 'NOTA 1º BIMESTRE' e 'NOTA 2º BIMESTRE'
df_notas_26['MEDIA_NOTAS'] = df_notas_26[['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE']].mean(axis=1, skipna=True)

# Quantidade de reprovações e aprovações por estudante:
# 1. Cria colunas temporárias booleanas (True se a condição for atendida, False se não)
df_notas_26["aprovado_temp"] = df_notas_26["MEDIA_NOTAS"] >= 6
df_notas_26["reprovado_temp"] = df_notas_26["MEDIA_NOTAS"] < 6


# Agrupa por 'MATRÍCULA ESTUDANTE' e soma as colunas temporárias
df_resumo_notas_26 = (
    df_notas_26.groupby("MATRÍCULA ESTUDANTE")[["aprovado_temp", "reprovado_temp"]]
    .sum()
    .reset_index()
)

# Renomeia as colunas
df_resumo_notas_26.rename(
    columns={
        "aprovado_temp": "Componentes_aprovados",
        "reprovado_temp": "Componentes_reprovados",
    },
    inplace=True,
)


# % de estudantes com ao menos 1 componente abaixo da média:
percentual_estudantes_1_reprovacao_26 = (df_resumo_notas_26['Componentes_reprovados'] >= 1).sum() / len(df_resumo_notas_26) * 100

print(f"Percentual de estudantes com ao menos 1 componente abaixo da média (2026): {percentual_estudantes_1_reprovacao_26:.2f}%")


############################################ % estudantes com ao menos 1 componente abaixo da média (aplicando RPP do estado) ####################################################
'''
Nesse caso são estudantes com 1 reprovação a mais do que o permitido para RAPP, logo, estudantes com 7 reprovações ou mais.
'''

######## 2025 ########:
# % de estudantes com ao menos 1 componente abaixo da média:
percentual_estudantes_7_reprovacoes_25 = (df_resumo_notas_25['Componentes_reprovados'] >= 7).sum() / len(df_resumo_notas_25) * 100

print(f"Percentual de estudantes com ao menos 7 componente abaixo da média (2025): {percentual_estudantes_7_reprovacoes_25:.2f}%")


######## 2026 ########:
# % de estudantes com ao menos 1 componente abaixo da média:
percentual_estudantes_7_reprovacoes_26 = (df_resumo_notas_26['Componentes_reprovados'] >= 7).sum() / len(df_resumo_notas_26) * 100

print(f"Percentual de estudantes com ao menos 7 componente abaixo da média (2026): {percentual_estudantes_7_reprovacoes_26:.2f}%")



############################################ % de lançamento de notas no sistema #####################################################
######## 2025 ########:
# Total de notas esperadas (2 por estudante, por serem 2 bimestres: 'NOTA '1º BIMESTRE' e 'NOTA 2º BIMESTRE')
total_esperado_25 = len(df_notas_25) * 2

# Total de notas lançadas (não nulas)
total_lancado_25 = df_notas_25[['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE']].notna().sum().sum()

# Percentual de notas lançadas
percentual_lancado_25 = (total_lancado_25 / total_esperado_25) * 100

print(f'Percentual de notas lançadas (2025): {percentual_lancado_25:.2f}%')


######## 2026 ########:
# Total de notas esperadas (2 por estudante, por serem 2 bimestres: 'NOTA '1º BIMESTRE' e 'NOTA 2º BIMESTRE')
total_esperado_26 = len(df_notas_26) * 2

# Total de notas lançadas (não nulas)
total_lancado_26 = df_notas_26[['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE']].notna().sum().sum()

# Percentual de notas lançadas
percentual_lancado_26 = (total_lancado_26 / total_esperado_26) * 100

print(f'Percentual de notas lançadas (2026): {percentual_lancado_26:.2f}%')


############################################ % de estudantes com 100% das notas lançadas #####################################################
######## 2025 ########:
# Verifica, em cada linha, se as duas notas estão preenchidas
linha_completa_25 = df_notas_25[['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE']].notna().all(axis=1)

# Para cada matrícula, verifica se TODAS as linhas estão completas
matricula_completa_25 = linha_completa_25.groupby(df_notas_25['MATRÍCULA ESTUDANTE']).all()

# Percentual de matrículas com 100% das notas lançadas
percentual_matriculas_completas_25 = matricula_completa_25.mean() * 100

print(f'Percentual de matrículas com 100% das notas lançadas (2025): {percentual_matriculas_completas_25:.2f}%')


######## 2026 ########:
# Verifica, em cada linha, se as duas notas estão preenchidas
linha_completa_26 = df_notas_26[['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE']].notna().all(axis=1)

# Para cada matrícula, verifica se TODAS as linhas estão completas
matricula_completa_26 = linha_completa_26.groupby(df_notas_26['MATRÍCULA ESTUDANTE']).all()

# Percentual de matrículas com 100% das notas lançadas
percentual_matriculas_completas_26 = matricula_completa_26.mean() * 100

print(f'Percentual de matrículas com 100% das notas lançadas (2026): {percentual_matriculas_completas_26:.2f}%')












