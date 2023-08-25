import pandas as pd

# Ler o arquivo Excel
df = pd.read_excel("C:\\Users\\italo.mendes\\Desktop\\Automação-Batimentos\\ARQUIVO BATIMENTO FILTRADO - 23.08.xlsx")

# Quantas linhas cada consignatária possui
consignataria_counts = df['CONSIGNATARIA'].value_counts()

# Quantos CPF's exclusivos estão no arquivo
unique_cpf_count = df['CPF'].nunique()

# Ler o novo arquivo Excel
df_new = pd.read_excel("C:\\Users\\italo.mendes\\Desktop\\Automação-Batimentos\\ARQUIVO BATIMENTO FILTRADO - 24.08.xlsx")

# Diferença no número total de registros entre os dois arquivos
total_difference = len(df_new) - len(df)

# Diferença no número de CPFs exclusivos entre os dois arquivos
unique_cpf_difference = df_new['CPF'].nunique() - df['CPF'].nunique()

# Novas consignatárias no arquivo mais recente que não estavam no anterior
new_consignatarias = set(df_new['CONSIGNATARIA']) - set(df['CONSIGNATARIA'])

# Consignatárias do arquivo anterior que não estão no mais recente
missing_consignatarias = set(df['CONSIGNATARIA']) - set(df_new['CONSIGNATARIA'])

# Identificar as linhas novas ou diferentes no arquivo mais recente
diff_rows = df_new[~df_new.apply(tuple, 1).isin(df.apply(tuple, 1))]

# Exportar as linhas diferentes para um novo arquivo Excel
diff_file_path = "C:\\Users\\italo.mendes\\Desktop\\Automação-Batimentos\\differences.xlsx"

# Ler o arquivo "PROPOSTAS YUPPI"
df_yuppie = pd.read_excel("C:\\Users\\italo.mendes\\Desktop\\Automação-Batimentos\\PROPOSTAS YUPPIE 25.08.xlsx")

# Renomear as colunas conforme solicitado
df_yuppie = df_yuppie.rename(columns={
    "CPF CLIENTE": "CPF",
    "DATA DIGITAÇÃO": "Data Digitação",
    "BANCO": "Consignataria"
})

# Comparação e atribuição das cores
df_yuppie["Color"] = ""

for index, row in df_yuppie.iterrows():
    # Filtrando a linha correspondente no dataframe de diferenças
    matching_row = diff_rows[
        (diff_rows["CPF"] == row["CPF"]) & 
        (diff_rows["CONSIGNATARIA"].str.contains(str(row["Consignataria"]), case=False, na=False))
    ]
    
    if matching_row.empty:
        # Caso o CPF não seja encontrado naquela consignatária
        df_yuppie.at[index, "Color"] = "red"
    else:
        # Calculando a diferença entre as datas
        date_diff = (matching_row["DEFERIMENTO"].iloc[0] - row["Data Digitação"]).days
        if date_diff <= 7:
            df_yuppie.at[index, "Color"] = "green"
        else:
            df_yuppie.at[index, "Color"] = "yellow"

# Exportando o arquivo com as alterações
output_path = "C:\\Users\\italo.mendes\\Desktop\\Automação-Batimentos\\BATIMENTO COMPLETO.xlsx"
df_yuppie.to_excel(output_path, index=False, freeze_panes=(1,0))
