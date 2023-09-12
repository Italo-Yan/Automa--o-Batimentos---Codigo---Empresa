import pandas as pd

# Ler o arquivo Excel
df = pd.read_excel("ARQUIVO FILTRADO 2")

# Quantas linhas cada consignatária possui
consignataria_counts = df['CONSIGNATARIA'].value_counts()
print(f"Quantas linhas cada consignatária possui: {consignataria_counts}")

# Quantos CPF's exclusivos estão no arquivo
unique_cpf_count = df['CPF'].nunique()
print(f"Quantas CPF's exclusivos estão no arquivo: {unique_cpf_count}")
print("-------------------------------------------")

# Ler o novo arquivo Excel
df_new = pd.read_excel("ARQUIVO FILTRADO 2")

# Diferença no número total de registros entre os dois arquivos
total_difference = len(df_new) - len(df)
print(f"Diferença no número total de registros: {total_difference}")

# Diferença no número de CPFs exclusivos entre os dois arquivos
unique_cpf_difference = df_new['CPF'].nunique() - df['CPF'].nunique()

# Novas consignatárias no arquivo mais recente que não estavam no anterior
new_consignatarias = set(df_new['CONSIGNATARIA']) - set(df['CONSIGNATARIA'])

# Consignatárias do arquivo anterior que não estão no mais recente
missing_consignatarias = set(df['CONSIGNATARIA']) - set(df_new['CONSIGNATARIA'])

# Formatando a saída
output = []

if unique_cpf_difference != 0:
    output.append(f"Diferença no número de CPFs exclusivos: {unique_cpf_difference}")

if new_consignatarias:
    output.append(f"Novas consignatárias no arquivo mais recente: {', '.join(new_consignatarias)}")

if missing_consignatarias:
    output.append(f"Consignatárias ausentes no arquivo mais recente: {', '.join(missing_consignatarias)}")

# Verificando se não há diferenças
if not output:
    print("Não há diferenças nas comparações.")
else:
    print("\n".join(output))

# Identificar as linhas novas ou diferentes no arquivo mais recente
diff_rows = df_new[~df_new.apply(tuple, 1).isin(df.apply(tuple, 1))]
print(f"Número de linhas novas ou diferentes no arquivo mais recente: {len(diff_rows)}")

# Ler o arquivo "PROPOSTAS"
df_arquivo = pd.read_excel("ARQUIVO EXTRATO")

# Renomear as colunas conforme solicitado
df_arquivo = df_arquivo.rename(columns={
    "CPF CLIENTE": "CPF",
    "DATA DIGITAÇÃO": "Data Digitação",
    "BANCO": "Consignataria"
})

# Aplica a coloração em toda a linha com base na coluna 'Color'.
def color_row_based_on_color_column(series):
    color = ""
    if series["Color"] == "red":
        color = 'background-color: red'
    elif series["Color"] == "green":
        color = 'background-color: green'
    elif series["Color"] == "yellow":
        color = 'background-color: yellow'
    else:
        color = ''
    return [color] * len(series)

# Comparação e atribuição das cores
df_arquivo["Color"] = ""

for index, row in df_arquivo.iterrows():
    # Filtrando a linha correspondente no dataframe de diferenças
    matching_row = diff_rows[
        (diff_rows["CPF"] == row["CPF"]) & 
        (diff_rows["CONSIGNATARIA"].str.contains(str(row["Consignataria"]), case=False, na=False))
    ]
    
    if matching_row.empty:
        # Caso o CPF não seja encontrado naquela consignatária
        df_arquivo.at[index, "Color"] = "red"
    else:
        # Calculando a diferença entre as datas
        date_diff = (matching_row["DEFERIMENTO"].iloc[0] - row["Data Digitação"]).days
        if date_diff <= 7:
            df_arquivo.at[index, "Color"] = "green"
        else:
            df_arquivo.at[index, "Color"] = "yellow"

# Aplicando a coloração e exportando o arquivo
styled_arquivo = df_arquivo.style.apply(color_row_based_on_color_column, axis=1)
styled_arquivo.to_excel("EXPORTAÇÃO DO ARQUIVO", engine='openpyxl', index=False)


#PROJETO DESATIVADO ! ! !
