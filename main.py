import pandas as pd

# Lire le fichier CSV avec une virgule comme délimiteur
csv_file = 'TEST_TECHNIQUE.txt'
df_csv = pd.read_csv(csv_file, delimiter=',')

# Chemin du fichier contenant les noms des colonnes
excel_columns_file_path = 'Copie de MACRO RéCUP SISLSOL.xls'

# Écrire le DataFrame du CSV dans un fichier Excel
excel_file = 'Excel.xlsx'
df_csv.to_excel(excel_file, index=False)

# Lire les colonnes depuis un fichier Excel existant
excel_read_new_columns = pd.read_excel(excel_columns_file_path)

# Lire le fichier Excel généré
df_excel = pd.read_excel(excel_file)

# Chemin du fichier contenant les nouveaux noms de colonnes
fichier_nouveaux_noms = './Copie de MACRO RéCUP SISLSOL.xls'

# Charger le fichier contenant les nouveaux noms de colonnes
df_nouveaux_noms = pd.read_excel(fichier_nouveaux_noms, header=None)

# Extraire les nouveaux noms de colonnes
nouveaux_noms = df_nouveaux_noms.iloc[0].tolist()

# Renommer les colonnes du DataFrame Excel
df_excel.columns = nouveaux_noms

# Afficher les nouveaux noms de colonnes
print("nouveaux noms depuis le fichier Excel :", nouveaux_noms)

# Obtenir les colonnes des deux DataFrames
csv_columns = set(df_csv.columns)
excel_columns = set(df_excel.columns)

# Trouver les différences entre les colonnes
columns_in_csv_not_in_excel = csv_columns - excel_columns
columns_in_excel_not_in_csv = excel_columns - csv_columns

# Afficher les longueurs des colonnes
print('len: csv.columns', len(csv_columns))
print('len: excel.columns', len(excel_columns))

# Afficher les différences entre les colonnes
print(f"Colonnes présentes dans le CSV mais pas dans l'Excel : {columns_in_csv_not_in_excel}")
print(f"Colonnes présentes dans l'Excel mais pas dans le CSV : {columns_in_excel_not_in_csv}")

# Vérifier si le nombre de colonnes est le même
if len(csv_columns) == len(excel_columns):
    # Afficher les nouveaux noms de colonnes
    print("Nouveaux noms de colonnes :")
    print(df_excel.columns)
    
    # Trier les données par colonnes spécifiées
    sort_columns = ["CODE BCR DE L'APERITEUR ", "REFERENCE CONTRAT DE L'APERITEUR   ", "REFERENCE SINISTRE DE L'APERITEUR "]
    df_excel = df_excel.sort_values(by=sort_columns)
else:
    print(f"Erreur : Le nombre de colonnes dans le CSV ({len(csv_columns)}) ne correspond pas au nombre de colonnes dans l'Excel ({len(excel_columns)}).")

# Filtrer les lignes où la colonne I vaut 15
df_excel_filtered = df_excel[df_excel.iloc[:, 8] == 15]

# Exclure les lignes où la colonne J vaut 15
df_excel_filtered = df_excel_filtered[df_excel_filtered.iloc[:, 9] != 15]

# Masquer les colonnes inutiles
columns_to_keep = df_excel_filtered.columns[11:]
df_excel_filtered = df_excel_filtered[columns_to_keep]

# Supprimer les colonnes de M à P
columns_to_drop = df_excel_filtered.columns[2:6]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Supprimer les colonnes de R à S
columns_to_drop = df_excel_filtered.columns[6:8]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Supprimer les colonnes de AE à AH
columns_to_drop = df_excel_filtered.columns[19:23]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Supprimer les colonnes de AJ à AK
columns_to_drop = df_excel_filtered.columns[23:25]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Supprimer les colonnes de AN à AQ
columns_to_drop = df_excel_filtered.columns[25:29]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Supprimer les colonnes de AR à AX
columns_to_drop = df_excel_filtered.columns[29:36]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Supprimer les colonnes de AZ à BN
columns_to_drop = df_excel_filtered.columns[36:51]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Supprimer les colonnes de BP à BR
columns_to_drop = df_excel_filtered.columns[51:54]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Supprimer les colonnes de BT à CF
columns_to_drop = df_excel_filtered.columns[54:66]
df_excel_filtered = df_excel_filtered.drop(columns=columns_to_drop)

# Convertir les colonnes spécifiées en chaînes de caractères
columns_to_convert = ['AY', 'BO', 'BS']

for column in columns_to_convert:
    if column in df_excel_filtered.columns:
        df_excel_filtered[column] = df_excel_filtered[column].astype(str)


# Sauvegarder le DataFrame filtré dans un nouveau fichier Excel
new_excel_file = 'Ouverture AXA 05-2024.xlsx'
df_excel_filtered.to_excel(new_excel_file, index=False, sheet_name='Ouverture AXA')
