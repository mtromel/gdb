import pandas as pd

# solicitar a quantidade de arquivos a serem tratados
while True:
    try:
        print("Comparador de Arquivos CSV gerados pelo ExportarEmailsRecursivo")
        num_files = int(input("Informe quantos arquivos CSV deseja comparar (mínimo 2 e máximo 4): "))
        if num_files < 2:
            print("Por favor, informe pelo menos 2 arquivos.")
            continue
        elif num_files > 4:
            print("Por favor, informe no máximo 4 arquivos.")
            continue
        break
    except ValueError:
        print("Entrada inválida. Por favor, informe um número inteiro.")


# solicitar os caminhos dos arquivos CSV
file_paths = []
for i in range(num_files):
    file_path = input(f"Informe o caminho e nome do arquivo CSV {i + 1}: ")
    file_paths.append(file_path)
    
if len(file_paths) == 2:
    csv1 = pd.read_csv(file_paths[0], sep=';')
    csv2 = pd.read_csv(file_paths[1], sep=';')
elif len(file_paths) == 3:
    csv1 = pd.read_csv(file_paths[0], sep=';')
    csv2 = pd.read_csv(file_paths[1], sep=';')
    csv3 = pd.read_csv(file_paths[2], sep=';')
elif len(file_paths) == 4:
    csv1 = pd.read_csv(file_paths[0], sep=';')
    csv2 = pd.read_csv(file_paths[1], sep=';')
    csv3 = pd.read_csv(file_paths[2], sep=';')
    csv4 = pd.read_csv(file_paths[4], sep=';')

# remover espaços vazios dos dados
csv1.columns.str.strip()
csv2.columns.str.strip()
if len(file_paths) >= 3:
    csv3.columns.str.strip()
if len(file_paths) == 4:
    csv4.columns.str.strip()

# exibir as primeiras linhas de cada arquivo
print("Arquivo CSV 1:")
print(csv1.head())
print(" ")
print("-" * 50)
print("Arquivo CSV 2:")
print(csv2.head())
print(" ")
print("-" * 50)
if len(file_paths) >= 3:
    print("Arquivo CSV 3:")
    print(csv3.head())
    print(" ")
    print("-" * 50)
if len(file_paths) == 4:
    print("Arquivo CSV 4:")
    print(csv4.head())
    print(" ")
    print("-" * 50)

# comparar os arquivos e exibir os resultados
print('')
print('-' * 50)
print(f'{num_files} arquivos importados com sucesso!')
print('')
print('Comparando os arquivos...')
duplicados = pd.merge(csv1, csv2, on=["Remetente", "Assunto", "DataRecebimento"])
print('')
print('-' * 50)
print('Comparação concluída!')
print(f'Foram encontrados {len(duplicados)} registros duplicados entre os arquivos')
print(duplicados.head())