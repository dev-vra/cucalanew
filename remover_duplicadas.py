import pandas as pd
import os

def remover_duplicatas_e_salvar_excel():
    """
    Esta função solicita ao usuário o caminho de um arquivo (CSV ou XLSX),
    remove linhas duplicadas com base em uma chave de concatenação 
    e salva o resultado em um novo arquivo XLSX (Excel).
    """
    # Solicita ao usuário o caminho do arquivo
    caminho_arquivo = input("Por favor, insira o caminho completo do seu arquivo (CSV ou XLSX) e pressione Enter: ")

    # Verifica se o arquivo existe
    if not os.path.exists(caminho_arquivo):
        print("Erro: O arquivo não foi encontrado no caminho especificado.")
        return

    df = None
    try:
        # Pega a extensão do arquivo para decidir como lê-lo
        extensao = os.path.splitext(caminho_arquivo)[1].lower()

        if extensao == '.csv':
            # Tenta ler o CSV com diferentes codificações
            codificacoes = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
            for enc in codificacoes:
                try:
                    df = pd.read_csv(caminho_arquivo, encoding=enc)
                    print(f"Arquivo CSV carregado com sucesso usando a codificação '{enc}'.")
                    break
                except UnicodeDecodeError:
                    continue
            if df is None:
                print("Erro: Não foi possível decodificar o arquivo com as codificações testadas.")
                return

        elif extensao in ['.xlsx', '.xls']:
            # Tenta ler o arquivo Excel
            df = pd.read_excel(caminho_arquivo)
            print("Arquivo Excel carregado com sucesso!")

        else:
            print(f"Erro: Formato de arquivo '{extensao}' não suportado. Por favor, use .csv ou .xlsx.")
            return

        # Define as colunas para a chave de duplicidade
        colunas_chave = ['CONT. REF', 'REF.CUCALA', 'NUMBER']

        # Verifica se as colunas necessárias existem no DataFrame
        for coluna in colunas_chave:
            if coluna not in df.columns:
                print(f"Erro: A coluna '{coluna}' não foi encontrada no arquivo.")
                return

        # Cria a chave de confirmação de duplicatas
        df['chave_duplicidade'] = df[colunas_chave[0]].astype(str) + '-' + \
                                   df[colunas_chave[1]].astype(str) + '-' + \
                                   df[colunas_chave[2]].astype(str)

        # Conta o número de linhas antes de remover as duplicatas
        linhas_antes = len(df)
        print(f"O arquivo original possui {linhas_antes} linhas.")

        # Remove as duplicatas com base na nova chave
        df_sem_duplicatas = df.drop_duplicates(subset=['chave_duplicidade'], keep='first')

        # Conta as linhas removidas
        linhas_depois = len(df_sem_duplicatas)
        linhas_removidas = linhas_antes - linhas_depois
        print(f"{linhas_removidas} linhas duplicadas foram removidas.")

        # Remove a coluna de chave auxiliar
        df_sem_duplicatas = df_sem_duplicatas.drop(columns=['chave_duplicidade'])

        # --- ALTERAÇÃO AQUI ---
        # Salva o resultado em um novo arquivo XLSX (Excel)
        nome_arquivo_saida = 'database_sem_duplicatas.xlsx'
        df_sem_duplicatas.to_excel(nome_arquivo_saida, index=False)

        print(f"\nProcesso concluído! O arquivo sem duplicatas foi salvo como: '{nome_arquivo_saida}'")

    except Exception as e:
        print(f"Ocorreu um erro inesperado ao processar o arquivo: {e}")

if __name__ == "__main__":
    remover_duplicatas_e_salvar_excel()