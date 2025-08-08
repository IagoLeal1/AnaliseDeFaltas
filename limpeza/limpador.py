import pandas as pd
import os

# --- Configurações ---
ARQUIVO_ENTRADA_EXCEL = 'Amplimed - Gestão de Clínicas (14).xlsx' 
ARQUIVO_SAIDA_EXCEL = 'dados_limpos.xlsx'
COLUNA_PARA_REMOVER = 'Data e Hora agendada'

def limpar_e_salvar_planilha_excel():
    """
    Lê o arquivo .xlsx, remove a coluna de data e salva a nova versão.
    """
    print("--- INICIANDO SCRIPT DE LIMPEZA (SEM DATA) ---")
    
    try:
        print(f"Lendo o arquivo Excel: '{ARQUIVO_ENTRADA_EXCEL}' com o leitor 'calamine'...")
        df = pd.read_excel(ARQUIVO_ENTRADA_EXCEL, engine='calamine')
        print("Arquivo Excel lido com sucesso!")

    except FileNotFoundError:
        print(f"\n[ERRO FATAL]: O arquivo '{ARQUIVO_ENTRADA_EXCEL}' não foi encontrado na pasta 'limpeza'.")
        return
    except Exception as e:
        print(f"\n[ERRO FATAL] Ocorreu um problema ao ler o arquivo Excel. Erro: {e}")
        return

    print("Realizando limpeza dos dados...")
    df.columns = df.columns.str.strip()
    
    # --- A MUDANÇA ESTÁ AQUI ---
    # Verificamos se a coluna de data existe e a removemos.
    if COLUNA_PARA_REMOVER in df.columns:
        df = df.drop(columns=[COLUNA_PARA_REMOVER])
        print(f"Coluna '{COLUNA_PARA_REMOVER}' removida com sucesso.")
    else:
        print(f"Aviso: A coluna '{COLUNA_PARA_REMOVER}' não foi encontrada para ser removida.")

    print(f"Total de registros a serem salvos: {len(df)}")

    caminho_saida = os.path.join('..', 'analise', ARQUIVO_SAIDA_EXCEL)
    try:
        os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
        print(f"Salvando a planilha limpa em: '{caminho_saida}'...")
        df.to_excel(caminho_saida, index=False, engine='openpyxl')
        print("-" * 40)
        print(" SUCESSO! O arquivo limpo (sem data) foi salvo na pasta 'analise'!")
        print("-" * 40)
    except Exception as e:
        print(f"[ERRO FATAL] Não foi possível salvar o novo arquivo Excel. Erro: {e}")

if __name__ == "__main__":
    limpar_e_salvar_planilha_excel()