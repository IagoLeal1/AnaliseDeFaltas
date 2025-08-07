import pandas as pd
import os   # Módulo para lidar com caminhos de arquivos

# --- Configurações ---
# O arquivo de entrada está na mesma pasta que este script
ARQUIVO_ENTRADA_EXCEL = 'MarinaSilva.xlsx'

# O arquivo de saída será salvo na pasta 'analise'
CAMINHO_SAIDA = os.path.join('..', 'analise', 'dados_limpos.xlsx')


def limpar_e_salvar_planilha_excel():
    print("--- INICIANDO SCRIPT DE LIMPEZA ---")
    
    try:
        print(f"Lendo o arquivo Excel: '{ARQUIVO_ENTRADA_EXCEL}'...")
        df = pd.read_excel(ARQUIVO_ENTRADA_EXCEL)
    except FileNotFoundError:
        print(f"\n[ERRO FATAL]: O arquivo '{ARQUIVO_ENTRADA_EXCEL}' não foi encontrado na pasta 'limpeza'.")
        return
    except Exception as e:
        print(f"[ERRO FATAL] Ocorreu um problema ao ler o arquivo Excel. Erro: {e}")
        return

    print("Realizando limpeza e padronização dos dados...")
    df.columns = df.columns.str.strip()
    data_col = 'Data e Hora agendada'
    
    if data_col not in df.columns:
        print(f"[ERRO FATAL] A coluna '{data_col}' não foi encontrada.")
        return
        
    df[data_col] = df[data_col].astype(str)
    df[data_col] = pd.to_datetime(df[data_col], dayfirst=True, errors='coerce')
    df.dropna(subset=[data_col], inplace=True)
    
    print(f"Total de registros limpos: {len(df)}")

    try:
        # Garante que o diretório de saída exista
        os.makedirs(os.path.dirname(CAMINHO_SAIDA), exist_ok=True)
        
        print(f"Salvando a planilha limpa em: '{CAMINHO_SAIDA}'...")
        df.to_excel(CAMINHO_SAIDA, index=False, engine='openpyxl')
        print("-" * 40)
        print(" SUCESSO! O arquivo limpo foi salvo na pasta 'analise'!")
        print("-" * 40)
    except Exception as e:
        print(f"[ERRO FATAL] Não foi possível salvar o novo arquivo Excel. Erro: {e}")

if __name__ == "__main__":
    limpar_e_salvar_planilha_excel()