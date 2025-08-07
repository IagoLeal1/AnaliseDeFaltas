import pandas as pd
import matplotlib.pyplot as plt
import os

# --- CONFIGURAÇÕES DA ANÁLISE ---
ARQUIVO_ENTRADA_LIMPO = 'dados_limpos.xlsx'

# Nomes das colunas
COLUNA_PACIENTE = 'Paciente'
COLUNA_STATUS = 'Status'

# Apelidos para cada status
STATUS_FALTOU = 'Ncompareceu'
STATUS_PRESENTE = 'Finalizado'
STATUS_CANCELADO = 'Cancelado'

def gerar_relatorio_paciente(df, nome_do_paciente):
    """
    Gera um relatório em texto, um arquivo Excel na pasta /relatorios
    e um gráfico PNG na pasta /graficos.
    """
    print(f"\n🔎 --- GERANDO RELATÓRIO PARA: {nome_do_paciente} --- 🔎")
    df_paciente = df[df[COLUNA_PACIENTE] == nome_do_paciente].copy()

    if df_paciente.empty:
        print(f"Paciente '{nome_do_paciente}' não encontrado.")
        return

    # Cálculos...
    presencas = (df_paciente[COLUNA_STATUS] == STATUS_PRESENTE).sum()
    faltas = (df_paciente[COLUNA_STATUS] == STATUS_FALTOU).sum()
    cancelados = (df_paciente[COLUNA_STATUS] == STATUS_CANCELADO).sum()
    total_valido = presencas + faltas
    taxa_de_falta = (faltas / total_valido) * 100 if total_valido > 0 else 0

    print(f"Consultas finalizadas (presenças): {presencas}")
    print(f"Consultas não comparecidas (faltas): {faltas}")
    print(f"Consultas canceladas: {cancelados}")
    print(f"📊 Taxa de Falta: {taxa_de_falta:.2f}%")
    
    # --- Geração dos arquivos de saída com nome e pastas dinâmicas ---
    
    # 1. Cria um nome de arquivo "seguro"
    nome_arquivo_seguro = nome_do_paciente.lower().replace(' ', '_')
    
    # 2. Define os caminhos de saída para as pastas corretas
    #    ../ significa "voltar uma pasta" para chegar na pasta 'faltas'
    caminho_relatorio_excel = os.path.join('..', 'relatorios', f'relatorio_{nome_arquivo_seguro}.xlsx')
    caminho_grafico = os.path.join('..', 'graficos', f'grafico_{nome_arquivo_seguro}.png')

    # 3. Garante que o diretório de relatórios exista e salva o Excel
    try:
        os.makedirs(os.path.dirname(caminho_relatorio_excel), exist_ok=True)
        df_paciente.to_excel(caminho_relatorio_excel, index=False, engine='openpyxl')
        print(f"\n✅ Relatório Excel salvo com sucesso em: '{caminho_relatorio_excel}'")
    except Exception as e:
        print(f"\n[ERRO] Não foi possível salvar o relatório Excel. Erro: {e}")

    # 4. Garante que o diretório de gráficos exista e salva o Gráfico
    if total_valido > 0:
        fig, ax = plt.subplots(figsize=(10, 7))
        labels = ['Presenças', 'Faltas', 'Cancelados']
        sizes = [presencas, faltas, cancelados]
        cores = ['#2E8B57', '#DC143C', '#A9A9A9']
        explode = (0, 0.1, 0)
        ax.pie(sizes, explode=explode, labels=labels, colors=cores, autopct='%1.1f%%',
               shadow=True, startangle=140, textprops={'fontsize': 12})
        ax.axis('equal')
        plt.title(f'Resumo de Consultas de\n{nome_do_paciente}', fontsize=16, weight='bold')
        
        try:
            os.makedirs(os.path.dirname(caminho_grafico), exist_ok=True)
            plt.savefig(caminho_grafico)
            print(f"✅ Gráfico salvo com sucesso em: '{caminho_grafico}'")
        except Exception as e:
            print(f"\n[ERRO] Não foi possível salvar o gráfico. Erro: {e}")
        
        plt.close(fig)

def rodar_analise_de_dados():
    try:
        df = pd.read_excel(ARQUIVO_ENTRADA_LIMPO)
    except FileNotFoundError:
        print(f"[ERRO] O arquivo '{ARQUIVO_ENTRADA_LIMPO}' não foi encontrado na pasta 'analise'.")
        print("Você executou o SCRIPT 1 ('limpador.py') primeiro?")
        return
    
    gerar_relatorio_paciente(df, 'Marina Silva')
    
    print("\nAnálise finalizada!")

if __name__ == "__main__":
    rodar_analise_de_dados()