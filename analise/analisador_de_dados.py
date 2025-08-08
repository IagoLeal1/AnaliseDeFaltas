import pandas as pd
import matplotlib.pyplot as plt
import os
import re

# --- CONFIGURAÇÕES DA ANÁLISE ---
ARQUIVO_ENTRADA_LIMPO = 'dados_limpos.xlsx'

# Nomes das colunas
COLUNA_PACIENTE = 'Paciente'
COLUNA_STATUS = 'Status'
COLUNA_PROCEDIMENTO = 'Procedimento'

# Apelidos para cada status
STATUS_FALTOU = 'Ncompareceu'
STATUS_PRESENTE = 'Finalizado'
STATUS_CANCELADO = 'Cancelado'

def gerar_relatorios_completos(df, nome_do_paciente):
    """
    Função principal que gera um kit completo de relatórios para um paciente:
    - Um .txt para cada procedimento.
    - Um .txt geral para a "chefia".
    - Um .xlsx e .png para cada procedimento.
    """
    print(f"\n🔎 --- GERANDO KIT COMPLETO DE RELATÓRIOS PARA: {nome_do_paciente} --- 🔎")
    
    # --- 1. PREPARAÇÃO ---
    df_paciente = df[df[COLUNA_PACIENTE] == nome_do_paciente].copy()

    if df_paciente.empty:
        print(f"Paciente '{nome_do_paciente}' não encontrado.")
        return

    nome_pasta_paciente = nome_do_paciente.lower().replace(' ', '_')
    caminho_pasta_relatorios = os.path.join('..', 'relatorios', nome_pasta_paciente)
    caminho_pasta_graficos = os.path.join('..', 'graficos', nome_pasta_paciente)
    os.makedirs(caminho_pasta_relatorios, exist_ok=True)
    os.makedirs(caminho_pasta_graficos, exist_ok=True)
    
    procedimentos_unicos = df_paciente[COLUNA_PROCEDIMENTO].unique()
    
    # Variáveis para guardar o resumo geral
    resumos_para_chefia = []
    total_presencas_geral = 0
    total_faltas_geral = 0
    total_cancelados_geral = 0

    # --- 2. LOOP POR PROCEDIMENTO ---
    for procedimento in procedimentos_unicos:
        df_procedimento = df_paciente[df_paciente[COLUNA_PROCEDIMENTO] == procedimento]

        presencas = (df_procedimento[COLUNA_STATUS] == STATUS_PRESENTE).sum()
        faltas = (df_procedimento[COLUNA_STATUS] == STATUS_FALTOU).sum()
        cancelados = (df_procedimento[COLUNA_STATUS] == STATUS_CANCELADO).sum()
        total_valido = presencas + faltas
        taxa_de_falta = (faltas / total_valido) * 100 if total_valido > 0 else 0

        # Acumula os totais para o resumo geral
        total_presencas_geral += presencas
        total_faltas_geral += faltas
        total_cancelados_geral += cancelados

        # Monta o texto do relatório para este procedimento
        texto_relatorio_procedimento = f"""
--- RELATÓRIO DO PROCEDIMENTO: {procedimento} ---
Paciente: {nome_do_paciente}
Data da Geração: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}

Consultas finalizadas (presenças): {presencas}
Consultas não comparecidas (faltas): {faltas}
Consultas canceladas: {cancelados}
-------------------------------------------------
📊 Taxa de Falta (sobre consultas válidas): {taxa_de_falta:.2f}%
"""
        # Guarda este resumo para o relatório final
        resumos_para_chefia.append(texto_relatorio_procedimento)

        # --- Geração dos arquivos de saída para o procedimento ---
        nome_arquivo_base = re.sub(r'[\\/*?:"<>|]',"", procedimento).lower().replace(' ', '_')
        
        # Salva o relatório .txt do procedimento
        try:
            caminho_txt = os.path.join(caminho_pasta_relatorios, f'relatorio_{nome_arquivo_base}.txt')
            with open(caminho_txt, 'w', encoding='utf-8') as f:
                f.write(texto_relatorio_procedimento.strip())
        except Exception as e:
            print(f"     [ERRO] Falha ao salvar .txt: {e}")

        # Salva o relatório .xlsx do procedimento
        caminho_excel = os.path.join(caminho_pasta_relatorios, f'relatorio_{nome_arquivo_base}.xlsx')
        df_procedimento.to_excel(caminho_excel, index=False, engine='openpyxl')

        # Gera e salva o Gráfico .png do procedimento
        if total_valido > 0:
            caminho_grafico = os.path.join(caminho_pasta_graficos, f'grafico_{nome_arquivo_base}.png')
            # ... (código do gráfico, sem alterações)
            fig, ax = plt.subplots()
            ax.pie([presencas, faltas, cancelados], labels=['Presenças', 'Faltas', 'Cancelados'], colors=['#2E8B57', '#DC143C', '#A9A9A9'], autopct='%1.1f%%', startangle=90)
            ax.axis('equal')
            plt.title(f'Resumo de: {procedimento}\nPaciente: {nome_do_paciente}')
            plt.savefig(caminho_grafico)
            plt.close(fig)

    # --- 3. GERAÇÃO DO RELATÓRIO MESTRE PARA A CHEFIA ---
    total_valido_geral = total_presencas_geral + total_faltas_geral
    taxa_falta_geral = (total_faltas_geral / total_valido_geral) * 100 if total_valido_geral > 0 else 0

    texto_chefe_cabecalho = f"""
=====================================================
    RELATÓRIO CONSOLIDADO PARA A CHEFIA
=====================================================
Paciente: {nome_do_paciente}
Data da Geração: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}

--- RESUMO GERAL (TODOS OS PROCEDIMENTOS) ---
Consultas finalizadas (presenças): {total_presencas_geral}
Consultas não comparecidas (faltas): {total_faltas_geral}
Consultas canceladas: {total_cancelados_geral}
-------------------------------------------------
📊 Taxa de Falta GERAL: {taxa_falta_geral:.2f}%
=====================================================

--- DETALHAMENTO POR PROCEDIMENTO ---
"""
    # Junta o cabeçalho com todos os resumos de procedimentos
    texto_chefe_final = texto_chefe_cabecalho + "\n".join(resumos_para_chefia)
    
    try:
        nome_arquivo_seguro = nome_do_paciente.lower().replace(' ', '_')
        caminho_chefe_txt = os.path.join(caminho_pasta_relatorios, f'resumo_chefe_{nome_arquivo_seguro}.txt')
        with open(caminho_chefe_txt, 'w', encoding='utf-8') as f:
            f.write(texto_chefe_final.strip())
        print(f"\n✅ Relatório para a chefia salvo com sucesso em: '{caminho_chefe_txt}'")
    except Exception as e:
        print(f"\n[ERRO] Falha ao salvar relatório da chefia: {e}")

def rodar_analise_automatica():
    # ... (esta função não precisa de nenhuma alteração)
    print("--- INICIANDO ANÁLISE AUTOMÁTICA PARA TODOS OS PACIENTES ---")
    try:
        df = pd.read_excel(ARQUIVO_ENTRADA_LIMPO)
    except FileNotFoundError:
        print(f"[ERRO] O arquivo '{ARQUIVO_ENTRADA_LIMPO}' não foi encontrado.")
        return
    
    lista_de_pacientes = df[COLUNA_PACIENTE].unique()
    print(f"Encontrados {len(lista_de_pacientes)} pacientes únicos na planilha.")
    
    for nome_paciente in lista_de_pacientes:
        print("-" * 50)
        gerar_relatorios_completos(df, nome_paciente)

    print("-" * 50)
    print("\nAnálise automática finalizada para todos os pacientes!")

if __name__ == "__main__":
    rodar_analise_automatica()