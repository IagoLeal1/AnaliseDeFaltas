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

# --- FUNÇÃO PARA RELATÓRIOS INDIVIDUAIS (ORIGINAL) ---
def gerar_relatorios_completos(df, nome_do_paciente):
    """
    Função que gera um kit completo de relatórios para um paciente específico.
    """
    print(f"\n🔎 --- GERANDO KIT COMPLETO DE RELATÓRIOS PARA: {nome_do_paciente} --- 🔎")
    
    # 1. PREPARAÇÃO
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
    
    resumos_para_chefia = []
    total_presencas_geral = 0
    total_faltas_geral = 0
    total_cancelados_geral = 0

    # 2. LOOP POR PROCEDIMENTO
    for procedimento in procedimentos_unicos:
        df_procedimento = df_paciente[df_paciente[COLUNA_PROCEDIMENTO] == procedimento]

        presencas = (df_procedimento[COLUNA_STATUS] == STATUS_PRESENTE).sum()
        faltas = (df_procedimento[COLUNA_STATUS] == STATUS_FALTOU).sum()
        cancelados = (df_procedimento[COLUNA_STATUS] == STATUS_CANCELADO).sum()
        total_valido = presencas + faltas
        taxa_de_falta = (faltas / total_valido) * 100 if total_valido > 0 else 0

        total_presencas_geral += presencas
        total_faltas_geral += faltas
        total_cancelados_geral += cancelados

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
        resumos_para_chefia.append(texto_relatorio_procedimento)

        nome_arquivo_base = re.sub(r'[\\/*?:"<>|]',"", procedimento).lower().replace(' ', '_')
        
        try:
            caminho_txt = os.path.join(caminho_pasta_relatorios, f'relatorio_{nome_arquivo_base}.txt')
            with open(caminho_txt, 'w', encoding='utf-8') as f:
                f.write(texto_relatorio_procedimento.strip())
        except Exception as e:
            print(f"     [ERRO] Falha ao salvar .txt: {e}")

        caminho_excel = os.path.join(caminho_pasta_relatorios, f'relatorio_{nome_arquivo_base}.xlsx')
        df_procedimento.to_excel(caminho_excel, index=False, engine='openpyxl')

        if total_valido > 0:
            caminho_grafico = os.path.join(caminho_pasta_graficos, f'grafico_{nome_arquivo_base}.png')
            fig, ax = plt.subplots()
            ax.pie([presencas, faltas, cancelados], labels=['Presenças', 'Faltas', 'Cancelados'], colors=['#2E8B57', '#DC143C', '#A9A9A9'], autopct='%1.1f%%', startangle=90)
            ax.axis('equal')
            plt.title(f'Resumo de: {procedimento}\nPaciente: {nome_do_paciente}')
            plt.savefig(caminho_grafico)
            plt.close(fig)

    # 3. GERAÇÃO DO RELATÓRIO MESTRE PARA A CHEFIA (DO PACIENTE)
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
    texto_chefe_final = texto_chefe_cabecalho + "\n".join(resumos_para_chefia)
    
    try:
        nome_arquivo_seguro = nome_do_paciente.lower().replace(' ', '_')
        caminho_chefe_txt = os.path.join(caminho_pasta_relatorios, f'resumo_chefe_{nome_arquivo_seguro}.txt')
        with open(caminho_chefe_txt, 'w', encoding='utf-8') as f:
            f.write(texto_chefe_final.strip())
        print(f"\n✅ Relatório para a chefia salvo com sucesso em: '{caminho_chefe_txt}'")
    except Exception as e:
        print(f"\n[ERRO] Falha ao salvar relatório da chefia: {e}")


# --- FUNÇÃO PARA RELATÓRIO GERAL (COM TODAS AS MELHORIAS) ---
def gerar_relatorio_geral_consolidado(df):
    """
    Gera um relatório consolidado com a análise de todos os pacientes.
    Cria um relatório .txt formatado para fácil leitura pela gestão.
    """
    print("\n🔎 --- GERANDO RELATÓRIO GERAL CONSOLIDADO (TODOS OS PACIENTES) --- 🔎")

    # 1. PREPARAÇÃO DOS CAMINHOS
    caminho_pasta_relatorios = os.path.join('..', 'relatorios', 'FALTAS_TOTAIS_PACIENTES')
    os.makedirs(caminho_pasta_relatorios, exist_ok=True)
    
    # 2. CÁLCULO GERAL
    total_presencas = (df[COLUNA_STATUS] == STATUS_PRESENTE).sum()
    total_faltas = (df[COLUNA_STATUS] == STATUS_FALTOU).sum()
    total_cancelados = (df[COLUNA_STATUS] == STATUS_CANCELADO).sum()
    total_valido = total_presencas + total_faltas
    taxa_falta_geral = (total_faltas / total_valido) * 100 if total_valido > 0 else 0

    # 3. ANÁLISE POR PACIENTE
    df_pacientes = df.groupby([COLUNA_PACIENTE, COLUNA_STATUS]).size().unstack(fill_value=0)
    for status in [STATUS_PRESENTE, STATUS_FALTOU, STATUS_CANCELADO]:
        if status not in df_pacientes.columns:
            df_pacientes[status] = 0
    df_pacientes['Total_Valido'] = df_pacientes[STATUS_PRESENTE] + df_pacientes[STATUS_FALTOU]
    df_pacientes['Taxa_Falta_%'] = (df_pacientes[STATUS_FALTOU] / df_pacientes['Total_Valido'] * 100).fillna(0)
    df_pacientes = df_pacientes.sort_values(by=STATUS_FALTOU, ascending=False)

    # 4. ANÁLISE POR PROCEDIMENTO
    df_procedimentos = df.groupby([COLUNA_PROCEDIMENTO, COLUNA_STATUS]).size().unstack(fill_value=0)
    for status in [STATUS_PRESENTE, STATUS_FALTOU, STATUS_CANCELADO]:
        if status not in df_procedimentos.columns:
            df_procedimentos[status] = 0
    df_procedimentos['Total_Valido'] = df_procedimentos[STATUS_PRESENTE] + df_procedimentos[STATUS_FALTOU]
    df_procedimentos['Taxa_Falta_%'] = (df_procedimentos[STATUS_FALTOU] / df_procedimentos['Total_Valido'] * 100).fillna(0)
    df_procedimentos = df_procedimentos.sort_values(by='Taxa_Falta_%', ascending=False)
    
    # 5. MONTAGEM DO RELATÓRIO DE TEXTO (.txt) COM FORMATAÇÃO MELHORADA
    
    # Bloco de Desempenho por Paciente
    texto_desempenho_pacientes = "--- DESEMPENHO POR PACIENTE (ORDENADO POR Nº DE FALTAS) ---\n\n"
    header_pac = f"{'PACIENTE':<40} | {'FALTAS':^10} | {'PRESENÇAS':^11} | {'TAXA DE FALTA':^15}\n"
    separator_pac = f"{'-'*40}+{'-'*12}+{'-'*13}+{'-'*17}\n"
    texto_desempenho_pacientes += header_pac + separator_pac
    for nome, row in df_pacientes.iterrows():
        nome_paciente_trunc = str(nome)[:39]
        linha = f"{nome_paciente_trunc:<40} | {row[STATUS_FALTOU]:^10} | {row[STATUS_PRESENTE]:^11} | {f'{row["Taxa_Falta_%"]:.1f}%':^15}\n"
        texto_desempenho_pacientes += linha

    # Bloco de Análise por Procedimento
    texto_analise_procedimentos = "--- ANÁLISE POR PROCEDIMENTO (ORDENADO POR TAXA DE FALTA) ---\n\n"
    header_proc = f"{'PROCEDIMENTO':<40} | {'FALTAS':^10} | {'PRESENÇAS':^11} | {'TAXA DE FALTA':^15}\n"
    separator_proc = f"{'-'*40}+{'-'*12}+{'-'*13}+{'-'*17}\n"
    texto_analise_procedimentos += header_proc + separator_proc
    for nome, row in df_procedimentos.iterrows():
        nome_proc_trunc = str(nome)[:39]
        linha = f"{nome_proc_trunc:<40} | {row[STATUS_FALTOU]:^10} | {row[STATUS_PRESENTE]:^11} | {f'{row["Taxa_Falta_%"]:.1f}%':^15}\n"
        texto_analise_procedimentos += linha
        
    # Montagem do texto final
    texto_relatorio_geral = f"""
=====================================================
    RELATÓRIO CONSOLIDADO GERAL PARA A CHEFIA
=====================================================
Data da Geração: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}

--- RESUMO GERAL (TODA A CLÍNICA) ---
Consultas finalizadas (presenças): {total_presencas}
Consultas não comparecidas (faltas): {total_faltas}
Consultas canceladas: {total_cancelados}
-----------------------------------------------------
📊 Taxa de Falta GERAL (sobre consultas válidas): {taxa_falta_geral:.2f}%
=====================================================

{texto_desempenho_pacientes}
=====================================================

{texto_analise_procedimentos}
"""
    
    try:
        caminho_txt_geral = os.path.join(caminho_pasta_relatorios, 'relatorio_consolidado_geral.txt')
        with open(caminho_txt_geral, 'w', encoding='utf-8') as f:
            f.write(texto_relatorio_geral.strip())
        print(f"\n✅ Relatório .txt consolidado e formatado salvo com sucesso em: '{caminho_txt_geral}'")
    except Exception as e:
        print(f"\n[ERRO] Falha ao salvar relatório .txt consolidado: {e}")

    # 6. GERAÇÃO DO EXCEL (.xlsx)
    try:
        caminho_excel_geral = os.path.join(caminho_pasta_relatorios, 'relatorio_consolidado_completo.xlsx')
        with pd.ExcelWriter(caminho_excel_geral, engine='openpyxl') as writer:
            df_pacientes_excel = df_pacientes.rename(columns={
                STATUS_FALTOU: 'Faltas',
                STATUS_PRESENTE: 'Presenças',
                STATUS_CANCELADO: 'Cancelados'
            })
            df_procedimentos_excel = df_procedimentos.rename(columns={
                STATUS_FALTOU: 'Faltas',
                STATUS_PRESENTE: 'Presenças',
                STATUS_CANCELADO: 'Cancelados'
            })
            df_pacientes_excel.to_excel(writer, sheet_name='Resumo_por_Paciente')
            df_procedimentos_excel.to_excel(writer, sheet_name='Resumo_por_Procedimento')
            df.to_excel(writer, sheet_name='Dados_Completos', index=False)
        print(f"✅ Relatório .xlsx consolidado salvo com sucesso em: '{caminho_excel_geral}'")
    except Exception as e:
        print(f"[ERRO] Falha ao salvar relatório .xlsx consolidado: {e}")


# --- FUNÇÕES PARA RODAR AS ANÁLISES ---
def rodar_analise_individual():
    """
    Roda a análise original, gerando um kit de relatório para cada paciente.
    """
    print("--- INICIANDO ANÁLISE INDIVIDUAL PARA CADA PACIENTE ---")
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
    print("\nAnálise individual finalizada para todos os pacientes!")


# --- BLOCO PRINCIPAL COM MENU DE ESCOLHA ---
if __name__ == "__main__":
    while True:
        print("\n" + "="*30)
        print("   MENU DE ANÁLISE DE FALTAS")
        print("="*30)
        print("1. Gerar relatórios INDIVIDUAIS por paciente")
        print("2. Gerar relatório GERAL CONSOLIDADO de todos os pacientes")
        print("3. Sair")
        
        escolha = input("\nDigite sua opção (1, 2 ou 3): ")

        if escolha == '1':
            rodar_analise_individual()
            break
        elif escolha == '2':
            try:
                df_geral = pd.read_excel(ARQUIVO_ENTRADA_LIMPO)
                gerar_relatorio_geral_consolidado(df_geral)
            except FileNotFoundError:
                print(f"[ERRO] O arquivo '{ARQUIVO_ENTRADA_LIMPO}' não foi encontrado.")
            break
        elif escolha == '3':
            print("Saindo do programa.")
            break
        else:
            print("[AVISO] Opção inválida. Por favor, escolha 1, 2 ou 3.")