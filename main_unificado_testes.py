import os
import logging
import pandas as pd
import datetime
import shutil
from de_para_grupos import DE_PARA_GRUPOS

# --- Configuração de Log ---
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# --- CONFIGURAÇÕES FIXAS ---
NOT_APPLICABLE_VALUE = '__N/A__'
PEDIDOS_PARA_FILTRAR = ['Pedido Atendido', 'Pedido em Aberto', 'Pedido Encerrado', 'Pedido em Aprovação']
ITENS_PARA_FILTRAR = ['Item Atendido', 'Item em Aberto', 'Aprovar Alçada','Parcialmente Atendido']

# --- NOVO: DICIONÁRIO DE SUBSTITUIÇÃO DE GRUPOS ---
# Coloque aqui: 'Nome Atual na Planilha': 'Nome Novo Desejado'


def consolidar_simples(especificacoes_entrada, caminho_saida, projetos_alvo):
    COLUNA_GRUPO = 'nome_do_grupo'
    list_df = []

    try:
        for spec in especificacoes_entrada:
            caminho = spec['nome']
            if not os.path.exists(caminho):
                logging.warning(f"Arquivo não encontrado: {caminho}")
                continue

            logging.info(f"Lendo: {caminho}")
            df = pd.read_excel(caminho)

            # 1. Tratamento de Nulos para garantir que vazios sejam considerados
            cols_check = ['projeto', 'situacao_do_pedido', 'situacao_do_item', COLUNA_GRUPO]
            for col in cols_check:
                if col in df.columns:
                    df[col] = df[col].fillna(NOT_APPLICABLE_VALUE).astype(str).str.strip()
            
            if 'valor_rateado' in df.columns:
                df['valor_rateado'] = pd.to_numeric(df['valor_rateado'], errors='coerce').fillna(0)

            # 2. Filtro de Projetos (Aceita a lista enviada)
            df = df[df['projeto'].isin(projetos_alvo)]

            # 3. Filtro de Situação (Pedido e Item) aceitando os vazios (__N/A__)
            cond_pedido = (df['situacao_do_pedido'].isin(PEDIDOS_PARA_FILTRAR)) | (df['situacao_do_pedido'] == NOT_APPLICABLE_VALUE)
            cond_item = (df['situacao_do_item'].isin(ITENS_PARA_FILTRAR)) | (df['situacao_do_item'] == NOT_APPLICABLE_VALUE)
            
            df_filtrado = df[cond_pedido & cond_item].copy()
            list_df.append(df_filtrado)

        if not list_df:
            logging.error("Nenhum dado encontrado para processar.")
            return

        # 4. Consolidação e Agrupamento Total
        df_total = pd.concat(list_df, ignore_index=True)
        
        # --- APLICAÇÃO DA SUBSTITUIÇÃO DE NOMES DE GRUPO ---
        if COLUNA_GRUPO in df_total.columns:
            df_total[COLUNA_GRUPO] = df_total[COLUNA_GRUPO].replace(DE_PARA_GRUPOS)

        # ALTERAÇÃO: Agrupa por 'projeto' e 'grupo' para que ambos apareçam no Excel final
        resumo_final = df_total.groupby(['projeto', COLUNA_GRUPO])['valor_rateado'].sum().reset_index()

        # 5. Exportação (Aba única)
        with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
            resumo_final.to_excel(writer, sheet_name='Resumo_Fafen', index=False, startcol=0)
            logging.info(f"✅ Sucesso! Total de {len(resumo_final)} linhas exportadas.")

    except Exception as e:
        logging.error(f"Erro no processamento: {e}")

# --- EXECUÇÃO ---
if __name__ == "__main__":
    ARQUIVOS = [{'nome': r'C:\Users\operacoes\OneDrive - ENGEMAN\01-(MATRIZ) SETOR DE OPERAÇÕES - Documentos\01. MATRIZ - RECIFE\CONTROLADORIA\Controle de Mobilizações\Fafen\Controle Fafen Oficial\testes\base_FAFENS_filtrada_TESTE.xlsx'}]
    
    # Exemplo com múltiplos projetos
    PROJETOS = [
        'PB - FAFEN O&M BA - OS 177/25',
        'PB - FAFEN O&M SE - OS 177/25'
    ]
    
    SAIDA = r'C:\Users\operacoes\OneDrive - ENGEMAN\01-(MATRIZ) SETOR DE OPERAÇÕES - Documentos\01. MATRIZ - RECIFE\CONTROLADORIA\Controle de Mobilizações\Fafen\Controle Fafen Oficial\testes\Resumo_FAFENS_BA_SE_testes.xlsx'

    consolidar_simples(ARQUIVOS, SAIDA, PROJETOS)