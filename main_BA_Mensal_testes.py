import os
import logging
import pandas as pd
from datetime import datetime
import shutil
from de_para_grupos import DE_PARA_GRUPOS
# --- Configuração de Log ---
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# --- CONFIGURAÇÕES FIXAS ---
NOT_APPLICABLE_VALUE = '__N/A__'
PEDIDOS_PARA_FILTRAR = ['Pedido Atendido', 'Pedido em Aberto', 'Pedido Encerrado', 'Pedido em Aprovação']
ITENS_PARA_FILTRAR = ['Item Atendido', 'Item em Aberto', 'Aprovar Alçada','Parcialmente Atendido']



def consolidar_simples(especificacoes_entrada, caminho_saida, projeto_alvo):
    COLUNA_GRUPO = 'nome_do_grupo'
    list_df = []
    
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year

    try:
        for spec in especificacoes_entrada:
            caminho = spec['nome']
            if not os.path.exists(caminho):
                logging.warning(f"Arquivo não encontrado: {caminho}")
                continue

            logging.info(f"Lendo: {caminho}")
            df = pd.read_excel(caminho)

            # --- CORREÇÃO: LÓGICA DE FILTRO INCLUSIVA (OU) ---
            col_doc = 'data_da_entrada_da_nota'
            col_emissao = 'dt_emissao_do_pedido' # ou 'data_emissao' conforme seu arquivo

            # Inicializamos as condições como Falsas (não filtram nada se a coluna não existir)
            condicao_data_doc = pd.Series(False, index=df.index)
            condicao_data_emissao = pd.Series(False, index=df.index)

            if col_doc in df.columns:
                df[col_doc] = pd.to_datetime(df[col_doc], errors='coerce')
                condicao_data_doc = (df[col_doc].dt.month == mes_atual) & (df[col_doc].dt.year == ano_atual)

            if col_emissao in df.columns:
                df[col_emissao] = pd.to_datetime(df[col_emissao], errors='coerce')
                condicao_data_emissao = (df[col_emissao].dt.month == mes_atual) & (df[col_emissao].dt.year == ano_atual)

            # Aplica o filtro: Se condicao_data_doc FOR VERDADEIRA OU condicao_data_emissao FOR VERDADEIRA
            df = df[condicao_data_doc | condicao_data_emissao]
            
            logging.info(f"Registros encontrados no mês {mes_atual}/{ano_atual}: {len(df)}")

            # --- TRATAMENTOS DE NULOS E STRINGS ---
            cols_check = ['projeto', 'situacao_do_pedido', 'situacao_do_item', COLUNA_GRUPO]
            for col in cols_check:
                if col in df.columns:
                    df[col] = df[col].fillna(NOT_APPLICABLE_VALUE).astype(str).str.strip()
            
            if 'valor_rateado' in df.columns:
                df['valor_rateado'] = pd.to_numeric(df['valor_rateado'], errors='coerce').fillna(0)

            # Filtro de Projeto
            df = df[df['projeto'] == projeto_alvo]

            # Filtro de Situação
            cond_pedido = (df['situacao_do_pedido'].isin(PEDIDOS_PARA_FILTRAR)) | (df['situacao_do_pedido'] == NOT_APPLICABLE_VALUE)
            cond_item = (df['situacao_do_item'].isin(ITENS_PARA_FILTRAR)) | (df['situacao_do_item'] == NOT_APPLICABLE_VALUE)
            
            df_filtrado = df[cond_pedido & cond_item].copy()
            list_df.append(df_filtrado)

        if not list_df or all(d.empty for d in list_df):
            logging.info("Nenhum dado encontrado após os filtros.")
            #return

        df_total = pd.concat(list_df, ignore_index=True)
        
        if COLUNA_GRUPO in df_total.columns:
            df_total[COLUNA_GRUPO] = df_total[COLUNA_GRUPO].replace(DE_PARA_GRUPOS)

        resumo_final = df_total.groupby(COLUNA_GRUPO)['valor_rateado'].sum().reset_index()

        with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
            resumo_final.to_excel(writer, sheet_name='Resumo_Fafen', index=False)
            logging.info(f"✅ Sucesso! Exportado para: {caminho_saida}")

    except Exception as e:
        logging.error(f"Erro no processamento: {e}")

if __name__ == "__main__":
    ARQUIVOS = [{'nome': r'C:\Users\operacoes\OneDrive - ENGEMAN\01-(MATRIZ) SETOR DE OPERAÇÕES - Documentos\01. MATRIZ - RECIFE\CONTROLADORIA\Controle de Mobilizações\Fafen\Controle Fafen Oficial\testes\base_FAFENS_filtrada_TESTE.xlsx'}]
    PROJETO = 'PB - FAFEN O&M BA - OS 177/25'
    SAIDA = r'C:\Users\operacoes\OneDrive - ENGEMAN\01-(MATRIZ) SETOR DE OPERAÇÕES - Documentos\01. MATRIZ - RECIFE\CONTROLADORIA\Controle de Mobilizações\Fafen\Controle Fafen Oficial\testes\Dados_FAFEN_BA_Mensal_testes.xlsx'
    consolidar_simples(ARQUIVOS, SAIDA, PROJETO)