import logging
import pandas as pd
from pathlib import Path
from typing import List, Tuple, Any, Dict, Optional, IO

# --- Função Auxiliar de Leitura de Linhas ---
def _processar_linhas_sped(
    f: IO[Any],
    dados_completos: List[Dict],
    dados_itens_sped: List[Dict],
    dados_analiticos_sped: List[Dict],
    dados_cte_sped_d190: List[Dict],
    chaves_com_c101: set # <--- NOVO: Recebe o conjunto para guardar chaves com DIFAL
) -> None:
    """Função auxiliar para processar as linhas de um arquivo SPED aberto."""

    # Variáveis de Estado
    current_invoice_data: Dict[str, Any] = {}
    current_cfops_nfe: set[str] = set()
    current_chv_nfe: str = ''
    current_chv_cte: str = ''
    current_chv_energia: str = ''
    current_chv_comunicacao: str = ''

    for linha in f:
        campos = linha.strip().split('|')
        reg_type = campos[1] if len(campos) > 1 else None

        # --- Bloco C (NF-e Mercadorias) ---
        if reg_type == 'C100':
            if current_invoice_data:
                current_invoice_data['CFOP_SPED'] = '/'.join(sorted(list(current_cfops_nfe))) if current_cfops_nfe else ''
                dados_completos.append(current_invoice_data)
            current_cfops_nfe = set(); current_chv_nfe = ''
            current_chv_cte = ''; current_chv_energia = ''; current_chv_comunicacao = ''

            if len(campos) > 27:
                current_invoice_data = {
                    'CHV_NFE': campos[9], 'VL_DOC_SPED': campos[12],
                    'ICMS_SPED': campos[22], 'ICMS_ST_SPED': campos[23],
                    'IPI_SPED': campos[25], 'PIS_SPED': campos[26], 'COFINS_SPED': campos[27],
                    'FCP_ST_SPED': '0,00', 'IPI_DEVOL_SPED': '0,00',
                    'ICMS_SN_SPED': '0,00', 'ICMS_MONO_SPED': '0,00',
                    'TIPO_NOTA_SPED': ''
                }
                current_chv_nfe = campos[9]
            else: current_invoice_data = {}

        # --- NOVO: Captura DIFAL (C101) ---
        elif reg_type == 'C101' and current_chv_nfe:
            # Se a nota tem registro C101, salvamos a chave dela
            chaves_com_c101.add(current_chv_nfe)

        elif reg_type == 'C170' and current_chv_nfe:
            if len(campos) > 11 and campos[11]: current_cfops_nfe.add(campos[11])
            if len(campos) > 11:
                vlr_ipi_item_sped = 0.0
                if len(campos) > 24:
                    try:
                        vl_ipi_str = campos[24]
                        if vl_ipi_str: vlr_ipi_item_sped = float(vl_ipi_str.replace(',', '.'))
                    except (ValueError, TypeError) as e:
                        logging.warning(f"Erro ao ler IPI: {e}")
                        vlr_ipi_item_sped = 0.0
                vlr_ipi_item_str = str(vlr_ipi_item_sped).replace('.', ',')

                dados_itens_sped.append({
                    'CHV_NFE': current_chv_nfe,
                    'N_ITEM_SPED': campos[2],
                    'COD_PROD_SPED': campos[3],
                    'CFOP_SPED_ITEM': campos[11],
                    'CST_ICMS_SPED_ITEM': campos[10] if len(campos) > 10 else '',
                    'VL_OPR_SPED_ITEM': campos[7] if len(campos) > 7 else '0,00',
                    'VL_BC_ICMS_SPED_ITEM': campos[13] if len(campos) > 13 else '0,00',
                    'VL_ICMS_SPED_ITEM': campos[15] if len(campos) > 15 else '0,00',
                    'VL_BC_ICMS_ST_SPED_ITEM': campos[16] if len(campos) > 16 else '0,00',
                    'VL_ICMS_ST_SPED_ITEM': campos[18] if len(campos) > 18 else '0,00',
                    'VLR_IPI_SPED_ITEM': vlr_ipi_item_str
                })

        elif reg_type == 'C190' and current_chv_nfe:
            if len(campos) > 11:
                if campos[3]: current_cfops_nfe.add(campos[3])
                dados_analiticos_sped.append({
                    'CHV_NFE': current_chv_nfe, 'CST_ICMS_SPED_ITEM': campos[2],
                    'CFOP_SPED_ITEM': campos[3], 'ALIQ_ICMS_SPED_ITEM': campos[4],
                    'VL_OPR_SPED_ITEM': campos[5], 'VL_BC_ICMS_SPED_ITEM': campos[6],
                    'VL_ICMS_SPED_ITEM': campos[7], 'VL_BC_ICMS_ST_SPED_ITEM': campos[8],
                    'VL_ICMS_ST_SPED_ITEM': campos[9], 'VLR_IPI_SPED_ITEM': campos[11]
                })

        # --- Bloco D (CT-e) ---
        elif reg_type == 'D100':
            if current_invoice_data:
                current_invoice_data['CFOP_SPED'] = '/'.join(sorted(list(current_cfops_nfe))) if current_cfops_nfe else ''
                dados_completos.append(current_invoice_data)
            current_invoice_data = {}; current_cfops_nfe = set(); current_chv_nfe = ''
            current_chv_cte = ''; current_chv_energia = ''; current_chv_comunicacao = ''
            if len(campos) > 9:
                current_chv_cte = campos[9]

        elif reg_type == 'D190' and current_chv_cte:
            if len(campos) > 9:
                dados_cte_sped_d190.append({
                    'CHV_CTE': current_chv_cte, 'CST_ICMS_SPED_D190': campos[2],
                    'CFOP_SPED_D190': campos[3], 'ALIQ_ICMS_SPED_D190': campos[4],
                    'VL_OPR_SPED_D190': campos[5], 'VL_BC_ICMS_SPED_D190': campos[6],
                    'VL_ICMS_SPED_D190': campos[7], 'VL_RED_BC_SPED_D190': campos[8],
                    'COD_OBS_SPED_D190': campos[9]
                })
                dados_analiticos_sped.append({
                    'CHV_NFE': current_chv_cte,
                    'CST_ICMS_SPED_ITEM': campos[2], 'CFOP_SPED_ITEM': campos[3],
                    'ALIQ_ICMS_SPED_ITEM': campos[4], 'VL_OPR_SPED_ITEM': campos[5],
                    'VL_BC_ICMS_SPED_ITEM': campos[6], 'VL_ICMS_SPED_ITEM': campos[7],
                    'VL_BC_ICMS_ST_SPED_ITEM': '0,00', 'VL_ICMS_ST_SPED_ITEM': '0,00',
                    'VLR_IPI_SPED_ITEM': '0,00'
                })

        # --- Bloco C (Energia) ---
        elif reg_type == 'C500':
            if current_invoice_data:
                current_invoice_data['CFOP_SPED'] = '/'.join(sorted(list(current_cfops_nfe))) if current_cfops_nfe else ''
                dados_completos.append(current_invoice_data)
            current_invoice_data = {}; current_cfops_nfe = set(); current_chv_nfe = ''
            current_chv_cte = ''; current_chv_energia = ''; current_chv_comunicacao = ''
            if len(campos) > 23:
                chv_energia_c500 = campos[10]
                current_chv_energia = chv_energia_c500 if chv_energia_c500 else f"Energia_{campos[6]}_{campos[9]}"
                dados_completos.append({
                    'CHV_NFE': current_chv_energia,
                    'VL_DOC_SPED': campos[12],
                    'ICMS_SPED': campos[18],
                    'ICMS_ST_SPED': '0,00', 'IPI_SPED': '0,00',
                    'PIS_SPED': campos[22], 'COFINS_SPED': campos[23],
                    'FCP_ST_SPED': '0,00', 'IPI_DEVOL_SPED': '0,00',
                    'ICMS_SN_SPED': '0,00', 'ICMS_MONO_SPED': '0,00',
                    'CFOP_SPED': campos[8],
                    'TIPO_NOTA_SPED': 'Energia Elétrica (C500)'
                })

        elif reg_type == 'C590' and current_chv_energia:
            if len(campos) > 10:
                dados_analiticos_sped.append({
                    'CHV_NFE': current_chv_energia,
                    'CST_ICMS_SPED_ITEM': campos[2], 'CFOP_SPED_ITEM': campos[3],
                    'ALIQ_ICMS_SPED_ITEM': campos[4], 'VL_OPR_SPED_ITEM': campos[5],
                    'VL_BC_ICMS_SPED_ITEM': campos[6], 'VL_ICMS_SPED_ITEM': campos[7],
                    'VL_BC_ICMS_ST_SPED_ITEM': campos[8], 'VL_ICMS_ST_SPED_ITEM': campos[9],
                    'VLR_IPI_SPED_ITEM': '0,00'
                })

        # --- Bloco D (Comunicação) ---
        elif reg_type == 'D500':
            if current_invoice_data:
                current_invoice_data['CFOP_SPED'] = '/'.join(sorted(list(current_cfops_nfe))) if current_cfops_nfe else ''
                dados_completos.append(current_invoice_data)
            current_invoice_data = {}; current_cfops_nfe = set(); current_chv_nfe = ''
            current_chv_cte = ''; current_chv_energia = ''; current_chv_comunicacao = ''
            if len(campos) > 21:
                current_chv_comunicacao = f"Comunicação_{campos[6]}_{campos[9]}"
                dados_completos.append({
                    'CHV_NFE': current_chv_comunicacao,
                    'VL_DOC_SPED': campos[11],
                    'ICMS_SPED': campos[17],
                    'ICMS_ST_SPED': '0,00', 'IPI_SPED': '0,00',
                    'PIS_SPED': campos[19], 'COFINS_SPED': campos[21],
                    'FCP_ST_SPED': '0,00', 'IPI_DEVOL_SPED': '0,00',
                    'ICMS_SN_SPED': '0,00', 'ICMS_MONO_SPED': '0,00',
                    'CFOP_SPED': campos[8],
                    'TIPO_NOTA_SPED': 'Comunicação (D500)'
                })

        elif reg_type == 'D590' and current_chv_comunicacao:
            if len(campos) > 10:
                dados_analiticos_sped.append({
                    'CHV_NFE': current_chv_comunicacao,
                    'CST_ICMS_SPED_ITEM': campos[2], 'CFOP_SPED_ITEM': campos[3],
                    'ALIQ_ICMS_SPED_ITEM': campos[4], 'VL_OPR_SPED_ITEM': campos[5],
                    'VL_BC_ICMS_SPED_ITEM': campos[6], 'VL_ICMS_SPED_ITEM': campos[7],
                    'VL_BC_ICMS_ST_SPED_ITEM': campos[8], 'VL_ICMS_ST_SPED_ITEM': campos[9],
                    'VLR_IPI_SPED_ITEM': '0,00'
                })

    if current_invoice_data:
        current_invoice_data['CFOP_SPED'] = '/'.join(sorted(list(current_cfops_nfe))) if current_cfops_nfe else ''
        dados_completos.append(current_invoice_data)


# --- Função Principal de Extração ---
def extrair_dados_sped(caminho_arquivo_sped: Path) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Retorna 5 DataFrames:
    1. df_sped (Cabeçalhos C100/D100/etc)
    2. df_sped_itens (C170)
    3. df_sped_analitico (C190/D190/etc)
    4. df_sped_cte (D190 específico CTE)
    5. df_chaves_difal (NOVO: Apenas chaves que têm C101)
    """
    logging.info('Lendo e processando arquivo SPED...')

    dados_completos: List[Dict[str, Any]] = []
    dados_itens_sped: List[Dict[str, Any]] = []
    dados_analiticos_sped: List[Dict[str, Any]] = []
    dados_cte_sped_d190: List[Dict[str, Any]] = []
    chaves_com_c101: set = set() # Set para evitar duplicatas

    encoding_to_try = 'latin-1'

    try:
        with open(caminho_arquivo_sped, 'r', encoding=encoding_to_try) as f:
            _processar_linhas_sped(f, dados_completos, dados_itens_sped, dados_analiticos_sped, dados_cte_sped_d190, chaves_com_c101)
    except UnicodeDecodeError:
        logging.warning(f"Falha ao ler SPED com {encoding_to_try}. Tentando utf-8...")
        encoding_to_try = 'utf-8'
        try:
            with open(caminho_arquivo_sped, 'r', encoding=encoding_to_try) as f:
                _processar_linhas_sped(f, dados_completos, dados_itens_sped, dados_analiticos_sped, dados_cte_sped_d190, chaves_com_c101)
        except Exception as e:
            raise Exception(f"Erro inesperado ao ler SPED (utf-8): {e}")
    except Exception as e:
        raise Exception(f"Erro inesperado ao ler SPED: {e}")

    # --- Criação dos DataFrames ---

    # 1. Cabeçalhos
    df_sped = pd.DataFrame(dados_completos)
    if not df_sped.empty:
        numeric_cols_sped = ['VL_DOC_SPED', 'ICMS_SPED', 'ICMS_ST_SPED', 'IPI_SPED', 'PIS_SPED', 'COFINS_SPED', 'IPI_DEVOL_SPED', 'FCP_ST_SPED', 'ICMS_SN_SPED', 'ICMS_MONO_SPED']
        for col in numeric_cols_sped:
            if col in df_sped.columns:
                df_sped[col] = pd.to_numeric(df_sped[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0).round(2)
            else:
                df_sped[col] = 0.0
        string_cols_sped = ['CHV_NFE', 'CFOP_SPED', 'TIPO_NOTA_SPED']
        for col in string_cols_sped:
            if col not in df_sped.columns: df_sped[col] = ''
        df_sped = df_sped.fillna('')
        if 'CHV_NFE' in df_sped.columns: df_sped.drop_duplicates(subset=['CHV_NFE'], keep='first', inplace=True)

    # 2. Itens (C170)
    df_sped_itens = pd.DataFrame(dados_itens_sped)
    numeric_sped_item_cols_c170 = ['VL_OPR_SPED_ITEM', 'VL_BC_ICMS_SPED_ITEM', 'VL_ICMS_SPED_ITEM', 'VL_BC_ICMS_ST_SPED_ITEM', 'VL_ICMS_ST_SPED_ITEM', 'VLR_IPI_SPED_ITEM']
    if not df_sped_itens.empty:
        df_sped_itens = df_sped_itens.fillna('')
        for col in numeric_sped_item_cols_c170:
             if col in df_sped_itens.columns:
                  df_sped_itens[col] = pd.to_numeric(df_sped_itens[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0).round(2)
             else:
                  df_sped_itens[col] = 0.0
        if 'CST_ICMS_SPED_ITEM' not in df_sped_itens.columns: df_sped_itens['CST_ICMS_SPED_ITEM'] = ''
        df_sped_itens['N_ITEM_SPED'] = df_sped_itens['N_ITEM_SPED'].astype(str)
        df_sped_itens.drop_duplicates(subset=['CHV_NFE', 'N_ITEM_SPED'], keep='first', inplace=True)

    # 3. Analíticos (C190)
    df_sped_analitico = pd.DataFrame(dados_analiticos_sped)
    numeric_analitico = ['ALIQ_ICMS_SPED_ITEM', 'VL_OPR_SPED_ITEM', 'VL_BC_ICMS_SPED_ITEM', 'VL_ICMS_SPED_ITEM', 'VL_BC_ICMS_ST_SPED_ITEM', 'VL_ICMS_ST_SPED_ITEM', 'VLR_IPI_SPED_ITEM']
    if not df_sped_analitico.empty:
         df_sped_analitico = df_sped_analitico.fillna('')
         for col in numeric_analitico:
             if col in df_sped_analitico.columns:
                 df_sped_analitico[col] = pd.to_numeric(df_sped_analitico[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0).round(2)
             else:
                 df_sped_analitico[col] = 0.0

    # 4. CTE Específico
    df_sped_cte = pd.DataFrame(dados_cte_sped_d190)
    numeric_cte = ['ALIQ_ICMS_SPED_D190', 'VL_OPR_SPED_D190', 'VL_BC_ICMS_SPED_D190', 'VL_ICMS_SPED_D190']
    if not df_sped_cte.empty:
         if 'VL_RED_BC_SPED_D190' in df_sped_cte.columns: df_sped_cte.drop(columns=['VL_RED_BC_SPED_D190'], inplace=True)
         if 'COD_OBS_SPED_D190' in df_sped_cte.columns: df_sped_cte.drop(columns=['COD_OBS_SPED_D190'], inplace=True)
         df_sped_cte = df_sped_cte.fillna('')
         for col in numeric_cte:
             if col in df_sped_cte.columns:
                  df_sped_cte[col] = pd.to_numeric(df_sped_cte[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0).round(2)
             else:
                  df_sped_cte[col] = 0.0

    # 5. NOVO: Chaves com DIFAL (C101)
    df_chaves_difal = pd.DataFrame(list(chaves_com_c101), columns=['CHV_NFE'])

    return df_sped, df_sped_itens, df_sped_analitico, df_sped_cte, df_chaves_difal