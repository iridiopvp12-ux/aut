# src/logic/fiscal_logic.py
import logging
import time
import os
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Tuple, Any, Dict, Optional, IO, Callable

# --- IMPORTAÇÕES DOS MÓDULOS ---
from .sped_parser import extrair_dados_sped
from .xml_parser import processar_pasta_xml
from .rules_parser import ler_regras_acumuladores
from .report_generator import gerar_relatorio_excel
from .core_logic import (
    get_acumulador,
    check_cfop_status,
    calcular_status_geral,
    _executar_analise_detalhada_interna,
    _calcular_totalizadores_cfop_cst
)

# Importa a lógica de apuração padrão (COMERCIO)
from .apuracao_logic import preencher_template_apuracao

# Variável global para os itens
df_itens_global: Optional[pd.DataFrame] = None

def setup_logging(base_path: Path, username: Optional[str] = None) -> Optional[Path]:
    logs_dir: Path = base_path / 'Logs_Analisador'
    log_filename_path: Optional[Path] = None
    try:
        logs_dir.mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        log_filename = f'analise_{username or "unknown"}_{timestamp}.log'
        log_filename_path = logs_dir / log_filename
    except OSError as e:
        print(f"Não foi possível criar a pasta de logs.\nCaminho: {logs_dir}\nErro: {e}")

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    if logger.hasHandlers(): logger.handlers.clear()

    if log_filename_path:
        try:
            file_handler = logging.FileHandler(log_filename_path, encoding='utf-8')
            file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(file_formatter)
            logger.addHandler(file_handler)
        except IOError as e:
            print(f"Não foi possível escrever no arquivo de log: {e}")
            log_filename_path = None

    if username: logging.info(f"Usuário '{username}' iniciou a sessão no programa.")
    return log_filename_path


# --- FUNÇÃO ORQUESTRADORA ---
def executar_analise_completa(
    caminho_sped: Path, pasta_xmls: Path, caminho_regras: Path, username: str,
    cfop_sem_credito_icms: List[str], cfop_sem_credito_ipi: List[str], tolerancia_valor: float,
    status_callback: Optional[Callable[[str], None]] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    done_callback: Optional[Callable[[Path, int], None]] = None,
    error_callback: Optional[Callable[[str], None]] = None,
    caminho_regras_detalhadas: Optional[Path] = None,
    template_apuracao_path: Optional[Path] = None,
    tipo_setor: str = 'Comercio',
    regras_cliente: Dict[str, Any] = None # <--- REGRAS DO CADASTRO DE CLIENTES
) -> None:

    global df_itens_global
    try:
        logging.info(f"Análise iniciada pelo usuário: {username}. Setor selecionado: {tipo_setor}")
        if status_callback: status_callback("Iniciando extração do SPED...")

        # 1. Aplicação das Regras do Cliente
        regras_cliente = regras_cliente or {}
        ignorar_pis_cofins = regras_cliente.get('nao_calcular_pis_cofins', False)
        exigir_acumulador = regras_cliente.get('exigir_acumulador', False)

        if ignorar_pis_cofins:
            logging.info("REGRA ATIVA: Não calcular PIS/COFINS (Simples Nacional).")
        if exigir_acumulador:
            logging.info("REGRA ATIVA: Exigir Acumulador preenchido.")

        # 2. Extração de dados
        logging.info("Iniciando extração do SPED...")
        df_sped, df_sped_itens, df_sped_analitico_combinado, df_sped_cte_d190, df_chaves_difal = extrair_dados_sped(caminho_sped)

        logging.info("Iniciando extração dos XMLs (NF-e e CT-e)...")
        if status_callback: status_callback("Processando XMLs...")
        df_xml_totais, df_xml_itens, df_xml_cte_totais = processar_pasta_xml(pasta_xmls, progress_callback)
        df_itens_global = df_xml_itens

        logging.info("Iniciando leitura das regras...")
        df_regras = ler_regras_acumuladores(caminho_regras)
        regras_map = df_regras.set_index(['CNPJ_CPF', 'CFOP'])['ACUMULADOR'].to_dict()

        df_recon_relatorio = pd.DataFrame()
        df_itens_aba = pd.DataFrame()
        df_aliquota_aba = pd.DataFrame()
        total_problemas = 0

        # -------------------------------------------------------------------------
        # 3. Conciliação TOTAL DA NOTA (C100, C500, D500 vs XML NF-e)
        # -------------------------------------------------------------------------
        logging.info('Cruzando dados SPED (C100, C500, D500) x XML (NF-e)...')
        if status_callback: status_callback("Cruzando dados SPED x XML...")

        df_recon = pd.merge(df_xml_totais, df_sped, on='CHV_NFE', how='outer', indicator=True)

        df_recon['SITUACAO_NOTA'] = np.select(
            [df_recon['_merge'] == 'left_only', df_recon['_merge'] == 'right_only'],
            ['FALTA NO SPED', 'FALTA XML'],
            default='OK'
        )
        df_recon.drop(columns=['_merge'], inplace=True)

        # Tratamento de Nulos
        numeric_cols = ['VL_DOC_XML', 'VL_DOC_SPED', 'ICMS_XML', 'ICMS_SPED', 'ICMS_ST_XML', 'ICMS_ST_SPED', 'FCP_ST_XML', 'FCP_ST_SPED', 'ICMS_SN_XML', 'ICMS_SN_SPED', 'ICMS_MONO_XML', 'ICMS_MONO_SPED', 'IPI_XML', 'IPI_SPED', 'IPI_DEVOL_XML', 'IPI_DEVOL_SPED', 'PIS_SPED', 'COFINS_SPED']
        string_cols = [ 'CHV_NFE', 'NUM_NF', 'CNPJ_EMITENTE', 'CFOP_XML', 'CFOP_SPED', 'CEST_XML', 'TIPO_NOTA', 'TIPO_NOTA_SPED' ]

        for col in numeric_cols:
            if col not in df_recon.columns: df_recon[col] = 0.0
        for col in string_cols:
            if col not in df_recon.columns: df_recon[col] = ''

        df_recon[numeric_cols] = df_recon[numeric_cols].fillna(0).round(2)
        df_recon[string_cols] = df_recon[string_cols].fillna('')

        df_recon['TIPO_NOTA'] = np.where(
            (df_recon['TIPO_NOTA'] == '') & (df_recon['TIPO_NOTA_SPED'] != ''),
            df_recon['TIPO_NOTA_SPED'],
            df_recon['TIPO_NOTA']
        )
        df_recon.loc[(df_recon['SITUACAO_NOTA'] == 'OK') & (df_recon['CNPJ_EMITENTE'] == ''), 'SITUACAO_NOTA'] = 'SEM CNPJ NO XML'

        logging.info('Aplicando regras de acumuladores (NF-e, C500, D500)...')
        df_recon['ACUMULADOR'] = df_recon.apply(get_acumulador, axis=1, regras_map=regras_map)

        df_recon['ICMS_TOTAL_XML'] = (df_recon['ICMS_XML'] + df_recon['ICMS_SN_XML']).round(2)
        df_recon['IPI_TOTAL_XML'] = (df_recon['IPI_XML'] + df_recon['IPI_DEVOL_XML']).round(2)

        # Ajuste IPI Devolução
        condicao_devolucao_ipi = (
            (df_recon['IPI_XML'] == 0) &
            (df_recon['IPI_DEVOL_XML'] > 0) &
            (df_recon['IPI_TOTAL_XML'] == df_recon['IPI_DEVOL_XML'])
        )
        df_recon['IPI_SPED'] = np.where(
            condicao_devolucao_ipi,
            df_recon['IPI_TOTAL_XML'],
            df_recon['IPI_SPED']
        )

        df_recon['STATUS_CFOP'] = df_recon.apply(check_cfop_status, axis=1)

        # Verificação de Impostos
        impostos_a_verificar = ['ICMS', 'ICMS_ST', 'IPI', 'FCP_ST', 'ICMS_MONO']
        for imposto in impostos_a_verificar:
            sped_col, status_col = f'{imposto}_SPED', f'STATUS_{imposto}'; xml_col = f'{imposto}_XML'; xml_total_col = f'{imposto}_TOTAL_XML' if imposto in ['ICMS', 'IPI'] else xml_col
            if sped_col not in df_recon.columns: df_recon[sped_col] = 0.0
            if xml_total_col not in df_recon.columns: df_recon[xml_total_col] = df_recon[xml_col] if xml_col in df_recon.columns else 0.0
            cond_cfop_sem_credito = pd.Series(False, index=df_recon.index)
            cfop_sped_col_exists = 'CFOP_SPED' in df_recon.columns
            if imposto == 'ICMS' and cfop_sem_credito_icms and cfop_sped_col_exists:
                cond_cfop_sem_credito = df_recon['CFOP_SPED'].apply(lambda x: isinstance(x, str) and any(cfop in x.split('/') for cfop in cfop_sem_credito_icms))
            elif imposto == 'IPI' and cfop_sem_credito_ipi and cfop_sped_col_exists:
                cond_cfop_sem_credito = df_recon['CFOP_SPED'].apply(lambda x: isinstance(x, str) and any(cfop in x.split('/') for cfop in cfop_sem_credito_ipi))
            cond_valores_iguais = (df_recon[xml_total_col] - df_recon[sped_col]).abs() <= tolerancia_valor
            df_recon[status_col] = np.where(cond_valores_iguais | cond_cfop_sem_credito, 'OK', 'DIVERGENTE')

        cond_valor_divergente = (df_recon['VL_DOC_XML'] - df_recon['VL_DOC_SPED']).abs() > tolerancia_valor
        df_recon['STATUS_VALOR'] = np.where(cond_valor_divergente & (df_recon['SITUACAO_NOTA'] == 'OK'), 'DIVERGENTE', 'OK')

        # PIS/COFINS Calculado
        if df_itens_global is not None and 'BC_PIS_COFINS_CALC' in df_itens_global.columns:
            df_itens_sum_bc = df_itens_global.groupby('CHV_NFE')['BC_PIS_COFINS_CALC'].sum().round(2).reset_index()
            df_recon = pd.merge(df_recon, df_itens_sum_bc, on='CHV_NFE', how='left')
            df_recon['BC_PIS_COFINS_CALC'] = df_recon['BC_PIS_COFINS_CALC'].fillna(0)
        else:
            df_recon['BC_PIS_COFINS_CALC'] = 0.0

        df_recon['BC_PIS_COFINS_CALC'] = df_recon['BC_PIS_COFINS_CALC'].apply(lambda x: max(x, 0))
        df_recon['PIS_CALC'] = (df_recon['BC_PIS_COFINS_CALC'] * 0.0165).round(2)
        df_recon['COFINS_CALC'] = (df_recon['BC_PIS_COFINS_CALC'] * 0.0760).round(2)

        df_recon['STATUS_PIS'] = np.where((df_recon['PIS_CALC'] - df_recon['PIS_SPED']).abs() <= tolerancia_valor, 'OK', 'DIVERGENTE')
        df_recon['STATUS_COFINS'] = np.where((df_recon['COFINS_CALC'] - df_recon['COFINS_SPED']).abs() <= tolerancia_valor, 'OK', 'DIVERGENTE')

        # --- APLICA REGRA: IGNORAR PIS/COFINS ---
        if ignorar_pis_cofins:
            df_recon['STATUS_PIS'] = 'N/A'
            df_recon['STATUS_COFINS'] = 'N/A'

        cond_energia_com = df_recon['TIPO_NOTA_SPED'].str.contains("Energia|Comunicação")
        df_recon.loc[cond_energia_com, ['STATUS_PIS', 'STATUS_COFINS']] = 'N/A'

        status_cols_to_na = [col for col in df_recon.columns if col.startswith('STATUS_')]
        df_recon.loc[df_recon['SITUACAO_NOTA'] != 'OK', status_cols_to_na] = 'N/A'

        df_recon['STATUS_GERAL'] = df_recon.apply(calcular_status_geral, axis=1)

        # --- APLICA REGRA: EXIGIR ACUMULADOR ---
        if exigir_acumulador:
            # Se ACUMULADOR for vazio, None ou 'REVISAR', marca STATUS_GERAL como REVISAR
            # (a menos que a nota falte no XML ou SPED, onde o status original prevalece)
            mask_falta_acumulador = (df_recon['ACUMULADOR'].isna()) | (df_recon['ACUMULADOR'] == '') | (df_recon['ACUMULADOR'] == 'REVISAR')
            mask_nota_existe = (df_recon['SITUACAO_NOTA'] == 'OK')
            df_recon.loc[mask_falta_acumulador & mask_nota_existe, 'STATUS_GERAL'] = 'REVISAR'

        # -------------------------------------------------------------------------
        # 4. Preparação dos Itens (C170)
        # -------------------------------------------------------------------------
        df_itens_final = df_itens_global.copy() if df_itens_global is not None else pd.DataFrame()
        if not df_itens_final.empty:

            def check_item_cfop(row: pd.Series) -> str:
                xml_cfop = str(row.get('CFOP', ''))
                sped_item_cfop = str(row.get('CFOP_SPED_ITEM', ''))
                if sped_item_cfop == 'N/A no SPED' or not sped_item_cfop: return 'REVISAR (Sem SPED)'
                if not xml_cfop: return 'REVISAR (Sem XML)'
                if xml_cfop == sped_item_cfop: return 'OK'
                expected_sped_cfop = xml_cfop
                if xml_cfop.startswith('5'): expected_sped_cfop = '1' + xml_cfop[1:]
                elif xml_cfop.startswith('6'): expected_sped_cfop = '2' + xml_cfop[1:]
                elif xml_cfop.startswith('7'): expected_sped_cfop = '3' + xml_cfop[1:]
                return 'OK' if sped_item_cfop == expected_sped_cfop else 'DIVERGENTE'

            if not df_sped_itens.empty:
                logging.info("Cruzando itens XML x SPED (C170) usando N_ITEM...")
                try:
                    df_itens_final['N_ITEM'] = pd.to_numeric(df_itens_final['N_ITEM'], errors='coerce').fillna(0).astype(int)
                    df_sped_itens['N_ITEM_SPED'] = pd.to_numeric(df_sped_itens['N_ITEM_SPED'], errors='coerce').fillna(0).astype(int)
                except Exception as e:
                    logging.warning(f"Falha ao converter N_ITEM/N_ITEM_SPED para inteiro: {e}")

                df_itens_final = pd.merge(df_itens_final, df_sped_itens,
                                        left_on=['CHV_NFE', 'N_ITEM'],
                                        right_on=['CHV_NFE', 'N_ITEM_SPED'],
                                        how='left')
                df_itens_final.drop(columns=['N_ITEM_SPED', 'COD_PROD_SPED'], inplace=True, errors='ignore')

                sped_c170_cols = ['CFOP_SPED_ITEM', 'CST_ICMS_SPED_ITEM', 'VL_OPR_SPED_ITEM', 'VL_BC_ICMS_SPED_ITEM',
                                    'VL_ICMS_SPED_ITEM', 'VL_BC_ICMS_ST_SPED_ITEM', 'VL_ICMS_ST_SPED_ITEM', 'VLR_IPI_SPED_ITEM']

                for col in sped_c170_cols:
                    if col not in df_itens_final.columns: df_itens_final[col] = np.nan

                df_itens_final['CFOP_SPED_ITEM'] = df_itens_final['CFOP_SPED_ITEM'].fillna('N/A no SPED')
                cols_to_fill_zero = [col for col in sped_c170_cols[1:]]
                df_itens_final[cols_to_fill_zero] = df_itens_final[cols_to_fill_zero].fillna(0.0)

            else:
                logging.warning("Itens SPED (C170) não encontrados. CFOP do item ficará 'N/A'.")
                df_itens_final['CFOP_SPED_ITEM'] = 'N/A no SPED'
                df_itens_final['VLR_IPI_SPED_ITEM'] = 0.0

            logging.info("Calculando status do CFOP a nível de item (NF-e)...")
            df_itens_final['STATUS_CFOP_ITEM'] = df_itens_final.apply(check_item_cfop, axis=1)

            if caminho_regras_detalhadas and not df_itens_final.empty:
                logging.info("Iniciando análise detalhada opcional (PROCV NF-e)...")
                try: df_itens_final = _executar_analise_detalhada_interna(df_itens_final, caminho_regras_detalhadas)
                except Exception as e: logging.warning(f"Falha na análise detalhada: {e}. Colunas das regras não adicionadas.")

            if not df_itens_final.empty and not df_recon.empty:
                # Merge com dados do cabeçalho
                recon_cols_to_merge = [
                    'CHV_NFE', 'NUM_NF', 'ACUMULADOR', 'SITUACAO_NOTA', 'STATUS_GERAL', 'TIPO_NOTA',
                    'STATUS_VALOR', 'VL_DOC_XML', 'VL_DOC_SPED', 'STATUS_CFOP',
                    'STATUS_ICMS', 'ICMS_SPED', 'ICMS_TOTAL_XML',
                    'STATUS_ICMS_ST', 'ICMS_ST_SPED', 'ICMS_ST_XML',
                    'STATUS_FCP_ST', 'FCP_ST_SPED', 'FCP_ST_XML',
                    'STATUS_IPI', 'IPI_TOTAL_XML',
                    'STATUS_ICMS_MONO', 'ICMS_MONO_SPED', 'ICMS_MONO_XML',
                    'STATUS_PIS', 'PIS_CALC', 'PIS_SPED',
                    'STATUS_COFINS', 'COFINS_CALC', 'COFINS_SPED', 'ICMS_SN_XML'
                ]

                cols_existentes_em_recon = [col for col in recon_cols_to_merge if col in df_recon.columns]
                cols_to_drop_from_itens = [
                    col for col in cols_existentes_em_recon
                    if col in df_itens_final.columns and
                    col not in ['CHV_NFE', 'BC_PIS_COFINS_CALC', 'PIS_CALC', 'COFINS_CALC']
                ]
                if cols_to_drop_from_itens:
                    df_itens_final = df_itens_final.drop(columns=cols_to_drop_from_itens)

                df_itens_final = pd.merge(
                    df_itens_final,
                    df_recon[cols_existentes_em_recon],
                    on='CHV_NFE',
                    how='left',
                    suffixes=('_ITEM', '_TOTAL_NOTA')
                )

                df_itens_final.rename(columns={
                    'BC_PIS_COFINS_CALC_ITEM': 'BC_PIS_COFINS_CALC', 'PIS_CALC_ITEM': 'PIS_CALC',
                    'COFINS_CALC_ITEM': 'COFINS_CALC', 'PIS_SPED_ITEM': 'PIS_SPED',
                    'COFINS_SPED_ITEM': 'COFINS_SPED', 'BC_PIS_COFINS_CALC_TOTAL_NOTA': 'BC_PIS_COFINS_CALC_TOTAL',
                    'PIS_CALC_TOTAL_NOTA': 'PIS_CALC_TOTAL', 'COFINS_CALC_TOTAL_NOTA': 'COFINS_CALC_TOTAL',
                    'PIS_SPED_TOTAL_NOTA': 'PIS_SPED_TOTAL', 'COFINS_SPED_TOTAL_NOTA': 'COFINS_SPED_TOTAL'
                }, inplace=True)

                # Preenchimento de Nulos após merge
                cols_preencher = [col for col in cols_existentes_em_recon if col != 'CHV_NFE']
                cols_preencher.extend(['PIS_CALC_TOTAL', 'COFINS_CALC_TOTAL', 'PIS_SPED_TOTAL', 'COFINS_SPED_TOTAL'])
                for col in cols_preencher:
                    if col in df_itens_final.columns:
                        if pd.api.types.is_numeric_dtype(df_itens_final[col]):
                            df_itens_final[col] = df_itens_final[col].fillna(0)
                        else:
                            df_itens_final[col] = df_itens_final[col].fillna('')

                logging.info("Calculando impostos proporcionais a nível de item (NF-e)...")
                df_itens_final['VLR_PROD'] = pd.to_numeric(df_itens_final['VLR_PROD'], errors='coerce').fillna(0)
                df_itens_final['VL_DOC_XML'] = pd.to_numeric(df_itens_final['VL_DOC_XML'], errors='coerce').fillna(0)
                df_itens_final['TAXA_PROPORCAO_ITEM'] = np.where(
                    df_itens_final['VL_DOC_XML'] > 0,
                    df_itens_final['VLR_PROD'] / df_itens_final['VL_DOC_XML'],
                    0
                )

                colunas_para_prorratear = [
                    'ICMS_SPED', 'ICMS_ST_SPED', 'ICMS_ST_XML', 'FCP_ST_SPED', 'FCP_ST_XML',
                    'ICMS_MONO_SPED', 'ICMS_MONO_XML', 'PIS_SPED_TOTAL', 'COFINS_SPED_TOTAL'
                ]

                for col in colunas_para_prorratear:
                    if col in df_itens_final.columns:
                        df_itens_final[col] = pd.to_numeric(df_itens_final[col], errors='coerce').fillna(0)
                        novo_nome_col = col.replace('_TOTAL', '')
                        df_itens_final[novo_nome_col] = (df_itens_final[col] * df_itens_final['TAXA_PROPORCAO_ITEM']).round(2)
                        if novo_nome_col != col:
                            df_itens_final.drop(columns=[col], inplace=True, errors='ignore')

                df_itens_final.drop(columns=['TAXA_PROPORCAO_ITEM', 'PIS_CALC_TOTAL', 'COFINS_CALC_TOTAL'], inplace=True, errors='ignore')

                if 'ICMS_TOTAL_XML' in df_itens_final.columns and 'VLR_ICMS' in df_itens_final.columns:
                    df_itens_final['ICMS_TOTAL_XML'] = df_itens_final['VLR_ICMS']
                if 'BC_PIS_COFINS_CALC' in df_itens_final.columns:
                    df_itens_final['PIS_CALC'] = (df_itens_final['BC_PIS_COFINS_CALC'] * 0.0165).round(2)
                    df_itens_final['COFINS_CALC'] = (df_itens_final['BC_PIS_COFINS_CALC'] * 0.0760).round(2)
                else:
                    df_itens_final['PIS_CALC'] = 0.0
                    df_itens_final['COFINS_CALC'] = 0.0

                if 'MVA ORIGINAL' in df_itens_final.columns:
                    df_itens_final['MVA ORIGINAL'] = pd.to_numeric(df_itens_final['MVA ORIGINAL'], errors='coerce').fillna(0)

                if 'VLR_ICMS' in df_itens_final.columns and 'VLR_ICMS_SN' in df_itens_final.columns and 'VLR_ICMS_MONO' in df_itens_final.columns:
                    df_itens_final['VLR_ICMS'] = pd.to_numeric(df_itens_final['VLR_ICMS'], errors='coerce').fillna(0)
                    df_itens_final['VLR_ICMS_SN'] = pd.to_numeric(df_itens_final['VLR_ICMS_SN'], errors='coerce').fillna(0)
                    df_itens_final['VLR_ICMS_MONO'] = pd.to_numeric(df_itens_final['VLR_ICMS_MONO'], errors='coerce').fillna(0)
                    df_itens_final['VLR_ICMS_TOTAL_ITEM'] = (df_itens_final['VLR_ICMS'] + df_itens_final['VLR_ICMS_SN'] + df_itens_final['VLR_ICMS_MONO']).round(2)
                else:
                    df_itens_final['VLR_ICMS_TOTAL_ITEM'] = df_itens_final['VLR_ICMS'] if 'VLR_ICMS' in df_itens_final.columns else 0.0

                if 'VL_DOC_XML' in df_itens_final.columns and 'VL_DOC_SPED' in df_itens_final.columns:
                    df_itens_final['DIF_VALOR_TOTAL'] = (df_itens_final['VL_DOC_XML'] - df_itens_final['VL_DOC_SPED']).round(2)
                else:
                    df_itens_final['DIF_VALOR_TOTAL'] = 0.0

        # -------------------------------------------------------------------------
        # 4. Preparação dos DataFrames para o Excel
        # -------------------------------------------------------------------------
        colunas_relatorio = [
            'STATUS_GERAL', 'SITUACAO_NOTA', 'CHV_NFE', 'NUM_NF', 'CNPJ_EMITENTE', 'ACUMULADOR',
            'TIPO_NOTA', 'STATUS_VALOR', 'VL_DOC_XML', 'VL_DOC_SPED',
            'STATUS_CFOP', 'CFOP_XML', 'CFOP_SPED', 'CEST_XML',
            'STATUS_ICMS', 'ICMS_TOTAL_XML', 'ICMS_SPED',
            'STATUS_ICMS_ST', 'ICMS_ST_XML', 'ICMS_ST_SPED',
            'STATUS_FCP_ST', 'FCP_ST_XML', 'FCP_ST_SPED',
            'STATUS_IPI', 'IPI_TOTAL_XML', 'IPI_SPED',
            'STATUS_ICMS_MONO', 'ICMS_MONO_XML', 'ICMS_MONO_SPED',
            'BC_PIS_COFINS_CALC', 'STATUS_PIS', 'PIS_CALC', 'PIS_SPED',
            'STATUS_COFINS', 'COFINS_CALC', 'COFINS_SPED',
        ]
        if not df_recon.empty:
            df_recon_relatorio = df_recon[[col for col in colunas_relatorio if col in df_recon.columns]]
            if 'STATUS_GERAL' in df_recon.columns:
                total_problemas = df_recon['STATUS_GERAL'].apply(lambda x: isinstance(x, str) and x != 'OK' and x != 'N/A').sum()

        if not df_itens_final.empty:
            colunas_itens_xml = [
                'STATUS_GERAL', 'SITUACAO_NOTA', 'TIPO_NOTA', 'CHV_NFE', 'NUM_NF', 'CNPJ_EMITENTE', 'ACUMULADOR', 'N_ITEM',
                'TIPO_DESTINATARIO',
                'COD_PROD', 'DESC_PROD', 'NCM', 'CEST',
                'STATUS_CFOP_ITEM', 'CFOP', 'CFOP_SPED_ITEM', 'CST_ICMS_SPED_ITEM',
                'STATUS_VALOR', 'VL_DOC_XML', 'VL_DOC_SPED', 'DIF_VALOR_TOTAL', 'cBenef',
                'QTD', 'UNID', 'VLR_UNIT', 'VLR_PROD', 'DESPESA_XML',
                'VLR_ICMS_TOTAL_ITEM', 'VLR_BC_ICMS_XML', 'pICMS_XML',
                'VLR_IPI', 'VLR_ICMS_MONO', 'BC_PIS_COFINS_CALC',
                'VL_OPR_SPED_ITEM', 'VL_BC_ICMS_SPED_ITEM', 'VL_ICMS_SPED_ITEM', 'VL_BC_ICMS_ST_SPED_ITEM', 'VL_ICMS_ST_SPED_ITEM',
                'STATUS_ICMS', 'ICMS_SPED',
                'STATUS_ICMS_ST', 'ICMS_ST_XML', 'ICMS_ST_SPED',
                'STATUS_FCP_ST', 'FCP_ST_XML', 'FCP_ST_SPED',
                'STATUS_IPI', 'VLR_IPI_SPED_ITEM',
                'STATUS_PIS', 'PIS_CALC', 'PIS_SPED', 'STATUS_COFINS', 'COFINS_CALC', 'COFINS_SPED',
                'PRODUTO', 'ST', 'REGIME_PIS_COFINS', 'MVA ORIGINAL'
            ]

            colunas_itens_existentes = [col for col in colunas_itens_xml if col in df_itens_final.columns]
            df_itens_aba = df_itens_final[colunas_itens_existentes].copy()
            df_itens_aba.rename(columns={'VLR_IPI_SPED_ITEM': 'IPI_SPED (Item C170)'}, inplace=True)

        if not df_itens_final.empty and caminho_regras_detalhadas:
            colunas_aliquota_xml = [
                'NUM_NF', 'TIPO_NOTA', 'COD_PROD', 'DESC_PROD', 'NCM', 'CEST', 'cBenef',
                'CFOP', 'CFOP_SPED_ITEM', 'VLR_TOTAL_NF', 'VLR_PROD',
                'CST_ICMS_XML', 'VLR_BC_ICMS_XML', 'VLR_ICMS', 'VLR_ICMS_ST', 'pICMS_XML'
            ]
            regras_ncm_cols = ['PRODUTO', 'ST', 'REGIME_PIS_COFINS', 'MVA ORIGINAL']
            for col in regras_ncm_cols:
                if col in df_itens_final.columns: colunas_aliquota_xml.append(col)

            colunas_aliq_existentes = [col for col in colunas_aliquota_xml if col in df_itens_final.columns]
            df_aliquota_aba = df_itens_final[colunas_aliq_existentes].copy()

            rename_map = {
                'VLR_ICMS_TOTAL_ITEM': 'VLR_ICMS_SOMA_SN', 'VLR_ICMS': 'VLR_ICMS',
                'pICMS_XML': 'Aliquota ICMS (XML)', 'PRODUTO': 'Produto (Regra)',
                'REGIME_PIS_COFINS': 'Regime PIS/COFINS (Regra)', 'MVA ORIGINAL': 'MVA Original (Regra)'
            }
            actual_rename_map = {k: v for k, v in rename_map.items() if k in df_aliquota_aba.columns}
            df_aliquota_aba.rename(columns=actual_rename_map, inplace=True)
            df_aliquota_aba.drop_duplicates(inplace=True)

        # 5. Totalizadores e Base DIFAL
        logging.info("Calculando totalizadores combinados (NF-e, CT-e, Energia, Com)...")
        df_totalizadores_cst = _calcular_totalizadores_cfop_cst(df_sped_analitico_combinado)

        if not df_totalizadores_cst.empty:
            cfop_str = df_totalizadores_cst['CFOP (SPED)'].astype(str)
            df_totalizadores_entrada = df_totalizadores_cst[cfop_str.str.startswith(('1', '2', '3'))].copy()
            df_totalizadores_saida = df_totalizadores_cst[cfop_str.str.startswith(('5', '6', '7'))].copy()
        else:
            df_totalizadores_entrada = pd.DataFrame()
            df_totalizadores_saida = pd.DataFrame()

        df_base_difal_por_cfop = pd.DataFrame()
        if not df_chaves_difal.empty and not df_sped_analitico_combinado.empty:
            logging.info("Calculando Base de Cálculo para abatimento de DIFAL (C101)...")
            df_analitico_difal = pd.merge(df_sped_analitico_combinado, df_chaves_difal, on='CHV_NFE', how='inner')
            if not df_analitico_difal.empty:
                df_base_difal_por_cfop = df_analitico_difal.groupby('CFOP_SPED_ITEM')['VL_BC_ICMS_SPED_ITEM'].sum().reset_index()
                df_base_difal_por_cfop.rename(columns={'CFOP_SPED_ITEM': 'CFOP', 'VL_BC_ICMS_SPED_ITEM': 'VALOR_BASE_DIFAL'}, inplace=True)

        # -------------------------------------------------------------------------
        # 6. Conciliação CT-e (Atualizado com Novas Colunas)
        # -------------------------------------------------------------------------
        logging.info("Iniciando conciliação de CT-e (XML vs SPED D190)...")
        df_report_cte = df_sped_cte_d190.copy()

        if 'CHV_CTE' in df_report_cte.columns:
            df_report_cte.rename(columns={'CHV_CTE': 'NUM_CTE_SPED'}, inplace=True)
        if not df_report_cte.empty and 'NUM_CTE_SPED' in df_report_cte.columns:
            df_report_cte['NUM_CTE_SPED'] = df_report_cte['NUM_CTE_SPED'].astype(str).str.strip()
        if not df_xml_cte_totais.empty and 'NUM_CTE_XML' in df_xml_cte_totais.columns:
            df_xml_cte_totais['NUM_CTE_XML'] = df_xml_cte_totais['NUM_CTE_XML'].astype(str).str.strip()

        if not df_xml_cte_totais.empty and not df_report_cte.empty and 'NUM_CTE_SPED' in df_report_cte.columns:
            df_sped_cte_agg = df_report_cte.groupby('NUM_CTE_SPED').agg(
                VL_OPR_SPED_SUM=pd.NamedAgg(column='VL_OPR_SPED_D190', aggfunc='sum'),
                VL_BC_ICMS_SPED_SUM=pd.NamedAgg(column='VL_BC_ICMS_SPED_D190', aggfunc='sum'),
                VL_ICMS_SPED_SUM=pd.NamedAgg(column='VL_ICMS_SPED_D190', aggfunc='sum'),
                CFOP_SPED_LIST=pd.NamedAgg(column='CFOP_SPED_D190', aggfunc=lambda x: sorted(list(x.unique()))),
            ).reset_index()

            # --- DEFINIÇÃO DAS COLUNAS DO XML PARA MERGE (Expandido) ---
            xml_cols_original = [
                'NUM_CTE_XML', 'CHV_CTE', 'CFOP_XML', 'CST_XML',
                'VL_OPR_XML', 'VL_BC_ICMS_XML', 'VL_ICMS_XML',

                # NOVAS COLUNAS DO XML PARSER
                'VL_TOTAL_CTE_XML', 'CNPJ_TRANSPORTADOR', 'IE_TRANSPORTADOR', 'UF_EMITENTE_CTE',
                'REMETENTE_NOME', 'DESTINATARIO_NOME', 'TOMADOR_CNPJ', 'TOMADOR_NOME',
                'MUN_ORIGEM', 'MUN_DESTINO', 'ALIQ_ICMS_XML', 'ITEM_PREDOMINANTE'
            ]

            xml_cols_to_merge = [col for col in xml_cols_original if col in df_xml_cte_totais.columns]

            df_cte_merge = pd.merge(
                df_xml_cte_totais[xml_cols_to_merge],
                df_sped_cte_agg,
                left_on='NUM_CTE_XML',
                right_on='NUM_CTE_SPED',
                how='outer',
                suffixes=('_XML', '_SPED_AGG'),
                indicator=True
            )

            df_cte_merge['SITUACAO_CTE'] = np.select(
                [df_cte_merge['_merge'] == 'left_only', df_cte_merge['_merge'] == 'right_only'],
                ['FALTA NO SPED', 'FALTA XML'],
                default='OK'
            )

            # Preenche zeros apenas nas colunas de SOMA do SPED
            numeric_cols_cte_agg = ['VL_OPR_SPED_SUM', 'VL_BC_ICMS_SPED_SUM', 'VL_ICMS_SPED_SUM']
            for col in numeric_cols_cte_agg:
                if col in df_cte_merge.columns: df_cte_merge[col] = df_cte_merge[col].fillna(0.0)

            # Status Valor (Compara Total XML vs Total Operação SPED ou Valor Operação XML vs SPED)
            # Preferimos VL_TOTAL_CTE_XML se existir, senão VL_OPR_XML
            col_valor_xml = 'VL_TOTAL_CTE_XML' if 'VL_TOTAL_CTE_XML' in df_cte_merge.columns else 'VL_OPR_XML'

            # Garante que col_valor_xml não é NaN para a comparação
            df_cte_merge[col_valor_xml] = df_cte_merge[col_valor_xml].fillna(0.0)

            df_cte_merge['STATUS_VALOR'] = np.where(
                (df_cte_merge[col_valor_xml] - df_cte_merge['VL_OPR_SPED_SUM']).abs() <= tolerancia_valor, 'OK', 'DIVERGENTE'
            )
            df_cte_merge['STATUS_BC_ICMS'] = np.where(
                (df_cte_merge['VL_BC_ICMS_XML'].fillna(0.0) - df_cte_merge['VL_BC_ICMS_SPED_SUM']).abs() <= tolerancia_valor, 'OK', 'DIVERGENTE'
            )
            df_cte_merge['STATUS_ICMS'] = np.where(
                (df_cte_merge['VL_ICMS_XML'].fillna(0.0) - df_cte_merge['VL_ICMS_SPED_SUM']).abs() <= tolerancia_valor, 'OK', 'DIVERGENTE'
            )

            df_cte_merge['CFOP_SPED_AGG'] = df_cte_merge['CFOP_SPED_LIST'].apply(lambda x: '/'.join(x) if isinstance(x, list) else '')
            df_cte_merge['STATUS_CFOP'] = np.where(
                df_cte_merge['CFOP_XML'] == df_cte_merge['CFOP_SPED_AGG'], 'OK', 'DIVERGENTE'
            )
            df_cte_merge.loc[df_cte_merge['CFOP_SPED_AGG'].str.contains('/'), 'STATUS_CFOP'] = 'REVISAR'
            df_cte_merge.loc[df_cte_merge['SITUACAO_CTE'] != 'OK', ['STATUS_VALOR', 'STATUS_BC_ICMS', 'STATUS_ICMS', 'STATUS_CFOP']] = 'N/A'

            # --- LISTA FINAL DE COLUNAS PARA O RELATÓRIO DE CT-e ---
            status_cols_to_add = [
                'NUM_CTE_SPED', 'NUM_CTE_XML', 'CHV_CTE',
                'SITUACAO_CTE', 'STATUS_VALOR', 'STATUS_BC_ICMS', 'STATUS_ICMS', 'STATUS_CFOP',
                'VL_OPR_XML', 'VL_BC_ICMS_XML', 'VL_ICMS_XML', 'CFOP_XML', 'CST_XML',
                # NOVAS COLUNAS
                'VL_TOTAL_CTE_XML', 'CNPJ_TRANSPORTADOR', 'IE_TRANSPORTADOR', 'UF_EMITENTE_CTE',
                'REMETENTE_NOME', 'DESTINATARIO_NOME', 'TOMADOR_CNPJ', 'TOMADOR_NOME',
                'MUN_ORIGEM', 'MUN_DESTINO', 'ALIQ_ICMS_XML', 'ITEM_PREDOMINANTE'
            ]
            status_cols_to_add = [col for col in status_cols_to_add if col in df_cte_merge.columns]

            logging.info("Mapeando status da conciliação de volta para os registros D190 originais...")
            df_report_cte = pd.merge(
                df_report_cte,
                df_cte_merge[status_cols_to_add],
                on='NUM_CTE_SPED',
                how='left'
            )
            df_report_cte.rename(columns={'NUM_CTE_SPED': 'CHV_CTE'}, inplace=True)

            df_report_cte['SITUACAO_CTE'] = df_report_cte['SITUACAO_CTE'].fillna('FALTA XML')
            cols_status = ['STATUS_VALOR', 'STATUS_BC_ICMS', 'STATUS_ICMS', 'STATUS_CFOP']
            df_report_cte[cols_status] = df_report_cte[cols_status].fillna('N/A')

            # Preenche NaN nas colunas novas com vazios ou zeros
            cols_numericas_cte = ['VL_OPR_XML', 'VL_BC_ICMS_XML', 'VL_ICMS_XML', 'VL_TOTAL_CTE_XML', 'ALIQ_ICMS_XML']
            for col in cols_numericas_cte:
                if col in df_report_cte.columns: df_report_cte[col] = df_report_cte[col].fillna(0.0)

            cols_texto_cte = ['CHV_CTE_XML', 'CFOP_XML', 'CST_XML', 'CNPJ_TRANSPORTADOR', 'IE_TRANSPORTADOR', 'UF_EMITENTE_CTE', 'REMETENTE_NOME', 'DESTINATARIO_NOME', 'TOMADOR_CNPJ', 'TOMADOR_NOME', 'MUN_ORIGEM', 'MUN_DESTINO', 'ITEM_PREDOMINANTE']
            for col in cols_texto_cte:
                if col in df_report_cte.columns: df_report_cte[col] = df_report_cte[col].fillna('')

            if 'CHV_CTE_y' in df_report_cte.columns: df_report_cte.rename(columns={'CHV_CTE_y': 'CHV_CTE_XML'}, inplace=True)

        elif df_report_cte.empty:
            pass
        else:
            if 'CHV_CTE' in df_report_cte.columns: df_report_cte.rename(columns={'CHV_CTE': 'NUM_CTE_SPED'}, inplace=True)
            df_report_cte['SITUACAO_CTE'] = 'FALTA XML'
            df_report_cte.rename(columns={'NUM_CTE_SPED': 'CHV_CTE'}, inplace=True)

        df_sped_cte_d190_final = df_report_cte

        # 7. Geração do Arquivo Excel
        caminho_saida = caminho_sped.parent / f'Relatorio_Conciliacao_Fiscal_{time.strftime("%Y%m%d_%H%M%S")}.xlsx'
        logging.info(f"Gerando relatório em Excel: {caminho_saida}")
        if status_callback: status_callback("Gerando relatório Excel...")

        gerar_relatorio_excel(
            caminho_saida,
            df_recon_relatorio,
            df_itens_aba,
            df_aliquota_aba,
            df_totalizadores_entrada,
            df_totalizadores_saida,
            df_sped_cte_d190_final
        )

        # 8. Preenchimento do Template de Apuração
        if template_apuracao_path:
            try:
                logging.info(f"Iniciando preenchimento do template de apuração (Setor: {tipo_setor})...")
                if status_callback: status_callback("Preenchendo template de apuração...")

                if tipo_setor == 'Moveleiro':
                    try:
                        from .apuracao_moveleiro import preencher_template_moveleiro
                        preencher_template_moveleiro(
                            template_apuracao_path,
                            df_totalizadores_entrada,
                            df_totalizadores_saida,
                            df_base_difal_por_cfop
                        )
                        logging.info("Template Moveleiro preenchido com sucesso.")
                    except ImportError:
                        logging.error("Módulo 'apuracao_moveleiro' não encontrado.")
                        if error_callback: error_callback("Módulo 'apuracao_moveleiro' não encontrado.")

                elif tipo_setor == 'E-commerce':
                    try:
                        from .apuracao_ecommerce import preencher_template_ecommerce
                        preencher_template_ecommerce(
                            template_apuracao_path,
                            df_totalizadores_entrada,
                            df_totalizadores_saida
                        )
                        logging.info("Template E-commerce preenchido.")
                    except ImportError:
                        logging.error("Módulo 'apuracao_ecommerce' não encontrado.")
                        if error_callback: error_callback("Módulo 'apuracao_ecommerce' não encontrado.")

                else:
                    # Padrão (Comercio)
                    preencher_template_apuracao(
                        template_apuracao_path,
                        df_totalizadores_entrada,
                        df_totalizadores_saida
                    )
                    logging.info("Preenchimento do template de apuração (Padrão/Comercio) concluído.")

            except Exception as e:
                logging.error(f"Falha ao preencher o template de apuração: {e}", exc_info=True)
                if error_callback: error_callback(f"Falha ao preencher template: {e}")

        logging.info("Relatório Excel gerado com sucesso.")
        if done_callback: done_callback(caminho_saida, total_problemas)

    except Exception as e:
        logging.exception("Ocorreu uma falha crítica na análise.")
        if error_callback: error_callback(f"Erro Crítico: {e}")
