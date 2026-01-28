import logging
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, Tuple, Optional, Any

# Importa as constantes da pasta local
from .constants import MAPA_CST_UNIFICADO

def get_acumulador(row: pd.Series, regras_map: Dict[Tuple[str, str], str]) -> str:

    cnpj_raw = row.get('CNPJ_EMITENTE', '')
    if not cnpj_raw or pd.isna(cnpj_raw): return ''
    cnpj = ''.join(filter(str.isdigit, str(cnpj_raw)))
    if not cnpj: return ''
    cfops_str = str(row.get('CFOP_SPED', '')) if pd.notna(row.get('CFOP_SPED')) and row.get('CFOP_SPED') else str(row.get('CFOP_XML', ''))
    cfops = set(filter(None, cfops_str.split('/')))
    if not cfops: return ''
    found_acumuladores: set[str] = set()
    for cfop in cfops:
        acumulador = regras_map.get((cnpj, str(cfop)))
        if acumulador: found_acumuladores.add(acumulador)
    if not found_acumuladores: return ''
    elif 'REVISAR' in found_acumuladores: return 'REVISAR'
    elif len(found_acumuladores) > 1:
        logging.warning(f"Múltiplos acumuladores ({found_acumuladores}) para CNPJ {cnpj}, CFOPs {cfops} na nota {row.get('CHV_NFE', '')}. Marcado REVISAR.")
        return 'REVISAR'
    else: return found_acumuladores.pop()


def check_cfop_status(row: pd.Series) -> str:

    xml_cfops_str, sped_cfops_str = str(row.get('CFOP_XML', '')), str(row.get('CFOP_SPED', ''))

    if not xml_cfops_str and sped_cfops_str:
        return 'N/A'

    xml_cfops = set(filter(None, xml_cfops_str.split('/')))
    sped_cfops = set(filter(None, sped_cfops_str.split('/')))
    if not xml_cfops and not sped_cfops: return 'N/A'
    if not xml_cfops or not sped_cfops: return 'DIVERGENTE'

    # --- INÍCIO DA MODIFICAÇÃO ---
    if len(xml_cfops) == 1 and len(sped_cfops) == 1:
        xml_cfop = list(xml_cfops)[0]
        sped_cfop = list(sped_cfops)[0]

        # 1. Checa se são idênticos (Ex: Saída 5102 vs SPED 5102)
        if xml_cfop == sped_cfop:
            return 'OK'

        # 2. Se não for idêntico, checa a transformação (Ex: Saída 5102 vs SPED 1102)
        expected_sped_cfop = xml_cfop
        if xml_cfop.startswith('5'): expected_sped_cfop = '1' + xml_cfop[1:]
        elif xml_cfop.startswith('6'): expected_sped_cfop = '2' + xml_cfop[1:]
        elif xml_cfop.startswith('7'): expected_sped_cfop = '3' + xml_cfop[1:]

        return 'OK' if sped_cfop == expected_sped_cfop else 'DIVERGENTE'

    # Lógica para múltiplos CFOPs
    # 1. Checa se são idênticos
    if xml_cfops == sped_cfops:
        return 'OK (Múltiplos)'

    # 2. Se não, checa a transformação
    expected_sped_equivalents = set()
    for cfop in xml_cfops:
        if cfop.startswith('5'): expected_sped_equivalents.add('1' + cfop[1:])
        elif cfop.startswith('6'): expected_sped_equivalents.add('2' + cfop[1:])
        elif cfop.startswith('7'): expected_sped_equivalents.add('3' + cfop[1:])
        else: expected_sped_equivalents.add(cfop)

    if sped_cfops == expected_sped_equivalents:
        return 'OK (Múltiplos)'
    else:
        return 'REVISAR (Múltiplos)'
    # --- FIM DA MODIFICAÇÃO ---


def calcular_status_geral(row: pd.Series) -> str:

    if row['SITUACAO_NOTA'] in ['FALTA XML', 'FALTA NO SPED']: return row['SITUACAO_NOTA']
    status_cols = [col for col in row.index if col.startswith('STATUS_')]
    all_status_values = row[status_cols].values
    if 'DIVERGENTE' in all_status_values: return 'DIVERGENTE'
    if 'REVISAR' in all_status_values or 'REVISAR (Múltiplos)' in all_status_values or row['SITUACAO_NOTA'] == 'SEM CNPJ NO XML': return 'REVISAR'
    return 'OK'


def _executar_analise_detalhada_interna(df_itens_xml: pd.DataFrame, arquivo_excel_regras: Path) -> pd.DataFrame:

    logging.info(f"Iniciando cruzamento detalhado com: {arquivo_excel_regras.name}")
    df_regras_detalhadas: Optional[pd.DataFrame] = None
    try:
        if not arquivo_excel_regras.exists(): raise FileNotFoundError(f"Arquivo de regras detalhadas não encontrado: {arquivo_excel_regras}")
        try: df_regras_detalhadas = pd.read_excel(arquivo_excel_regras, sheet_name='Planilha1')
        except Exception:
            try: df_regras_detalhadas = pd.read_excel(arquivo_excel_regras); logging.warning(f"'Planilha1' não encontrada. Lendo a primeira aba.")
            except Exception as e_inner: raise ValueError(f"Erro ao ler o arquivo de regras Excel: {e_inner}")
        if df_regras_detalhadas is None or df_regras_detalhadas.empty: raise ValueError("Arquivo de regras detalhadas vazio ou inválido.")
        logging.info(f"[DEBUG] Colunas originais lidas das regras: {df_regras_detalhadas.columns.tolist()}")
        df_regras_detalhadas.columns = df_regras_detalhadas.columns.str.strip()
        logging.info(f"[DEBUG] Colunas das regras após limpeza (.strip()): {df_regras_detalhadas.columns.tolist()}")
        if 'NCM' not in df_regras_detalhadas.columns: raise ValueError("Coluna 'NCM' não encontrada no arquivo de regras detalhadas.")
        df_itens_xml['NCM'] = df_itens_xml['NCM'].astype(str).str.strip()
        df_regras_detalhadas['NCM'] = df_regras_detalhadas['NCM'].astype(str).str.strip()
        if df_regras_detalhadas.duplicated(subset=['NCM']).any():
            logging.warning("NCMs duplicados encontrados nas regras. Mantendo apenas a primeira ocorrência.")
            df_regras_detalhadas.drop_duplicates(subset=['NCM'], keep='first', inplace=True)
        colunas_regras_para_usar = ['NCM', 'PRODUTO', 'ST', 'CST PIS/COFINS', 'MVA ORIGINAL']
        colunas_regras_existentes = [col for col in colunas_regras_para_usar if col in df_regras_detalhadas.columns]
        logging.info(f"[DEBUG] Colunas selecionadas das regras para o merge: {colunas_regras_existentes}")
        df_analise = pd.merge(df_itens_xml, df_regras_detalhadas[colunas_regras_existentes], on='NCM', how='left', suffixes=('_XML', '_REGRA'))
        logging.info(f"[DEBUG] Colunas no dataframe após o merge: {df_analise.columns.tolist()}")
        coluna_regras_pis_original = 'CST PIS/COFINS'
        if coluna_regras_pis_original in df_analise.columns:
            logging.info(f"Traduzindo a coluna '{coluna_regras_pis_original}'...")
            df_analise[coluna_regras_pis_original] = df_analise[coluna_regras_pis_original].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            conditions = [df_analise[coluna_regras_pis_original] == '4', df_analise[coluna_regras_pis_original] == '6', df_analise[coluna_regras_pis_original] == '-']
            choices = ['MONOFÁSICO', 'ALÍQUOTA 0', 'NORMAL']
            df_analise['REGIME_PIS_COFINS'] = np.select(conditions, choices, default='NORMAL')
            df_analise.loc[df_analise[coluna_regras_pis_original].isna() | (df_analise[coluna_regras_pis_original] == ''), 'REGIME_PIS_COFINS'] = 'N/A'
            df_analise['REGIME_PIS_COFINS'] = df_analise['REGIME_PIS_COFINS'].fillna('N/A')
        else:
            logging.warning(f"Coluna '{coluna_regras_pis_original}' não encontrada nas regras. Tradução PIS/COFINS ignorada."); df_analise['REGIME_PIS_COFINS'] = 'N/A'
        logging.info("Merge detalhado (PROCV por NCM) aplicado.")
        return df_analise
    except Exception as e:
        logging.error(f"Erro durante a análise detalhada interna: {e}")
        return df_itens_xml


def _calcular_totalizadores_cfop_cst(df_analitico_combinado: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula o totalizador CONSOLIDADO por CFOP (SPED), CST (SPED) e Alíquota (SPED),
    e traduz o CST para sua descrição legal.
    FONTE: SPED C190, D190, C590, D590 combinados.
    """

    if df_analitico_combinado is None or df_analitico_combinado.empty:
        logging.warning("DataFrame analítico (C190/D190/C590/D590) vazio. Não é possível calcular totalizadores.")
        return pd.DataFrame()

    df_calc = df_analitico_combinado.copy()

    sped_value_cols = ['VL_OPR_SPED_ITEM', 'VL_BC_ICMS_SPED_ITEM', 'VL_ICMS_SPED_ITEM', 'VL_BC_ICMS_ST_SPED_ITEM', 'VL_ICMS_ST_SPED_ITEM', 'VLR_IPI_SPED_ITEM']
    for col in sped_value_cols:
        if col not in df_calc.columns:
            df_calc[col] = 0.0
        else:
            df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').fillna(0.0)

    if 'CFOP_SPED_ITEM' not in df_calc.columns: df_calc['CFOP_SPED_ITEM'] = 'N/A'
    if 'CST_ICMS_SPED_ITEM' not in df_calc.columns: df_calc['CST_ICMS_SPED_ITEM'] = 'N/A'
    if 'ALIQ_ICMS_SPED_ITEM' not in df_calc.columns: df_calc['ALIQ_ICMS_SPED_ITEM'] = 0.0
    else: df_calc['ALIQ_ICMS_SPED_ITEM'] = pd.to_numeric(df_calc['ALIQ_ICMS_SPED_ITEM'], errors='coerce').fillna(0.0)

    df_calc['CFOP_SPED_ITEM'] = df_calc['CFOP_SPED_ITEM'].astype(str).str.strip()
    df_calc['CST_ICMS_SPED_ITEM'] = df_calc['CST_ICMS_SPED_ITEM'].astype(str).str.strip()

    group_cols = ['CFOP_SPED_ITEM', 'CST_ICMS_SPED_ITEM', 'ALIQ_ICMS_SPED_ITEM']

    df_totalizadores = df_calc.groupby(group_cols).agg(
        QTD_DOCUMENTOS=('CHV_NFE', pd.Series.nunique),
        Total_Operacao=('VL_OPR_SPED_ITEM', 'sum'),
        Base_de_Calculo_ICMS=('VL_BC_ICMS_SPED_ITEM', 'sum'),
        Total_ICMS=('VL_ICMS_SPED_ITEM', 'sum'),
        Base_de_Calculo_ICMS_ST=('VL_BC_ICMS_ST_SPED_ITEM', 'sum'),
        Total_ICMS_ST=('VL_ICMS_ST_SPED_ITEM', 'sum'),
        Total_IPI=('VLR_IPI_SPED_ITEM', 'sum'),
    ).reset_index()

    nova_base_calculo = (
        df_totalizadores['Total_Operacao'] -
        df_totalizadores['Total_IPI'] -
        df_totalizadores['Total_ICMS_ST']
    )

    df_totalizadores['Alíquota ICMS'] = np.where(
        nova_base_calculo > 0,
        (df_totalizadores['Total_ICMS'] / nova_base_calculo) * 100,
        0.0
    ).round(2)

    df_final = df_totalizadores.rename(columns={
        'CFOP_SPED_ITEM': 'CFOP (SPED)',
        'CST_ICMS_SPED_ITEM': 'CST (SPED)',
        'ALIQ_ICMS_SPED_ITEM': 'Alíquota (SPED)',
        'QTD_DOCUMENTOS': 'QTD Documentos',
        'Total_Operacao': 'Total Operação',
        'Base_de_Calculo_ICMS': 'Base de Cálculo ICMS',
        'Total_ICMS': 'Total ICMS',
        'Base_de_Calculo_ICMS_ST': 'Base de Cálculo ICMS ST',
        'Total_ICMS_ST': 'Total ICMS ST',
        'Total_IPI': 'Total IPI'
    })

    df_final['CST (SPED)'] = df_final['CST (SPED)'].astype(str).str.strip()
    df_final['Descricao CST'] = df_final['CST (SPED)'].map(MAPA_CST_UNIFICADO).fillna(df_final['CST (SPED)'])

    colunas_ordenadas = [
        'CFOP (SPED)',
        'CST (SPED)',
        'Descricao CST',
        'Alíquota (SPED)',
        'Alíquota ICMS',
        'Total Operação', 'Base de Cálculo ICMS',
        'Total ICMS', 'Base de Cálculo ICMS ST', 'Total ICMS ST', 'Total IPI',
        'QTD Documentos'
    ]
    df_final = df_final[[col for col in colunas_ordenadas if col in df_final.columns]]

    cols_to_round = ['Total Operação', 'Base de Cálculo ICMS', 'Total ICMS', 'Base de Cálculo ICMS ST', 'Total ICMS ST', 'Total IPI']
    for col in cols_to_round:
        if col in df_final.columns: df_final[col] = df_final[col].round(2)

    return df_final.sort_values(by=['CFOP (SPED)', 'CST (SPED)', 'Alíquota (SPED)'])