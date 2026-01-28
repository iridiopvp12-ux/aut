import logging
import pandas as pd
from pathlib import Path

def ler_regras_acumuladores(caminho_regras: Path) -> pd.DataFrame:
    logging.info('Lendo arquivo de regras de acumuladores...')
    try:
        dtype_map = {'CNPJ_CPF': str, 'CFOP': str, 'ACUMULADOR': str}
        suffix = caminho_regras.suffix.lower()
        if suffix == '.csv':

            try:
                df = pd.read_csv(caminho_regras, dtype=dtype_map, sep=None, engine='python', encoding='utf-8-sig')
                if len(df.columns) < 3: raise ValueError("Poucas colunas detectadas.")
            except Exception:
                logging.warning("Detecção automática falhou. Tentando ',' e utf-8...")
                try:
                    df = pd.read_csv(caminho_regras, dtype=dtype_map, sep=',', encoding='utf-8-sig')
                except Exception:
                    logging.warning("Falha com ','. Tentando ';' e latin-1...")
                    df = pd.read_csv(caminho_regras, dtype=dtype_map, sep=';', encoding='latin-1')
        elif suffix in ['.xlsx', '.xls']:

            try: df = pd.read_excel(caminho_regras, sheet_name='CNPJ/CFOP (SPED)', dtype=dtype_map)
            except Exception: logging.warning("'CNPJ/CFOP (SPED)' sheet not found. Reading the first sheet."); df = pd.read_excel(caminho_regras, dtype=dtype_map)
        else: raise ValueError("Formato de arquivo de regras não suportado. Use .csv ou .xlsx/.xls")

        required_cols = ['CNPJ_CPF', 'CFOP', 'ACUMULADOR']

        df.columns = df.columns.str.strip()
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols: raise ValueError(f"O arquivo de regras não contém as colunas obrigatórias: {', '.join(missing_cols)}. Colunas encontradas: {df.columns.tolist()}")

        for col in required_cols: df[col] = df[col].astype(str).str.strip()

        def _normalize_cnpj_cpf(text: str) -> str:
            if not isinstance(text, str): return ''
            return ''.join(filter(str.isdigit, text))
        df['CNPJ_CPF'] = df['CNPJ_CPF'].apply(_normalize_cnpj_cpf)

        df.dropna(subset=required_cols, inplace=True)

        df['ACUMULADOR'] = df['ACUMULADOR'].apply(
            lambda x: str(int(float(x))) if isinstance(x, str) and x.replace('.', '', 1).isdigit() and float(x) == int(float(x)) else str(x)
        )
        df['ACUMULADOR'] = df['ACUMULADOR'].str.replace(r'\.0$', '', regex=True)

        duplicates = df.duplicated(subset=['CNPJ_CPF', 'CFOP'], keep=False)
        if duplicates.any():
            logging.warning("Regras duplicadas (mesmo CNPJ_CPF e CFOP) encontradas. Marcando como 'REVISAR'.")
            df.loc[duplicates, 'ACUMULADOR'] = 'REVISAR'
        df.drop_duplicates(subset=['CNPJ_CPF', 'CFOP'], keep='first', inplace=True)

        logging.info(f"Encontradas {len(df)} regras de acumuladores únicas.")
        return df
    except FileNotFoundError: raise Exception(f"Arquivo de regras não encontrado em: {caminho_regras}")
    except ImportError:
        msg = "A biblioteca 'openpyxl' (para .xlsx) ou 'xlrd' (para .xls) é necessária.\n\nInstale com:\npip install openpyxl xlrd"
        logging.error(msg)
        raise ImportError(msg)
    except Exception as e: logging.error(f"Erro ao processar o arquivo de regras: {e}"); raise