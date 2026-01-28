MAPA_FINNFE = {
    '1': 'Normal',
    '2': 'Complementar',
    '3': 'Ajuste',
    '4': 'Devolução'
}

# MAPA DE TRADUÇÃO DAS REGRAS DO CST
# Tabela B - Tributação do ICMS (Regime Normal - Lucro Real/Presumido)
MAPA_TRIBUTACAO_ICMS = {
    '00': '00 - Tributada integralmente',
    '10': '10 - Tributada e com cobrança do ICMS por ST',
    '20': '20 - Com redução de base de cálculo',
    '30': '30 - Isenta/Não tributada e com cobrança do ICMS por ST',
    '40': '40 - Isenta',
    '41': '41 - Não tributada',
    '50': '50 - Suspensão',
    '51': '51 - Diferimento',
    '60': '60 - ICMS cobrado anteriormente por ST',
    '70': '70 - Com redução de BC e cobrança do ICMS por ST',
    '90': '90 - Outras'
}
# Tabela B - CSOSN (Simples Nacional)
MAPA_CSOSN = {
    '101': '101 - Tributada pelo Simples Nacional com permissão de crédito',
    '102': '102 - Tributada pelo Simples Nacional sem permissão de crédito',
    '103': '103 - Isenção do ICMS no Simples Nacional (faixa de receita)',
    '201': '201 - Tributada pelo Simples Nacional com permissão de crédito e com ST',
    '202': '202 - Tributada pelo Simples Nacional sem permissão de crédito e com ST',
    '203': '203 - Isenção do ICMS no Simples Nacional (faixa de receita) e com ST',
    '300': '300 - Imune',
    '400': '400 - Não tributada pelo Simples Nacional',
    '500': '500 - ICMS cobrado anteriormente por ST (substituto) ou antecipação',
    '900': '900 - Outros'
}

def criar_mapa_cst_completo():
    mapa_completo = {}
    for k, v in MAPA_CSOSN.items():
        mapa_completo[k] = v
    for origem in range(9): # 0 a 8
        for cst, desc in MAPA_TRIBUTACAO_ICMS.items():
            chave = f"{origem}{cst}"
            if chave not in mapa_completo:
                mapa_completo[chave] = desc
    return mapa_completo

MAPA_CST_UNIFICADO = criar_mapa_cst_completo()