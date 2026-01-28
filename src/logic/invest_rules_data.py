# app/fiscal/invest_rules_data.py

"""
LISTA DE EXCEÇÕES - APURAÇÃO INVEST
===================================
Estes são os NCMs que NÃO possuem benefício (INVEST = NÃO),
independente do código do produto começar com 'A' ou '1'.

Lógica aplicada no sistema:
1. Produto começa com 'A' ou '1'? -> Candidato a SIM.
2. Produto está nesta lista abaixo? -> Vira NÃO (Exceção).
"""

# Adicione ou remova NCMs aqui.
# O sistema remove pontos e espaços automaticamente.
NCMS_SEM_BENEFICIO = {
    "38249941",
    "28539019",
    "38112140",
    "39100090",
    "34053000",
    "33074900",
    # Adicione novos NCMs proibidos aqui conforme a planilha for atualizada
}