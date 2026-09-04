"""
Concilia+ — configuração de clientes.

Cada bloco descreve, para UM lojista, os nomes reais de aba/coluna dos
relatórios dele e a estratégia de matching a usar. Adicionar um cliente novo
= acrescentar um bloco aqui (e a senha dele em st.secrets — nunca aqui).

matching_strategy:
  "valor_data" -> casa vl_final + data_recebimento (Sistema) com
                  valor + data_pagamento (Máquina). Use quando os
                  identificadores de Sistema e Máquina NÃO se correspondem
                  — é o caso mais comum, validado com dados reais da HOH
                  Clínica Integrada (os IDs de cada sistema vivem em faixas
                  numéricas completamente diferentes).
  "id_direto"  -> casa id (Sistema) diretamente com id (Máquina). Só use se
                  você confirmar, nos arquivos reais do cliente, que o PDV e
                  a adquirente exportam o MESMO identificador — é raro.

Em qualquer estratégia, pagamentos que a adquirente agrupa em lote (soma de
2+ parcelas pagas na mesma data, creditadas de uma vez) são detectados
automaticamente sobre o que sobrar sem correspondência direta — isso não
depende da estratégia escolhida.

Campos de cada bloco:
  display_name              -> nome mostrado na tela após o login
  sistema_sheet              -> nome da aba no arquivo do Sistema
  sistema_header_row         -> linha do cabeçalho (0-indexada) ou None para detectar automaticamente
  sistema_cols                -> mapa papel canônico -> nome real da coluna
  maquina_sheet               -> nome da aba no arquivo da Máquina
  maquina_header_row          -> linha do cabeçalho (0-indexada) ou None para detectar automaticamente
  maquina_cols                 -> mapa papel canônico -> nome real da coluna
  valor_moeda_br               -> True se "Valor do Pagamento" vem como texto "R$ 1.234,56"
  data_dayfirst                 -> True se as datas do Sistema vêm em dd/mm/aaaa
  matching_strategy             -> "valor_data" ou "id_direto"
  status_liquidado_sistema      -> lista de status do Sistema que NÃO contam como estorno
  status_liquidado_maquina      -> lista de status da Máquina que NÃO contam como estorno
  ativo                          -> True/False — controla acesso independente da senha
"""

CLIENTS = {
    # Cliente de exemplo — configuração e matching JÁ VALIDADOS com dados
    # reais (arquivos Sistema.xlsx / Maquina.xlsx testados nesta conversa).
    "hoh-clinica-integrada": {
        "display_name": "HOH Clínica Integrada",
        "sistema_sheet": "Relatório",
        "sistema_header_row": None,  # cabeçalho na linha 1, detectado automaticamente
        "sistema_cols": {
            "id": "Transação",
            "valor": "Valor",
            "taxas": "Taxas",
            "vl_final": "Vl Final",
            "status": "Status",
            "data": "Data",
            "forma": "Forma",
            "data_recebimento": "Recebimento",
            "parcela": "Parcela",
        },
        "maquina_sheet": "Recebimentos",
        "maquina_header_row": 2,  # cabeçalho real na linha 3 (índice 2) — confirmado pelo cliente
        "maquina_cols": {
            "id": "Nº Compromisso",
            "data_pagamento": "Data de Pagamento",
            "valor": "Valor do Pagamento",
            "status": "Status",
            "contrato": "Contrato",
        },
        "valor_moeda_br": True,
        "data_dayfirst": True,
        "matching_strategy": "valor_data",
        "status_liquidado_sistema": ["LIQUIDADO"],
        "status_liquidado_maquina": ["CREDITO EFETUADO", "ANTECIPACAO - CREDITO EFETUADO"],
        "ativo": True,
    },

    # Próximo lojista entra aqui como um novo bloco. Exemplo ilustrativo
    # (apagar/ajustar quando um cliente real desse perfil chegar):
    # "clinica-sorriso": {
    #     "display_name": "Clínica Sorriso",
    #     "sistema_sheet": "Relatório",
    #     "sistema_header_row": None,
    #     "sistema_cols": {
    #         "id": "Transação", "valor": "Valor", "taxas": "Taxas",
    #         "vl_final": "Vl Final", "status": "Status", "data": "Data",
    #         "forma": "Forma", "data_recebimento": "Recebimento",
    #     },
    #     "maquina_sheet": "Recebimentos",
    #     "maquina_header_row": 2,
    #     "maquina_cols": {
    #         "id": "Nº Compromisso", "data_pagamento": "Data de Pagamento",
    #         "valor": "Valor do Pagamento", "status": "Status",
    #     },
    #     "valor_moeda_br": True,
    #     "data_dayfirst": True,
    #     "matching_strategy": "valor_data",
    #     "status_liquidado_sistema": ["LIQUIDADO"],
    #     "status_liquidado_maquina": ["CREDITO EFETUADO", "ANTECIPACAO - CREDITO EFETUADO"],
    #     "ativo": True,
    # },
}
