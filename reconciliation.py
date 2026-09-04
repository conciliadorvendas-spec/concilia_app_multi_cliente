"""
Concilia+ — motor de conciliação.

Recebe os dois arquivos e a config do cliente (clients_config.py) e devolve
um resultado estruturado (dataframes + resumo). Não gera Excel — isso é
responsabilidade de report_builder.py. Nenhum nome de coluna fica fixo aqui:
tudo vem da config.
"""

from collections import defaultdict
from itertools import combinations

import pandas as pd

MAX_BATCH_GROUP_SIZE = 6
VALUE_TOLERANCE = 0.01


class ReconciliationError(Exception):
    """Erro esperado (aba/coluna ausente, formato incompatível) — vira mensagem amigável na UI."""


# --------------------------------------------------------------------------- #
# Carregamento e normalização
# --------------------------------------------------------------------------- #

def _find_header_row(file_obj, sheet_name, expected_cols, max_scan=10):
    file_obj.seek(0)
    try:
        raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None, nrows=max_scan)
    except ValueError as exc:
        raise ReconciliationError(f"Não encontrei a aba '{sheet_name}' no arquivo.") from exc

    best_row, best_overlap = None, 0
    for i in range(len(raw)):
        row_values = {str(v).strip() for v in raw.iloc[i].tolist()}
        overlap = len(row_values & expected_cols)
        if overlap > best_overlap:
            best_overlap, best_row = overlap, i

    if best_row is None or best_overlap < max(2, len(expected_cols) // 2):
        raise ReconciliationError(
            f"Não encontrei o cabeçalho esperado na aba '{sheet_name}'. "
            f"Colunas esperadas: {sorted(expected_cols)}"
        )
    return best_row


def _parse_brl(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace("R$", "").replace("\xa0", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _load_and_rename(file_obj, sheet, header_row, cols_map, role_label):
    expected_cols = set(cols_map.values())
    if header_row is None:
        header_row = _find_header_row(file_obj, sheet, expected_cols)
    file_obj.seek(0)
    try:
        df = pd.read_excel(file_obj, sheet_name=sheet, header=header_row)
    except ValueError as exc:
        raise ReconciliationError(f"Não encontrei a aba '{sheet}' no arquivo do {role_label}.") from exc

    missing = expected_cols - set(df.columns)
    if missing:
        raise ReconciliationError(
            f"Não encontrei a(s) coluna(s) {sorted(missing)} na aba '{sheet}' do {role_label}."
        )

    rename_map = {real_name: canonico for canonico, real_name in cols_map.items()}
    return df.rename(columns=rename_map).reset_index(drop=True)


def load_sistema(file_obj, config):
    df = _load_and_rename(
        file_obj, config["sistema_sheet"], config.get("sistema_header_row"),
        config["sistema_cols"], "Sistema",
    )
    df["data_recebimento"] = pd.to_datetime(
        df["data_recebimento"], dayfirst=config.get("data_dayfirst", True), errors="coerce"
    )
    df["vl_final"] = pd.to_numeric(df["vl_final"], errors="coerce")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce") if "valor" in df.columns else df["vl_final"]
    df["taxas"] = pd.to_numeric(df["taxas"], errors="coerce") if "taxas" in df.columns else 0.0
    if "parcela" not in df.columns:
        df["parcela"] = 1
    if "forma" not in df.columns:
        df["forma"] = ""
    if "data" not in df.columns:
        df["data"] = ""
    return df


def load_maquina(file_obj, config):
    df = _load_and_rename(
        file_obj, config["maquina_sheet"], config.get("maquina_header_row"),
        config["maquina_cols"], "Máquina",
    )
    df["data_pagamento"] = pd.to_datetime(df["data_pagamento"], errors="coerce")
    if config.get("valor_moeda_br", False):
        df["valor"] = df["valor"].apply(_parse_brl)
    else:
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    if "contrato" not in df.columns:
        df["contrato"] = ""
    if "id" not in df.columns:
        df["id"] = ""
    return df


# --------------------------------------------------------------------------- #
# Matching — duas estratégias, mesmo formato de saída (conjuntos de índices)
# --------------------------------------------------------------------------- #

def _match_by_composite_key(sistema, maquina, s_key_cols, m_key_cols):
    """Casa por chave composta (arredondada a 2 casas para valores)."""
    def build_key(row, cols):
        parts = []
        for c in cols:
            v = row[c]
            if isinstance(v, float):
                v = round(v, 2)
            parts.append(str(v))
        return "_".join(parts)

    s_groups, m_groups = defaultdict(list), defaultdict(list)
    for idx, row in sistema.iterrows():
        s_groups[build_key(row, s_key_cols)].append(idx)
    for idx, row in maquina.iterrows():
        m_groups[build_key(row, m_key_cols)].append(idx)

    s_matched, m_matched = set(), set()
    for key, sidxs in s_groups.items():
        midxs = m_groups.get(key, [])
        for i in range(min(len(sidxs), len(midxs))):
            s_matched.add(sidxs[i])
            m_matched.add(midxs[i])
    return s_matched, m_matched


def _match_valor_data(sistema, maquina):
    return _match_by_composite_key(
        sistema, maquina,
        s_key_cols=["vl_final", "data_recebimento"],
        m_key_cols=["valor", "data_pagamento"],
    )


def _match_id_direto(sistema, maquina):
    return _match_by_composite_key(
        sistema, maquina,
        s_key_cols=["id"],
        m_key_cols=["id"],
    )


def _match(sistema, maquina, config):
    strategy = config.get("matching_strategy", "valor_data")
    if strategy == "id_direto":
        return _match_id_direto(sistema, maquina)
    if strategy == "valor_data":
        return _match_valor_data(sistema, maquina)
    raise ReconciliationError(f"matching_strategy '{strategy}' desconhecida.")


# --------------------------------------------------------------------------- #
# Resolução de pagamentos agrupados em lote (independe da estratégia acima)
# --------------------------------------------------------------------------- #

def _resolve_batches(sistema, maquina, s_matched, m_matched):
    sistema_unmatched = sistema.loc[~sistema.index.isin(s_matched)].copy()
    maquina_unmatched = maquina.loc[~maquina.index.isin(m_matched)].copy()

    grouped_resolutions = []
    still_unmatched_sistema = sistema_unmatched.copy()
    unresolved_maquina_rows = []

    for _, mrow in maquina_unmatched.iterrows():
        same_date = still_unmatched_sistema[still_unmatched_sistema["data_recebimento"] == mrow["data_pagamento"]]
        idx_list = list(same_date.index)
        found = False
        for r in range(2, min(len(idx_list), MAX_BATCH_GROUP_SIZE) + 1):
            for combo in combinations(idx_list, r):
                total = round(sistema.loc[list(combo), "vl_final"].sum(), 2)
                if abs(total - round(mrow["valor"], 2)) < VALUE_TOLERANCE:
                    grouped_resolutions.append({
                        "id_maquina": mrow["id"],
                        "valor_pago": mrow["valor"],
                        "data": mrow["data_pagamento"],
                        "ids_sistema": sistema.loc[list(combo), "id"].tolist(),
                        "parcelas": sistema.loc[list(combo), "parcela"].tolist(),
                        "valores": sistema.loc[list(combo), "vl_final"].tolist(),
                    })
                    still_unmatched_sistema = still_unmatched_sistema.drop(index=list(combo))
                    found = True
                    break
            if found:
                break
        if not found:
            unresolved_maquina_rows.append(mrow)

    maquina_orfaos = pd.DataFrame(unresolved_maquina_rows) if unresolved_maquina_rows else maquina.iloc[0:0].copy()
    return still_unmatched_sistema, maquina_orfaos, grouped_resolutions


# --------------------------------------------------------------------------- #
# Resumo financeiro
# --------------------------------------------------------------------------- #

def _compute_summary(sistema, maquina, sistema_orfaos, maquina_orfaos, grouped):
    total_taxas = float(sistema["taxas"].sum())
    total_vlfinal_all = float(sistema["vl_final"].sum())
    total_orfaos_sistema = float(sistema_orfaos["vl_final"].sum()) if not sistema_orfaos.empty else 0.0
    total_entradas = round(total_vlfinal_all - total_orfaos_sistema, 2)

    return {
        "total_entradas": total_entradas,
        "total_taxas": round(total_taxas, 2),
        "n_sistema": int(len(sistema)),
        "n_maquina": int(len(maquina)),
        "n_conciliadas": int(len(sistema) - len(sistema_orfaos)),
        "n_lotes": int(len(grouped)),
        "n_orfaos_sistema": int(len(sistema_orfaos)),
        "n_orfaos_maquina": int(len(maquina_orfaos)),
    }


def _compute_estornos(sistema, config):
    status_ok = {s.upper() for s in config.get("status_liquidado_sistema", [])}
    status_col = sistema["status"].astype(str).str.upper().str.strip()
    return round(float(sistema.loc[~status_col.isin(status_ok), "vl_final"].sum()), 2)


# --------------------------------------------------------------------------- #
# Orquestração
# --------------------------------------------------------------------------- #

def reconcile(sistema_file, maquina_file, config):
    """
    Retorna um dicionário com tudo que report_builder.py precisa:
      summary, sistema_orfaos, maquina_orfaos, grouped
    """
    sistema = load_sistema(sistema_file, config)
    maquina = load_maquina(maquina_file, config)

    s_matched, m_matched = _match(sistema, maquina, config)
    sistema_orfaos, maquina_orfaos, grouped = _resolve_batches(sistema, maquina, s_matched, m_matched)

    summary = _compute_summary(sistema, maquina, sistema_orfaos, maquina_orfaos, grouped)
    summary["total_estornos"] = _compute_estornos(sistema, config)

    return {
        "summary": summary,
        "sistema_orfaos": sistema_orfaos,
        "maquina_orfaos": maquina_orfaos,
        "grouped": grouped,
    }
