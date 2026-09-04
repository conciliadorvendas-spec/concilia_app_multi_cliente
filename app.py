"""
Concilia+ — app multi-cliente.

Um único app, um único link. Login por usuário + senha: o "usuário" é a
chave do cliente em clients_config.CLIENTS; a senha vive em st.secrets
(Streamlit Cloud > Settings > Secrets), nunca no repositório. Depois do
login, a sessão só enxerga a config do cliente autenticado.
"""

from datetime import datetime

import streamlit as st

from clients_config import CLIENTS
from reconciliation import ReconciliationError, reconcile
from report_builder import build_workbook

st.set_page_config(page_title="Concilia+", page_icon="🧾", layout="centered")

NAVY = "#1A3A5C"
GREEN = "#2ECC71"


def fmt_brl(value):
    s = f"{value:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def show_logo(subtitle=None):
    st.markdown(
        f"<h1 style='color:{NAVY}; margin-bottom:0;'>Concilia<span style='color:{GREEN};'>+</span></h1>",
        unsafe_allow_html=True,
    )
    if subtitle:
        st.caption(subtitle)


def check_credentials(username, password):
    """Retorna (config_do_cliente, motivo_do_erro). config é None se falhar."""
    config = CLIENTS.get(username)
    if config is None:
        return None, "Usuário ou senha inválidos."

    credentials = st.secrets.get("credentials", {})
    senha_correta = credentials.get(username)
    if senha_correta is None or password != senha_correta:
        return None, "Usuário ou senha inválidos."

    if not config.get("ativo", True):
        return None, "Acesso suspenso. Fale com a Concilia+ para regularizar sua assinatura."

    return config, None


def login_screen():
    show_logo("Conciliação automática de vendas")
    st.write("")
    with st.form("login"):
        username = st.text_input("Usuário")
        password = st.text_input("Senha", type="password")
        submitted = st.form_submit_button("Entrar", use_container_width=True)

    if submitted:
        config, error = check_credentials(username.strip(), password)
        if error:
            st.error(error)
        else:
            st.session_state["username"] = username.strip()
            st.session_state["client_config"] = config
            st.rerun()


def reset_session():
    for key in ("username", "client_config"):
        st.session_state.pop(key, None)


def app_screen():
    config = st.session_state["client_config"]

    col1, col2 = st.columns([5, 1])
    with col1:
        show_logo(config["display_name"])
    with col2:
        st.write("")
        if st.button("Sair"):
            reset_session()
            st.rerun()

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        sistema_file = st.file_uploader("Relatório do Sistema (ERP)", type=["xlsx"], key="sistema_upl")
    with c2:
        maquina_file = st.file_uploader("Relatório da Máquina (Adquirente)", type=["xlsx"], key="maquina_upl")

    for f in (sistema_file, maquina_file):
        if f is not None and f.size > 10 * 1024 * 1024:
            st.error(f"O arquivo '{f.name}' passa de 10 MB. Envie um arquivo menor.")
            return

    conciliar = st.button(
        "Conciliar vendas",
        type="primary",
        use_container_width=True,
        disabled=not (sistema_file and maquina_file),
    )

    if conciliar:
        with st.spinner("Conciliando vendas..."):
            try:
                result = reconcile(sistema_file, maquina_file, config)
                workbook_bytes = build_workbook(result, config)
            except ReconciliationError as exc:
                st.error(str(exc))
                return
            except Exception as exc:  # noqa: BLE001
                st.error("Erro inesperado ao processar os arquivos. Fale com a Concilia+.")
                st.exception(exc)
                return

        summary = result["summary"]
        st.success("Conciliação concluída!")

        m1, m2, m3 = st.columns(3)
        m1.metric("Total de entradas", fmt_brl(summary["total_entradas"]))
        m2.metric("Total em taxas", fmt_brl(summary["total_taxas"]))
        m3.metric("Estornos / cancelamentos", fmt_brl(summary["total_estornos"]))

        with st.expander("Ver indicadores de conciliação"):
            st.write(f"- Parcelas no Sistema: **{summary['n_sistema']}**")
            st.write(f"- Pagamentos na Máquina: **{summary['n_maquina']}**")
            st.write(f"- Conciliadas com sucesso: **{summary['n_conciliadas']}**")
            st.write(f"- Pagamentos agrupados (lote) validados: **{summary['n_lotes']}**")
            st.write(f"- Vendas sem correspondência na Máquina: **{summary['n_orfaos_sistema']}**")
            st.write(f"- Pagamentos sem correspondência no Sistema: **{summary['n_orfaos_maquina']}**")

        nome_arquivo = config["display_name"].lower().replace(" ", "_")
        st.download_button(
            "Baixar Relatório",
            data=workbook_bytes,
            file_name=f"conciliacao_{nome_arquivo}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        if st.button("Nova conciliação"):
            st.session_state.pop("sistema_upl", None)
            st.session_state.pop("maquina_upl", None)
            st.rerun()


if "client_config" not in st.session_state:
    login_screen()
else:
    app_screen()
