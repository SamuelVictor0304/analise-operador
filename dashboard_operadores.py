from pathlib import Path
import re

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).parent
EVENTOS_FILE = BASE_DIR / "Cobmais-Eventos-908-2026050417.xlsx"
RESULTADOS_FILE = BASE_DIR / "NOVA BASE RESULTADOS 2026.xlsm"


def latest_file(pattern):
    files = [p for p in BASE_DIR.glob(pattern) if not p.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo encontrado para o padrão: {pattern}")
    return max(files, key=lambda p: p.stat().st_mtime)


st.set_page_config(
    page_title="Performance Operacional",
    layout="wide",
    initial_sidebar_state="expanded",
)


CORP_PALETTE = ["#213547", "#2f6f73", "#b7791f", "#6b7280", "#8a3ffc", "#bf616a"]

FIELD_HELP = {
    "OPERADOR": ("Operador", "Negociador responsável pelo evento ou acordo."),
    "FAIXA_ATRASO": ("Faixa de atraso", "Agrupamento do atraso/DPD do contrato em faixas gerenciais."),
    "REGIÃO": ("Região", "Região cadastrada na base de resultados."),
    "clientes": ("Clientes", "Quantidade distinta de contratos/clientes no agrupamento."),
    "clientes_trabalhados": ("Clientes trabalhados", "Quantidade distinta de contratos acionados pelo operador."),
    "acionamentos": ("Acionamentos", "Total de eventos válidos, sem AUTO/importação por padrão."),
    "contatos_efetivos": ("Contatos efetivos", "Mesmo critério de CPC: eventos iniciados por 02, 03, 04 ou 05."),
    "contatos_cliente": ("Contatos cliente", "Eventos iniciados por 02 ou 03, contato direto com cliente."),
    "cpcs": ("CPCs", "Eventos produtivos iniciados por 02, 03, 04 ou 05."),
    "clientes_cpc": ("Clientes com CPC", "Contratos distintos que tiveram ao menos um CPC."),
    "contratos_cpc": ("Contratos com CPC", "Contratos distintos que tiveram ao menos um CPC."),
    "acordos": ("Acordos", "Quantidade de acordos localizados na base de resultados."),
    "pagamentos": ("Pagamentos", "Acordos com status PAGOU ou data de pagamento preenchida."),
    "acordos_sem_pagamento": ("Acordos sem pagamento", "Acordos sem status pago e sem data de pagamento."),
    "acordos_em_aberto": ("Acordos em aberto", "Acordos com status EM ABERTO na base de resultados."),
    "acordos_nao_pagou": ("Acordos não pagos", "Acordos com status NÃO PAGOU na base de resultados."),
    "tx_contato": ("Taxa de CPC", "CPCs divididos pelo total de acionamentos."),
    "tx_acordo": ("Taxa CPC -> acordo", "Acordos divididos pelo total de CPCs do operador."),
    "tx_acordo_cliente_cpc": ("Taxa contrato CPC -> acordo", "Acordos divididos pelos contratos distintos com CPC."),
    "tx_pagamento": ("Taxa acordo -> pagamento", "Pagamentos divididos por acordos."),
    "tx_pagamento_cpc": ("Taxa CPC -> pagamento", "Pagamentos divididos pelo total de CPCs do operador."),
    "tx_sem_pagamento": ("Taxa sem pagamento", "Acordos sem pagamento divididos pelo total de acordos."),
    "tx_cpc_acordo": ("Taxa CPC -> acordo", "Contratos com CPC que geraram acordo, divididos pelos contratos com CPC."),
    "tx_cpc_pagamento": ("Taxa CPC -> pagamento", "Contratos com CPC que geraram pagamento, divididos pelos contratos com CPC."),
    "tx_acordo_pagamento": ("Taxa acordo -> pagamento", "Pagamentos divididos pelos acordos originados em contratos com CPC."),
    "tx_acordo_sem_pagamento": ("Taxa acordo sem pagamento", "Acordos sem pagamento divididos pelos acordos originados em contratos com CPC."),
    "valor_negociado": ("Valor negociado", "Soma da coluna VALOR DO BANCO - META da base de resultados."),
    "valor_pago": ("Valor recebido", "Valor negociado dos acordos pagos; acordos não pagos entram como R$ 0,00."),
    "valor_em_aberto": ("Valor em aberto", "Valor negociado dos acordos com status EM ABERTO."),
    "valor_nao_pagou": ("Valor não pago", "Valor negociado dos acordos com status NÃO PAGOU."),
    "ticket_medio": ("Ticket médio", "Valor negociado médio dos acordos."),
    "recuperacao": ("% recuperação", "Valor recebido dividido pelo valor negociado."),
    "score": ("Score", "Índice composto que pondera contato, acordo, pagamento, valor recebido e volume."),
}


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_contract(value):
    if pd.isna(value):
        return np.nan
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits or np.nan


def normalize_operator(value):
    text = normalize_text(value).lower()
    return text if text else np.nan


def atraso_faixa(days):
    if pd.isna(days):
        return "Sem atraso"
    days = float(days)
    if days <= 30:
        return "000-030"
    if days <= 60:
        return "031-060"
    if days <= 90:
        return "061-090"
    if days <= 120:
        return "091-120"
    if days <= 180:
        return "121-180"
    if days <= 360:
        return "181-360"
    return "361+"


def money_fmt(value):
    value = 0 if pd.isna(value) else float(value)
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct_fmt(value):
    value = 0 if pd.isna(value) or np.isinf(value) else float(value)
    return f"{value:.1%}".replace(".", ",")


def num_fmt(value):
    value = 0 if pd.isna(value) else float(value)
    return f"{value:,.0f}".replace(",", ".")


def safe_div(num, den):
    return np.where(den == 0, 0, num / den)


@st.cache_data(show_spinner=False)
def load_data():
    eventos = pd.read_excel(EVENTOS_FILE, sheet_name="Eventos")
    clientes_file = latest_file("Pesquisa-Cliente-908-*.xlsx")
    contratos = pd.read_excel(clientes_file, sheet_name="Contratos")
    resultados = pd.read_excel(RESULTADOS_FILE, sheet_name="BASE", usecols="A:Z")

    eventos.columns = [normalize_text(c).upper() for c in eventos.columns]
    contratos.columns = [normalize_text(c).upper() for c in contratos.columns]
    resultados.columns = [normalize_text(c).upper() for c in resultados.columns]

    eventos = eventos.dropna(how="all").copy()
    contratos = contratos.dropna(how="all").copy()
    resultados = resultados.dropna(subset=["Nº CONTRATO", "ACORDO POR"], how="all").copy()

    eventos["CONTRATO_KEY"] = eventos["CONTRATO"].map(normalize_contract)
    eventos["OPERADOR"] = eventos["OPERADOR"].map(normalize_operator)
    eventos["DATA"] = pd.to_datetime(eventos["DATA"], dayfirst=True, errors="coerce")
    eventos["EVENTO_TXT"] = eventos["EVENTO"].map(normalize_text)
    eventos["EVENTO_UPPER"] = eventos["EVENTO_TXT"].str.upper()
    eventos["DESCRICAO_UPPER"] = eventos["DESCRIÇÃO"].map(normalize_text).str.upper()
    eventos["TIPO DE ACIONAMENTO"] = eventos["TIPO DE ACIONAMENTO"].map(normalize_text)

    contratos["CONTRATO_KEY"] = contratos["CONTRATO"].map(normalize_contract)
    contratos["ATRASO"] = pd.to_numeric(contratos["ATRASO"], errors="coerce")
    contratos["TOTAL ABERTO"] = pd.to_numeric(contratos["TOTAL ABERTO"], errors="coerce")
    contratos["FAIXA_ATRASO"] = contratos["ATRASO"].map(atraso_faixa)

    contrato_cols = [
        "CONTRATO_KEY",
        "PRODUTO",
        "REGIAO",
        "FILIAL",
        "ESTAGIO",
        "ATRASO",
        "TOTAL ABERTO",
        "FAIXA_ATRASO",
    ]
    contrato_cols = [c for c in contrato_cols if c in contratos.columns]
    contratos_lookup = contratos[contrato_cols].drop_duplicates("CONTRATO_KEY")
    eventos = eventos.merge(contratos_lookup, on="CONTRATO_KEY", how="left", suffixes=("", "_CONTRATO"))

    eventos["IS_AUTO"] = eventos["OPERADOR"].eq("auto")
    eventos["IS_IMPORTACAO"] = eventos["EVENTO_UPPER"].str.contains("IMPORTACAO|IMPORTAÇÃO", na=False)
    eventos["IS_ACIONAMENTO"] = ~(eventos["IS_AUTO"] | eventos["IS_IMPORTACAO"])
    eventos["IS_CPC"] = eventos["EVENTO_UPPER"].str.match(r"^\s*(02|03|04|05)\b", na=False)
    eventos["IS_CONTATO_EFETIVO"] = eventos["IS_CPC"]
    eventos["IS_CONTATO_CLIENTE"] = eventos["EVENTO_UPPER"].str.match(r"^\s*(02|03)\b", na=False)
    eventos["MES"] = eventos["DATA"].dt.to_period("M").astype(str)

    resultados["CONTRATO_KEY"] = resultados["Nº CONTRATO"].map(normalize_contract)
    resultados["OPERADOR"] = resultados["ACORDO POR"].map(normalize_operator)
    resultados["DATA_ACORDO"] = pd.to_datetime(resultados["EMISSÃO"], errors="coerce")
    resultados["DATA_PAGAMENTO"] = pd.to_datetime(resultados["DATA DO PAGAMENTO"], errors="coerce")
    resultados["VALOR_NEGOCIADO"] = pd.to_numeric(resultados["VALOR DO BANCO - META"], errors="coerce").fillna(0)
    resultados["HONORARIOS"] = pd.to_numeric(resultados["HONORÁRIOS %"], errors="coerce").fillna(0)
    resultados["DPD"] = pd.to_numeric(resultados["DPD"], errors="coerce")
    resultados["FAIXA_ATRASO"] = resultados["DPD"].map(atraso_faixa)
    resultados["REGIÃO"] = resultados["REGIÃO"].fillna(resultados.get("UF", "Sem região")).map(normalize_text)
    resultados["UF"] = resultados["UF"].map(normalize_text)
    resultados["CAMPANHA"] = resultados["CAMPANHA"].map(normalize_text)
    resultados["TIPO DE ACORDO"] = resultados["TIPO DE ACORDO"].map(normalize_text)
    resultados["STATUS"] = resultados["STATUS"].map(normalize_text).str.upper()
    resultados["IS_ACORDO"] = resultados["CONTRATO_KEY"].notna() & resultados["OPERADOR"].notna()
    resultados["IS_EM_ABERTO"] = resultados["STATUS"].eq("EM ABERTO")
    resultados["IS_NAO_PAGOU"] = resultados["STATUS"].eq("NÃO PAGOU")
    resultados["IS_PAGO"] = resultados["STATUS"].eq("PAGOU") | (resultados["DATA_PAGAMENTO"].notna())
    resultados["VALOR_PAGO"] = np.where(resultados["IS_PAGO"], resultados["VALOR_NEGOCIADO"], 0)
    resultados["VALOR_EM_ABERTO"] = np.where(resultados["IS_EM_ABERTO"], resultados["VALOR_NEGOCIADO"], 0)
    resultados["VALOR_NAO_PAGOU"] = np.where(resultados["IS_NAO_PAGOU"], resultados["VALOR_NEGOCIADO"], 0)
    resultados["MES_RESULTADO"] = resultados["MÊS"].map(normalize_text).str.upper()
    resultados["MES_NUM"] = pd.to_numeric(resultados["Nº CORRESPONDENTE AO MÊS"], errors="coerce")
    resultados["MES"] = resultados["DATA_ACORDO"].dt.to_period("M").astype(str)
    resultados = resultados[resultados["IS_ACORDO"]].copy()

    return eventos, contratos, resultados


def apply_filters(eventos, resultados):
    st.sidebar.title("Filtros")

    operadores = sorted(set(eventos["OPERADOR"].dropna()) | set(resultados["OPERADOR"].dropna()))
    regioes = sorted(resultados["REGIÃO"].replace("", np.nan).dropna().unique())
    faixas = ["000-030", "031-060", "061-090", "091-120", "121-180", "181-360", "361+", "Sem atraso"]
    campanhas = sorted(resultados["CAMPANHA"].replace("", np.nan).dropna().unique())
    produtos = sorted(eventos.get("PRODUTO", pd.Series(dtype=str)).replace("", np.nan).dropna().unique())
    meses_df = (
        resultados[["MES_RESULTADO", "MES_NUM"]]
        .replace("", np.nan)
        .dropna(subset=["MES_RESULTADO"])
        .drop_duplicates()
        .sort_values(["MES_NUM", "MES_RESULTADO"])
    )
    meses = meses_df["MES_RESULTADO"].tolist()
    mes_padrao = []
    if not meses_df.empty:
        mes_padrao = [meses_df.sort_values(["MES_NUM", "MES_RESULTADO"]).iloc[-1]["MES_RESULTADO"]]

    operador_sel = st.sidebar.multiselect("Operador", operadores)
    mes_sel = st.sidebar.multiselect("Mês do resultado", meses, default=mes_padrao)
    regiao_sel = st.sidebar.multiselect("Região", regioes)
    faixa_sel = st.sidebar.multiselect("Faixa de atraso", faixas)
    campanha_sel = st.sidebar.multiselect("Campanha", campanhas)
    produto_sel = st.sidebar.multiselect("Produto", produtos)
    incluir_auto = st.sidebar.toggle("Incluir AUTO/importação nos acionamentos", value=False)

    min_data = eventos["DATA"].min()
    max_data = eventos["DATA"].max()
    data_range = st.sidebar.date_input(
        "Período dos eventos",
        value=(min_data.date(), max_data.date()),
        min_value=min_data.date(),
        max_value=max_data.date(),
    )
    if isinstance(data_range, tuple) and len(data_range) == 2:
        inicio, fim = pd.Timestamp(data_range[0]), pd.Timestamp(data_range[1]) + pd.Timedelta(days=1)
        eventos = eventos[(eventos["DATA"].isna()) | ((eventos["DATA"] >= inicio) & (eventos["DATA"] < fim))]

    if operador_sel:
        eventos = eventos[eventos["OPERADOR"].isin(operador_sel)]
        resultados = resultados[resultados["OPERADOR"].isin(operador_sel)]
    if mes_sel:
        resultados = resultados[resultados["MES_RESULTADO"].isin(mes_sel)]
    if regiao_sel:
        resultados = resultados[resultados["REGIÃO"].isin(regiao_sel)]
    if faixa_sel:
        eventos = eventos[eventos["FAIXA_ATRASO"].isin(faixa_sel)]
        resultados = resultados[resultados["FAIXA_ATRASO"].isin(faixa_sel)]
    if campanha_sel:
        resultados = resultados[resultados["CAMPANHA"].isin(campanha_sel)]
        contratos_campanha = set(resultados["CONTRATO_KEY"].dropna())
        eventos = eventos[eventos["CONTRATO_KEY"].isin(contratos_campanha)]
    if produto_sel and "PRODUTO" in eventos.columns:
        eventos = eventos[eventos["PRODUTO"].isin(produto_sel)]
    if not incluir_auto:
        eventos = eventos[eventos["IS_ACIONAMENTO"]]

    return eventos, resultados


def aggregate_operator(eventos, resultados):
    ev = eventos.groupby("OPERADOR", dropna=True).agg(
        acionamentos=("EVENTO_TXT", "size"),
        clientes_trabalhados=("CONTRATO_KEY", "nunique"),
        contatos_efetivos=("IS_CONTATO_EFETIVO", "sum"),
        contatos_cliente=("IS_CONTATO_CLIENTE", "sum"),
        cpcs=("IS_CPC", "sum"),
        clientes_cpc=("CONTRATO_KEY", lambda s: s[eventos.loc[s.index, "IS_CPC"]].nunique()),
    )
    rs = resultados.groupby("OPERADOR", dropna=True).agg(
        acordos=("CONTRATO_KEY", "count"),
        pagamentos=("IS_PAGO", "sum"),
        acordos_sem_pagamento=("IS_PAGO", lambda s: (~s).sum()),
        acordos_em_aberto=("IS_EM_ABERTO", "sum"),
        acordos_nao_pagou=("IS_NAO_PAGOU", "sum"),
        valor_negociado=("VALOR_NEGOCIADO", "sum"),
        valor_pago=("VALOR_PAGO", "sum"),
        valor_em_aberto=("VALOR_EM_ABERTO", "sum"),
        valor_nao_pagou=("VALOR_NAO_PAGOU", "sum"),
        ticket_medio=("VALOR_NEGOCIADO", "mean"),
    )
    df = ev.join(rs, how="outer").fillna(0).reset_index()
    df["tx_contato"] = safe_div(df["contatos_efetivos"], df["acionamentos"])
    df["tx_acordo"] = safe_div(df["acordos"], df["cpcs"])
    df["tx_acordo_cliente_cpc"] = safe_div(df["acordos"], df["clientes_cpc"])
    df["tx_pagamento"] = safe_div(df["pagamentos"], df["acordos"])
    df["tx_pagamento_cpc"] = safe_div(df["pagamentos"], df["cpcs"])
    df["tx_sem_pagamento"] = safe_div(df["acordos_sem_pagamento"], df["acordos"])
    df["recuperacao"] = safe_div(df["valor_pago"], df["valor_negociado"])
    df["score"] = (
        df["tx_contato"].rank(pct=True) * 0.15
        + df["tx_acordo"].rank(pct=True) * 0.25
        + df["tx_pagamento"].rank(pct=True) * 0.25
        + df["valor_pago"].rank(pct=True) * 0.25
        + df["clientes_trabalhados"].rank(pct=True) * 0.10
    )
    return df.sort_values(["score", "valor_pago"], ascending=False)


def aggregate_cpc_operator(eventos, resultados):
    cpc_eventos = eventos[eventos["IS_CPC"]].copy()
    ev = cpc_eventos.groupby("OPERADOR", dropna=True).agg(
        cpcs=("EVENTO_TXT", "size"),
        clientes_cpc=("CONTRATO_KEY", "nunique"),
        contratos_cpc=("CONTRATO_KEY", "nunique"),
    )
    cpc_contratos = cpc_eventos[["OPERADOR", "CONTRATO_KEY"]].dropna().drop_duplicates()
    resultado_contrato = resultados.groupby(["OPERADOR", "CONTRATO_KEY"], dropna=True).agg(
        qtd_acordos=("CONTRATO_KEY", "count"),
        teve_pagamento=("IS_PAGO", "max"),
        teve_em_aberto=("IS_EM_ABERTO", "max"),
        teve_nao_pagou=("IS_NAO_PAGOU", "max"),
        valor_negociado=("VALOR_NEGOCIADO", "sum"),
        valor_pago=("VALOR_PAGO", "sum"),
        valor_em_aberto=("VALOR_EM_ABERTO", "sum"),
        valor_nao_pagou=("VALOR_NAO_PAGOU", "sum"),
    ).reset_index()
    cpc_resultado = cpc_contratos.merge(resultado_contrato, on=["OPERADOR", "CONTRATO_KEY"], how="left")
    cpc_resultado["qtd_acordos"] = cpc_resultado["qtd_acordos"].fillna(0)
    cpc_resultado["teve_pagamento"] = cpc_resultado["teve_pagamento"].fillna(False).astype(bool)
    cpc_resultado["teve_em_aberto"] = cpc_resultado["teve_em_aberto"].fillna(False).astype(bool)
    cpc_resultado["teve_nao_pagou"] = cpc_resultado["teve_nao_pagou"].fillna(False).astype(bool)
    cpc_resultado["valor_negociado"] = cpc_resultado["valor_negociado"].fillna(0)
    cpc_resultado["valor_pago"] = cpc_resultado["valor_pago"].fillna(0)
    cpc_resultado["valor_em_aberto"] = cpc_resultado["valor_em_aberto"].fillna(0)
    cpc_resultado["valor_nao_pagou"] = cpc_resultado["valor_nao_pagou"].fillna(0)
    cpc_resultado["teve_acordo"] = cpc_resultado["qtd_acordos"] > 0
    cpc_resultado["acordo_sem_pagamento"] = cpc_resultado["teve_acordo"] & ~cpc_resultado["teve_pagamento"]

    rs = cpc_resultado.groupby("OPERADOR", dropna=True).agg(
        acordos=("teve_acordo", "sum"),
        pagamentos=("teve_pagamento", "sum"),
        acordos_sem_pagamento=("acordo_sem_pagamento", "sum"),
        acordos_em_aberto=("teve_em_aberto", "sum"),
        acordos_nao_pagou=("teve_nao_pagou", "sum"),
        valor_negociado=("valor_negociado", "sum"),
        valor_pago=("valor_pago", "sum"),
        valor_em_aberto=("valor_em_aberto", "sum"),
        valor_nao_pagou=("valor_nao_pagou", "sum"),
    )
    df = ev.join(rs, how="left").fillna(0).reset_index()
    df["ticket_medio"] = safe_div(df["valor_negociado"], df["acordos"])
    df["tx_cpc_acordo"] = safe_div(df["acordos"], df["contratos_cpc"])
    df["tx_cpc_pagamento"] = safe_div(df["pagamentos"], df["contratos_cpc"])
    df["tx_acordo_pagamento"] = safe_div(df["pagamentos"], df["acordos"])
    df["tx_acordo_sem_pagamento"] = safe_div(df["acordos_sem_pagamento"], df["acordos"])
    df["recuperacao"] = safe_div(df["valor_pago"], df["valor_negociado"])
    return df.sort_values(["pagamentos", "valor_pago", "tx_cpc_pagamento"], ascending=False)


def aggregate_resultados(resultados, dimension):
    df = resultados.groupby(dimension, dropna=False).agg(
        clientes=("CONTRATO_KEY", "nunique"),
        acordos=("CONTRATO_KEY", "count"),
        pagamentos=("IS_PAGO", "sum"),
        acordos_em_aberto=("IS_EM_ABERTO", "sum"),
        acordos_nao_pagou=("IS_NAO_PAGOU", "sum"),
        valor_negociado=("VALOR_NEGOCIADO", "sum"),
        valor_pago=("VALOR_PAGO", "sum"),
        valor_em_aberto=("VALOR_EM_ABERTO", "sum"),
        valor_nao_pagou=("VALOR_NAO_PAGOU", "sum"),
        ticket_medio=("VALOR_NEGOCIADO", "mean"),
    ).reset_index()
    df["tx_pagamento"] = safe_div(df["pagamentos"], df["acordos"])
    df["recuperacao"] = safe_div(df["valor_pago"], df["valor_negociado"])
    return df


def metric_card(label, value, help_text=None):
    st.metric(label, value, help=help_text)


def bar_chart(df, x, y, color=None, tooltip=None, title=None, sort="-x", height=320):
    if df.empty:
        st.info("Sem dados para os filtros selecionados.")
        return
    x_encoding = x if not isinstance(x, str) else alt.X(x, title=None)
    y_encoding = y if not isinstance(y, str) else alt.Y(y, title=None, sort=sort)
    chart = (
        alt.Chart(df)
        .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3)
        .encode(
            x=x_encoding,
            y=y_encoding,
            color=alt.Color(color, scale=alt.Scale(range=CORP_PALETTE), legend=None) if color else alt.value("#2f6f73"),
            tooltip=tooltip or list(df.columns),
        )
        .properties(height=height, title=title)
    )
    st.altair_chart(chart, use_container_width=True)


def line_chart(df, x, y, color, title):
    if df.empty:
        st.info("Sem dados para os filtros selecionados.")
        return
    chart = (
        alt.Chart(df)
        .mark_line(point=True)
        .encode(
            x=alt.X(x, title=None),
            y=alt.Y(y, title=None),
            color=alt.Color(color, scale=alt.Scale(range=CORP_PALETTE)),
            tooltip=list(df.columns),
        )
        .properties(height=280, title=title)
    )
    st.altair_chart(chart, use_container_width=True)


def heatmap(df, x, y, metric, title):
    if df.empty:
        st.info("Sem dados para os filtros selecionados.")
        return
    tooltip_cols = [
        col for col in df.columns
        if not (col in {"valor_negociado", "valor_pago", "ticket_medio"} and f"{col}_br" in df.columns)
    ]
    chart = (
        alt.Chart(df)
        .mark_rect()
        .encode(
            x=alt.X(x, title=None),
            y=alt.Y(y, title=None),
            color=alt.Color(metric, scale=alt.Scale(scheme="tealblues"), title=None),
            tooltip=tooltip_cols,
        )
        .properties(height=420, title=title)
    )
    st.altair_chart(chart, use_container_width=True)


def display_fields(df):
    out = df.copy()
    money_cols = ["valor_negociado", "valor_pago", "valor_em_aberto", "valor_nao_pagou", "ticket_medio"]
    pct_cols = [
        "tx_contato",
        "tx_acordo",
        "tx_acordo_cliente_cpc",
        "tx_pagamento",
        "tx_pagamento_cpc",
        "tx_sem_pagamento",
        "tx_cpc_acordo",
        "tx_cpc_pagamento",
        "tx_acordo_pagamento",
        "tx_acordo_sem_pagamento",
        "recuperacao",
        "score",
    ]
    num_cols = [
        "acionamentos",
        "clientes_trabalhados",
        "contatos_efetivos",
        "contatos_cliente",
        "cpcs",
        "clientes_cpc",
        "contratos_cpc",
        "acordos",
        "pagamentos",
        "acordos_sem_pagamento",
        "acordos_em_aberto",
        "acordos_nao_pagou",
        "clientes",
    ]
    for col in money_cols:
        if col in out:
            out[f"{col}_br"] = out[col].map(money_fmt)
    for col in pct_cols:
        if col in out:
            out[f"{col}_br"] = out[col].map(pct_fmt)
    for col in num_cols:
        if col in out:
            out[f"{col}_br"] = out[col].map(num_fmt)
    return out


def formatted_table(df):
    out = df.copy()
    for col in ["valor_negociado", "valor_pago", "valor_em_aberto", "valor_nao_pagou", "ticket_medio"]:
        if col in out:
            out[col] = out[col].map(money_fmt)
    for col in [
        "tx_contato",
        "tx_acordo",
        "tx_acordo_cliente_cpc",
        "tx_pagamento",
        "tx_pagamento_cpc",
        "tx_sem_pagamento",
        "tx_cpc_acordo",
        "tx_cpc_pagamento",
        "tx_acordo_pagamento",
        "tx_acordo_sem_pagamento",
        "recuperacao",
        "score",
    ]:
        if col in out:
            out[col] = out[col].map(pct_fmt)
    for col in [
        "acionamentos",
        "clientes_trabalhados",
        "contatos_efetivos",
        "contatos_cliente",
        "cpcs",
        "clientes_cpc",
        "contratos_cpc",
        "acordos",
        "pagamentos",
        "acordos_sem_pagamento",
        "acordos_em_aberto",
        "acordos_nao_pagou",
        "clientes",
    ]:
        if col in out:
            out[col] = out[col].map(num_fmt)
    return out


def help_config(df):
    config = {}
    for col in df.columns:
        if col in FIELD_HELP:
            label, help_text = FIELD_HELP[col]
            config[col] = st.column_config.TextColumn(label=label, help=help_text)
    return config


def data_table(df, **kwargs):
    formatted = formatted_table(df)
    st.dataframe(
        formatted,
        column_config=help_config(formatted),
        use_container_width=True,
        hide_index=True,
        **kwargs,
    )


def glossary():
    with st.expander("Glossário dos indicadores"):
        items = [
            ("CPC", "Eventos iniciados por 02, 03, 04 ou 05."),
            ("Valor negociado", "Soma da coluna VALOR DO BANCO - META."),
            ("Valor recebido", "Valor negociado apenas dos acordos pagos; não pagos entram como R$ 0,00."),
            ("Taxa CPC -> acordo", "Contratos com CPC que geraram acordo / contratos com CPC."),
            ("Taxa CPC -> pagamento", "Contratos com CPC que geraram pagamento / contratos com CPC."),
            ("Acordos sem pagamento", "Acordos que ainda não possuem status PAGOU nem data de pagamento."),
            ("Em aberto", "Acordos ainda pendentes, com status EM ABERTO."),
            ("Não pagou", "Acordos vencidos/sem efetivação, com status NÃO PAGOU."),
            ("Recuperação", "Valor recebido / valor negociado."),
        ]
        for name, desc in items:
            st.markdown(f"**{name}:** {desc}")


eventos_raw, contratos_raw, resultados_raw = load_data()
eventos, resultados = apply_filters(eventos_raw, resultados_raw)
operador_df = aggregate_operator(eventos, resultados)
cpc_df = aggregate_cpc_operator(eventos, resultados)

st.title("Performance Operacional por Operador")
st.caption("Análise executiva de acionamentos, contatos, acordos, pagamentos, recuperação e eficiência por segmento.")
glossary()

total_clientes = eventos["CONTRATO_KEY"].nunique()
total_acionamentos = len(eventos)
total_contatos = int(eventos["IS_CONTATO_EFETIVO"].sum())
total_acordos = len(resultados)
total_pagamentos = int(resultados["IS_PAGO"].sum())
total_em_aberto = int(resultados["IS_EM_ABERTO"].sum())
total_nao_pagou = int(resultados["IS_NAO_PAGOU"].sum())
valor_negociado = resultados["VALOR_NEGOCIADO"].sum()
valor_pago = resultados["VALOR_PAGO"].sum()
valor_em_aberto = resultados["VALOR_EM_ABERTO"].sum()
valor_nao_pagou = resultados["VALOR_NAO_PAGOU"].sum()

kpi_cols = st.columns(10)
with kpi_cols[0]:
    metric_card("Clientes", num_fmt(total_clientes))
with kpi_cols[1]:
    metric_card("Acionamentos", num_fmt(total_acionamentos))
with kpi_cols[2]:
    metric_card("Contatos efetivos", num_fmt(total_contatos))
with kpi_cols[3]:
    metric_card("Acordos", num_fmt(total_acordos))
with kpi_cols[4]:
    metric_card("Pagamentos", num_fmt(total_pagamentos))
with kpi_cols[5]:
    metric_card("Em aberto", num_fmt(total_em_aberto))
with kpi_cols[6]:
    metric_card("Não pagou", num_fmt(total_nao_pagou))
with kpi_cols[7]:
    metric_card("Negociado", money_fmt(valor_negociado))
with kpi_cols[8]:
    metric_card("Recebido", money_fmt(valor_pago))
with kpi_cols[9]:
    metric_card("Recuperação", pct_fmt(valor_pago / valor_negociado if valor_negociado else 0))

tabs = st.tabs(["Visão Geral", "Operadores", "CPC", "Faixa de Atraso", "Região", "Matriz", "Insights"])

with tabs[0]:
    c1, c2 = st.columns([1.2, 1])
    with c1:
        top = display_fields(operador_df.head(12))
        bar_chart(
            top,
            x=alt.X("score:Q", axis=alt.Axis(format="%")),
            y="OPERADOR:N",
            tooltip=[
                "OPERADOR",
                "score_br",
                "acionamentos_br",
                "cpcs_br",
                "acordos_br",
                "pagamentos_br",
                "valor_pago_br",
            ],
            title="Ranking geral de performance",
        )
    with c2:
        funil = pd.DataFrame(
            {
                "Etapa": ["Acionamentos", "Contatos efetivos", "Acordos", "Pagamentos", "Em aberto", "Não pagou"],
                "Volume": [total_acionamentos, total_contatos, total_acordos, total_pagamentos, total_em_aberto, total_nao_pagou],
            }
        )
        bar_chart(funil, x="Volume:Q", y="Etapa:N", title="Funil operacional", sort=None, height=320)

    status_df = pd.DataFrame(
        {
            "Status": ["Pagou", "Em aberto", "Não pagou", "Outros sem pagamento"],
            "Contratos": [
                total_pagamentos,
                total_em_aberto,
                total_nao_pagou,
                max(total_acordos - total_pagamentos - total_em_aberto - total_nao_pagou, 0),
            ],
            "Valor": [
                valor_pago,
                valor_em_aberto,
                valor_nao_pagou,
                max(valor_negociado - valor_pago - valor_em_aberto - valor_nao_pagou, 0),
            ],
        }
    )
    status_df = display_fields(status_df.rename(columns={"Contratos": "clientes", "Valor": "valor_negociado"}))
    bar_chart(
        status_df,
        x="valor_negociado:Q",
        y="Status:N",
        color="Status:N",
        tooltip=["Status", "clientes_br", "valor_negociado_br"],
        title="Distribuição financeira por status",
        sort=None,
        height=260,
    )

    by_mes_eventos = eventos.groupby("MES").agg(acionamentos=("EVENTO_TXT", "size"), contatos=("IS_CONTATO_EFETIVO", "sum")).reset_index()
    by_mes_result = resultados.groupby("MES").agg(acordos=("CONTRATO_KEY", "count"), pagamentos=("IS_PAGO", "sum")).reset_index()
    by_mes = by_mes_eventos.merge(by_mes_result, on="MES", how="outer").fillna(0)
    trend = by_mes.melt("MES", value_vars=["acionamentos", "contatos", "acordos", "pagamentos"], var_name="Indicador", value_name="Volume")
    line_chart(trend, "MES:N", "Volume:Q", "Indicador:N", "Evolução mensal")

with tabs[1]:
    c1, c2 = st.columns(2)
    with c1:
        chart_operador = display_fields(operador_df.head(15))
        bar_chart(
            chart_operador,
            x="valor_pago:Q",
            y="OPERADOR:N",
            tooltip=["OPERADOR", "valor_pago_br", "pagamentos_br", "tx_pagamento_br", "score_br"],
            title="Valor recuperado por operador",
        )
    with c2:
        volume_eficiencia = display_fields(operador_df[operador_df["acionamentos"] > 0])
        scatter = (
            alt.Chart(volume_eficiencia)
            .mark_circle(size=120, opacity=0.78)
            .encode(
                x=alt.X("acionamentos:Q", title="Acionamentos"),
                y=alt.Y("tx_acordo:Q", title="Conversão contato/acordo", axis=alt.Axis(format="%")),
                size=alt.Size("valor_pago:Q", title="Valor recebido"),
                color=alt.Color("tx_pagamento:Q", scale=alt.Scale(scheme="tealblues"), title="Acordo/pagamento"),
                tooltip=[
                    "OPERADOR",
                    "acionamentos_br",
                    "cpcs_br",
                    "acordos_br",
                    "pagamentos_br",
                    "valor_pago_br",
                    "tx_acordo_br",
                    "tx_pagamento_br",
                ],
            )
            .properties(height=320, title="Volume versus eficiência")
        )
        st.altair_chart(scatter, use_container_width=True)

    st.subheader("Ranking detalhado")
    cols = [
        "OPERADOR",
        "clientes_trabalhados",
        "acionamentos",
        "cpcs",
        "clientes_cpc",
        "acordos",
        "pagamentos",
        "acordos_sem_pagamento",
        "acordos_em_aberto",
        "acordos_nao_pagou",
        "tx_contato",
        "tx_acordo",
        "tx_pagamento_cpc",
        "tx_pagamento",
        "tx_sem_pagamento",
        "valor_negociado",
        "valor_pago",
        "valor_em_aberto",
        "valor_nao_pagou",
        "ticket_medio",
        "recuperacao",
        "score",
    ]
    data_table(operador_df[cols])

with tabs[2]:
    st.subheader("Conversão CPC para acordos e pagamentos")
    st.caption("CPC considerado pelos eventos iniciados por 02, 03, 04 e 05.")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        metric_card("CPCs", num_fmt(cpc_df["cpcs"].sum()))
    with c2:
        metric_card("Acordos após CPC", num_fmt(cpc_df["acordos"].sum()))
    with c3:
        metric_card("Pagamentos", num_fmt(cpc_df["pagamentos"].sum()))
    with c4:
        metric_card("Acordos sem pagamento", num_fmt(cpc_df["acordos_sem_pagamento"].sum()))
    with c5:
        metric_card("Em aberto", num_fmt(cpc_df["acordos_em_aberto"].sum()))
    with c6:
        metric_card("Não pagou", num_fmt(cpc_df["acordos_nao_pagou"].sum()))

    c1, c2 = st.columns(2)
    with c1:
        cpc_chart = display_fields(cpc_df[cpc_df["cpcs"] > 0].sort_values("tx_cpc_acordo", ascending=False).head(15))
        bar_chart(
            cpc_chart,
            x=alt.X("tx_cpc_acordo:Q", axis=alt.Axis(format="%")),
            y="OPERADOR:N",
            tooltip=["OPERADOR", "cpcs_br", "acordos_br", "tx_cpc_acordo_br", "valor_negociado_br"],
            title="Conversão CPC -> acordo por negociador",
        )
    with c2:
        cpc_pag_chart = display_fields(cpc_df[cpc_df["cpcs"] > 0].sort_values("tx_cpc_pagamento", ascending=False).head(15))
        bar_chart(
            cpc_pag_chart,
            x=alt.X("tx_cpc_pagamento:Q", axis=alt.Axis(format="%")),
            y="OPERADOR:N",
            tooltip=["OPERADOR", "cpcs_br", "pagamentos_br", "tx_cpc_pagamento_br", "valor_pago_br"],
            title="Conversão CPC -> pagamento por negociador",
        )

    c1, c2 = st.columns(2)
    with c1:
        sem_pg = display_fields(cpc_df[cpc_df["acordos_sem_pagamento"] > 0].sort_values("acordos_sem_pagamento", ascending=False).head(15))
        bar_chart(
            sem_pg,
            x="acordos_sem_pagamento:Q",
            y="OPERADOR:N",
            tooltip=[
                "OPERADOR",
                "acordos_br",
                "acordos_sem_pagamento_br",
                "tx_acordo_sem_pagamento_br",
                "valor_negociado_br",
            ],
            title="Acordos convertidos sem pagamento",
        )
    with c2:
        cpc_scatter = display_fields(cpc_df[(cpc_df["cpcs"] > 0) & (cpc_df["acordos"] > 0)])
        scatter = (
            alt.Chart(cpc_scatter)
            .mark_circle(size=130, opacity=0.78)
            .encode(
                x=alt.X("tx_cpc_acordo:Q", title="CPC -> acordo", axis=alt.Axis(format="%")),
                y=alt.Y("tx_acordo_pagamento:Q", title="Acordo -> pagamento", axis=alt.Axis(format="%")),
                size=alt.Size("valor_pago:Q", title="Valor recebido"),
                color=alt.Color("acordos_sem_pagamento:Q", scale=alt.Scale(scheme="orangered"), title="Sem pagamento"),
                tooltip=[
                    "OPERADOR",
                    "cpcs_br",
                    "acordos_br",
                    "pagamentos_br",
                    "acordos_sem_pagamento_br",
                    "tx_cpc_acordo_br",
                    "tx_acordo_pagamento_br",
                    "valor_pago_br",
                ],
            )
            .properties(height=320, title="Qualidade da conversão")
        )
        st.altair_chart(scatter, use_container_width=True)

    st.subheader("Tabela analítica CPC")
    cpc_cols = [
        "OPERADOR",
        "cpcs",
        "clientes_cpc",
        "acordos",
        "pagamentos",
        "acordos_sem_pagamento",
        "acordos_em_aberto",
        "acordos_nao_pagou",
        "tx_cpc_acordo",
        "tx_cpc_pagamento",
        "tx_acordo_pagamento",
        "tx_acordo_sem_pagamento",
        "valor_negociado",
        "valor_pago",
        "valor_em_aberto",
        "valor_nao_pagou",
        "ticket_medio",
        "recuperacao",
    ]
    data_table(cpc_df[cpc_cols])

with tabs[3]:
    faixa_df = aggregate_resultados(resultados, "FAIXA_ATRASO")
    ev_faixa = eventos.groupby("FAIXA_ATRASO").agg(acionamentos=("EVENTO_TXT", "size"), contatos_efetivos=("IS_CONTATO_EFETIVO", "sum")).reset_index()
    faixa_df = faixa_df.merge(ev_faixa, on="FAIXA_ATRASO", how="outer").fillna(0)
    faixa_df["tx_contato"] = safe_div(faixa_df["contatos_efetivos"], faixa_df["acionamentos"])
    faixa_chart = display_fields(faixa_df)

    c1, c2 = st.columns(2)
    with c1:
        bar_chart(
            faixa_chart,
            x="valor_pago:Q",
            y="FAIXA_ATRASO:N",
            tooltip=["FAIXA_ATRASO", "valor_pago_br", "valor_negociado_br", "acordos_br", "pagamentos_br"],
            title="Valor recuperado por faixa",
            sort=None,
        )
    with c2:
        bar_chart(
            faixa_chart,
            x=alt.X("tx_pagamento:Q", axis=alt.Axis(format="%")),
            y="FAIXA_ATRASO:N",
            tooltip=["FAIXA_ATRASO", "tx_pagamento_br", "acordos_br", "pagamentos_br", "recuperacao_br"],
            title="Conversão acordo/pagamento por faixa",
            sort=None,
        )

    best_faixa = (
        resultados.groupby(["FAIXA_ATRASO", "OPERADOR"])
        .agg(
            acordos=("CONTRATO_KEY", "count"),
            pagamentos=("IS_PAGO", "sum"),
            acordos_em_aberto=("IS_EM_ABERTO", "sum"),
            acordos_nao_pagou=("IS_NAO_PAGOU", "sum"),
            valor_pago=("VALOR_PAGO", "sum"),
            valor_em_aberto=("VALOR_EM_ABERTO", "sum"),
            valor_nao_pagou=("VALOR_NAO_PAGOU", "sum"),
        )
        .reset_index()
    )
    best_faixa["tx_pagamento"] = safe_div(best_faixa["pagamentos"], best_faixa["acordos"])
    best_faixa = best_faixa.sort_values(["FAIXA_ATRASO", "valor_pago", "tx_pagamento"], ascending=[True, False, False]).groupby("FAIXA_ATRASO").head(1)
    st.subheader("Melhor operador por faixa")
    data_table(best_faixa)

with tabs[4]:
    regiao_df = aggregate_resultados(resultados, "REGIÃO").sort_values("valor_pago", ascending=False)
    regiao_chart = display_fields(regiao_df)
    c1, c2 = st.columns(2)
    with c1:
        bar_chart(
            regiao_chart,
            x="valor_pago:Q",
            y="REGIÃO:N",
            tooltip=["REGIÃO", "valor_pago_br", "valor_negociado_br", "acordos_br", "pagamentos_br"],
            title="Valor recebido por região",
        )
    with c2:
        bar_chart(
            regiao_chart,
            x=alt.X("recuperacao:Q", axis=alt.Axis(format="%")),
            y="REGIÃO:N",
            tooltip=["REGIÃO", "recuperacao_br", "valor_pago_br", "valor_negociado_br"],
            title="Recuperação por região",
        )

    best_regiao = (
        resultados.groupby(["REGIÃO", "OPERADOR"])
        .agg(
            acordos=("CONTRATO_KEY", "count"),
            pagamentos=("IS_PAGO", "sum"),
            acordos_em_aberto=("IS_EM_ABERTO", "sum"),
            acordos_nao_pagou=("IS_NAO_PAGOU", "sum"),
            valor_pago=("VALOR_PAGO", "sum"),
            valor_em_aberto=("VALOR_EM_ABERTO", "sum"),
            valor_nao_pagou=("VALOR_NAO_PAGOU", "sum"),
            valor_negociado=("VALOR_NEGOCIADO", "sum"),
        )
        .reset_index()
    )
    best_regiao["recuperacao"] = safe_div(best_regiao["valor_pago"], best_regiao["valor_negociado"])
    best_regiao = best_regiao.sort_values(["REGIÃO", "valor_pago", "recuperacao"], ascending=[True, False, False]).groupby("REGIÃO").head(3)
    st.subheader("Top operadores por região")
    data_table(best_regiao)

with tabs[5]:
    matrix = (
        resultados.groupby(["OPERADOR", "FAIXA_ATRASO", "REGIÃO"])
        .agg(
            acordos=("CONTRATO_KEY", "count"),
            pagamentos=("IS_PAGO", "sum"),
            acordos_em_aberto=("IS_EM_ABERTO", "sum"),
            acordos_nao_pagou=("IS_NAO_PAGOU", "sum"),
            valor_pago=("VALOR_PAGO", "sum"),
            valor_em_aberto=("VALOR_EM_ABERTO", "sum"),
            valor_nao_pagou=("VALOR_NAO_PAGOU", "sum"),
            valor_negociado=("VALOR_NEGOCIADO", "sum"),
        )
        .reset_index()
    )
    matrix["tx_pagamento"] = safe_div(matrix["pagamentos"], matrix["acordos"])
    matrix["recuperacao"] = safe_div(matrix["valor_pago"], matrix["valor_negociado"])

    metric_choice = st.selectbox(
        "Métrica da matriz",
        ["valor_pago", "valor_em_aberto", "valor_nao_pagou", "tx_pagamento", "recuperacao", "acordos", "pagamentos", "acordos_em_aberto", "acordos_nao_pagou"],
        index=0,
    )
    heat_data = matrix.groupby(["OPERADOR", "FAIXA_ATRASO"]).agg(
        {metric_choice: "sum" if metric_choice in ["valor_pago", "valor_em_aberto", "valor_nao_pagou", "acordos", "pagamentos", "acordos_em_aberto", "acordos_nao_pagou"] else "mean"}
    ).reset_index()
    heatmap(display_fields(heat_data), "FAIXA_ATRASO:N", "OPERADOR:N", f"{metric_choice}:Q", "Operador x faixa de atraso")

    st.subheader("Matriz analítica por operador, faixa e região")
    data_table(matrix.sort_values(["valor_pago", "pagamentos"], ascending=False))

with tabs[6]:
    avg_score = operador_df["score"].mean() if not operador_df.empty else 0
    oportunidades = operador_df[(operador_df["acionamentos"] >= operador_df["acionamentos"].median()) & (operador_df["score"] < avg_score)].sort_values("acionamentos", ascending=False)
    destaques = operador_df[operador_df["score"] >= operador_df["score"].quantile(0.75)].sort_values("score", ascending=False)
    faixas_oportunidade = aggregate_resultados(resultados, "FAIXA_ATRASO")
    faixas_oportunidade = faixas_oportunidade.sort_values(["valor_negociado", "recuperacao"], ascending=[False, True])

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Operadores para priorização")
        data_table(destaques[["OPERADOR", "score", "valor_pago", "tx_pagamento", "recuperacao", "clientes_trabalhados"]].head(10))
    with c2:
        st.subheader("Alto volume com eficiência abaixo da média")
        data_table(oportunidades[["OPERADOR", "score", "acionamentos", "tx_contato", "tx_acordo", "valor_pago"]].head(10))

    st.subheader("Faixas com maior oportunidade de recuperação")
    data_table(faixas_oportunidade[["FAIXA_ATRASO", "clientes", "acordos", "pagamentos", "acordos_em_aberto", "acordos_nao_pagou", "valor_negociado", "valor_pago", "valor_em_aberto", "valor_nao_pagou", "recuperacao"]])
