from pathlib import Path
from html import escape
from io import BytesIO
import math
import os
import re
import unicodedata

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).parent
EVENTOS_FILE = BASE_DIR / "Cobmais-Eventos-908-2026050417.xlsx"
RESULTADOS_FILE = BASE_DIR / "NOVA BASE RESULTADOS 2026.xlsm"
COLABORADORES_FILE = BASE_DIR / "Base de colaboradores.xlsx"
EXCLUDED_OPERATORS = {"samuel.levi"}
EXCLUDED_OPERATOR_PREFIXES = ("mauricio",)
POSTGRES_DEFAULTS = {
    "host": "",
    "port": 5432,
    "database": "",
    "user": "",
    "password": "",
    "schema": "workplan",
    "table": "casos_workplan",
}


def latest_file(pattern):
    files = [p for p in BASE_DIR.glob(pattern) if not p.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo encontrado para o padrão: {pattern}")
    return max(files, key=lambda p: p.stat().st_mtime)


def file_version(path):
    path = Path(path)
    stat = path.stat()
    return (path.name, stat.st_size, stat.st_mtime_ns)


def data_file_versions():
    return (
        file_version(latest_file("Cobmais-Eventos-908-*.xlsx")),
        file_version(latest_file("Pesquisa-Cliente-908-*.xlsx")),
        file_version(RESULTADOS_FILE),
    )


st.set_page_config(
    page_title="Performance Operacional",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    div[data-testid="stMetric"] {
        overflow: visible;
    }
    .metric-card {
        min-height: 96px;
        border: 1px solid rgba(148, 163, 184, 0.24);
        border-radius: 8px;
        padding: 12px 14px;
        background: rgba(15, 23, 42, 0.18);
    }
    .metric-card__label {
        margin-bottom: 8px;
        color: rgba(255, 255, 255, 0.82);
        font-size: 0.82rem;
        font-weight: 700;
        line-height: 1.2;
    }
    .metric-card__value {
        color: #ffffff;
        font-size: 1.48rem;
        font-weight: 650;
        line-height: 1.18;
        white-space: normal;
        overflow-wrap: anywhere;
        word-break: break-word;
    }
    .metric-card--compact .metric-card__value {
        font-size: 1.28rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


CORP_PALETTE = ["#213547", "#2f6f73", "#b7791f", "#6b7280", "#8a3ffc", "#bf616a"]
MONTH_NAMES_PT = {
    1: "JANEIRO",
    2: "FEVEREIRO",
    3: "MARÇO",
    4: "ABRIL",
    5: "MAIO",
    6: "JUNHO",
    7: "JULHO",
    8: "AGOSTO",
    9: "SETEMBRO",
    10: "OUTUBRO",
    11: "NOVEMBRO",
    12: "DEZEMBRO",
}

FIELD_HELP = {
    "OPERADOR": ("Operador", "Negociador responsável pelo evento ou acordo."),
    "FAIXA_ATRASO": ("Faixa de atraso", "Agrupamento do atraso/DPD do contrato em faixas gerenciais."),
    "SEGMENTO_DPD": ("Segmento DPD", "Classificação do contrato pela coluna Y/DPD Formula: POTLOSS 1-720, SALVAGE 721-1440 e SALVAGE + acima de 1440."),
    "prioridade_workplan": ("Prioridade", "Classificação do contrato no Workplan pelo score de recuperação."),
    "score_recuperacao": ("Score recuperação", "Score de priorização futura combinando segmento, valor, histórico de CPC e tempo sem contato. Contratos com pagamento ou acordo em aberto são excluídos da recomendação."),
    "motivo_priorizacao": ("Motivo priorização", "Principais fatores que aumentaram a prioridade do contrato."),
    "cpf_cnpj": ("CPF/CNPJ", "Documento do cliente conforme Workplan."),
    "PRODUTO": ("Produto", "Produto/carteira localizado no histórico de contratos."),
    "total_amount_due": ("Valor em aberto", "Valor total em aberto do contrato no Workplan."),
    "dpd": ("DPD", "Dias de atraso do contrato no Workplan."),
    "dias_sem_contato": ("Dias sem contato", "Dias desde o último contato registrado no Workplan ou último acionamento no histórico."),
    "acionamentos_hist": ("Acionamentos históricos", "Quantidade de acionamentos históricos localizados para o contrato."),
    "cpcs_hist": ("CPCs históricos", "Quantidade de CPCs históricos localizados para o contrato."),
    "acordos_hist": ("Acordos históricos", "Quantidade de acordos históricos localizados para o contrato."),
    "pagamentos_hist": ("Pagamentos históricos", "Quantidade de pagamentos históricos localizados para o contrato."),
    "carteira_elegivel": ("Carteira elegível", "Soma do valor em aberto dos contratos elegíveis no grupo."),
    "contratos_elegiveis": ("Contratos elegíveis", "Quantidade de contratos elegíveis no grupo."),
    "recuperacao_esperada": ("Recuperação esperada", "Carteira elegível ponderada pelas taxas históricas de contato, CPC, acordo, pagamento e percentual recuperado."),
    "recuperacao_esperada_pct_carteira": ("% esperado da carteira", "Recuperação esperada dividida pela carteira elegível."),
    "taxa_contato": ("Taxa acionamento -> contato", "Contatos com cliente divididos pelos acionamentos históricos no grupo."),
    "taxa_cpc": ("Taxa contato -> CPC", "CPCs históricos divididos pelos contatos com cliente no grupo."),
    "taxa_acordo": ("Taxa CPC -> acordo", "Acordos históricos divididos pelos CPCs históricos no grupo."),
    "taxa_pagamento": ("Taxa acordo -> pagamento", "Pagamentos históricos divididos pelos acordos históricos no grupo."),
    "percentual_medio_recuperado": ("% médio recuperado", "Valor pago histórico dividido pelo valor negociado histórico no grupo."),
    "base_taxas": ("Base das taxas", "Indica se as taxas foram calculadas pelo grupo ou pela média geral por baixa amostra."),
    "flag_cobravel": ("Cobravel", "Indica se o contrato está marcado como cobrável no Workplan."),
    "status_cpc": ("Status CPC", "Status de CPC disponível no Workplan."),
    "REGIÃO": ("Região", "Região cadastrada na base de resultados."),
    "clientes": ("Clientes", "Quantidade distinta de contratos/clientes no agrupamento."),
    "clientes_trabalhados": ("Clientes trabalhados", "Quantidade distinta de contratos acionados pelo operador."),
    "acionamentos": ("Acionamentos", "Total de eventos válidos, sem AUTO/importação por padrão."),
    "contatos_efetivos": ("Contatos efetivos", "Mesmo critério de CPC: eventos iniciados por 02, 03, 04 ou 05."),
    "contatos_cliente": ("Contatos cliente", "Eventos iniciados por 02 ou 03, contato direto com cliente."),
    "cpcs": ("CPCs", "Eventos produtivos iniciados por 02, 03, 04 ou 05."),
    "cpcs_unicos": ("CPCs únicos", "Contratos distintos com pelo menos um CPC por operador. Remove acionamentos repetidos do mesmo contrato."),
    "clientes_cpc": ("Clientes com CPC", "Contratos distintos que tiveram ao menos um CPC."),
    "contratos_cpc": ("Contratos com CPC", "Contratos distintos que tiveram ao menos um CPC."),
    "acordos": ("Acordos", "Quantidade de acordos localizados na base de resultados."),
    "pagamentos": ("Pagamentos", "Acordos com status PAGOU ou data de pagamento preenchida."),
    "acordos_sem_pagamento": ("Acordos sem pagamento", "Acordos sem status pago e sem data de pagamento."),
    "acordos_em_aberto": ("Acordos em aberto", "Acordos com status EM ABERTO ou status vazio sem data de pagamento."),
    "acordos_nao_pagou": ("Acordos não pagos", "Acordos com status NÃO PAGOU na base de resultados."),
    "pct_quebra": ("% quebras", "Acordos não pagos divididos pelo total de acordos."),
    "tx_contato": ("Taxa de CPC", "CPCs divididos pelo total de acionamentos."),
    "tx_acordo": ("Taxa CPC -> acordo", "Acordos divididos pelo total de CPCs do operador."),
    "tx_acordo_cliente_cpc": ("Taxa contrato CPC -> acordo", "Acordos divididos pelos contratos distintos com CPC."),
    "tx_pagamento": ("Taxa acordo -> pagamento", "Pagamentos divididos por acordos."),
    "efetividade_pagamento": ("Efetividade pagamento", "Pagamentos divididos por pagamentos mais acordos quebrados/não pagos."),
    "tx_pagamento_cpc": ("Taxa CPC -> pagamento", "Pagamentos divididos pelo total de CPCs do operador."),
    "tx_sem_pagamento": ("Taxa sem pagamento", "Acordos sem pagamento divididos pelo total de acordos."),
    "tx_cpc_acordo": ("Taxa CPC -> acordo", "Contratos com CPC que geraram acordo, divididos pelos contratos com CPC."),
    "tx_cpc_pagamento": ("Taxa CPC -> pagamento", "Contratos com CPC que geraram pagamento, divididos pelos contratos com CPC."),
    "tx_cpc_unico_acordo": ("Taxa CPC único -> acordo", "Contratos únicos com CPC que geraram acordo, divididos pelos contratos únicos com CPC."),
    "tx_cpc_unico_pagamento": ("Taxa CPC único -> pagamento", "Contratos únicos com CPC que geraram pagamento, divididos pelos contratos únicos com CPC."),
    "tx_acordo_pagamento": ("Taxa acordo -> pagamento", "Pagamentos divididos pelos acordos originados em contratos com CPC."),
    "tx_acordo_sem_pagamento": ("Taxa acordo sem pagamento", "Acordos sem pagamento divididos pelos acordos originados em contratos com CPC."),
    "valor_negociado": ("Valor negociado", "Soma da coluna VALOR DO BANCO - META da base de resultados."),
    "valor_pago": ("Valor recebido", "Valor negociado dos acordos pagos; acordos não pagos entram como R$ 0,00."),
    "valor_em_aberto": ("Valor em aberto", "Valor negociado dos acordos com status EM ABERTO ou status vazio sem data de pagamento."),
    "valor_nao_pagou": ("Valor não pago", "Valor negociado dos acordos com status NÃO PAGOU."),
    "valor_quebra": ("Valor quebras", "Valor negociado dos acordos não pagos."),
    "ticket_medio": ("Ticket médio", "Valor negociado médio dos acordos."),
    "recuperacao": ("% recuperação", "Valor recebido dividido pelo valor negociado."),
    "score": ("Score", "Índice composto que pondera contato, acordo, pagamento, valor recebido e volume."),
    "meta_individual": ("Meta individual", "Meta mensal do negociador: R$ 150 mil, ou R$ 300 mil para Ana Karolina e Luiz Mauro."),
    "atingimento_meta_individual": ("% meta individual", "Valor recebido dividido pela meta individual do negociador."),
    "pct_aberto_meta_individual": ("% aberto/meta individual", "Valor em aberto dividido pela meta individual do negociador."),
    "saldo_meta_individual": ("Saldo meta individual", "Valor recebido menos meta individual. Negativo indica falta para bater meta."),
    "meta_geral_escritorio": ("Meta geral escritório", "Meta geral do escritório para o mês, lida na aba METAS."),
    "participacao_meta_geral": ("% meta geral", "Quanto o operador contribuiu para a meta geral do escritório."),
    "quartil_meta_individual": ("Quartil meta individual", "Quartil do atingimento da meta individual. Q4 é o melhor grupo."),
    "quartil_meta_geral": ("Quartil meta geral", "Quartil da participação na meta geral do escritório. Q4 é o melhor grupo."),
    "diagnostico_meta": ("Diagnóstico meta", "Leitura gerencial combinando atingimento individual e contribuição na meta geral."),
    "nome_colaborador": ("Nome colaborador", "Nome completo do colaborador conforme Base de colaboradores."),
    "base_colaborador": ("Base colaborador", "Aba/carteira da Base de colaboradores onde o login foi encontrado."),
    "cargo_colaborador": ("Cargo", "Cargo do colaborador conforme Base de colaboradores."),
    "negociador_cadastrado": ("Cadastro colaborador", "Indica se o operador está cadastrado como negociador na Base de colaboradores."),
}


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_status(value):
    text = normalize_text(value).upper()
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


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


def is_excluded_operator(value):
    text = normalize_operator(value)
    if pd.isna(text):
        return False
    return text in EXCLUDED_OPERATORS or any(text.startswith(prefix) for prefix in EXCLUDED_OPERATOR_PREFIXES)


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


def segmento_dpd(days):
    if pd.isna(days):
        return "Sem DPD"
    text = normalize_text(days).upper()
    if text:
        compact = re.sub(r"[^A-Z0-9+]", "", text)
        if compact == "POTLOSS":
            return "POTLOSS"
        if compact == "SALVAGE":
            return "SALVAGE"
        if compact in {"SALVAGE+", "SALVAGEPLUS"}:
            return "SALVAGE +"
    numeric_days = pd.to_numeric(days, errors="coerce")
    if pd.isna(numeric_days):
        return "Sem DPD"
    if numeric_days < 1:
        return "Sem DPD"
    if numeric_days <= 720:
        return "POTLOSS"
    if numeric_days <= 1440:
        return "SALVAGE"
    return "SALVAGE +"


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


def scalar_safe_div(num, den):
    return 0 if pd.isna(den) or den == 0 else num / den


def streamlit_secret_section(name):
    try:
        return st.secrets.get(name, {})
    except Exception:
        return {}


def streamlit_secret_value(name, default=""):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def postgres_config():
    secrets = streamlit_secret_section("postgres")
    database_url = (
        os.getenv("SUPABASE_DB_URL")
        or os.getenv("DATABASE_URL")
        or streamlit_secret_value("SUPABASE_DB_URL")
        or streamlit_secret_value("supabase_db_url")
        or streamlit_secret_value("DATABASE_URL")
        or streamlit_secret_value("database_url")
        or secrets.get("database_url", secrets.get("url", ""))
    )
    has_connection_parts = any(
        [
            os.getenv("PGHOST"),
            os.getenv("PGDATABASE"),
            os.getenv("PGUSER"),
            os.getenv("PGPASSWORD"),
            secrets.get("host"),
            secrets.get("database"),
            secrets.get("user"),
            secrets.get("password"),
        ]
    )
    return {
        "database_url": database_url,
        "configured": bool(database_url or has_connection_parts),
        "host": os.getenv("PGHOST", secrets.get("host", POSTGRES_DEFAULTS["host"])),
        "port": int(os.getenv("PGPORT", secrets.get("port", POSTGRES_DEFAULTS["port"]))),
        "database": os.getenv("PGDATABASE", secrets.get("database", POSTGRES_DEFAULTS["database"])),
        "user": os.getenv("PGUSER", secrets.get("user", POSTGRES_DEFAULTS["user"])),
        "password": os.getenv("PGPASSWORD", secrets.get("password", POSTGRES_DEFAULTS["password"])),
        "schema": os.getenv("PGSCHEMA", secrets.get("schema", POSTGRES_DEFAULTS["schema"])),
        "table": os.getenv("PGTABLE", secrets.get("table", POSTGRES_DEFAULTS["table"])),
    }


def quartile_label(series, higher_is_better=True, min_series=None, min_value=1):
    values = pd.to_numeric(series, errors="coerce")
    valid = values.notna()
    if min_series is not None:
        valid = valid & (pd.to_numeric(min_series, errors="coerce").fillna(0) >= min_value)

    labels = pd.Series("Sem base", index=series.index, dtype="object")
    if valid.sum() == 0:
        return labels

    ranked = values[valid].rank(method="average", pct=True)
    if not higher_is_better:
        ranked = 1 - ranked + (1 / valid.sum())

    labels.loc[ranked[ranked <= 0.25].index] = "Q1 - crítico"
    labels.loc[ranked[(ranked > 0.25) & (ranked <= 0.50)].index] = "Q2 - atenção"
    labels.loc[ranked[(ranked > 0.50) & (ranked <= 0.75)].index] = "Q3 - bom"
    labels.loc[ranked[ranked > 0.75].index] = "Q4 - destaque"
    return labels


def selected_months(resultados):
    meses = resultados[["MES_RESULTADO", "MES_NUM"]].replace("", np.nan).dropna(subset=["MES_RESULTADO"]).drop_duplicates()
    return meses.sort_values(["MES_NUM", "MES_RESULTADO"])["MES_RESULTADO"].tolist()


@st.cache_data(show_spinner=False)
def load_collaborators():
    if not COLABORADORES_FILE.exists():
        return pd.DataFrame(columns=["OPERADOR", "nome_colaborador", "base_colaborador", "cargo_colaborador"])

    sheets = pd.read_excel(COLABORADORES_FILE, sheet_name=None)
    frames = []
    for sheet_name, df in sheets.items():
        df.columns = [normalize_text(c).upper() for c in df.columns]
        if "LOGIN COBMAIS" not in df.columns:
            continue
        if "ATIVO" in df.columns:
            df = df[df["ATIVO"].map(normalize_text).str.upper().eq("SIM")].copy()
        base = pd.DataFrame(
            {
                "OPERADOR": df["LOGIN COBMAIS"].map(normalize_operator),
                "nome_colaborador": df.get("NOME COLABORADOR", pd.Series(index=df.index, dtype=object)).map(normalize_text),
                "base_colaborador": sheet_name,
                "cargo_colaborador": df.get("CARGO", pd.Series(index=df.index, dtype=object)).map(normalize_text).str.upper(),
            }
        )
        base = base[base["OPERADOR"].notna()]
        base = base[~base["OPERADOR"].isin(["escritório", "escritorio"])]
        base = base[~base["OPERADOR"].map(is_excluded_operator)]
        frames.append(base)

    if not frames:
        return pd.DataFrame(columns=["OPERADOR", "nome_colaborador", "base_colaborador", "cargo_colaborador"])

    colaboradores = pd.concat(frames, ignore_index=True).drop_duplicates(["OPERADOR", "base_colaborador"])
    judicial = colaboradores[colaboradores["base_colaborador"].eq("JUDICIAL")].copy()
    if judicial.empty:
        judicial = colaboradores[colaboradores["cargo_colaborador"].eq("NEGOCIADOR")].copy()
    return judicial.drop_duplicates("OPERADOR")


@st.cache_data(show_spinner=False)
def load_office_goals():
    metas = pd.read_excel(RESULTADOS_FILE, sheet_name="METAS", header=None)
    header_idx = metas.index[metas.iloc[:, 0].astype(str).str.strip().str.upper().eq("NEGOCIADOR")]
    if len(header_idx) == 0:
        return {}

    header_row = header_idx[0]
    headers = metas.iloc[header_row].map(normalize_text).str.upper().tolist()
    totals = metas.iloc[header_row + 1:].copy()
    total_rows = totals[totals.iloc[:, 0].astype(str).str.strip().str.upper().eq("TOTAL")]
    if total_rows.empty:
        return {}

    total_row = total_rows.iloc[0]
    goals = {}
    month_abbr = {
        "JANEIRO": "JAN",
        "FEVEREIRO": "FEV",
        "MARÇO": "MAR",
        "ABRIL": "ABR",
        "MAIO": "MAI",
        "JUNHO": "JUN",
        "JULHO": "JUL",
        "AGOSTO": "AGO",
        "SETEMBRO": "SET",
        "OUTUBRO": "OUT",
        "NOVEMBRO": "NOV",
        "DEZEMBRO": "DEZ",
    }
    for month, abbr in month_abbr.items():
        target_col = None
        for i, header in enumerate(headers):
            if header == f"META {abbr}":
                target_col = i
                break
        if target_col is not None:
            goals[month] = pd.to_numeric(total_row.iloc[target_col], errors="coerce")
    return {k: float(v) for k, v in goals.items() if pd.notna(v)}


@st.cache_data(show_spinner=False)
def load_office_received():
    metas = pd.read_excel(RESULTADOS_FILE, sheet_name="METAS", header=None)
    header_idx = metas.index[metas.iloc[:, 0].astype(str).str.strip().str.upper().eq("NEGOCIADOR")]
    if len(header_idx) == 0:
        return {}

    header_row = header_idx[0]
    headers = metas.iloc[header_row].map(normalize_text).str.upper().tolist()
    totals = metas.iloc[header_row + 1:].copy()
    total_rows = totals[totals.iloc[:, 0].astype(str).str.strip().str.upper().eq("TOTAL")]
    if total_rows.empty:
        return {}

    total_row = total_rows.iloc[0]
    received = {}
    months = [
        "JANEIRO",
        "FEVEREIRO",
        "MARÇO",
        "ABRIL",
        "MAIO",
        "JUNHO",
        "JULHO",
        "AGOSTO",
        "SETEMBRO",
        "OUTUBRO",
        "NOVEMBRO",
        "DEZEMBRO",
    ]
    for month in months:
        target_col = None
        for i, header in enumerate(headers):
            if header == month:
                target_col = i
                break
        if target_col is not None:
            received[month] = pd.to_numeric(total_row.iloc[target_col], errors="coerce")
    return {k: float(v) for k, v in received.items() if pd.notna(v)}


def build_meta_analysis(operador_df, resultados, operadores_scope=None):
    metas_gerais = load_office_goals()
    colaboradores = load_collaborators()
    if operadores_scope:
        colaboradores = colaboradores[colaboradores["OPERADOR"].isin(operadores_scope)].copy()
    meses = selected_months(resultados)
    meses_count = max(len(meses), 1)
    meta_geral = sum(metas_gerais.get(mes, 0) for mes in meses)

    operadores_base = colaboradores[["OPERADOR"]].drop_duplicates()
    df = operadores_base.merge(operador_df, on="OPERADOR", how="left")
    df = df.merge(colaboradores, on="OPERADOR", how="left")
    metric_cols = [
        "acionamentos",
        "clientes_trabalhados",
        "contatos_efetivos",
        "contatos_cliente",
        "cpcs",
        "clientes_cpc",
        "acordos",
        "pagamentos",
        "acordos_sem_pagamento",
        "acordos_em_aberto",
        "acordos_nao_pagou",
        "pct_quebra",
        "valor_negociado",
        "valor_pago",
        "valor_em_aberto",
        "valor_nao_pagou",
        "valor_quebra",
        "ticket_medio",
        "tx_contato",
        "tx_acordo",
        "tx_acordo_cliente_cpc",
        "tx_pagamento",
        "efetividade_pagamento",
        "tx_pagamento_cpc",
        "tx_sem_pagamento",
        "recuperacao",
        "score",
    ]
    for col in metric_cols:
        if col in df:
            df[col] = df[col].fillna(0)
    df["negociador_cadastrado"] = np.where(df["nome_colaborador"].notna(), "Sim", "Não")
    df["nome_colaborador"] = df["nome_colaborador"].fillna("Não localizado na base")
    df["base_colaborador"] = df["base_colaborador"].fillna("Fora da base")
    df["cargo_colaborador"] = df["cargo_colaborador"].fillna("Não localizado")

    pos_retomado = {"ana.karolina.oliveira", "luiz.mauro"}
    df["meta_individual"] = np.where(df["OPERADOR"].isin(pos_retomado), 300000, 150000) * meses_count
    df["meta_geral_escritorio"] = meta_geral
    df["atingimento_meta_individual"] = safe_div(df["valor_pago"], df["meta_individual"])
    df["pct_aberto_meta_individual"] = safe_div(df["valor_em_aberto"], df["meta_individual"])
    df["saldo_meta_individual"] = df["valor_pago"] - df["meta_individual"]
    df["participacao_meta_geral"] = safe_div(df["valor_pago"], df["meta_geral_escritorio"])
    df["quartil_meta_individual"] = quartile_label(df["atingimento_meta_individual"], min_series=df["meta_individual"], min_value=1)
    df["quartil_meta_geral"] = quartile_label(df["participacao_meta_geral"], min_series=df["meta_individual"], min_value=1)
    df["diagnostico_meta"] = np.select(
        [
            df["atingimento_meta_individual"] >= 1,
            df["atingimento_meta_individual"] >= 0.75,
            df["atingimento_meta_individual"] >= 0.50,
        ],
        ["Meta batida", "Próximo da meta", "Atenção"],
        default="Crítico",
    )
    return df.sort_values(["atingimento_meta_individual", "valor_pago"], ascending=False), meses, meta_geral


def meta_operator_groups(df):
    metrics = [
        ("Atingimento da meta individual", "quartil_meta_individual"),
        ("Participação na meta geral", "quartil_meta_geral"),
    ]
    groups = ["Q4 - destaque", "Q3 - bom", "Q2 - atenção", "Q1 - crítico", "Sem base"]
    rows = []
    for metric_label, metric_col in metrics:
        for group in groups:
            operadores = sorted(df.loc[df[metric_col].eq(group), "OPERADOR"].dropna().astype(str).tolist())
            rows.append(
                {
                    "Métrica": metric_label,
                    "Grupo": group,
                    "Qtd. operadores": len(operadores),
                    "Operadores": ", ".join(operadores) if operadores else "-",
                }
            )
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def load_workplan():
    try:
        import psycopg2
    except ImportError:
        return pd.DataFrame(), "Driver psycopg2-binary não instalado."

    cfg = postgres_config()
    if not cfg["configured"]:
        return (
            pd.DataFrame(),
            "Workplan nao configurado. Configure SUPABASE_DB_URL/DATABASE_URL ou PGHOST/PGDATABASE/PGUSER/PGPASSWORD nas variaveis de ambiente ou em st.secrets.",
        )

    query = f"""
        SELECT
            agreement_no,
            cust_name,
            cpf_cnpj,
            dpd,
            total_amount_due,
            last_contact_date,
            allocation_date,
            last_marking_date,
            city,
            state,
            last_marking_value,
            pct_of_margin_money,
            no_first_ins_unpaid,
            status,
            flag_cobravel,
            status_base,
            flag_cpc,
            status_cpc,
            probabilidade,
            faixa_atraso,
            regiao,
            uf
        FROM "{cfg['schema']}"."{cfg['table']}"
    """
    try:
        if cfg["database_url"]:
            conn = psycopg2.connect(cfg["database_url"])
        else:
            conn = psycopg2.connect(
                host=cfg["host"],
                port=cfg["port"],
                dbname=cfg["database"],
                user=cfg["user"],
                password=cfg["password"],
            )
        with conn:
            df = pd.read_sql_query(query, conn)
    except Exception as exc:
        host = str(cfg.get("host", "")).lower()
        message = str(exc)
        if host in {"localhost", "127.0.0.1", "::1"} and "connection refused" in message.lower():
            return (
                pd.DataFrame(),
                "PostgreSQL local nao esta acessivel a partir deste ambiente. Em deploy, localhost aponta para o servidor do Streamlit; configure SUPABASE_DB_URL ou DATABASE_URL com a connection string do Supabase.",
            )
        return pd.DataFrame(), f"Não foi possível carregar o Workplan: {exc}"
    finally:
        if "conn" in locals():
            conn.close()

    df.columns = [normalize_text(c).lower() for c in df.columns]
    df["CONTRATO_KEY"] = df["agreement_no"].map(normalize_contract)
    df["cpf_cnpj"] = df["cpf_cnpj"].map(normalize_text)
    df["dpd"] = pd.to_numeric(df["dpd"], errors="coerce")
    df["total_amount_due"] = pd.to_numeric(df["total_amount_due"], errors="coerce").fillna(0)
    df["pct_of_margin_money"] = pd.to_numeric(df["pct_of_margin_money"], errors="coerce")
    df["no_first_ins_unpaid"] = pd.to_numeric(df["no_first_ins_unpaid"], errors="coerce")
    for col in ["last_contact_date", "allocation_date", "last_marking_date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    df["SEGMENTO_DPD"] = df["dpd"].map(segmento_dpd)
    df["FAIXA_ATRASO"] = df["faixa_atraso"].map(normalize_text)
    df["FAIXA_ATRASO"] = df["FAIXA_ATRASO"].where(df["FAIXA_ATRASO"].ne(""), df["dpd"].map(atraso_faixa))
    df["REGIÃO"] = df["regiao"].map(normalize_text)
    df["flag_cobravel"] = df["flag_cobravel"].map(normalize_text).str.upper()
    df["flag_cpc"] = df["flag_cpc"].map(normalize_text).str.upper()
    df["status_cpc"] = df["status_cpc"].map(normalize_text)
    df["status_base"] = df["status_base"].map(normalize_text)
    return df[df["CONTRATO_KEY"].notna()].copy(), None


def build_workplan_analysis(workplan, eventos_hist, resultados_hist):
    if workplan.empty:
        return workplan

    eventos_contrato = eventos_hist.groupby("CONTRATO_KEY", dropna=True).agg(
        acionamentos_hist=("EVENTO_TXT", "size"),
        cpcs_hist=("IS_CPC", "sum"),
        ultimo_acionamento=("DATA", "max"),
    ).reset_index()
    eventos_perfil_cols = ["CONTRATO_KEY"] + [col for col in ["PRODUTO"] if col in eventos_hist.columns]
    eventos_perfil = eventos_hist[eventos_perfil_cols].dropna(subset=["CONTRATO_KEY"]).drop_duplicates("CONTRATO_KEY")
    resultados_contrato = resultados_hist.groupby("CONTRATO_KEY", dropna=True).agg(
        acordos_hist=("CONTRATO_KEY", "count"),
        pagamentos_hist=("IS_PAGO", "sum"),
        acordos_em_aberto_hist=("IS_EM_ABERTO", "sum"),
        acordos_nao_pagou_hist=("IS_NAO_PAGOU", "sum"),
        valor_negociado_hist=("VALOR_NEGOCIADO", "sum"),
        valor_pago_hist=("VALOR_PAGO", "sum"),
    ).reset_index()

    df = workplan.merge(eventos_contrato, on="CONTRATO_KEY", how="left")
    df = df.merge(eventos_perfil, on="CONTRATO_KEY", how="left")
    df = df.merge(resultados_contrato, on="CONTRATO_KEY", how="left")
    fill_zero = [
        "acionamentos_hist",
        "cpcs_hist",
        "acordos_hist",
        "pagamentos_hist",
        "acordos_em_aberto_hist",
        "acordos_nao_pagou_hist",
        "valor_negociado_hist",
        "valor_pago_hist",
    ]
    for col in fill_zero:
        df[col] = df[col].fillna(0)

    ultimo_contato = df[["last_contact_date", "ultimo_acionamento"]].max(axis=1)
    hoje = pd.Timestamp.today().normalize()
    df["dias_sem_contato"] = (hoje - ultimo_contato).dt.days
    df["dias_sem_contato"] = df["dias_sem_contato"].fillna(999).clip(lower=0)

    segmento_score = df["SEGMENTO_DPD"].map({"POTLOSS": 0.25, "SALVAGE": 0.16, "SALVAGE +": 0.08}).fillna(0)
    valor_score = pd.to_numeric(df["total_amount_due"], errors="coerce").rank(pct=True).fillna(0) * 0.22
    cpc_score = np.where((df["cpcs_hist"] > 0) | df["flag_cpc"].eq("SIM"), 0.16, 0)
    contato_score = np.select(
        [
            df["dias_sem_contato"] >= 30,
            df["dias_sem_contato"] >= 15,
            df["dias_sem_contato"] >= 7,
        ],
        [0.20, 0.14, 0.08],
        default=0.03,
    )
    cobravel_score = np.where(df["flag_cobravel"].eq("SIM"), 0.05, -0.20)
    df["score_recuperacao"] = (segmento_score + valor_score + cpc_score + contato_score + cobravel_score).clip(0, 1)
    df["prioridade_workplan"] = np.select(
        [
            df["score_recuperacao"] >= 0.70,
            df["score_recuperacao"] >= 0.45,
        ],
        ["Alta", "Média"],
        default="Baixa",
    )

    motivos = []
    for _, row in df.iterrows():
        parts = []
        if row["SEGMENTO_DPD"] in {"POTLOSS", "SALVAGE"}:
            parts.append(row["SEGMENTO_DPD"])
        if row["total_amount_due"] >= df["total_amount_due"].quantile(0.75):
            parts.append("alto valor")
        if row["cpcs_hist"] > 0 or row["flag_cpc"] == "SIM":
            parts.append("histórico de CPC")
        if row["dias_sem_contato"] >= 15:
            parts.append("sem contato recente")
        motivos.append(", ".join(parts) if parts else "baixo sinal histórico")
    df["motivo_priorizacao"] = motivos
    return df.sort_values(["score_recuperacao", "total_amount_due"], ascending=False)


@st.cache_data(show_spinner=False)
def load_data(data_version):
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
    eventos = eventos[~eventos["OPERADOR"].map(is_excluded_operator)].copy()
    eventos["DATA"] = pd.to_datetime(eventos["DATA"], dayfirst=True, errors="coerce")
    eventos["EVENTO_TXT"] = eventos["EVENTO"].map(normalize_text)
    eventos["EVENTO_UPPER"] = eventos["EVENTO_TXT"].str.upper()
    eventos["DESCRICAO_UPPER"] = eventos["DESCRIÇÃO"].map(normalize_text).str.upper()
    eventos["TIPO DE ACIONAMENTO"] = eventos["TIPO DE ACIONAMENTO"].map(normalize_text)

    contratos["CONTRATO_KEY"] = contratos["CONTRATO"].map(normalize_contract)
    contratos["ATRASO"] = pd.to_numeric(contratos["ATRASO"], errors="coerce")
    contratos["TOTAL ABERTO"] = pd.to_numeric(contratos["TOTAL ABERTO"], errors="coerce")
    contratos["FAIXA_ATRASO"] = contratos["ATRASO"].map(atraso_faixa)
    contratos["SEGMENTO_DPD"] = contratos["ATRASO"].map(segmento_dpd)

    contrato_cols = [
        "CONTRATO_KEY",
        "PRODUTO",
        "REGIAO",
        "FILIAL",
        "ESTAGIO",
        "ATRASO",
        "TOTAL ABERTO",
        "FAIXA_ATRASO",
        "SEGMENTO_DPD",
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
    resultados = resultados[~resultados["OPERADOR"].map(is_excluded_operator)].copy()
    resultados["DATA_ACORDO"] = pd.to_datetime(resultados["EMISSÃO"], errors="coerce")
    resultados["DATA_PAGAMENTO"] = pd.to_datetime(resultados["DATA DO PAGAMENTO"], errors="coerce")
    resultados["DATA_VENCIMENTO"] = pd.to_datetime(resultados["DATA DE VENCIMENTO"], errors="coerce")
    resultados["VALOR_NEGOCIADO"] = pd.to_numeric(resultados["VALOR DO BANCO - META"], errors="coerce").fillna(0)
    resultados["HONORARIOS"] = pd.to_numeric(resultados["HONORÁRIOS %"], errors="coerce").fillna(0)
    resultados["DPD"] = pd.to_numeric(resultados["DPD"], errors="coerce")
    resultados["FAIXA_ATRASO"] = resultados["DPD"].map(atraso_faixa)
    resultados["SEGMENTO_DPD"] = resultados["DPD FORMULA"].map(segmento_dpd)
    resultados["REGIÃO"] = resultados["REGIÃO"].fillna(resultados.get("UF", "Sem região")).map(normalize_text)
    resultados["UF"] = resultados["UF"].map(normalize_text)
    resultados["CAMPANHA"] = resultados["CAMPANHA"].map(normalize_text)
    resultados["TIPO DE ACORDO"] = resultados["TIPO DE ACORDO"].map(normalize_text)
    resultados["STATUS"] = resultados["STATUS"].map(normalize_text).str.upper()
    resultados["STATUS_KEY"] = resultados["STATUS"].map(normalize_status)
    resultados["IS_ACORDO"] = resultados["CONTRATO_KEY"].notna() & resultados["OPERADOR"].notna()
    resultados["IS_EM_ABERTO"] = resultados["STATUS_KEY"].eq("EM ABERTO") | (
        resultados["STATUS_KEY"].eq("") & resultados["DATA_PAGAMENTO"].isna()
    )
    resultados["IS_NAO_PAGOU"] = resultados["STATUS_KEY"].eq("NAO PAGOU")
    resultados["IS_PAGO"] = resultados["STATUS_KEY"].eq("PAGOU") | (resultados["DATA_PAGAMENTO"].notna())
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
    segmentos_dpd = ["POTLOSS", "SALVAGE", "SALVAGE +"]
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
        mes_atual = MONTH_NAMES_PT.get(pd.Timestamp.today().month)
        if mes_atual in meses:
            mes_padrao = [mes_atual]
        else:
            meses_validos = meses_df[meses_df["MES_NUM"] <= pd.Timestamp.today().month]
            base_mes = meses_validos if not meses_validos.empty else meses_df
            mes_padrao = [base_mes.sort_values(["MES_NUM", "MES_RESULTADO"]).iloc[-1]["MES_RESULTADO"]]

    operador_sel = st.sidebar.multiselect("Operador", operadores)
    mes_sel = st.sidebar.multiselect("Mês do resultado", meses, default=mes_padrao)
    regiao_sel = st.sidebar.multiselect("Região", regioes)
    faixa_sel = st.sidebar.multiselect("Faixa de atraso", faixas)
    segmento_dpd_sel = st.sidebar.multiselect("Segmento DPD", segmentos_dpd)
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
    if segmento_dpd_sel:
        eventos = eventos[eventos["SEGMENTO_DPD"].isin(segmento_dpd_sel)]
        resultados = resultados[resultados["SEGMENTO_DPD"].isin(segmento_dpd_sel)]
    if campanha_sel:
        resultados = resultados[resultados["CAMPANHA"].isin(campanha_sel)]
        contratos_campanha = set(resultados["CONTRATO_KEY"].dropna())
        eventos = eventos[eventos["CONTRATO_KEY"].isin(contratos_campanha)]
    if produto_sel and "PRODUTO" in eventos.columns:
        eventos = eventos[eventos["PRODUTO"].isin(produto_sel)]
    if not incluir_auto:
        eventos = eventos[eventos["IS_ACIONAMENTO"]]

    return eventos, resultados, operador_sel


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
    df["efetividade_pagamento"] = safe_div(df["pagamentos"], df["pagamentos"] + df["acordos_nao_pagou"])
    df["tx_pagamento_cpc"] = safe_div(df["pagamentos"], df["cpcs"])
    df["tx_sem_pagamento"] = safe_div(df["acordos_sem_pagamento"], df["acordos"])
    df["pct_quebra"] = safe_div(df["acordos_nao_pagou"], df["acordos"])
    df["valor_quebra"] = df["valor_nao_pagou"]
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
    ev["cpcs_unicos"] = ev["contratos_cpc"]
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
    df["tx_cpc_unico_acordo"] = safe_div(df["acordos"], df["cpcs_unicos"])
    df["tx_cpc_unico_pagamento"] = safe_div(df["pagamentos"], df["cpcs_unicos"])
    df["tx_acordo_pagamento"] = safe_div(df["pagamentos"], df["acordos"])
    df["efetividade_pagamento"] = safe_div(df["pagamentos"], df["pagamentos"] + df["acordos_nao_pagou"])
    df["tx_acordo_sem_pagamento"] = safe_div(df["acordos_sem_pagamento"], df["acordos"])
    df["pct_quebra"] = safe_div(df["acordos_nao_pagou"], df["acordos"])
    df["valor_quebra"] = df["valor_nao_pagou"]
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
    df["efetividade_pagamento"] = safe_div(df["pagamentos"], df["pagamentos"] + df["acordos_nao_pagou"])
    df["pct_quebra"] = safe_div(df["acordos_nao_pagou"], df["acordos"])
    df["valor_quebra"] = df["valor_nao_pagou"]
    df["recuperacao"] = safe_div(df["valor_pago"], df["valor_negociado"])
    return df


def group_label(df, columns):
    cols = [columns] if isinstance(columns, str) else list(columns)
    cols = [col for col in cols if col in df.columns]
    if not cols:
        return pd.Series("Sem grupo", index=df.index, dtype="object")
    labels = df[cols].apply(lambda row: " | ".join(normalize_text(value) or "Sem grupo" for value in row), axis=1)
    return labels.replace("", "Sem grupo")


def valid_group_mask(df, columns):
    cols = [columns] if isinstance(columns, str) else list(columns)
    cols = [col for col in cols if col in df.columns]
    if not cols:
        return pd.Series(False, index=df.index)
    mask = pd.Series(True, index=df.index)
    for col in cols:
        mask = mask & df[col].map(normalize_text).ne("")
    return mask


def add_projection_group(df, columns):
    grouped = df[valid_group_mask(df, columns)].copy()
    if grouped.empty:
        return grouped
    grouped["grupo"] = group_label(grouped, columns)
    return grouped


def expected_recovery_by_group(workplan_view, eventos_hist, resultados_hist, workplan_col, eventos_col, resultados_col):
    workplan_cols = [workplan_col] if isinstance(workplan_col, str) else list(workplan_col)
    if workplan_view.empty or not any(col in workplan_view.columns for col in workplan_cols):
        return pd.DataFrame()

    carteira = add_projection_group(workplan_view, workplan_col)
    if carteira.empty:
        return pd.DataFrame()
    carteira_df = (
        carteira.groupby("grupo", dropna=False)
        .agg(
            contratos_elegiveis=("CONTRATO_KEY", "nunique"),
            carteira_elegivel=("total_amount_due", "sum"),
        )
        .reset_index()
    )

    eventos_base = eventos_hist[eventos_hist["IS_ACIONAMENTO"]].copy()
    eventos_cols = [eventos_col] if isinstance(eventos_col, str) else list(eventos_col)
    if any(col in eventos_base.columns for col in eventos_cols):
        eventos_base = eventos_base[valid_group_mask(eventos_base, eventos_col)].copy()
        eventos_base["grupo"] = group_label(eventos_base, eventos_col)
        eventos_df = (
            eventos_base.groupby("grupo", dropna=False)
            .agg(
                acionamentos_hist=("EVENTO_TXT", "size"),
                contatos_cliente_hist=("IS_CONTATO_CLIENTE", "sum"),
                cpcs_hist=("IS_CPC", "sum"),
            )
            .reset_index()
        )
    else:
        eventos_df = pd.DataFrame(columns=["grupo", "acionamentos_hist", "contatos_cliente_hist", "cpcs_hist"])

    resultados_cols = [resultados_col] if isinstance(resultados_col, str) else list(resultados_col)
    if any(col in resultados_hist.columns for col in resultados_cols):
        resultados_base = resultados_hist[valid_group_mask(resultados_hist, resultados_col)].copy()
        resultados_base["grupo"] = group_label(resultados_base, resultados_col)
        resultados_df = (
            resultados_base.groupby("grupo", dropna=False)
            .agg(
                acordos_hist=("CONTRATO_KEY", "count"),
                pagamentos_hist=("IS_PAGO", "sum"),
                valor_negociado_hist=("VALOR_NEGOCIADO", "sum"),
                valor_pago_hist=("VALOR_PAGO", "sum"),
            )
            .reset_index()
        )
    else:
        resultados_df = pd.DataFrame(columns=["grupo", "acordos_hist", "pagamentos_hist", "valor_negociado_hist", "valor_pago_hist"])

    df = carteira_df.merge(eventos_df, on="grupo", how="left").merge(resultados_df, on="grupo", how="left")
    fill_cols = [
        "acionamentos_hist",
        "contatos_cliente_hist",
        "cpcs_hist",
        "acordos_hist",
        "pagamentos_hist",
        "valor_negociado_hist",
        "valor_pago_hist",
    ]
    for col in fill_cols:
        df[col] = df[col].fillna(0)

    global_rates = {
        "taxa_contato": scalar_safe_div(eventos_base["IS_CONTATO_CLIENTE"].sum(), len(eventos_base)),
        "taxa_cpc": scalar_safe_div(eventos_base["IS_CPC"].sum(), eventos_base["IS_CONTATO_CLIENTE"].sum()),
        "taxa_acordo": scalar_safe_div(len(resultados_hist), eventos_base["IS_CPC"].sum()),
        "taxa_pagamento": scalar_safe_div(resultados_hist["IS_PAGO"].sum(), len(resultados_hist)),
        "percentual_medio_recuperado": scalar_safe_div(resultados_hist["VALOR_PAGO"].sum(), resultados_hist["VALOR_NEGOCIADO"].sum()),
    }

    df["taxa_contato_grupo"] = safe_div(df["contatos_cliente_hist"], df["acionamentos_hist"])
    df["taxa_cpc_grupo"] = safe_div(df["cpcs_hist"], df["contatos_cliente_hist"])
    df["taxa_acordo_grupo"] = safe_div(df["acordos_hist"], df["cpcs_hist"])
    df["taxa_pagamento_grupo"] = safe_div(df["pagamentos_hist"], df["acordos_hist"])
    df["percentual_medio_recuperado_grupo"] = safe_div(df["valor_pago_hist"], df["valor_negociado_hist"])

    usar_media_contato = df["acionamentos_hist"] < 30
    usar_media_cpc = df["contatos_cliente_hist"] < 10
    usar_media_acordo = df["cpcs_hist"] < 10
    usar_media_pagamento = df["acordos_hist"] < 5
    usar_media_recuperado = df["valor_negociado_hist"] <= 0

    df["taxa_contato"] = np.where(usar_media_contato, global_rates["taxa_contato"], df["taxa_contato_grupo"])
    df["taxa_cpc"] = np.where(usar_media_cpc, global_rates["taxa_cpc"], df["taxa_cpc_grupo"])
    df["taxa_acordo"] = np.where(usar_media_acordo, global_rates["taxa_acordo"], df["taxa_acordo_grupo"])
    df["taxa_pagamento"] = np.where(usar_media_pagamento, global_rates["taxa_pagamento"], df["taxa_pagamento_grupo"])
    df["percentual_medio_recuperado"] = np.where(
        usar_media_recuperado,
        global_rates["percentual_medio_recuperado"],
        df["percentual_medio_recuperado_grupo"],
    )
    for col in ["taxa_contato", "taxa_cpc", "taxa_acordo", "taxa_pagamento", "percentual_medio_recuperado"]:
        df[col] = pd.Series(df[col]).fillna(0).clip(lower=0, upper=1)

    df["recuperacao_esperada"] = (
        df["carteira_elegivel"]
        * df["taxa_contato"]
        * df["taxa_cpc"]
        * df["taxa_acordo"]
        * df["taxa_pagamento"]
        * df["percentual_medio_recuperado"]
    )
    df["recuperacao_esperada_pct_carteira"] = safe_div(df["recuperacao_esperada"], df["carteira_elegivel"])
    usa_alguma_media = usar_media_contato | usar_media_cpc | usar_media_acordo | usar_media_pagamento | usar_media_recuperado
    usa_todas_medias = usar_media_contato & usar_media_cpc & usar_media_acordo & usar_media_pagamento & usar_media_recuperado
    df["base_taxas"] = np.select(
        [usa_todas_medias, usa_alguma_media],
        ["Média geral", "Grupo + média geral"],
        default="Grupo",
    )
    return df.sort_values("recuperacao_esperada", ascending=False)


def metric_card(label, value, help_text=None):
    value_text = str(value)
    title = help_text or f"{label}: {value_text}"
    compact_class = " metric-card--compact" if len(value_text) >= 14 else ""
    st.markdown(
        f"""
        <div class="metric-card{compact_class}" title="{escape(title)}">
            <div class="metric-card__label">{escape(str(label))}</div>
            <div class="metric-card__value">{escape(value_text)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def gauge_point(cx, cy, radius, ratio):
    angle = math.pi * (1 - ratio)
    return cx + radius * math.cos(angle), cy - radius * math.sin(angle)


def gauge_path(cx, cy, radius, start_ratio, end_ratio):
    start_x, start_y = gauge_point(cx, cy, radius, start_ratio)
    end_x, end_y = gauge_point(cx, cy, radius, end_ratio)
    return f"M {start_x:.2f} {start_y:.2f} A {radius} {radius} 0 0 1 {end_x:.2f} {end_y:.2f}"


def hex_to_rgb(color):
    color = color.lstrip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def blend_hex(start, end, ratio):
    ratio = min(max(ratio, 0), 1)
    start_rgb = hex_to_rgb(start)
    end_rgb = hex_to_rgb(end)
    blended = [round(a + (b - a) * ratio) for a, b in zip(start_rgb, end_rgb)]
    return "#{:02x}{:02x}{:02x}".format(*blended)


def meta_progress_color(pct):
    pct = 0 if pd.isna(pct) else min(max(float(pct), 0), 1)
    if pct <= 0.40:
        return blend_hex("#9f1d2f", "#de6a76", pct / 0.40)
    if pct <= 0.70:
        return blend_hex("#c95616", "#f2a23a", (pct - 0.40) / 0.30)
    if pct <= 0.80:
        return blend_hex("#d4a514", "#ffe071", (pct - 0.70) / 0.10)
    return blend_hex("#1f7a4d", "#72d391", (pct - 0.80) / 0.20)


def meta_gauge(value, target, month_label, open_today_count, open_today_value, open_today_rows=None, title="Recebimento Total"):
    value = 0 if pd.isna(value) else float(value)
    target = 0 if pd.isna(target) else float(target)
    open_today_count = 0 if pd.isna(open_today_count) else int(open_today_count)
    open_today_value = 0 if pd.isna(open_today_value) else float(open_today_value)
    if target <= 0:
        st.info("Sem meta geral cadastrada para montar o indicador.")
        return

    gap = max(target - value, 0)
    progress = min(max(value / target, 0), 1)
    pct_target = value / target if target else 0
    color = meta_progress_color(pct_target)
    open_today_rows = open_today_rows if open_today_rows is not None else pd.DataFrame()
    if open_today_rows.empty:
        open_today_html = '<div class="meta-tooltip__empty">Sem boletos em aberto para hoje.</div>'
    else:
        items = []
        for _, row in open_today_rows.sort_values(["OPERADOR", "NOME DO CLIENTE"]).iterrows():
            cliente = normalize_text(row.get("NOME DO CLIENTE", "Cliente sem nome"))
            operador = normalize_text(row.get("OPERADOR", ""))
            valor = money_fmt(row.get("VALOR_EM_ABERTO", 0))
            items.append(
                f'<div class="meta-tooltip__row"><span>{escape(cliente)}</span><span>{escape(operador)}</span><strong>{escape(valor)}</strong></div>'
            )
        open_today_html = "".join(items)

    cx, cy, radius = 320, 230, 185
    bg_path = gauge_path(cx, cy, radius, 0, 1)
    value_path = gauge_path(cx, cy, radius, 0, progress)

    st.markdown(
        f"""
        <style>
        .meta-panel__side {{
            width: 190px;
            display: flex;
            flex-direction: column;
            gap: 10px;
        }}
        .meta-panel__card {{
            border: 1px solid rgba(125, 211, 252, .24);
            border-radius: 7px;
            background: rgba(2, 47, 63, .72);
            box-shadow: inset 0 1px 0 rgba(255,255,255,.08);
            color: #f8fafc;
            overflow: visible;
        }}
        .meta-panel__card-title {{
            padding: 7px 10px;
            border-bottom: 1px solid rgba(125, 211, 252, .18);
            color: rgba(255,255,255,.86);
            font-size: .76rem;
            font-weight: 800;
            text-align: center;
            text-transform: uppercase;
        }}
        .meta-panel__value {{
            padding: 10px 10px;
            text-align: center;
            font-size: 1.02rem;
            font-weight: 750;
        }}
        .meta-panel__open {{
            position: relative;
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 8px;
            padding: 8px 10px;
            border-bottom: 1px solid rgba(125, 211, 252, .18);
            cursor: default;
            font-size: .86rem;
        }}
        .meta-panel__open span {{
            color: rgba(255,255,255,.82);
        }}
        .meta-panel__today-value {{
            padding: 9px 10px;
            text-align: center;
            color: #9ee7ef;
            font-weight: 760;
        }}
        .meta-tooltip {{
            display: none;
            position: absolute;
            top: 34px;
            right: 0;
            z-index: 20;
            width: 460px;
            max-height: 260px;
            overflow: auto;
            padding: 10px;
            border: 1px solid rgba(125, 211, 252, .34);
            border-radius: 7px;
            background: #07111f;
            box-shadow: 0 18px 34px rgba(0,0,0,.35);
        }}
        .meta-panel__open:hover .meta-tooltip {{
            display: block;
        }}
        .meta-tooltip__title {{
            margin-bottom: 8px;
            color: #f8fafc;
            font-size: .82rem;
            font-weight: 800;
        }}
        .meta-tooltip__row {{
            display: grid;
            grid-template-columns: minmax(150px, 1fr) 120px auto;
            gap: 8px;
            align-items: center;
            padding: 6px 0;
            border-top: 1px solid rgba(148,163,184,.18);
            color: rgba(255,255,255,.84);
            font-size: .78rem;
        }}
        .meta-tooltip__row span {{
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        .meta-tooltip__row strong {{
            color: #9ee7ef;
        }}
        .meta-tooltip__empty {{
            color: rgba(255,255,255,.78);
            font-size: .8rem;
        }}
        </style>
        <div style="border:1px solid rgba(47,111,115,.55);border-radius:8px;padding:12px 14px;background:rgba(15,23,42,.10);max-width:940px;margin:12px auto 18px;">
            <div style="display:grid;grid-template-columns:1fr auto;gap:16px;align-items:start;">
                <div>
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:2px;">
                        <div style="color:#ffffff;font-size:1.65rem;font-weight:650;line-height:1.1;">{escape(title)}</div>
                        <div style="color:#ffffff;font-size:1.25rem;font-weight:650;line-height:1.1;">{escape(month_label)}</div>
                    </div>
                    <svg viewBox="0 0 640 290" width="100%" height="260" style="display:block;max-height:260px;" role="img" aria-label="{escape(title)}">
                        <path d="{bg_path}" fill="none" stroke="rgba(255,255,255,.14)" stroke-width="70" stroke-linecap="butt"/>
                        <path d="{value_path}" fill="none" stroke="{color}" stroke-width="70" stroke-linecap="butt"/>
                        <rect x="260" y="145" width="120" height="54" rx="10" fill="rgba(255,255,255,.72)"/>
                        <text x="320" y="181" text-anchor="middle" fill="#2f6f73" font-size="28" font-weight="700">{escape(pct_fmt(pct_target))}</text>
                        <text x="320" y="246" text-anchor="middle" fill="#ffffff" font-size="34" font-weight="500">{escape(money_fmt(value))}</text>
                        <text x="118" y="274" fill="rgba(255,255,255,.72)" font-size="15">{escape(money_fmt(0))}</text>
                        <text x="466" y="274" fill="rgba(255,255,255,.72)" font-size="15">{escape(money_fmt(target))}</text>
                    </svg>
                </div>
                <div class="meta-panel__side">
                    <div class="meta-panel__card">
                        <div class="meta-panel__card-title">GAP</div>
                        <div class="meta-panel__value">{escape(money_fmt(gap))}</div>
                    </div>
                    <div class="meta-panel__card">
                        <div class="meta-panel__card-title">Hoje</div>
                        <div class="meta-panel__open">
                            <span>Em aberto</span><strong>{escape(num_fmt(open_today_count))}</strong>
                            <div class="meta-tooltip">
                                <div class="meta-tooltip__title">Boletos em aberto hoje</div>
                                {open_today_html}
                            </div>
                        </div>
                        <div class="meta-panel__today-value">{escape(money_fmt(open_today_value))}</div>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


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
    money_cols = [
        "valor_negociado",
        "valor_pago",
        "valor_em_aberto",
        "valor_nao_pagou",
        "valor_quebra",
        "ticket_medio",
        "meta_individual",
        "saldo_meta_individual",
        "meta_geral_escritorio",
        "total_amount_due",
        "valor_negociado_hist",
        "valor_pago_hist",
        "carteira_elegivel",
        "recuperacao_esperada",
    ]
    pct_cols = [
        "tx_contato",
        "tx_acordo",
        "tx_acordo_cliente_cpc",
        "tx_pagamento",
        "efetividade_pagamento",
        "pct_quebra",
        "tx_pagamento_cpc",
        "tx_sem_pagamento",
        "tx_cpc_acordo",
        "tx_cpc_pagamento",
        "tx_cpc_unico_acordo",
        "tx_cpc_unico_pagamento",
        "tx_acordo_pagamento",
        "tx_acordo_sem_pagamento",
        "recuperacao",
        "score",
        "atingimento_meta_individual",
        "pct_aberto_meta_individual",
        "participacao_meta_geral",
        "score_recuperacao",
        "taxa_contato",
        "taxa_cpc",
        "taxa_acordo",
        "taxa_pagamento",
        "percentual_medio_recuperado",
        "recuperacao_esperada_pct_carteira",
    ]
    num_cols = [
        "acionamentos",
        "clientes_trabalhados",
        "contatos_efetivos",
        "contatos_cliente",
        "cpcs",
        "cpcs_unicos",
        "clientes_cpc",
        "contratos_cpc",
        "acordos",
        "pagamentos",
        "acordos_sem_pagamento",
        "acordos_em_aberto",
        "acordos_nao_pagou",
        "clientes",
        "contratos_elegiveis",
        "dpd",
        "dias_sem_contato",
        "acionamentos_hist",
        "contatos_cliente_hist",
        "cpcs_hist",
        "acordos_hist",
        "pagamentos_hist",
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
    for col in [
        "valor_negociado",
        "valor_pago",
        "valor_em_aberto",
        "valor_nao_pagou",
        "valor_quebra",
        "ticket_medio",
        "meta_individual",
        "saldo_meta_individual",
        "meta_geral_escritorio",
        "total_amount_due",
        "valor_negociado_hist",
        "valor_pago_hist",
        "carteira_elegivel",
        "recuperacao_esperada",
    ]:
        if col in out:
            out[col] = out[col].map(money_fmt)
    for col in [
        "tx_contato",
        "tx_acordo",
        "tx_acordo_cliente_cpc",
        "tx_pagamento",
        "efetividade_pagamento",
        "pct_quebra",
        "tx_pagamento_cpc",
        "tx_sem_pagamento",
        "tx_cpc_acordo",
        "tx_cpc_pagamento",
        "tx_cpc_unico_acordo",
        "tx_cpc_unico_pagamento",
        "tx_acordo_pagamento",
        "tx_acordo_sem_pagamento",
        "recuperacao",
        "score",
        "atingimento_meta_individual",
        "pct_aberto_meta_individual",
        "participacao_meta_geral",
        "score_recuperacao",
        "taxa_contato",
        "taxa_cpc",
        "taxa_acordo",
        "taxa_pagamento",
        "percentual_medio_recuperado",
        "recuperacao_esperada_pct_carteira",
    ]:
        if col in out:
            out[col] = out[col].map(pct_fmt)
    for col in [
        "acionamentos",
        "clientes_trabalhados",
        "contatos_efetivos",
        "contatos_cliente",
        "cpcs",
        "cpcs_unicos",
        "clientes_cpc",
        "contratos_cpc",
        "acordos",
        "pagamentos",
        "acordos_sem_pagamento",
        "acordos_em_aberto",
        "acordos_nao_pagou",
        "clientes",
        "contratos_elegiveis",
        "dpd",
        "dias_sem_contato",
        "acionamentos_hist",
        "contatos_cliente_hist",
        "cpcs_hist",
        "acordos_hist",
        "pagamentos_hist",
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


def dataframe_to_excel_bytes(df, sheet_name="Workplan", extra_sheets=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        for extra_sheet_name, extra_df in (extra_sheets or {}).items():
            extra_df.to_excel(writer, index=False, sheet_name=extra_sheet_name)
    return output.getvalue()


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


eventos_raw, contratos_raw, resultados_raw = load_data(data_file_versions())
workplan_raw, workplan_error = load_workplan()
eventos, resultados, operadores_filtrados = apply_filters(eventos_raw, resultados_raw)
operador_df = aggregate_operator(eventos, resultados)
cpc_df = aggregate_cpc_operator(eventos, resultados)
workplan_df = build_workplan_analysis(workplan_raw, eventos_raw[eventos_raw["IS_ACIONAMENTO"]], resultados_raw)

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
efetividade_pagamento_geral = total_pagamentos / (total_pagamentos + total_nao_pagou) if (total_pagamentos + total_nao_pagou) else 0
pct_quebra_geral = total_nao_pagou / total_acordos if total_acordos else 0

kpi_row1 = st.columns(5)
kpi_row2 = st.columns(6)
kpi_row3 = st.columns(3)
with kpi_row1[0]:
    metric_card("Clientes", num_fmt(total_clientes))
with kpi_row1[1]:
    metric_card("Acionamentos", num_fmt(total_acionamentos))
with kpi_row1[2]:
    metric_card("Contatos efetivos", num_fmt(total_contatos))
with kpi_row1[3]:
    metric_card("Acordos", num_fmt(total_acordos))
with kpi_row1[4]:
    metric_card("Pagamentos", num_fmt(total_pagamentos))
with kpi_row2[0]:
    metric_card("Em aberto", num_fmt(total_em_aberto))
with kpi_row2[1]:
    metric_card("Não pagou", num_fmt(total_nao_pagou))
with kpi_row2[2]:
    metric_card("Negociado", money_fmt(valor_negociado))
with kpi_row2[3]:
    metric_card("Recebido", money_fmt(valor_pago))
with kpi_row2[4]:
    metric_card("Recuperação", pct_fmt(valor_pago / valor_negociado if valor_negociado else 0))

tabs = st.tabs(["Visão Geral", "Operadores", "CPC", "Faixa de Atraso", "DPD", "Região", "Matriz", "Metas", "Workplan", "Insights"])

with kpi_row3[0]:
    metric_card("% quebras", pct_fmt(pct_quebra_geral))
with kpi_row3[1]:
    metric_card("Valor quebras", money_fmt(valor_nao_pagou))
with kpi_row3[2]:
    metric_card("Base quebras", f"{num_fmt(total_nao_pagou)} de {num_fmt(total_acordos)}")

with kpi_row2[5]:
    metric_card("Efetividade pgto", pct_fmt(efetividade_pagamento_geral))

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
        "pct_quebra",
        "tx_contato",
        "tx_acordo",
        "tx_pagamento_cpc",
        "tx_pagamento",
        "efetividade_pagamento",
        "tx_sem_pagamento",
        "valor_negociado",
        "valor_pago",
        "valor_em_aberto",
        "valor_nao_pagou",
        "valor_quebra",
        "ticket_medio",
        "recuperacao",
        "score",
    ]
    data_table(operador_df[cols])

with tabs[2]:
    st.subheader("Conversão CPC para acordos e pagamentos")
    st.caption("CPC considerado pelos eventos iniciados por 02, 03, 04 e 05.")

    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    with c1:
        metric_card("CPCs", num_fmt(cpc_df["cpcs"].sum()))
    with c2:
        metric_card("CPCs únicos", num_fmt(cpc_df["cpcs_unicos"].sum()))
    with c3:
        metric_card("Acordos após CPC", num_fmt(cpc_df["acordos"].sum()))
    with c4:
        metric_card("Pagamentos", num_fmt(cpc_df["pagamentos"].sum()))
    with c5:
        metric_card("Acordos sem pagamento", num_fmt(cpc_df["acordos_sem_pagamento"].sum()))
    with c6:
        metric_card("Em aberto", num_fmt(cpc_df["acordos_em_aberto"].sum()))
    with c7:
        metric_card("Não pagou", num_fmt(cpc_df["acordos_nao_pagou"].sum()))

    c1, c2 = st.columns(2)
    with c1:
        cpc_chart = display_fields(cpc_df[cpc_df["cpcs_unicos"] > 0].sort_values("tx_cpc_unico_acordo", ascending=False).head(15))
        bar_chart(
            cpc_chart,
            x=alt.X("tx_cpc_unico_acordo:Q", axis=alt.Axis(format="%")),
            y="OPERADOR:N",
            tooltip=["OPERADOR", "cpcs_br", "cpcs_unicos_br", "acordos_br", "tx_cpc_unico_acordo_br", "valor_negociado_br"],
            title="Conversão CPC único -> acordo por negociador",
        )
    with c2:
        cpc_pag_chart = display_fields(cpc_df[cpc_df["cpcs_unicos"] > 0].sort_values("tx_cpc_unico_pagamento", ascending=False).head(15))
        bar_chart(
            cpc_pag_chart,
            x=alt.X("tx_cpc_unico_pagamento:Q", axis=alt.Axis(format="%")),
            y="OPERADOR:N",
            tooltip=["OPERADOR", "cpcs_br", "cpcs_unicos_br", "pagamentos_br", "tx_cpc_unico_pagamento_br", "valor_pago_br"],
            title="Conversão CPC único -> pagamento por negociador",
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
        cpc_scatter = display_fields(cpc_df[(cpc_df["cpcs_unicos"] > 0) & (cpc_df["acordos"] > 0)])
        scatter = (
            alt.Chart(cpc_scatter)
            .mark_circle(size=130, opacity=0.78)
            .encode(
                x=alt.X("tx_cpc_unico_acordo:Q", title="CPC único -> acordo", axis=alt.Axis(format="%")),
                y=alt.Y("tx_acordo_pagamento:Q", title="Acordo -> pagamento", axis=alt.Axis(format="%")),
                size=alt.Size("valor_pago:Q", title="Valor recebido"),
                color=alt.Color("acordos_sem_pagamento:Q", scale=alt.Scale(scheme="orangered"), title="Sem pagamento"),
                tooltip=[
                    "OPERADOR",
                    "cpcs_br",
                    "cpcs_unicos_br",
                    "acordos_br",
                    "pagamentos_br",
                    "acordos_sem_pagamento_br",
                    "tx_cpc_unico_acordo_br",
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
        "cpcs_unicos",
        "clientes_cpc",
        "acordos",
        "pagamentos",
        "acordos_sem_pagamento",
        "acordos_em_aberto",
        "acordos_nao_pagou",
        "pct_quebra",
        "tx_cpc_acordo",
        "tx_cpc_pagamento",
        "tx_cpc_unico_acordo",
        "tx_cpc_unico_pagamento",
        "tx_acordo_pagamento",
        "efetividade_pagamento",
        "tx_acordo_sem_pagamento",
        "valor_negociado",
        "valor_pago",
        "valor_em_aberto",
        "valor_nao_pagou",
        "valor_quebra",
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
    best_faixa["efetividade_pagamento"] = safe_div(best_faixa["pagamentos"], best_faixa["pagamentos"] + best_faixa["acordos_nao_pagou"])
    best_faixa["pct_quebra"] = safe_div(best_faixa["acordos_nao_pagou"], best_faixa["acordos"])
    best_faixa["valor_quebra"] = best_faixa["valor_nao_pagou"]
    best_faixa = best_faixa.sort_values(["FAIXA_ATRASO", "valor_pago", "tx_pagamento"], ascending=[True, False, False]).groupby("FAIXA_ATRASO").head(1)
    st.subheader("Melhor operador por faixa")
    data_table(best_faixa)

with tabs[4]:
    resultados_segmento = resultados[resultados["SEGMENTO_DPD"].ne("Sem DPD")].copy()
    eventos_segmento = eventos[eventos["SEGMENTO_DPD"].ne("Sem DPD")].copy()
    segmento_df = aggregate_resultados(resultados_segmento, "SEGMENTO_DPD")
    ev_segmento = eventos_segmento.groupby("SEGMENTO_DPD").agg(
        acionamentos=("EVENTO_TXT", "size"),
        contatos_efetivos=("IS_CONTATO_EFETIVO", "sum"),
    ).reset_index()
    segmento_df = segmento_df.merge(ev_segmento, on="SEGMENTO_DPD", how="outer").fillna(0)
    segmento_df["tx_contato"] = safe_div(segmento_df["contatos_efetivos"], segmento_df["acionamentos"])
    segmento_order = ["POTLOSS", "SALVAGE", "SALVAGE +"]
    segmento_df["SEGMENTO_DPD"] = pd.Categorical(segmento_df["SEGMENTO_DPD"], categories=segmento_order, ordered=True)
    segmento_df = segmento_df.sort_values("SEGMENTO_DPD")
    segmento_chart = display_fields(segmento_df)

    c1, c2 = st.columns(2)
    with c1:
        bar_chart(
            segmento_chart,
            x="valor_pago:Q",
            y="SEGMENTO_DPD:N",
            tooltip=["SEGMENTO_DPD", "valor_pago_br", "valor_negociado_br", "acordos_br", "pagamentos_br", "recuperacao_br"],
            title="Valor recebido por segmento DPD",
            sort=segmento_order,
        )
    with c2:
        bar_chart(
            segmento_chart,
            x=alt.X("tx_pagamento:Q", axis=alt.Axis(format="%")),
            y="SEGMENTO_DPD:N",
            tooltip=["SEGMENTO_DPD", "tx_pagamento_br", "acordos_br", "pagamentos_br", "tx_contato_br"],
            title="Conversão acordo/pagamento por segmento DPD",
            sort=segmento_order,
        )

    best_segmento = (
        resultados_segmento.groupby(["SEGMENTO_DPD", "OPERADOR"])
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
    best_segmento["tx_pagamento"] = safe_div(best_segmento["pagamentos"], best_segmento["acordos"])
    best_segmento["efetividade_pagamento"] = safe_div(best_segmento["pagamentos"], best_segmento["pagamentos"] + best_segmento["acordos_nao_pagou"])
    best_segmento["pct_quebra"] = safe_div(best_segmento["acordos_nao_pagou"], best_segmento["acordos"])
    best_segmento["valor_quebra"] = best_segmento["valor_nao_pagou"]
    best_segmento["recuperacao"] = safe_div(best_segmento["valor_pago"], best_segmento["valor_negociado"])
    best_segmento = best_segmento.sort_values(["SEGMENTO_DPD", "valor_pago", "tx_pagamento"], ascending=[True, False, False]).groupby("SEGMENTO_DPD").head(3)
    st.subheader("Top operadores por segmento DPD")
    data_table(best_segmento)

    st.subheader("Resumo por segmento DPD")
    data_table(segmento_df[["SEGMENTO_DPD", "clientes", "acionamentos", "contatos_efetivos", "tx_contato", "acordos", "pagamentos", "acordos_em_aberto", "acordos_nao_pagou", "pct_quebra", "efetividade_pagamento", "valor_negociado", "valor_pago", "valor_em_aberto", "valor_nao_pagou", "valor_quebra", "recuperacao"]])

with tabs[5]:
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
    best_regiao["efetividade_pagamento"] = safe_div(best_regiao["pagamentos"], best_regiao["pagamentos"] + best_regiao["acordos_nao_pagou"])
    best_regiao["pct_quebra"] = safe_div(best_regiao["acordos_nao_pagou"], best_regiao["acordos"])
    best_regiao["valor_quebra"] = best_regiao["valor_nao_pagou"]
    best_regiao["recuperacao"] = safe_div(best_regiao["valor_pago"], best_regiao["valor_negociado"])
    best_regiao = best_regiao.sort_values(["REGIÃO", "valor_pago", "recuperacao"], ascending=[True, False, False]).groupby("REGIÃO").head(3)
    st.subheader("Top operadores por região")
    data_table(best_regiao)

with tabs[6]:
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
    matrix["efetividade_pagamento"] = safe_div(matrix["pagamentos"], matrix["pagamentos"] + matrix["acordos_nao_pagou"])
    matrix["pct_quebra"] = safe_div(matrix["acordos_nao_pagou"], matrix["acordos"])
    matrix["valor_quebra"] = matrix["valor_nao_pagou"]
    matrix["recuperacao"] = safe_div(matrix["valor_pago"], matrix["valor_negociado"])

    metric_choice = st.selectbox(
        "Métrica da matriz",
        ["valor_pago", "valor_em_aberto", "valor_nao_pagou", "valor_quebra", "tx_pagamento", "efetividade_pagamento", "pct_quebra", "recuperacao", "acordos", "pagamentos", "acordos_em_aberto", "acordos_nao_pagou"],
        index=0,
    )
    heat_data = matrix.groupby(["OPERADOR", "FAIXA_ATRASO"]).agg(
        {metric_choice: "sum" if metric_choice in ["valor_pago", "valor_em_aberto", "valor_nao_pagou", "valor_quebra", "acordos", "pagamentos", "acordos_em_aberto", "acordos_nao_pagou"] else "mean"}
    ).reset_index()
    heatmap(display_fields(heat_data), "FAIXA_ATRASO:N", "OPERADOR:N", f"{metric_choice}:Q", "Operador x faixa de atraso")

    st.subheader("Matriz analítica por operador, faixa e região")
    data_table(matrix.sort_values(["valor_pago", "pagamentos"], ascending=False))

with tabs[7]:
    st.subheader("Metas e quartis de atingimento")
    st.caption("Meta mensal: R$ 150.000 por negociador. Ana Karolina e Luiz Mauro usam R$ 300.000 por cuidarem de pós retomado.")

    metas_df, meses_meta, meta_geral = build_meta_analysis(operador_df, resultados, operadores_filtrados)
    recebido_meta_geral = resultados["VALOR_PAGO"].sum()
    valor_aberto_meta_geral = resultados["VALOR_EM_ABERTO"].sum()
    hoje = pd.Timestamp.today().normalize()
    abertos_hoje = resultados[resultados["IS_EM_ABERTO"] & resultados["DATA_VENCIMENTO"].dt.normalize().eq(hoje)]
    boletos_abertos_hoje = len(abertos_hoje)
    valor_aberto_hoje = abertos_hoje["VALOR_EM_ABERTO"].sum()
    meses_texto = ", ".join(meses_meta) if meses_meta else "Sem mês filtrado"

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        metric_card("Mês analisado", meses_texto)
    with c2:
        metric_card("Meta geral escritório", money_fmt(meta_geral))
    with c3:
        metric_card("Recebido", money_fmt(recebido_meta_geral))
    with c4:
        metric_card("% meta geral", pct_fmt(recebido_meta_geral / meta_geral if meta_geral else 0))
    with c5:
        metric_card("Em aberto", money_fmt(valor_aberto_meta_geral))
    with c6:
        metric_card("% aberto/meta", pct_fmt(valor_aberto_meta_geral / meta_geral if meta_geral else 0))

    meta_gauge(recebido_meta_geral, meta_geral, meses_texto, boletos_abertos_hoje, valor_aberto_hoje, abertos_hoje)

    meta_resumo = metas_df["diagnostico_meta"].value_counts().reset_index()
    meta_resumo.columns = ["Diagnóstico", "Operadores"]
    c1, c2 = st.columns([1, 1.3])
    with c1:
        st.subheader("Resumo por diagnóstico")
        st.dataframe(meta_resumo, use_container_width=True, hide_index=True)
    with c2:
        chart_meta = display_fields(metas_df.sort_values("atingimento_meta_individual", ascending=False).head(15))
        bar_chart(
            chart_meta,
            x=alt.X("atingimento_meta_individual:Q", axis=alt.Axis(format="%")),
            y="OPERADOR:N",
            color="diagnostico_meta:N",
            tooltip=[
                "OPERADOR",
                "valor_pago_br",
                "valor_em_aberto_br",
                "meta_individual_br",
                "atingimento_meta_individual_br",
                "pct_aberto_meta_individual_br",
                "saldo_meta_individual_br",
                "participacao_meta_geral_br",
                "diagnostico_meta",
            ],
            title="Atingimento da meta individual",
            height=360,
        )

    st.subheader("Operadores em cada quartil")
    grupos_meta = meta_operator_groups(metas_df)
    st.dataframe(
        grupos_meta,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Métrica": st.column_config.TextColumn("Métrica", help="Indicador usado para montar o quartil."),
            "Grupo": st.column_config.TextColumn("Grupo", help="Grupo do quartil. Q4 é destaque; Q1 é crítico."),
            "Qtd. operadores": st.column_config.NumberColumn("Qtd. operadores", help="Quantidade de operadores no grupo."),
            "Operadores": st.column_config.TextColumn("Operadores", help="Operadores classificados nesse grupo."),
        },
    )

    st.subheader("Tabela de metas por operador")
    meta_cols = [
        "OPERADOR",
        "nome_colaborador",
        "base_colaborador",
        "cargo_colaborador",
        "negociador_cadastrado",
        "valor_pago",
        "valor_em_aberto",
        "meta_individual",
        "atingimento_meta_individual",
        "pct_aberto_meta_individual",
        "efetividade_pagamento",
        "pct_quebra",
        "valor_quebra",
        "saldo_meta_individual",
        "meta_geral_escritorio",
        "participacao_meta_geral",
        "quartil_meta_individual",
        "quartil_meta_geral",
        "diagnostico_meta",
        "pagamentos",
        "acordos_nao_pagou",
        "acordos",
    ]
    data_table(metas_df[meta_cols])

with tabs[8]:
    st.subheader("Workplan e priorização futura")
    if workplan_error:
        st.warning(workplan_error)
    elif workplan_df.empty:
        st.info("Sem dados do Workplan para exibir.")
    else:
        elegivel_df = workplan_df[
            workplan_df["pagamentos_hist"].eq(0)
            & workplan_df["acordos_em_aberto_hist"].eq(0)
        ].copy()
        cobravel_df = elegivel_df[elegivel_df["flag_cobravel"].eq("SIM")].copy()
        base_workplan = cobravel_df if not cobravel_df.empty else elegivel_df

        if base_workplan.empty:
            st.info("Sem contratos sem pagamento histórico e sem acordo em aberto para recomendar.")

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            metric_card("Contratos elegíveis", num_fmt(len(base_workplan)))
        with c2:
            metric_card("Valor em aberto", money_fmt(base_workplan["total_amount_due"].sum()))
        with c3:
            metric_card("Prioridade alta", num_fmt(base_workplan["prioridade_workplan"].eq("Alta").sum()))
        with c4:
            metric_card("Com CPC histórico", num_fmt((base_workplan["cpcs_hist"] > 0).sum()))
        with c5:
            excluidos = workplan_df["pagamentos_hist"].gt(0) | workplan_df["acordos_em_aberto_hist"].gt(0)
            metric_card("Pagos/em aberto excluídos", num_fmt(excluidos.sum()))

        prioridade_sel = st.multiselect("Prioridade Workplan", ["Alta", "Média", "Baixa"], default=["Alta", "Média"])
        segmento_sel = st.multiselect("Segmento Workplan", ["POTLOSS", "SALVAGE", "SALVAGE +"])
        workplan_view = base_workplan.copy()
        if prioridade_sel:
            workplan_view = workplan_view[workplan_view["prioridade_workplan"].isin(prioridade_sel)]
        if segmento_sel:
            workplan_view = workplan_view[workplan_view["SEGMENTO_DPD"].isin(segmento_sel)]

        c1, c2 = st.columns(2)
        with c1:
            prioridade_df = (
                workplan_view.groupby("prioridade_workplan")
                .agg(contratos=("CONTRATO_KEY", "nunique"), total_amount_due=("total_amount_due", "sum"), score_recuperacao=("score_recuperacao", "mean"))
                .reset_index()
            )
            prioridade_df = display_fields(prioridade_df.rename(columns={"contratos": "clientes"}))
            bar_chart(
                prioridade_df,
                x="total_amount_due:Q",
                y=alt.Y("prioridade_workplan:N", sort=["Alta", "Média", "Baixa"], title=None),
                color="prioridade_workplan:N",
                tooltip=["prioridade_workplan", "clientes_br", "total_amount_due_br", "score_recuperacao_br"],
                title="Valor em aberto por prioridade",
                sort=None,
            )
        with c2:
            segmento_workplan = (
                workplan_view.groupby("SEGMENTO_DPD")
                .agg(clientes=("CONTRATO_KEY", "nunique"), total_amount_due=("total_amount_due", "sum"), score_recuperacao=("score_recuperacao", "mean"))
                .reset_index()
            )
            segmento_workplan = display_fields(segmento_workplan)
            bar_chart(
                segmento_workplan,
                x="total_amount_due:Q",
                y=alt.Y("SEGMENTO_DPD:N", sort=["POTLOSS", "SALVAGE", "SALVAGE +", "Sem DPD"], title=None),
                tooltip=["SEGMENTO_DPD", "clientes_br", "total_amount_due_br", "score_recuperacao_br"],
                title="Valor em aberto por segmento",
                sort=None,
            )

        st.subheader("Recuperação esperada ponderada")
        st.caption("Carteira elegível x taxa de contato x taxa de CPC x taxa de acordo x taxa de pagamento x percentual médio recuperado.")
        projection_dims = {
            "Segmento DPD": ("SEGMENTO_DPD", "SEGMENTO_DPD", "SEGMENTO_DPD"),
            "Faixa de atraso": ("FAIXA_ATRASO", "FAIXA_ATRASO", "FAIXA_ATRASO"),
            "Região": ("REGIÃO", "REGIAO", "REGIÃO"),
            "UF/Estado": ("uf", "UF", "UF"),
            "Segmento + faixa": (
                ["SEGMENTO_DPD", "FAIXA_ATRASO"],
                ["SEGMENTO_DPD", "FAIXA_ATRASO"],
                ["SEGMENTO_DPD", "FAIXA_ATRASO"],
            ),
            "Segmento + região": (
                ["SEGMENTO_DPD", "REGIÃO"],
                ["SEGMENTO_DPD", "REGIAO"],
                ["SEGMENTO_DPD", "REGIÃO"],
            ),
            "Faixa + região": (
                ["FAIXA_ATRASO", "REGIÃO"],
                ["FAIXA_ATRASO", "REGIAO"],
                ["FAIXA_ATRASO", "REGIÃO"],
            ),
            "Segmento + faixa + região": (
                ["SEGMENTO_DPD", "FAIXA_ATRASO", "REGIÃO"],
                ["SEGMENTO_DPD", "FAIXA_ATRASO", "REGIAO"],
                ["SEGMENTO_DPD", "FAIXA_ATRASO", "REGIÃO"],
            ),
        }
        if "PRODUTO" in workplan_view.columns:
            projection_dims["Produto"] = ("PRODUTO", "PRODUTO", "PRODUTO")
            projection_dims["Produto + segmento"] = (
                ["PRODUTO", "SEGMENTO_DPD"],
                ["PRODUTO", "SEGMENTO_DPD"],
                ["PRODUTO", "SEGMENTO_DPD"],
            )
        projection_group = st.selectbox("Agrupamento da projeção", list(projection_dims.keys()))
        projection_df = expected_recovery_by_group(
            workplan_view,
            eventos_raw,
            resultados_raw,
            *projection_dims[projection_group],
        )
        if projection_df.empty:
            st.info("Sem base suficiente para calcular a recuperação esperada nos filtros atuais.")
        else:
            total_carteira_proj = projection_df["carteira_elegivel"].sum()
            total_recuperacao_proj = projection_df["recuperacao_esperada"].sum()
            c1, c2, c3 = st.columns(3)
            with c1:
                metric_card("Carteira elegível projetada", money_fmt(total_carteira_proj))
            with c2:
                metric_card("Recuperação esperada", money_fmt(total_recuperacao_proj))
            with c3:
                metric_card("% esperado da carteira", pct_fmt(total_recuperacao_proj / total_carteira_proj if total_carteira_proj else 0))

            projection_display = display_fields(projection_df)
            bar_chart(
                projection_display.head(12),
                x="recuperacao_esperada:Q",
                y="grupo:N",
                tooltip=[
                    "grupo",
                    "contratos_elegiveis_br",
                    "carteira_elegivel_br",
                    "recuperacao_esperada_br",
                    "recuperacao_esperada_pct_carteira_br",
                    "taxa_contato_br",
                    "taxa_cpc_br",
                    "taxa_acordo_br",
                    "taxa_pagamento_br",
                    "percentual_medio_recuperado_br",
                    "base_taxas",
                ],
                title=f"Recuperação esperada por {projection_group.lower()}",
            )
            projection_cols = [
                "grupo",
                "contratos_elegiveis",
                "carteira_elegivel",
                "recuperacao_esperada",
                "recuperacao_esperada_pct_carteira",
                "taxa_contato",
                "taxa_cpc",
                "taxa_acordo",
                "taxa_pagamento",
                "percentual_medio_recuperado",
                "acionamentos_hist",
                "contatos_cliente_hist",
                "cpcs_hist",
                "acordos_hist",
                "pagamentos_hist",
                "base_taxas",
            ]
            data_table(projection_df[projection_cols])
            projection_workplan_cols = projection_dims[projection_group][0]
            projection_clients = add_projection_group(workplan_view, projection_workplan_cols)
            projection_client_cols = [
                "grupo",
                "agreement_no",
                "cust_name",
                "cpf_cnpj",
                "PRODUTO",
                "prioridade_workplan",
                "score_recuperacao",
                "SEGMENTO_DPD",
                "FAIXA_ATRASO",
                "REGIÃO",
                "uf",
                "dpd",
                "total_amount_due",
                "dias_sem_contato",
                "acionamentos_hist",
                "cpcs_hist",
                "acordos_hist",
                "pagamentos_hist",
                "status_base",
                "status_cpc",
                "city",
                "state",
            ]
            projection_client_cols = [col for col in projection_client_cols if col in projection_clients.columns]
            st.download_button(
                "Exportar projeção Excel",
                data=dataframe_to_excel_bytes(
                    projection_df[projection_cols],
                    sheet_name="Resumo",
                    extra_sheets={"Clientes": projection_clients[projection_client_cols]},
                ),
                file_name="recuperacao_esperada_workplan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.subheader("Contratos recomendados para novos acordos")
        workplan_cols = [
            "agreement_no",
            "cust_name",
            "cpf_cnpj",
            "PRODUTO",
            "prioridade_workplan",
            "score_recuperacao",
            "motivo_priorizacao",
            "SEGMENTO_DPD",
            "FAIXA_ATRASO",
            "REGIÃO",
            "uf",
            "dpd",
            "total_amount_due",
            "dias_sem_contato",
            "acionamentos_hist",
            "cpcs_hist",
            "acordos_hist",
            "pagamentos_hist",
            "status_base",
            "status_cpc",
            "city",
            "state",
        ]
        workplan_cols = [col for col in workplan_cols if col in workplan_view.columns]
        selected_workplan_cols = st.multiselect(
            "Colunas para visualizar/exportar",
            workplan_cols,
            default=workplan_cols,
            format_func=lambda col: FIELD_HELP.get(col, (col, ""))[0],
        )
        if not selected_workplan_cols:
            st.info("Selecione pelo menos uma coluna para visualizar/exportar.")
        else:
            export_view = workplan_view[selected_workplan_cols].copy()
            st.caption(f"A visualização mostra os 100 primeiros contratos. A exportação usa todos os {num_fmt(len(export_view))} contratos filtrados.")
            data_table(export_view.head(100))
            st.download_button(
                "Exportar Excel",
                data=dataframe_to_excel_bytes(export_view),
                file_name="contratos_recomendados_workplan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

with tabs[9]:
    avg_score = operador_df["score"].mean() if not operador_df.empty else 0
    oportunidades = operador_df[(operador_df["acionamentos"] >= operador_df["acionamentos"].median()) & (operador_df["score"] < avg_score)].sort_values("acionamentos", ascending=False)
    destaques = operador_df[operador_df["score"] >= operador_df["score"].quantile(0.75)].sort_values("score", ascending=False)
    faixas_oportunidade = aggregate_resultados(resultados, "FAIXA_ATRASO")
    faixas_oportunidade = faixas_oportunidade.sort_values(["valor_negociado", "recuperacao"], ascending=[False, True])
    segmentos_oportunidade = aggregate_resultados(resultados[resultados["SEGMENTO_DPD"].ne("Sem DPD")], "SEGMENTO_DPD")
    segmentos_oportunidade = segmentos_oportunidade.sort_values(["valor_negociado", "recuperacao"], ascending=[False, True])

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Operadores para priorização")
    data_table(destaques[["OPERADOR", "score", "valor_pago", "tx_pagamento", "efetividade_pagamento", "recuperacao", "clientes_trabalhados"]].head(10))
    with c2:
        st.subheader("Alto volume com eficiência abaixo da média")
        data_table(oportunidades[["OPERADOR", "score", "acionamentos", "tx_contato", "tx_acordo", "valor_pago"]].head(10))

    st.subheader("Faixas com maior oportunidade de recuperação")
    data_table(faixas_oportunidade[["FAIXA_ATRASO", "clientes", "acordos", "pagamentos", "acordos_em_aberto", "acordos_nao_pagou", "pct_quebra", "efetividade_pagamento", "valor_negociado", "valor_pago", "valor_em_aberto", "valor_nao_pagou", "valor_quebra", "recuperacao"]])

    st.subheader("Segmentos DPD com maior oportunidade de recuperação")
    data_table(segmentos_oportunidade[["SEGMENTO_DPD", "clientes", "acordos", "pagamentos", "acordos_em_aberto", "acordos_nao_pagou", "pct_quebra", "efetividade_pagamento", "valor_negociado", "valor_pago", "valor_em_aberto", "valor_nao_pagou", "valor_quebra", "recuperacao"]])
