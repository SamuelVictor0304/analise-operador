"""Análise por caso e mês da apreensão, com resultados de pós da aba BASE."""
from pathlib import Path
import re
import unicodedata

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st


MONTHS = dict(enumerate([
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
], 1))
CASE_COLUMNS = {
    "mes": "Mês da apreensão", "contrato": "Contrato", "cliente": "Cliente",
    "situacao": "Situação do caso", "pago": "Pagamento de pós", "tipo": "Judicial / DA",
    "data_retomada": "Data da retomada", "uf": "UF", "regiao": "Região", "placa": "Placa",
    "tem_acordo": "Acordo fechado", "acordo_apreensao": "Acordo em APREENSÃO",
    "venda_apreensao": "Venda em APREENSÃO", "negociador": "Negociação em APREENSÃO",
    "status_base": "Status na BASE", "valor_pago": "Valor pago de pós",
    "conferencia": "Conferência dos dados",
}


def text(value):
    return "" if pd.isna(value) else str(value).strip()


def key(value):
    normalized = unicodedata.normalize("NFKD", text(value).upper())
    return re.sub(r"[^A-Z0-9]", "", normalized.upper())


def contract_key(value):
    return re.sub(r"\D", "", text(value).removesuffix(".0"))


def format_brl(value):
    if pd.isna(value):
        return "—"
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def prepare_pos_retomado_cases(apreensoes, resultados):
    """Uma linha por retomada; parcelas da BASE não multiplicam os casos."""
    apreensoes = apreensoes.rename(columns=lambda c: key(c)).copy()
    resultados = resultados.rename(columns=lambda c: key(c)).copy()
    required = {"CONTRATO", "MES", "DATADARETOMADA", "JUDICIALDA", "ACORDO", "VENDA", "NEGOCIACAO"}
    missing = required - set(apreensoes.columns)
    base_required = {"NOCONTRATO", "EMISSAO", "DATADOPAGAMENTO", "STATUS", "CAMPANHA", "VALORDOBANCOMETA"}
    if missing or base_required - set(resultados.columns):
        raise ValueError("Colunas ausentes: " + ", ".join(sorted(missing | (base_required - set(resultados.columns)))))

    cases = pd.DataFrame(index=apreensoes.index)
    cases["contrato"] = apreensoes["CONTRATO"].map(contract_key)
    cases = cases[cases["contrato"].ne("")].copy()
    source = apreensoes.loc[cases.index]
    cases["data_retomada"] = pd.to_datetime(source["DATADARETOMADA"], errors="coerce", format="mixed", dayfirst=True)
    month_numbers = {key(name): number for number, name in MONTHS.items()}
    month_numbers.update({key(name)[:3]: number for number, name in MONTHS.items()})
    cases["mes_num"] = source["MES"].map(key).map(month_numbers)
    cases["mes_num"] = cases["mes_num"].fillna(cases["data_retomada"].dt.month)
    cases["mes"] = cases["mes_num"].map(MONTHS).fillna("SEM MÊS")
    cases["tipo"] = source["JUDICIALDA"].map(key).replace({"DA": "D.A"})
    cases["acordo_apreensao"] = source["ACORDO"].map(key).map({"SIM": "Sim", "NAO": "Não"}).fillna("Sem informação")
    cases["venda_apreensao"] = source["VENDA"].map(key).map({"SIM": "Sim", "NAO": "Não"}).fillna("Sem informação")
    cases["negociador"] = source["NEGOCIACAO"].map(text)
    for target, origin in {"cliente": "CLIENTE", "placa": "PLACA", "uf": "UF", "regiao": "REGIAO"}.items():
        cases[target] = source.get(origin, pd.Series("", index=source.index)).map(text).replace("", "Sem informação")
    cases = cases.drop_duplicates(["contrato", "data_retomada", "mes_num"], keep="last")
    cases = cases.sort_values(["contrato", "data_retomada"], na_position="last").reset_index(drop=True)
    cases["caso_id"] = cases.index
    cases["proxima_retomada"] = cases.groupby("contrato")["data_retomada"].shift(-1)

    base = pd.DataFrame(index=resultados.index)
    base["contrato"] = resultados["NOCONTRATO"].map(contract_key)
    pos_keys = {"POSRETOMADO", "POSRETOMADA", "POSTRETOMADO", "POSTRETOMADA"}
    base = base[resultados["CAMPANHA"].map(key).isin(pos_keys) & base["contrato"].ne("")].copy()
    raw = resultados.loc[base.index]
    base["data_acordo"] = pd.to_datetime(raw["EMISSAO"], errors="coerce", format="mixed", dayfirst=True)
    base["data_pagamento"] = pd.to_datetime(raw["DATADOPAGAMENTO"], errors="coerce", format="mixed", dayfirst=True)
    base["data_resultado"] = base["data_acordo"].fillna(base["data_pagamento"])
    base["status_key"] = raw["STATUS"].map(key)
    base["status_original"] = raw["STATUS"].map(text).replace("", "Sem status")
    base["valor"] = pd.to_numeric(raw["VALORDOBANCOMETA"], errors="coerce").fillna(0)
    matches = cases[["caso_id", "contrato", "data_retomada", "proxima_retomada"]].merge(base, on="contrato", how="inner")
    # O resultado deve pertencer a esta retomada, e não a uma retomada anterior.
    valid = matches["data_resultado"].ge(matches["data_retomada"]) & (
        matches["proxima_retomada"].isna() | matches["data_resultado"].lt(matches["proxima_retomada"])
    )
    matches = matches[valid].copy()
    matches["pago"] = (matches["status_key"].eq("PAGOU") | matches["data_pagamento"].notna()) & (
        matches["data_pagamento"].isna() | matches["data_pagamento"].ge(matches["data_retomada"])
    )
    matches["aberto"] = matches["status_key"].eq("EMABERTO") & ~matches["pago"]
    matches["nao_pagou"] = matches["status_key"].eq("NAOPAGOU") & ~matches["pago"]
    matches["valor_pago"] = matches["valor"].where(matches["pago"], 0)
    totals = matches.groupby("caso_id").agg(
        registros_base=("contrato", "size"), pago=("pago", "any"),
        aberto=("aberto", "any"), nao_pagou=("nao_pagou", "any"),
        valor_pago=("valor_pago", "sum"),
        status_base=("status_original", lambda s: ", ".join(sorted(set(s)))),
    )
    cases = cases.join(totals, on="caso_id")
    for flag in ["pago", "aberto", "nao_pagou"]:
        cases[flag] = cases[flag].eq(True)
    cases["registros_base"] = cases["registros_base"].fillna(0).astype(int)
    cases["valor_pago"] = cases["valor_pago"].fillna(0)
    cases["status_base"] = cases["status_base"].fillna("Sem resultado de pós localizado")
    cases["tem_acordo"] = cases["registros_base"].gt(0)
    cases["situacao"] = np.select(
        [cases["pago"], cases["aberto"], cases["nao_pagou"], cases["registros_base"].gt(0),
         cases["venda_apreensao"].eq("Sim")],
        ["Pago", "Em aberto", "Não pagou", "Outro status na BASE", "Vendido"],
        default="Sem acordo na BASE",
    )
    cases["conferencia"] = ""
    checks = [
        (cases["data_retomada"].isna(), "Data de retomada inválida ou ausente"),
        (cases["data_retomada"].notna() & cases["mes_num"].ne(cases["data_retomada"].dt.month), "Mês diferente da data de retomada"),
        (cases["registros_base"].gt(0) & cases["acordo_apreensao"].ne("Sim"), "Pós na BASE sem ACORDO = SIM em APREENSÃO"),
        (cases["registros_base"].eq(0) & cases["acordo_apreensao"].eq("Sim"), "ACORDO = SIM em APREENSÃO sem pós confirmado na BASE"),
        (cases["tipo"].eq("D.A"), "Caso D.A: fora dos indicadores JUDICIAL"),
    ]
    for mask, message in checks:
        cases.loc[mask, "conferencia"] += message + "; "
    cases["conferencia"] = cases["conferencia"].str.rstrip("; ")
    return cases.sort_values(["mes_num", "data_retomada", "contrato"], na_position="last").reset_index(drop=True)


@st.cache_data(show_spinner=False, max_entries=2)
def load_pos_retomado_cases(path, version):
    """Carrega apreensões e confirma acordos/pagamentos somente pela aba BASE."""
    with pd.ExcelFile(path) as book:
        sheet = next((name for name in book.sheet_names if key(name) == "APREENSAO"), None)
        if sheet is None:
            raise ValueError("A aba APREENSÃO não foi encontrada na base de resultados.")
        apreensoes = pd.read_excel(book, sheet_name=sheet, header=2, usecols=range(20))
        base_cols = {"NOCONTRATO", "EMISSAO", "DATADOPAGAMENTO", "STATUS", "CAMPANHA", "VALORDOBANCOMETA"}
        resultados = pd.read_excel(book, sheet_name="BASE", usecols=lambda col: key(col) in base_cols)
    return prepare_pos_retomado_cases(apreensoes, resultados)


def pos_retomado_summary(cases, converted="pago"):
    judicial = cases[cases["tipo"].eq("JUDICIAL")].copy()
    monthly = judicial.groupby(["mes_num", "mes"], dropna=False, observed=True).agg(
        casos=("caso_id", "size"), convertidos=(converted, "sum"),
        acordos=("tem_acordo", "sum"), pagos=("pago", "sum"), valor_pago=("valor_pago", "sum"),
    ).reset_index().sort_values("mes_num")
    monthly["ticket_medio_pos"] = monthly["valor_pago"] / monthly["pagos"].replace(0, np.nan)
    monthly["pct_conversao"] = monthly["convertidos"] / monthly["casos"]
    statuses = judicial.groupby(["mes_num", "mes", "situacao"], dropna=False, observed=True).size().rename("casos").reset_index()
    statuses["pct_mes"] = statuses["casos"] / statuses.groupby(["mes_num", "mes"], dropna=False)["casos"].transform("sum")
    consolidated = judicial.groupby("situacao", observed=True).size().rename("casos").reset_index()
    consolidated["pct_total"] = consolidated["casos"] / len(judicial) if len(judicial) else np.nan
    return monthly, statuses, consolidated.sort_values("casos", ascending=False)


def show_case_table(cases, download_key):
    view = cases[list(CASE_COLUMNS)].rename(columns=CASE_COLUMNS).copy()
    for col in ["Acordo fechado", "Pagamento de pós"]:
        view[col] = view[col].map({True: "Sim", False: "Não"})
    st.dataframe(view.style.format({"Valor pago de pós": format_brl}), hide_index=True, width="stretch", column_config={
        "Data da retomada": st.column_config.DateColumn(format="DD/MM/YYYY"),
    })
    st.download_button("Baixar casos em CSV", view.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                       file_name=f"pos_retomado_{download_key}.csv", mime="text/csv", key=f"download_pos_{download_key}")


def render_pos_retomado(path):
    path = Path(path)
    st.title("Pós Retomado")
    st.caption("Casos da aba APREENSÃO cruzados com a campanha Pós Retomado da aba BASE, por contrato e data da retomada.")
    try:
        with st.spinner("Conferindo apreensões e resultados de pós..."):
            cases = load_pos_retomado_cases(str(path), path.stat().st_mtime_ns)
    except (OSError, ValueError, KeyError) as error:
        st.error(f"Não foi possível carregar a análise de pós-retomado: {error}")
        return

    st.sidebar.title("Filtros do pós-retomado")
    filter_keys = ["pos_months", "pos_regions", "pos_ufs"]
    if st.sidebar.button("Limpar filtros do pós", width="stretch"):
        for name in filter_keys:
            st.session_state.pop(name, None)
        st.rerun()
    months = cases["mes"].drop_duplicates().tolist()
    selected_months = st.sidebar.multiselect("Mês da apreensão", months, key="pos_months")
    regions = st.sidebar.multiselect("Região", sorted(cases["regiao"].unique()), key="pos_regions")
    ufs = st.sidebar.multiselect("UF", sorted(cases["uf"].unique()), key="pos_ufs")
    for col, values in [("mes", selected_months), ("regiao", regions), ("uf", ufs)]:
        if values:
            cases = cases[cases[col].isin(values)]
    st.caption("Período: " + (", ".join(selected_months) if selected_months else "todos os meses da apreensão") +
               ". Indicadores: somente JUDICIAL. Casos D.A destacados ao final.")
    criterion = st.radio("Critério de conversão", ["Pagamento realizado", "Acordo fechado"], horizontal=True, key="pos_conversion")
    converted = "pago" if criterion == "Pagamento realizado" else "tem_acordo"
    judicial = cases[cases["tipo"].eq("JUDICIAL")]
    da = cases[cases["tipo"].eq("D.A")]
    monthly, statuses, consolidated = pos_retomado_summary(cases, converted)
    total = len(judicial)
    converted_count = int(judicial[converted].sum())
    pct = lambda value: f"{value:.1%}".replace(".", ",") if pd.notna(value) else "—"
    cards = st.columns(5)
    for col, label, value in zip(cards, ["Casos judiciais", "Casos convertidos", "% de conversão", "Com acordo fechado", "Casos D.A à parte"],
                                 [total, converted_count, pct(converted_count / total if total else np.nan), int(judicial["tem_acordo"].sum()), len(da)]):
        col.metric(label, value)
    st.caption("Conversão = casos convertidos ÷ casos JUDICIAL do recorte. Cada retomada conta uma vez, mesmo com várias parcelas.")
    with st.expander("Como os casos e as situações são apurados"):
        st.write("Pagamento realizado: status PAGOU ou data de pagamento na campanha de pós da BASE. "
                 "Acordo fechado: resultado de pós localizado exclusivamente na aba BASE, pelo contrato, após a retomada. "
                 "A coluna ACORDO de APREENSÃO serve apenas para conferência e não confirma conversão. "
                 "São usados o mês cadastrado em APREENSÃO e os desfechos disponíveis na atualização da planilha, "
                 "mesmo quando o pagamento ocorreu em mês posterior.")
        st.write("Situação por caso: Pago; se não houver pagamento, Em aberto; depois Não pagou; "
                 "outro status na BASE; vendido; sem acordo na BASE. "
                 "As situações são exclusivas e seus percentuais somam 100% em cada mês. "
                 "NEGOCIAÇÃO é exibida como responsável cadastrado, sem presumir um status de andamento.")
        st.write("Resultados anteriores à retomada e campanhas diferentes de pós não entram. "
                 "Sem data válida da retomada, o pagamento não pode ser confirmado pelo cruzamento. "
                 "Casos D.A não entram nos indicadores de JUDICIAL.")

    if judicial.empty:
        st.info("Nenhum caso JUDICIAL nos filtros selecionados.")
    else:
        st.subheader("Conversão por mês da apreensão")
        chart = alt.Chart(monthly).mark_bar(color="#2f6f73").encode(
            x=alt.X("mes:N", title="Mês da apreensão", sort=monthly["mes"].tolist()),
            y=alt.Y("pct_conversao:Q", title="Conversão", axis=alt.Axis(format=".0%"), scale=alt.Scale(domain=[0, 1])),
            tooltip=[alt.Tooltip("mes:N", title="Mês"), alt.Tooltip("casos:Q", title="Casos judiciais"),
                     alt.Tooltip("convertidos:Q", title="Convertidos"), alt.Tooltip("pct_conversao:Q", title="Conversão", format=".1%")],
        ).properties(height=260)
        st.altair_chart(chart, width="stretch")
        table = monthly.drop(columns="mes_num").rename(columns={"mes": "Mês", "casos": "Casos judiciais", "convertidos": "Convertidos",
            "acordos": "Com acordo", "pagos": "Com pagamento", "valor_pago": "Valor pago de pós",
            "ticket_medio_pos": "Ticket médio de pós", "pct_conversao": "% conversão"})
        table["% conversão"] = table["% conversão"].map(pct)
        st.dataframe(table.style.format({"Valor pago de pós": format_brl, "Ticket médio de pós": format_brl}), hide_index=True, width="stretch")
        st.caption("Ticket médio de pós = valor pago de pós ÷ casos com pagamento em cada mês da apreensão. Meses sem pagamento aparecem como —.")

        st.subheader("Situação dos casos judiciais")
        left, right = st.columns([2, 3])
        with left:
            summary = consolidated.rename(columns={"situacao": "Situação", "casos": "Casos", "pct_total": "% do total"})
            summary["% do total"] = summary["% do total"].map(pct)
            st.dataframe(summary, hide_index=True, width="stretch")
        with right:
            chart = alt.Chart(statuses).mark_bar().encode(
                x=alt.X("mes:N", sort=monthly["mes"].tolist(), title="Mês da apreensão"),
                y=alt.Y("pct_mes:Q", title="% dos casos do mês", axis=alt.Axis(format=".0%")),
                color=alt.Color("situacao:N", title="Situação"),
                tooltip=[alt.Tooltip("mes:N", title="Mês"), alt.Tooltip("situacao:N", title="Situação"),
                         alt.Tooltip("casos:Q", title="Casos"), alt.Tooltip("pct_mes:Q", title="% do mês", format=".1%")],
            ).properties(height=280)
            st.altair_chart(chart, width="stretch")
        detail = statuses.drop(columns="mes_num").rename(columns={"mes": "Mês", "situacao": "Situação", "casos": "Casos", "pct_mes": "% do mês"})
        detail["% do mês"] = detail["% do mês"].map(pct)
        st.dataframe(detail, hide_index=True, width="stretch")

        mismatches = judicial["tem_acordo"].ne(judicial["acordo_apreensao"].eq("Sim"))
        if mismatches.any():
            st.info(f"{int(mismatches.sum())} casos têm divergência entre a marcação ACORDO de APREENSÃO e o cruzamento com a BASE. Os indicadores de acordo e pagamento usam somente a BASE. Veja a conferência por caso abaixo.")
        date_issues = judicial["conferencia"].str.contains("Data de retomada|Mês diferente", regex=True)
        if date_issues.any():
            st.warning(f"{int(date_issues.sum())} casos precisam de conferência de mês/data da retomada. Os registros estão sinalizados no detalhe.")
        st.subheader("Situação de cada caso JUDICIAL")
        chosen = st.multiselect("Filtrar o detalhe por situação", sorted(judicial["situacao"].unique()), key="pos_case_status")
        show_case_table(judicial[judicial["situacao"].isin(chosen)] if chosen else judicial, "judicial")

    st.subheader("Casos D.A destacados")
    st.caption("D.A identificados na coluna JUDICIAL / DA da apreensão, excluídos dos totais e percentuais judiciais acima.")
    if da.empty:
        st.info("Nenhum caso D.A nos filtros selecionados.")
    else:
        da_monthly = da.groupby(["mes_num", "mes"], dropna=False).size().rename("Casos D.A").reset_index().sort_values("mes_num")
        st.dataframe(da_monthly[["mes", "Casos D.A"]].rename(columns={"mes": "Mês"}), hide_index=True, width="stretch")
        show_case_table(da, "da")
    other = cases[~cases["tipo"].isin(["JUDICIAL", "D.A"])]
    if not other.empty:
        st.warning(f"{len(other)} casos com classificação diferente de JUDICIAL ou D.A foram excluídos dos indicadores.")
    st.caption(f"Base atualizada em {pd.Timestamp.fromtimestamp(path.stat().st_mtime):%d/%m/%Y %H:%M}.")
