import io
import unicodedata
from datetime import datetime, date

import pandas as pd
import streamlit as st

# =====================================================
# CONSTANTES DE COLUNAS
# =====================================================

COL_DATA = "DATA ORIGINAL"
COL_DEP = "DEP."
COL_MOV = "TMV"
COL_QTD = "QUANTIDADE"
COL_PROC = "PROCESSO"

# =====================================================
# ABAS QUE SERÃO PROCESSADAS
# =====================================================

ABAS_PERMITIDAS = [
    "AnteLiga",
    "Metal SF",
    "Escoria",
    "Consumo Sucata",
    "Metal Fusao",
    "Solidos",
    "Geracao Sucata Fundicao",
]

# =====================================================
# CSS / LAYOUT
# =====================================================

st.set_page_config(page_title="Indicadores Fundição", layout="wide")

st.markdown("""
<style>
section[data-testid="stSidebar"] { background-color: #113CEA; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: white; }
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label { color: white !important; font-weight: 500; }
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] select,
section[data-testid="stSidebar"] textarea,
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: white !important; color: black !important; border-radius: 8px;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] span { color: black !important; }
div[data-baseweb="popover"] span { color: black !important; }
section[data-testid="stSidebar"] div[role="radiogroup"] * { color: white !important; }
section[data-testid="stSidebar"] button {
    background-color: #0D2BC7; color: white !important; font-weight: bold; border-radius: 10px;
}
.kpi-card {
    background-color: white; padding: 20px; border-radius: 16px;
    box-shadow: 0 6px 16px rgba(0,0,0,0.08); text-align: center;
}
.kpi-title { font-size: 14px; color: #6b7280; }
.kpi-value { font-size: 28px; font-weight: 700; color: #113CEA; }
.badge {
    display: inline-block; padding: 8px 14px; margin: 4px; border-radius: 20px;
    background-color: #EEF2FF; color: #113CEA; font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================


def normalizar(texto):
    texto = str(texto).upper().strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.replace(" ", "")


def _tratar_dataframe(df):
    """Aplica toda a limpeza/tipagem em uma aba recém-lida do Excel."""

    df = df.dropna(how="all").dropna(axis=1, how="all")

    df.columns = df.columns.astype(str).str.strip().str.upper()

    colunas_obrigatorias = [COL_DATA, COL_QTD, COL_PROC]
    colunas_faltantes = [c for c in colunas_obrigatorias if c not in df.columns]

    if colunas_faltantes:
        return None, colunas_faltantes

    df[COL_QTD] = pd.to_numeric(
        df[COL_QTD].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    ) * -1 / 1000

    df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce")

    if COL_DEP not in df.columns:
        df[COL_DEP] = pd.NA

    if COL_MOV not in df.columns:
        df[COL_MOV] = pd.NA

    return df, None


# =====================================================
# CARGA DAS ABAS A PARTIR DO ARQUIVO ENVIADO
# =====================================================

@st.cache_data(show_spinner=False)
def carregar_todas_abas(conteudo_excel):
    """Lê o Excel enviado (bytes) uma única vez e trata só as abas necessárias."""

    dfs = {}
    avisos = []

    xls = pd.ExcelFile(io.BytesIO(conteudo_excel), engine="openpyxl")
    mapa_abas = {normalizar(aba): aba for aba in xls.sheet_names}

    for aba_desejada in ABAS_PERMITIDAS:

        chave = normalizar(aba_desejada)

        if chave not in mapa_abas:
            avisos.append(
                ("warning", f"⚠️ A aba **{aba_desejada}** não foi encontrada no arquivo Excel.")
            )
            continue

        aba_real = mapa_abas[chave]

        df = xls.parse(sheet_name=aba_real)
        df, colunas_faltantes = _tratar_dataframe(df)

        if colunas_faltantes:
            avisos.append((
                "warning",
                f"⚠️ A aba **{aba_real}** foi ignorada porque não possui as colunas: "
                f"**{', '.join(colunas_faltantes)}**",
            ))
            continue

        dfs[aba_desejada] = df

    return dfs, avisos


# =====================================================
# ANÁLISE
# =====================================================

def analisar(df, dt_ini, dt_fim):

    df = df[
        (df[COL_DATA] >= dt_ini)
        & (df[COL_DATA] <= dt_fim)
        & (df[COL_PROC] != "APONTAMENTO MANUAL")
    ]

    produtivo = df[df[COL_PROC].notna() & (df[COL_PROC] != "NÃO ENCONTRADO")]

    processos = (
        produtivo
        .groupby(COL_PROC, as_index=False)[COL_QTD]
        .sum()
        .sort_values(COL_QTD, ascending=False)
    )

    total_prod = processos[COL_QTD].sum()
    total_nao = df[df[COL_PROC] == "NÃO ENCONTRADO"][COL_QTD].sum()
    total_geral = total_prod + total_nao
    perc_nao = (total_nao / total_geral * 100) if total_geral else 0

    depositos = (
        pd.to_numeric(df[COL_DEP], errors="coerce").dropna().astype(int).unique().tolist()
    )
    movimentos = (
        pd.to_numeric(df[COL_MOV], errors="coerce").dropna().astype(int).unique().tolist()
    )

    return (
        processos,
        sorted(depositos),
        sorted(movimentos),
        total_prod,
        total_nao,
        total_geral,
        perc_nao,
    )


def kpi(titulo, valor):
    return f"""
    <div class='kpi-card'>
        <div class='kpi-title'>{titulo}</div>
        <div class='kpi-value'>{valor}</div>
    </div>
    """


# =====================================================
# UI — UPLOAD DO ARQUIVO
# =====================================================

st.title("📊 Indicadores de Movimentação — Fundição")

st.sidebar.title("📁 Arquivo")

arquivo = st.sidebar.file_uploader(
    "Selecione o arquivo 'Informações SAP'",
    type=["xlsm", "xlsx"],
)

if arquivo is None:
    st.info("👈 Selecione o arquivo **Informações SAP.xlsm** na barra lateral para começar. "
            "O arquivo é processado no seu navegador e não é enviado a nenhum servidor.")
    st.session_state.pop("dados_abas", None)
    st.session_state.pop("arquivo_id", None)
    st.stop()

# Recarrega as bases se o usuário trocar de arquivo
arquivo_id = f"{arquivo.name}_{arquivo.size}"

if st.session_state.get("arquivo_id") != arquivo_id:
    st.session_state.pop("dados_abas", None)
    st.session_state["arquivo_id"] = arquivo_id

if "dados_abas" not in st.session_state:

    with st.spinner("🔄 Carregando bases para análise (pode levar alguns instantes)..."):
        st.session_state.dados_abas, avisos = carregar_todas_abas(arquivo.getvalue())

    for tipo, msg in avisos:
        getattr(st, tipo)(msg)

    st.success(f"✅ Bases carregadas com sucesso ({len(st.session_state.dados_abas)} abas)")

if not st.session_state.dados_abas:
    st.error("❌ Nenhuma das abas configuradas foi encontrada ou possui as colunas necessárias.")
    st.stop()

# =====================================================
# SIDEBAR — FILTROS
# =====================================================

st.sidebar.title("🎛️ Filtros")

abas = list(st.session_state.dados_abas.keys())
aba = st.sidebar.selectbox("📄 Aba", abas)
df_base = st.session_state.dados_abas[aba]

tipo_filtro = st.sidebar.radio(
    "📅 Tipo de período",
    ["Ano completo", "Trimestre", "Período personalizado"],
)

anos_disponiveis = sorted(df_base[COL_DATA].dropna().dt.year.unique().tolist())

if not anos_disponiveis:
    st.error(f"❌ A aba **{aba}** não possui datas válidas.")
    st.stop()

if tipo_filtro == "Ano completo":

    ano = st.sidebar.selectbox("Ano", anos_disponiveis)
    dt_ini = datetime(ano, 1, 1)
    dt_fim = datetime(ano, 12, 31)

elif tipo_filtro == "Trimestre":

    ano = st.sidebar.selectbox("Ano", anos_disponiveis)
    trimestre = st.sidebar.selectbox(
        "Trimestre",
        [
            "1º Trimestre (Jan–Mar)",
            "2º Trimestre (Abr–Jun)",
            "3º Trimestre (Jul–Set)",
            "4º Trimestre (Out–Dez)",
        ],
    )

    if trimestre.startswith("1º"):
        dt_ini, dt_fim = datetime(ano, 1, 1), datetime(ano, 3, 31)
    elif trimestre.startswith("2º"):
        dt_ini, dt_fim = datetime(ano, 4, 1), datetime(ano, 6, 30)
    elif trimestre.startswith("3º"):
        dt_ini, dt_fim = datetime(ano, 7, 1), datetime(ano, 9, 30)
    else:
        dt_ini, dt_fim = datetime(ano, 10, 1), datetime(ano, 12, 31)

else:

    dt_ini = pd.to_datetime(
        st.sidebar.date_input("Data inicial", value=date.today().replace(month=1, day=1))
    )
    dt_fim = pd.to_datetime(st.sidebar.date_input("Data final", value=date.today()))

executar = st.sidebar.button("▶️ Analisar")

with st.sidebar.expander("⚙️ Avançado"):
    if st.button("🗑️ Limpar cache e recarregar"):
        st.cache_data.clear()
        st.session_state.pop("dados_abas", None)
        st.rerun()

# =====================================================
# RESULTADOS
# =====================================================

if executar:

    (
        processos,
        depositos,
        movimentos,
        total_prod,
        total_nao,
        total_geral,
        perc_nao,
    ) = analisar(df_base, dt_ini, dt_fim)

    st.subheader("📌 Visão Geral")

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi("Total Produtivo", f"{total_prod:,.2f}"), unsafe_allow_html=True)
    k2.markdown(kpi("Não Encontrado", f"{total_nao:,.2f}"), unsafe_allow_html=True)
    k3.markdown(kpi("Total Geral", f"{total_geral:,.2f}"), unsafe_allow_html=True)
    k4.markdown(kpi("% Sem Classificação", f"{perc_nao:.2f}%"), unsafe_allow_html=True)

    st.divider()

    st.subheader("📦 Depósitos")
    if depositos:
        st.markdown(
            "".join(f"<span class='badge'>{d}</span>" for d in depositos),
            unsafe_allow_html=True,
        )
    else:
        st.info("Nenhum depósito encontrado no período.")

    st.subheader("🔄 Movimentos (TMV)")
    if movimentos:
        st.markdown(
            "".join(f"<span class='badge'>{m}</span>" for m in movimentos),
            unsafe_allow_html=True,
        )
    else:
        st.info("Nenhum movimento encontrado no período.")

    st.divider()

    st.subheader("📊 Volume por Processo")
    st.dataframe(processos, use_container_width=True)

    st.divider()

    st.subheader("📋 Base de Dados Completa")

    df_filtro = df_base[(df_base[COL_DATA] >= dt_ini) & (df_base[COL_DATA] <= dt_fim)]
    st.dataframe(df_filtro, use_container_width=True)

    # =================================================
    # ANÁLISES EXCLUSIVAS DA ABA SOLIDOS
    # =================================================

    if normalizar(aba) == "SOLIDOS":

        col_material = None
        for c in ["DESCRIÇÃO MATERIAL", "DENOMINAÇÃO", "MATERIAL"]:
            if c in df_base.columns:
                col_material = c
                break

        if col_material:

            st.subheader("📊 Lingote Importado FCI e MP 3º (t)")

            material_upper = df_filtro[col_material].fillna("").astype(str).str.upper()

            df_solidos = df_filtro[
                material_upper.str.startswith("LINGOTE AL")
                & df_filtro[COL_PROC].notna()
                & (df_filtro[COL_PROC] != "NÃO ENCONTRADO")
            ]

            if not df_solidos.empty:

                processos_solidos = (
                    df_solidos
                    .groupby(COL_PROC, as_index=False)[COL_QTD]
                    .sum()
                    .sort_values(COL_QTD, ascending=False)
                )

                st.dataframe(processos_solidos, use_container_width=True)

                total_solidos = processos_solidos[COL_QTD].sum()
                st.markdown(
                    kpi("Total Volume LINGOTE AL", f"{total_solidos:,.2f}"),
                    unsafe_allow_html=True,
                )
            else:
                st.info("Nenhum registro encontrado para materiais que começam com 'LINGOTE AL'.")

        if "DENOMINAÇÃO" in df_filtro.columns:

            denominacao_upper = df_filtro["DENOMINAÇÃO"].fillna("").astype(str).str.upper()

            df_sf = df_filtro[
                denominacao_upper.str.startswith("SF")
                & df_filtro[COL_PROC].notna()
                & (df_filtro[COL_PROC] != "NÃO ENCONTRADO")
            ]

            if not df_sf.empty:

                st.subheader("📊 Outros Sólidos (SF) por Processo Lingote beneficiamento")

                processos_sf = (
                    df_sf
                    .groupby(COL_PROC, as_index=False)[COL_QTD]
                    .sum()
                    .sort_values(COL_QTD, ascending=False)
                )

                st.dataframe(processos_sf, use_container_width=True)

                total_sf = processos_sf[COL_QTD].sum()
                st.markdown(kpi("Total Consumo SF", f"{total_sf:,.2f}"), unsafe_allow_html=True)
            else:
                st.info("Nenhum consumo encontrado para materiais SF.")