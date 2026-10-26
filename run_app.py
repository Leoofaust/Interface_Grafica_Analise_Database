import io
import unicodedata
from datetime import datetime, date
import pandas as pd
import streamlit as st

# =====================================================
# CONSTANTES DE COLUNAS — INFORMAÇÕES SAP
# =====================================================

COL_DATA = "DATA ORIGINAL"
COL_DEP = "DEP."
COL_MOV = "TMV"
COL_QTD = "QUANTIDADE"
COL_PROC = "PROCESSO"

# =====================================================
# CONSTANTES DE COLUNAS — DADOS GN (GÁS)
# =====================================================

GN_ABA = "Dados GN"
GN_DATA = "DATA"
GN_MAQ = "MÁQUINA"
GN_CONSUMO = "CONSUMO"
GN_FILTRO = "FILTRO DE MÁQUINA"
GN_PROC = "PROCESSO"
GN_DESC = "DESCRITIVOS"

# =====================================================
# ABAS QUE SERÃO PROCESSADAS (SAP)
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
# FUNÇÕES AUXILIARES (COMPARTILHADAS)
# =====================================================


def normalizar(texto):
    texto = str(texto).upper().strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.replace(" ", "")


def kpi(titulo, valor):
    return f"""
    <div class='kpi-card'>
        <div class='kpi-title'>{titulo}</div>
        <div class='kpi-value'>{valor}</div>
    </div>
    """


def filtros_periodo(anos_disponiveis, prefixo):
    """Renderiza os filtros de período na sidebar e devolve (dt_ini, dt_fim)."""

    tipo_filtro = st.sidebar.radio(
        "📅 Tipo de período",
        ["Ano completo", "Trimestre", "Período personalizado"],
        key=f"{prefixo}_tipo",
    )

    if tipo_filtro == "Ano completo":

        ano = st.sidebar.selectbox("Ano", anos_disponiveis, key=f"{prefixo}_ano")
        return datetime(ano, 1, 1), datetime(ano, 12, 31)

    if tipo_filtro == "Trimestre":

        ano = st.sidebar.selectbox("Ano", anos_disponiveis, key=f"{prefixo}_ano")
        trimestre = st.sidebar.selectbox(
            "Trimestre",
            [
                "1º Trimestre (Jan–Mar)",
                "2º Trimestre (Abr–Jun)",
                "3º Trimestre (Jul–Set)",
                "4º Trimestre (Out–Dez)",
            ],
            key=f"{prefixo}_tri",
        )

        if trimestre.startswith("1º"):
            return datetime(ano, 1, 1), datetime(ano, 3, 31)
        if trimestre.startswith("2º"):
            return datetime(ano, 4, 1), datetime(ano, 6, 30)
        if trimestre.startswith("3º"):
            return datetime(ano, 7, 1), datetime(ano, 9, 30)
        return datetime(ano, 10, 1), datetime(ano, 12, 31)

    dt_ini = pd.to_datetime(
        st.sidebar.date_input(
            "Data inicial",
            value=date.today().replace(month=1, day=1),
            key=f"{prefixo}_dini",
        )
    )
    dt_fim = pd.to_datetime(
        st.sidebar.date_input("Data final", value=date.today(), key=f"{prefixo}_dfim")
    )
    return dt_ini, dt_fim


# =====================================================
# ============  INFORMAÇÕES SAP  ======================
# =====================================================


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


def pagina_sap():

    st.title("📊 Indicadores de Movimentação — Fundição")

    st.sidebar.title("📁 Arquivo")

    arquivo = st.sidebar.file_uploader(
        "Selecione o arquivo 'Informações SAP'",
        type=["xlsm", "xlsx"],
        key="upload_sap",
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

    # ---------------- SIDEBAR — FILTROS ----------------

    st.sidebar.title("🎛️ Filtros")

    abas = list(st.session_state.dados_abas.keys())
    aba = st.sidebar.selectbox("📄 Aba", abas, key="sap_aba")
    df_base = st.session_state.dados_abas[aba]

    anos_disponiveis = sorted(df_base[COL_DATA].dropna().dt.year.unique().tolist())

    if not anos_disponiveis:
        st.error(f"❌ A aba **{aba}** não possui datas válidas.")
        st.stop()

    dt_ini, dt_fim = filtros_periodo(anos_disponiveis, "sap")

    executar = st.sidebar.button("▶️ Analisar", key="exec_sap")

    with st.sidebar.expander("⚙️ Avançado"):
        if st.button("🗑️ Limpar cache e recarregar", key="limpar_sap"):
            st.cache_data.clear()
            st.session_state.pop("dados_abas", None)
            st.rerun()

    # ---------------- RESULTADOS ----------------

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

        # ----------- ANÁLISES EXCLUSIVAS DA ABA SOLIDOS -----------

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


# =====================================================
# ============  DADOS GN (GÁS)  =======================
# =====================================================


def _converter_numero(serie):
    """Converte para número aceitando formatos '1234.5', '1234,5' e '1.234,5'."""

    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")

    def _conv(v):
        if pd.isna(v):
            return None
        s = str(v).strip()
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        return s

    return pd.to_numeric(serie.map(_conv), errors="coerce")


@st.cache_data(show_spinner=False)
def carregar_dados_gn(conteudo_excel):
    """Lê a aba 'Dados GN' do Excel de gás e faz a limpeza/tipagem."""

    xls = pd.ExcelFile(io.BytesIO(conteudo_excel), engine="openpyxl")
    mapa_abas = {normalizar(aba): aba for aba in xls.sheet_names}

    chave = normalizar(GN_ABA)
    if chave not in mapa_abas:
        return None, f"A aba **{GN_ABA}** não foi encontrada no arquivo Excel."

    df = xls.parse(sheet_name=mapa_abas[chave])
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = df.columns.astype(str).str.strip().str.upper()

    obrigatorias = [GN_DATA, GN_CONSUMO, GN_PROC]
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        return None, f"A aba **{GN_ABA}** não possui as colunas: **{', '.join(faltantes)}**"

    df[GN_CONSUMO] = _converter_numero(df[GN_CONSUMO])
    df[GN_DATA] = pd.to_datetime(df[GN_DATA], errors="coerce", dayfirst=True)
    df[GN_PROC] = df[GN_PROC].fillna("NÃO INFORMADO").astype(str).str.strip()

    if GN_MAQ not in df.columns:
        df[GN_MAQ] = pd.NA

    return df, None


def pagina_gn():

    st.title("🔥 Indicadores de Consumo de Gás — Dados GN")

    st.sidebar.title("📁 Arquivo")

    arquivo = st.sidebar.file_uploader(
        "Selecione o arquivo 'Dados GN'",
        type=["xlsm", "xlsx"],
        key="upload_gn",
    )

    if arquivo is None:
        st.info("👈 Selecione o arquivo **Dados GN** na barra lateral para começar. "
                "O arquivo precisa ter a aba **Dados GN**.")
        st.stop()

    with st.spinner("🔄 Carregando dados de gás..."):
        df_base, erro = carregar_dados_gn(arquivo.getvalue())

    if erro:
        st.error(f"❌ {erro}")
        st.stop()

    anos_disponiveis = sorted(df_base[GN_DATA].dropna().dt.year.unique().tolist())

    if not anos_disponiveis:
        st.error(f"❌ A aba **{GN_ABA}** não possui datas válidas.")
        st.stop()

    st.sidebar.title("🎛️ Filtros")

    dt_ini, dt_fim = filtros_periodo(anos_disponiveis, "gn")

    executar = st.sidebar.button("▶️ Analisar", key="exec_gn")

    with st.sidebar.expander("⚙️ Avançado"):
        if st.button("🗑️ Limpar cache e recarregar", key="limpar_gn"):
            st.cache_data.clear()
            st.rerun()

    if not executar:
        return

    # Compara pela data (sem hora) para incluir o último dia inteiro
    datas = df_base[GN_DATA].dt.normalize()
    df_filtro = df_base[(datas >= dt_ini) & (datas <= dt_fim)]

    if df_filtro.empty:
        st.info("Nenhum registro de gás encontrado no período.")
        return

    processos = (
        df_filtro
        .groupby(GN_PROC, as_index=False)[GN_CONSUMO]
        .sum()
        .sort_values(GN_CONSUMO, ascending=False)
    )

    total_consumo = df_filtro[GN_CONSUMO].sum()
    qtd_registros = len(df_filtro)
    qtd_maquinas = df_filtro[GN_MAQ].nunique(dropna=True)
    qtd_processos = len(processos)

    st.subheader("📌 Visão Geral")

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi("Consumo Total de Gás", f"{total_consumo:,.2f}"), unsafe_allow_html=True)
    k2.markdown(kpi("Registros", f"{qtd_registros:,}"), unsafe_allow_html=True)
    k3.markdown(kpi("Máquinas", f"{qtd_maquinas:,}"), unsafe_allow_html=True)
    k4.markdown(kpi("Processos", f"{qtd_processos:,}"), unsafe_allow_html=True)

    st.divider()

    st.subheader("📊 Volume de Gás por Processo")

    c1, c2 = st.columns([1, 1])
    c1.dataframe(processos, use_container_width=True)
    c2.bar_chart(processos.set_index(GN_PROC)[GN_CONSUMO])


# =====================================================
# SELETOR DE FONTE DE DADOS
# =====================================================

st.sidebar.title("🗂️ Fonte de dados")

fonte = st.sidebar.radio(
    "Escolha a fonte",
    ["Informações SAP", "Dados GN"],
    key="fonte_dados",
)

if fonte == "Informações SAP":
    pagina_sap()
else:
    pagina_gn()
