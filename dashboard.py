import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# ---------------------------------------------------------
# CONFIGURAÇÃO BASE DO APP
# ---------------------------------------------------------
st.set_page_config(
    page_title="ContaUni - Indicadores RH",
    layout="wide",
)

st.markdown("""
<style>
.block-container {
    padding-top: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# CAMINHOS DOS ARQUIVOS (CORREÇÃO FINAL)
# ---------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DADOS_DIR = BASE_DIR / "dados"   # CORRETO PARA DEPLOY NO STREAMLIT CLOUD

def ler_excel(nome_arquivo: str) -> pd.DataFrame:
    """Lê um arquivo Excel da pasta dados com segurança."""
    caminho = DADOS_DIR / nome_arquivo
    try:
        df = pd.read_excel(caminho)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar {nome_arquivo}: {e}")
        return pd.DataFrame()

# ---------------------------------------------------------
# CARREGANDO OS ARQUIVOS
# ---------------------------------------------------------
admissoes = ler_excel("admissoes.2025.xlsx")
demissoes = ler_excel("demissoes.rescisoes.2025.xlsx")
exames = ler_excel("exames.2025.xlsx")
epi = ler_excel("epi.uniformes.2025.xlsx")
produtores = ler_excel("produtores.2025.xlsx")
adt13 = ler_excel("adt.13.ferias.xlsx")

# ---------------------------------------------------------
# TRATAMENTO DE DADOS (PADRONIZAÇÃO)
# ---------------------------------------------------------
def tratar(df):
    if df.empty:
        return df
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
    return df

admissoes = tratar(admissoes)
demissoes = tratar(demissoes)
exames = tratar(exames)
epi = tratar(epi)
produtores = tratar(produtores)
adt13 = tratar(adt13)

# ---------------------------------------------------------
# UNIFICAÇÃO BASE GERAL
# ---------------------------------------------------------
def montar_base_indicadores():

    # Admissões
    adm = admissoes.copy()
    if not adm.empty:
        adm["TipoMovimento"] = "Admissao"
        adm["ValorLiquido"] = 0

    # Demissões / Rescisões
    dem = demissoes.copy()
    if not dem.empty:
        dem["TipoMovimento"] = "Demissao"
        dem["ValorLiquido"] = dem.get("ValorLiquido", 0)
        dem["MultaFGTS"] = dem.get("ValorMultaFGTS", 0)

    # EPI e Uniformes
    epi_df = epi.copy()
    if not epi_df.empty:
        epi_df["TipoMovimento"] = "EPI"
        epi_df.rename(columns={"Valor": "ValorEPI"}, inplace=True)

    # Exames
    exams = exames.copy()
    if not exams.empty:
        exams["TipoMovimento"] = "Exame"
        exams.rename(columns={"Valor": "ValorExame"}, inplace=True)

    # ADT13, 13º e Férias
    adt = adt13.copy()
    if not adt.empty:
        adt["TipoMovimento"] = adt["TipoLancamento"]

        adt["ValorADT13"] = adt["ValorLiquido"].where(adt["TipoLancamento"] == "ADT13", 0)
        adt["Valor13"] = adt["ValorLiquido"].where(adt["TipoLancamento"] == "13", 0)
        adt["ValorFerias"] = adt["ValorLiquido"].where(adt["TipoLancamento"] == "Ferias", 0)

    # Unificando tudo
    frames = [adm, dem, exams, epi_df, adt]
    base = pd.concat([f for f in frames if not f.empty], ignore_index=True)

    return base

base = montar_base_indicadores()

# ---------------------------------------------------------
# SIDEBAR – FILTROS
# ---------------------------------------------------------
st.sidebar.title("Filtros")

def filtro_multiselect(label, coluna):
    if coluna not in base.columns:
        return []
    return st.sidebar.multiselect(
        label,
        sorted(base[coluna].dropna().unique().tolist())
    )

produtor_sel = filtro_multiselect("Produtor Rural", "ProdutorRural")
filial_sel = filtro_multiselect("Filial", "Filial")
unidade_sel = filtro_multiselect("Unidade", "Unidade")
setor_sel = filtro_multiselect("Setor", "Setor")
mes_sel = filtro_multiselect("Mês", "Mes")
tipo_sel = filtro_multiselect("Tipo de movimento", "TipoMovimento")
extra_sel = filtro_multiselect("Tipo Extra (ADT13 / 13 / Férias)", "TipoLancamento")

# Filtros aplicados
filtro = base.copy()

def aplicar(coluna, selecao):
    if selecao:
        return filtro[coluna].isin(selecao)
    return True

try:
    filtro = filtro[
        aplicar("ProdutorRural", produtor_sel)
        & aplicar("Filial", filial_sel)
        & aplicar("Unidade", unidade_sel)
        & aplicar("Setor", setor_sel)
        & aplicar("Mes", mes_sel)
        & aplicar("TipoMovimento", tipo_sel)
        & aplicar("TipoLancamento", extra_sel)
    ]
except:
    pass

# ---------------------------------------------------------
# CABEÇALHO
# ---------------------------------------------------------
st.markdown("<h1 style='text-align:center;'>ContaUni</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center;'>Contabilidade Unindo Sonhos com Resultados!</h3>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align:center;'>Sistema de Indicadores do Departamento Pessoal</h4>", unsafe_allow_html=True)
st.markdown("---")

# ---------------------------------------------------------
# CARDS DE MÉTRICAS
# ---------------------------------------------------------
col1, col2, col3, col4, col5, col6 = st.columns(6)

col1.metric("Admissões", filtro[filtro["TipoMovimento"] == "Admissao"].shape[0])
col2.metric("Demissões", filtro[filtro["TipoMovimento"] == "Demissao"].shape[0])
col3.metric("Rescisão líquida R$", round(filtro["ValorLiquido"].sum(), 2))
col4.metric("Multa FGTS R$", round(filtro.get("MultaFGTS", 0).sum(), 2))
col5.metric("Exames R$", round(filtro.get("ValorExame", 0).sum(), 2))
col6.metric("Uniformes e EPI R$", round(filtro.get("ValorEPI", 0).sum(), 2))

col7, col8, col9 = st.columns(3)
col7.metric("Férias R$", round(filtro.get("ValorFerias", 0).sum(), 2))
col8.metric("ADT13 R$", round(filtro.get("ValorADT13", 0).sum(), 2))
col9.metric("13º R$", round(filtro.get("Valor13", 0).sum(), 2))

st.markdown("---")

# ---------------------------------------------------------
# GRÁFICO – CUSTOS POR PRODUTOR
# ---------------------------------------------------------
st.subheader("Custos por Produtor Rural")

if "ProdutorRural" in filtro.columns and not filtro.empty:
    soma_prod = filtro.groupby("ProdutorRural")[["ValorLiquido", "ValorExame", "ValorEPI"]].sum().reset_index()
    fig = px.bar(soma_prod, x="ProdutorRural", y="ValorLiquido", title="Custos de Rescisão por Produtor Rural")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Não há dados filtrados para exibir o gráfico.")

# ---------------------------------------------------------
# TABELA DETALHADA
# ---------------------------------------------------------
st.subheader("Detalhamento dos registros")
st.dataframe(filtro, use_container_width=True)
