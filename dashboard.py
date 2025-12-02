import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# ---------------------------------------------------------
# Configuração geral da página
# ---------------------------------------------------------
st.set_page_config(page_title="RH ContaUni", layout="wide")

st.markdown(
    "<h1 style='text-align:center;'>ContaUni</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    "<h3 style='text-align:center;'>Contabilidade Unindo Sonhos com Resultados!</h3>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align:center;'>Sistema de Indicadores do Departamento Pessoal</p>",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# Caminhos dos arquivos
# ---------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DADOS_DIR = BASE_DIR.parent / "dados"  # ..\dados a partir da pasta app


def ler_excel(nome_arquivo: str) -> pd.DataFrame:
    """Lê um arquivo Excel da pasta dados e retorna DataFrame vazio em caso de erro."""
    caminho = DADOS_DIR / nome_arquivo
    try:
        df = pd.read_excel(caminho)
        return df
    except Exception:
        return pd.DataFrame()


# ---------------------------------------------------------
# Carregar todas as planilhas
# ---------------------------------------------------------
adm = ler_excel("admissoes.2025.xlsx")
dem = ler_excel("demissoes.rescisoes.2025.xlsx")
exa = ler_excel("exames.2025.xlsx")
epi = ler_excel("epi.uniformes.2025.xlsx")
prod = ler_excel("produtores.2025.xlsx")
adt = ler_excel("adt.13.ferias.xlsx")

# ---------------------------------------------------------
# Funções de apoio
# ---------------------------------------------------------
def garantir_colunas(df: pd.DataFrame, colunas: list[str]) -> pd.DataFrame:
    """Garante existência das colunas informadas em um DF."""
    for c in colunas:
        if c not in df.columns:
            df[c] = None
    return df


# ---------------------------------------------------------
# Produtores
# ---------------------------------------------------------
if not prod.empty:
    for col in ["Filial", "Unidade", "ProdutorRural"]:
        if col not in prod.columns:
            prod[col] = ""
else:
    prod = pd.DataFrame(columns=["Filial", "Unidade", "ProdutorRural"])


def adicionar_produtor(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    if "Filial" not in df.columns or "Unidade" not in df.columns:
        return df

    base_local = df.copy()
    prod_reduzido = prod[["Filial", "Unidade", "ProdutorRural"]].drop_duplicates()

    base_local = base_local.merge(
        prod_reduzido,
        on=["Filial", "Unidade"],
        how="left",
        suffixes=("", "_prod"),
    )
    return base_local


# ---------------------------------------------------------
# Admissões
# ---------------------------------------------------------
adm = garantir_colunas(
    adm,
    ["Funcionario", "DataAdmissao", "Setor", "Filial", "Unidade", "Salario"],
)
adm = adicionar_produtor(adm)

# ---------------------------------------------------------
# Demissões e rescisões
# ---------------------------------------------------------
dem = garantir_colunas(
    dem,
    [
        "Funcionario",
        "DataDemissao",
        "Motivo",
        "Setor",
        "Filial",
        "Unidade",
        "ValorBrutoRescisao",
        "Encargos",
        "ValorLiquidoRescisao",
        "MultaFGTS",
        "ValorMultaFGTS",
    ],
)

# Se tiver apenas ValorMultaFGTS, usa para preencher MultaFGTS
if "MultaFGTS" not in dem.columns or dem["MultaFGTS"].isna().all():
    if "ValorMultaFGTS" in dem.columns:
        dem["MultaFGTS"] = dem["ValorMultaFGTS"]

dem["ValorBrutoRescisao"] = pd.to_numeric(
    dem["ValorBrutoRescisao"], errors="coerce"
).fillna(0)
dem["Encargos"] = pd.to_numeric(dem["Encargos"], errors="coerce").fillna(0)
dem["ValorLiquidoRescisao"] = pd.to_numeric(
    dem["ValorLiquidoRescisao"], errors="coerce"
).fillna(0)
dem["MultaFGTS"] = pd.to_numeric(dem["MultaFGTS"], errors="coerce").fillna(0)

dem = adicionar_produtor(dem)

# ---------------------------------------------------------
# Exames
# ---------------------------------------------------------
# Estrutura esperada da planilha:
# Funcionario, CPF, TipoVinculo, TipoExame,
# DataExame, Competencia, ValorExame, Setor, Filial, Unidade
exa = garantir_colunas(
    exa,
    ["Funcionario", "DataExame", "TipoExame", "ValorExame", "Valor", "Filial", "Unidade", "Setor"],
)

# Usa ValorExame como valor padrão
if "ValorExame" in exa.columns:
    exa["Valor"] = pd.to_numeric(exa["ValorExame"], errors="coerce").fillna(0)
else:
    exa["Valor"] = pd.to_numeric(exa["Valor"], errors="coerce").fillna(0)

exa = adicionar_produtor(exa)

# ---------------------------------------------------------
# EPI e uniformes
# ---------------------------------------------------------
# Estrutura esperada da planilha:
# Funcionario, CPF, TipoEntrega, TipoItem, DescricaoItem,
# Quantidade, ValorItem, DataEntrega, Setor, Filial, Unidade
epi = garantir_colunas(
    epi,
    ["Funcionario", "Filial", "Unidade", "Setor", "Quantidade", "ValorItem", "DataEntrega"],
)

# Data
if "Data" not in epi.columns:
    epi["Data"] = pd.to_datetime(epi["DataEntrega"], errors="coerce")

# Valor total por linha = Quantidade x ValorItem
valor_item = pd.to_numeric(epi["ValorItem"], errors="coerce").fillna(0)
quantidade = pd.to_numeric(epi["Quantidade"], errors="coerce").fillna(0)
epi["Valor"] = (valor_item * quantidade).fillna(0)

epi = adicionar_produtor(epi)

# ---------------------------------------------------------
# ADT, 13 e férias
# ---------------------------------------------------------
adt = garantir_colunas(
    adt,
    ["Funcionario", "TipoLancamento", "ValorLiquido", "Mes", "Filial", "Unidade", "Setor"],
)
adt["ValorLiquido"] = pd.to_numeric(adt["ValorLiquido"], errors="coerce").fillna(0)
adt = adicionar_produtor(adt)

# ---------------------------------------------------------
# Montar base única de indicadores
# ---------------------------------------------------------
bases = []

# Admissões
if not adm.empty:
    base_adm = pd.DataFrame()
    base_adm["Origem"] = "Admissao"
    base_adm["Funcionario"] = adm["Funcionario"]
    base_adm["Data"] = pd.to_datetime(adm["DataAdmissao"], errors="coerce")
    base_adm["Mes"] = base_adm["Data"].dt.month
    base_adm["Filial"] = adm["Filial"]
    base_adm["Unidade"] = adm["Unidade"]
    base_adm["Setor"] = adm["Setor"]
    base_adm["ProdutorRural"] = adm["ProdutorRural"]
    base_adm["TipoMovimento"] = "Admissao"
    base_adm["TipoLancExtra"] = None
    base_adm["CustoRescisao"] = 0.0
    base_adm["MultaFGTS"] = 0.0
    base_adm["CustoExames"] = 0.0
    base_adm["CustoEPI"] = 0.0
    base_adm["CustoADT13"] = 0.0
    base_adm["Custo13"] = 0.0
    base_adm["CustoFerias"] = 0.0
    bases.append(base_adm)

# Demissões e rescisões
if not dem.empty:
    base_dem = pd.DataFrame()
    base_dem["Origem"] = "Rescisao"
    base_dem["Funcionario"] = dem["Funcionario"]
    base_dem["Data"] = pd.to_datetime(dem["DataDemissao"], errors="coerce")
    base_dem["Mes"] = base_dem["Data"].dt.month
    base_dem["Filial"] = dem["Filial"]
    base_dem["Unidade"] = dem["Unidade"]
    base_dem["Setor"] = dem["Setor"]
    base_dem["ProdutorRural"] = dem["ProdutorRural"]
    base_dem["TipoMovimento"] = "Demissao"
    base_dem["TipoLancExtra"] = None
    base_dem["CustoRescisao"] = dem["ValorLiquidoRescisao"]
    base_dem["MultaFGTS"] = dem["MultaFGTS"]
    base_dem["CustoExames"] = 0.0
    base_dem["CustoEPI"] = 0.0
    base_dem["CustoADT13"] = 0.0
    base_dem["Custo13"] = 0.0
    base_dem["CustoFerias"] = 0.0
    bases.append(base_dem)

# Exames
if not exa.empty:
    base_exa = pd.DataFrame()
    base_exa["Origem"] = "Exame"
    base_exa["Funcionario"] = exa["Funcionario"]
    base_exa["Data"] = pd.to_datetime(exa["DataExame"], errors="coerce")
    base_exa["Mes"] = base_exa["Data"].dt.month
    base_exa["Filial"] = exa["Filial"]
    base_exa["Unidade"] = exa["Unidade"]
    base_exa["Setor"] = exa["Setor"]
    base_exa["ProdutorRural"] = exa["ProdutorRural"]
    base_exa["TipoMovimento"] = None
    base_exa["TipoLancExtra"] = None
    base_exa["CustoRescisao"] = 0.0
    base_exa["MultaFGTS"] = 0.0
    base_exa["CustoExames"] = exa["Valor"]
    base_exa["CustoEPI"] = 0.0
    base_exa["CustoADT13"] = 0.0
    base_exa["Custo13"] = 0.0
    base_exa["CustoFerias"] = 0.0
    bases.append(base_exa)

# EPI e uniformes
if not epi.empty:
    base_epi = pd.DataFrame()
    base_epi["Origem"] = "EPI"
    base_epi["Funcionario"] = epi["Funcionario"]
    base_epi["Data"] = pd.to_datetime(epi["Data"], errors="coerce")
    base_epi["Mes"] = base_epi["Data"].dt.month
    base_epi["Filial"] = epi["Filial"]
    base_epi["Unidade"] = epi["Unidade"]
    base_epi["Setor"] = epi["Setor"]
    base_epi["ProdutorRural"] = epi["ProdutorRural"]
    base_epi["TipoMovimento"] = None
    base_epi["TipoLancExtra"] = None
    base_epi["CustoRescisao"] = 0.0
    base_epi["MultaFGTS"] = 0.0
    base_epi["CustoExames"] = 0.0
    base_epi["CustoEPI"] = epi["Valor"]
    base_epi["CustoADT13"] = 0.0
    base_epi["Custo13"] = 0.0
    base_epi["CustoFerias"] = 0.0
    bases.append(base_epi)

# ADT, 13 e férias
if not adt.empty:
    base_adt = pd.DataFrame()
    base_adt["Origem"] = "LancExtra"
    base_adt["Funcionario"] = adt["Funcionario"]
    base_adt["Data"] = pd.NaT
    base_adt["Mes"] = pd.to_numeric(adt["Mes"], errors="coerce")
    base_adt["Filial"] = adt["Filial"]
    base_adt["Unidade"] = adt["Unidade"]
    base_adt["Setor"] = adt["Setor"]
    base_adt["ProdutorRural"] = adt["ProdutorRural"]
    base_adt["TipoMovimento"] = None
    base_adt["TipoLancExtra"] = adt["TipoLancamento"]

    base_adt["CustoRescisao"] = 0.0
    base_adt["MultaFGTS"] = 0.0
    base_adt["CustoExames"] = 0.0
    base_adt["CustoEPI"] = 0.0
    base_adt["CustoADT13"] = 0.0
    base_adt["Custo13"] = 0.0
    base_adt["CustoFerias"] = 0.0

    mask_adt13 = base_adt["TipoLancExtra"] == "ADT13"
    mask_13 = base_adt["TipoLancExtra"] == "13"
    mask_ferias = base_adt["TipoLancExtra"] == "Ferias"

    base_adt.loc[mask_adt13, "CustoADT13"] = adt.loc[mask_adt13, "ValorLiquido"]
    base_adt.loc[mask_13, "Custo13"] = adt.loc[mask_13, "ValorLiquido"]
    base_adt.loc[mask_ferias, "CustoFerias"] = adt.loc[mask_ferias, "ValorLiquido"]

    bases.append(base_adt)

# Concatenar tudo
if bases:
    base = pd.concat(bases, ignore_index=True)
else:
    base = pd.DataFrame(
        columns=[
            "Origem",
            "Funcionario",
            "Data",
            "Mes",
            "Filial",
            "Unidade",
            "Setor",
            "ProdutorRural",
            "TipoMovimento",
            "TipoLancExtra",
            "CustoRescisao",
            "MultaFGTS",
            "CustoExames",
            "CustoEPI",
            "CustoADT13",
            "Custo13",
            "CustoFerias",
        ]
    )

# ---------------------------------------------------------
# Filtros na sidebar
# ---------------------------------------------------------
st.sidebar.title("Filtros")


def opcoes(col):
    if col in base.columns:
        vals = sorted(base[col].dropna().unique())
        return vals
    return []


produtores_op = opcoes("ProdutorRural")
filiais_op = opcoes("Filial")
unidades_op = opcoes("Unidade")
setores_op = opcoes("Setor")
meses_op = sorted([int(m) for m in base["Mes"].dropna().unique()]) if "Mes" in base.columns else []

tipos_mov_op = sorted(
    [v for v in base["TipoMovimento"].dropna().unique()]
) if "TipoMovimento" in base.columns else []

tipos_lanc_extra_op = sorted(
    [v for v in base["TipoLancExtra"].dropna().unique()]
) if "TipoLancExtra" in base.columns else []

f_prod = st.sidebar.multiselect("Produtor Rural", produtores_op, default=produtores_op)
f_filial = st.sidebar.multiselect("Filial", filiais_op, default=filiais_op)
f_unidade = st.sidebar.multiselect("Unidade", unidades_op, default=unidades_op)
f_setor = st.sidebar.multiselect("Setor", setores_op, default=setores_op)
f_mes = st.sidebar.multiselect("Mês", meses_op, default=meses_op)
f_tipomov = st.sidebar.multiselect("Tipo de movimento", tipos_mov_op, default=tipos_mov_op)
f_tipolanc = st.sidebar.multiselect(
    "Tipo de lançamento extra ADT13  13  Férias", tipos_lanc_extra_op, default=tipos_lanc_extra_op
)

tipo_grafico = st.sidebar.selectbox("Tipo de gráfico", ["Barras", "Pizza"])
tipo_custo_graf = st.sidebar.selectbox(
    "Custo para o gráfico",
    [
        "CustoRescisao",
        "MultaFGTS",
        "CustoExames",
        "CustoEPI",
        "CustoADT13",
        "Custo13",
        "CustoFerias",
    ],
)

# ---------------------------------------------------------
# Aplicar filtros
# ---------------------------------------------------------
df = base.copy()

if f_prod:
    df = df[df["ProdutorRural"].isin(f_prod)]
if f_filial:
    df = df[df["Filial"].isin(f_filial)]
if f_unidade:
    df = df[df["Unidade"].isin(f_unidade)]
if f_setor:
    df = df[df["Setor"].isin(f_setor)]
if f_mes:
    df = df[df["Mes"].isin(f_mes)]
if f_tipomov:
    df = df[df["TipoMovimento"].isin(f_tipomov) | df["TipoMovimento"].isna()]
if f_tipolanc:
    df = df[df["TipoLancExtra"].isin(f_tipolanc) | df["TipoLancExtra"].isna()]

# ---------------------------------------------------------
# Métricas (cards)
# ---------------------------------------------------------
def soma_col(df_local: pd.DataFrame, col: str) -> float:
    if col in df_local.columns:
        return float(pd.to_numeric(df_local[col], errors="coerce").fillna(0).sum())
    return 0.0


qtd_adm = len(base[base["TipoMovimento"] == "Admissao"])
qtd_dem = len(base[base["TipoMovimento"] == "Demissao"])

total_rescisao = soma_col(base, "CustoRescisao")
total_multa_fgts = soma_col(base, "MultaFGTS")
total_exames = soma_col(base, "CustoExames")
total_epi = soma_col(base, "CustoEPI")
total_adt13 = soma_col(base, "CustoADT13")
total_13 = soma_col(base, "Custo13")
total_ferias = soma_col(base, "CustoFerias")


def fmt_val(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


c1, c2, c3, c4 = st.columns(4)
c1.metric("Admissões", qtd_adm)
c2.metric("Demissões", qtd_dem)
c3.metric("Rescisão líquida R$", fmt_val(total_rescisao))
c4.metric("Multa FGTS R$", fmt_val(total_multa_fgts))

c5, c6, c7, c8 = st.columns(4)
c5.metric("Exames R$", fmt_val(total_exames))
c6.metric("Uniformes e EPI R$", fmt_val(total_epi))
c7.metric("ADT13 R$", fmt_val(total_adt13))
c8.metric("13º R$", fmt_val(total_13))

c9, _, _, _ = st.columns(4)
c9.metric("Férias R$", fmt_val(total_ferias))

st.markdown("---")

# ---------------------------------------------------------
# Gráfico por produtor
# ---------------------------------------------------------
st.subheader("Custos por Produtor Rural")

if not df.empty and "ProdutorRural" in df.columns:
    df_graf = df.copy()
    df_graf[tipo_custo_graf] = pd.to_numeric(
        df_graf[tipo_custo_graf], errors="coerce"
    ).fillna(0)

    agg = df_graf.groupby("ProdutorRural")[tipo_custo_graf].sum().reset_index()
    agg = agg[agg[tipo_custo_graf] > 0]

    if not agg.empty:
        if tipo_grafico == "Barras":
            fig = px.bar(agg, x="ProdutorRural", y=tipo_custo_graf, text_auto=True)
        else:
            fig = px.pie(agg, names="ProdutorRural", values=tipo_custo_graf)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Não há dados para exibir no gráfico com os filtros atuais.")
else:
    st.info("Não há dados de ProdutorRural para exibir o gráfico.")

st.markdown("---")
st.subheader("Custos por Setor")

if not df.empty and "Setor" in df.columns:
    df_setor = df.copy()
    df_setor[tipo_custo_graf] = pd.to_numeric(
        df_setor[tipo_custo_graf], errors="coerce"
    ).fillna(0)

    agg_setor = df_setor.groupby("Setor")[tipo_custo_graf].sum().reset_index()
    agg_setor = agg_setor[agg_setor[tipo_custo_graf] > 0]

    if not agg_setor.empty:
        if tipo_grafico == "Barras":
            fig2 = px.bar(agg_setor, x="Setor", y=tipo_custo_graf, text_auto=True)
        else:
            fig2 = px.pie(agg_setor, names="Setor", values=tipo_custo_graf)
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Não há dados por setor com os filtros atuais.")
else:
    st.info("Não há dados de Setor para exibir o gráfico.")

# ---------------------------------------------------------
# Tabela detalhada
# ---------------------------------------------------------
st.markdown("---")
st.subheader("Detalhamento dos registros filtrados")

if df.empty:
    st.warning("Nenhum registro encontrado com os filtros selecionados.")
else:
    st.dataframe(df)
