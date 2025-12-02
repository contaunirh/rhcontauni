import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="ContaUni - Indicadores RH", layout="wide")

# ==============================
# FUNÇÃO PARA CARREGAR PLANILHAS
# ==============================
@st.cache_data
def carregar_dados():
    base_path = "dados"

    adm = pd.read_excel(os.path.join(base_path, "admissoes.2025.xlsx"))
    dem = pd.read_excel(os.path.join(base_path, "demissoes.rescisoes.2025.xlsx"))
    exames = pd.read_excel(os.path.join(base_path, "exames.2025.xlsx"))
    epi = pd.read_excel(os.path.join(base_path, "epi.uniformes.2025.xlsx"))
    adt13 = pd.read_excel(os.path.join(base_path, "adt.13.ferias.xlsx"))
    produtores = pd.read_excel(os.path.join(base_path, "produtores.2025.xlsx"))

    return adm, dem, exames, epi, adt13, produtores


# ===============================
# MONTA A BASE UNIFICADA
# ===============================
def montar_base(adm, dem, exames, epi, adt13, produtores):
    # Tratamento de produtor
    prod = produtores.rename(columns={"ProdutorRural": "Produtor", "Filial": "FilialProdutor"})

    # ADM
    adm_base = adm.copy()
    # Converte coluna de data para formato sem hora
    if "DataAdmissao" in adm_base.columns:
        adm_base["DataAdmissao"] = pd.to_datetime(adm_base["DataAdmissao"]).dt.date
    elif "Data" in adm_base.columns:
        adm_base["Data"] = pd.to_datetime(adm_base["Data"]).dt.date
    
    adm_base["TipoMovimento"] = "Admissao"
    adm_base["ValorExame"] = 0
    adm_base["ValorEPI"] = 0
    adm_base["Rescisao"] = 0
    adm_base["MultaFGTS"] = 0
    adm_base["ADT13"] = 0
    adm_base["Decimo13"] = 0
    adm_base["Ferias"] = 0

    # DEMISSÕES / RESCISÕES
    dem_base = dem.copy()
    # Converte coluna de data para formato sem hora
    if "DataDemissao" in dem_base.columns:
        dem_base["DataDemissao"] = pd.to_datetime(dem_base["DataDemissao"]).dt.date
    elif "Data" in dem_base.columns:
        dem_base["Data"] = pd.to_datetime(dem_base["Data"]).dt.date
    
    dem_base["TipoMovimento"] = "Demissao"
    dem_base.rename(columns={"LiquidoRescisao": "Rescisao", "MultaFGTS": "MultaFGTS"}, inplace=True)
    dem_base["ValorExame"] = 0
    dem_base["ValorEPI"] = 0
    dem_base["ADT13"] = 0
    dem_base["Decimo13"] = 0
    dem_base["Ferias"] = 0

    # EXAMES
    exames_base = exames.copy()
    # Converte coluna de data para formato sem hora
    if "DataExame" in exames_base.columns:
        exames_base["DataExame"] = pd.to_datetime(exames_base["DataExame"]).dt.date
    elif "Data" in exames_base.columns:
        exames_base["Data"] = pd.to_datetime(exames_base["Data"]).dt.date
    
    exames_base["TipoMovimento"] = "Exame"
    exames_base.rename(columns={"ValorExame": "ValorExame"}, inplace=True)
    exames_base["Rescisao"] = 0
    exames_base["ValorEPI"] = 0
    exames_base["MultaFGTS"] = 0
    exames_base["ADT13"] = 0
    exames_base["Decimo13"] = 0
    exames_base["Ferias"] = 0

    # EPI / UNIFORMES
    epi_base = epi.copy()
    # Converte coluna de data para formato sem hora
    if "DataEntrega" in epi_base.columns:
        epi_base["DataEntrega"] = pd.to_datetime(epi_base["DataEntrega"]).dt.date
    elif "Data" in epi_base.columns:
        epi_base["Data"] = pd.to_datetime(epi_base["Data"]).dt.date
    
    epi_base["TipoMovimento"] = "EPI"
    epi_base.rename(columns={"ValorEPI": "ValorEPI"}, inplace=True)
    epi_base["Rescisao"] = 0
    epi_base["ValorExame"] = 0
    epi_base["MultaFGTS"] = 0
    epi_base["ADT13"] = 0
    epi_base["Decimo13"] = 0
    epi_base["Ferias"] = 0

    # ADT 13 / FERIAS
    adt_base = adt13.copy()
    # Converte coluna de data para formato sem hora
    if "DataPagamento" in adt_base.columns:
        adt_base["DataPagamento"] = pd.to_datetime(adt_base["DataPagamento"]).dt.date
    elif "Data" in adt_base.columns:
        adt_base["Data"] = pd.to_datetime(adt_base["Data"]).dt.date
    
    # Identifica qual coluna tem o tipo de lançamento
    coluna_tipo = None
    for col in ["Lancamento", "TipoLancamento", "Tipo", "TipoExtra"]:
        if col in adt_base.columns:
            coluna_tipo = col
            break
    
    # Identifica qual coluna tem o valor
    coluna_valor = None
    for col in ["ValorLiquido", "Valor", "ValorTotal", "Total"]:
        if col in adt_base.columns:
            coluna_valor = col
            break
    
    # Se não encontrou as colunas necessárias, cria valores zerados
    if coluna_tipo is None or coluna_valor is None:
        adt_base["TipoExtra"] = ""
        adt_base["TipoMovimento"] = ""
        adt_base["ValorExame"] = 0
        adt_base["ValorEPI"] = 0
        adt_base["Rescisao"] = 0
        adt_base["MultaFGTS"] = 0
        adt_base["ADT13"] = 0
        adt_base["Decimo13"] = 0
        adt_base["Ferias"] = 0
    else:
        # Cria a coluna TipoExtra se não existir
        if coluna_tipo != "TipoExtra":
            adt_base["TipoExtra"] = adt_base[coluna_tipo]
        
        # Define o TipoMovimento baseado no TipoExtra
        adt_base["TipoMovimento"] = adt_base["TipoExtra"].apply(
            lambda x: "ADT 13º" if x == "ADT13" else ("13º Salário" if x == "13" else "Férias")
        )
        
        adt_base["ValorExame"] = 0
        adt_base["ValorEPI"] = 0
        adt_base["MultaFGTS"] = 0

        # Cria as colunas de valores baseadas no tipo de lançamento
        adt_base["ADT13"] = adt_base.apply(
            lambda x: x[coluna_valor] if x["TipoExtra"] == "ADT13" else 0, axis=1
        )
        adt_base["Decimo13"] = adt_base.apply(
            lambda x: x[coluna_valor] if x["TipoExtra"] == "13" else 0, axis=1
        )
        adt_base["Ferias"] = adt_base.apply(
            lambda x: x[coluna_valor] if x["TipoExtra"] == "Ferias" else 0, axis=1
        )
        
        # Zera Rescisao pois será usado apenas nas demissões
        adt_base["Rescisao"] = 0

    # JUNÇÃO FINAL
    base = pd.concat([adm_base, dem_base, exames_base, epi_base, adt_base], ignore_index=True)

    # Verifica se existe coluna Produtor na base antes de fazer merge
    if "Produtor" in base.columns and "Produtor" in prod.columns:
        base = base.merge(prod, left_on="Produtor", right_on="Produtor", how="left")
    else:
        # Se não existe Produtor, adiciona colunas vazias do arquivo de produtores
        for col in prod.columns:
            if col not in base.columns:
                base[col] = None

    return base.fillna(0)


# ===============================
# CARREGA DADOS COM TRATAMENTO DE ERRO
# ===============================
try:
    adm, dem, exames, epi, adt13, produtores = carregar_dados()
    base = montar_base(adm, dem, exames, epi, adt13, produtores)

    st.title("ContaUni")
    st.subheader("Contabilidade Unindo Sonhos com Resultados!")
    st.markdown("### Sistema de Indicadores do Departamento Pessoal")
    st.write("---")

    # ===============================
    # FILTROS NA SIDEBAR
    # ===============================
    st.sidebar.header("Filtros")

    colunas_filtro = {
        "Produtor Rural": "Produtor",
        "Filial": "FilialProdutor",
        "Unidade": "Unidade",
        "Setor": "Setor",
        "Mês": "Mes",
        "Tipo de movimento": "TipoMovimento",
        "Tipo Extra (ADT13 / 13 / Férias)": "TipoExtra"
    }

    filtros = {}

    for label, coluna in colunas_filtro.items():
        if coluna in base.columns:
            # Remove valores nulos e converte tudo para string para evitar erro de comparação
            valores_unicos = base[coluna].dropna().unique()
            # Converte para string e depois ordena
            valores_unicos = sorted([str(v) for v in valores_unicos])
            selecao = st.sidebar.multiselect(label, valores_unicos)
            # Converte de volta para o tipo original ao filtrar
            if selecao:
                # Tenta manter o tipo original
                try:
                    # Se a coluna original tinha números, converte de volta
                    if base[coluna].dtype in ['int64', 'float64']:
                        selecao = [float(v) if '.' in v else int(v) for v in selecao]
                except:
                    pass  # Mantém como string se não conseguir converter
            filtros[coluna] = selecao

    # ===============================
    # APLICAR FILTROS
    # ===============================
    filtro = base.copy()

    for coluna, valores in filtros.items():
        if valores:
            # Converte os valores da coluna para string para comparação
            filtro = filtro[filtro[coluna].astype(str).isin([str(v) for v in valores])]

    # ===============================
    # CARDS RESUMO
    # ===============================
    col1, col2, col3, col4, col5, col6 = st.columns(6)

    col1.metric("Admissões", filtro[filtro["TipoMovimento"] == "Admissao"].shape[0])
    col2.metric("Demissões", filtro[filtro["TipoMovimento"] == "Demissao"].shape[0])
    col3.metric("Rescisão R$", f"{filtro['Rescisao'].sum():,.2f}")
    col4.metric("Multa FGTS R$", f"{filtro['MultaFGTS'].sum():,.2f}")
    col5.metric("Exames R$", f"{filtro['ValorExame'].sum():,.2f}")
    col6.metric("Uniformes e EPI R$", f"{filtro['ValorEPI'].sum():,.2f}")

    col7, col8, col9 = st.columns(3)
    col7.metric("Férias R$", f"{filtro['Ferias'].sum():,.2f}")
    col8.metric("ADT 13º R$", f"{filtro['ADT13'].sum():,.2f}")
    col9.metric("13º Salário R$", f"{filtro['Decimo13'].sum():,.2f}")

    # ===============================
    # GRÁFICO POR PRODUTOR
    # ===============================
    st.write("---")
    st.subheader("Custos por Produtor Rural")

    if not filtro.empty:
        graf = (
            filtro.groupby("Produtor")[["Rescisao", "ValorExame", "ValorEPI",
                                        "ADT13", "Decimo13", "Ferias", "MultaFGTS"]]
            .sum()
            .reset_index()
        )

        # Renomeia as colunas para exibição no gráfico
        graf_display = graf.rename(columns={
            "Rescisao": "Rescisão Líquida",
            "ValorExame": "Exames",
            "ValorEPI": "EPI/Uniformes",
            "ADT13": "ADT 13º",
            "Decimo13": "13º Salário",
            "Ferias": "Férias",
            "MultaFGTS": "Multa FGTS"
        })

        fig = px.bar(
            graf_display,
            x="Produtor",
            y=["Rescisão Líquida", "Exames", "EPI/Uniformes", "ADT 13º", "13º Salário", "Férias", "Multa FGTS"],
            barmode="group",
            title="Custos por Produtor",
            height=500,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Nenhum dado encontrado para os filtros selecionados.")

    # ===============================
    # TABELA DETALHADA
    # ===============================
    st.write("---")
    st.subheader("Detalhamento dos registros")
    st.dataframe(filtro)

except FileNotFoundError as e:
    st.error(f"❌ Erro ao carregar arquivos: {e}")
    st.info("Verifique se todos os arquivos Excel estão na pasta 'dados/'")
except Exception as e:
    st.error(f"❌ Erro no sistema: {str(e)}")
    st.info("Entre em contato com o suporte técnico.")
    
    # Botão para debug (remover em produção)
    if st.checkbox("Mostrar detalhes do erro (debug)"):
        st.exception(e)
