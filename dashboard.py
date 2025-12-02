import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import datetime

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
    # Tratamento de produtor - mapeia Filial para ProdutorRural
    prod_map = produtores.set_index('Filial')['ProdutorRural'].to_dict()

    # ===== ADMISSÕES =====
    adm_base = adm.copy()
    if "DataAdmissao" in adm_base.columns:
        adm_base["DataAdmissao"] = pd.to_datetime(adm_base["DataAdmissao"]).dt.date
        adm_base["Mes"] = pd.to_datetime(adm_base["DataAdmissao"]).dt.strftime('%m/%Y')
    
    adm_base["TipoMovimento"] = "Admissão"
    adm_base["Produtor"] = adm_base["Filial"].map(prod_map)
    adm_base["Rescisao"] = 0
    adm_base["MultaFGTS"] = 0
    adm_base["ValorExame"] = 0
    adm_base["ValorEPI"] = 0
    adm_base["ADT13"] = 0
    adm_base["Decimo13"] = 0
    adm_base["Ferias"] = 0

    # ===== DEMISSÕES/RESCISÕES =====
    dem_base = dem.copy()
    if "DataDemissao" in dem_base.columns:
        dem_base["DataDemissao"] = pd.to_datetime(dem_base["DataDemissao"]).dt.date
        dem_base["Mes"] = pd.to_datetime(dem_base["DataDemissao"]).dt.strftime('%m/%Y')
    
    dem_base["TipoMovimento"] = "Demissão"
    dem_base["Produtor"] = dem_base["Filial"].map(prod_map)
    
    # Mapeamento correto da coluna de rescisão - testa várias possibilidades
    rescisao_encontrada = False
    for col_name in ["LiquidoRe", "LiquidoRescisao", "Liquido", "ValorLiquido", "Rescisao"]:
        if col_name in dem_base.columns:
            dem_base["Rescisao"] = dem_base[col_name]
            rescisao_encontrada = True
            break
    
    if not rescisao_encontrada:
        dem_base["Rescisao"] = 0
    
    # Garante que MultaFGTS existe
    if "MultaFGTS" not in dem_base.columns:
        dem_base["MultaFGTS"] = 0
    
    dem_base["ValorExame"] = 0
    dem_base["ValorEPI"] = 0
    dem_base["ADT13"] = 0
    dem_base["Decimo13"] = 0
    dem_base["Ferias"] = 0

    # ===== EXAMES =====
    exames_base = exames.copy()
    if "DataExame" in exames_base.columns:
        exames_base["DataExame"] = pd.to_datetime(exames_base["DataExame"]).dt.date
        exames_base["Mes"] = pd.to_datetime(exames_base["DataExame"]).dt.strftime('%m/%Y')
    
    exames_base["TipoMovimento"] = "Exame"
    exames_base["Produtor"] = exames_base["Filial"].map(prod_map)
    exames_base.rename(columns={"ValorExame": "ValorExame"}, inplace=True)
    exames_base["Rescisao"] = 0
    exames_base["MultaFGTS"] = 0
    exames_base["ValorEPI"] = 0
    exames_base["ADT13"] = 0
    exames_base["Decimo13"] = 0
    exames_base["Ferias"] = 0

    # ===== EPI/UNIFORMES =====
    epi_base = epi.copy()
    if "DataEntrega" in epi_base.columns:
        epi_base["DataEntrega"] = pd.to_datetime(epi_base["DataEntrega"]).dt.date
        epi_base["Mes"] = pd.to_datetime(epi_base["DataEntrega"]).dt.strftime('%m/%Y')
    
    epi_base["TipoMovimento"] = "EPI/Uniforme"
    epi_base["Produtor"] = epi_base["Filial"].map(prod_map)
    epi_base.rename(columns={"ValorItem": "ValorEPI"}, inplace=True)
    
    if "ValorEPI" not in epi_base.columns:
        epi_base["ValorEPI"] = 0
    
    epi_base["Rescisao"] = 0
    epi_base["MultaFGTS"] = 0
    epi_base["ValorExame"] = 0
    epi_base["ADT13"] = 0
    epi_base["Decimo13"] = 0
    epi_base["Ferias"] = 0

    # ===== ADT13/13º/FÉRIAS =====
    adt_base = adt13.copy()
    
    # Cria coluna Mes se não existir
    if "Mes" not in adt_base.columns and "Data" in adt_base.columns:
        adt_base["Mes"] = pd.to_datetime(adt_base["Data"]).dt.strftime('%m/%Y')
    
    adt_base["Produtor"] = adt_base["Filial"].map(prod_map)
    
    # Identifica tipo de lançamento
    if "Lancamento" in adt_base.columns:
        adt_base["TipoExtra"] = adt_base["Lancamento"]
        adt_base["TipoMovimento"] = adt_base["Lancamento"].apply(
            lambda x: "ADT 13º" if x == "ADT13" else ("13º Salário" if x == "13" else "Férias")
        )
    
    # Distribui valores
    if "ValorLiquido" in adt_base.columns:
        adt_base["ADT13"] = adt_base.apply(
            lambda x: x["ValorLiquido"] if x.get("Lancamento") == "ADT13" else 0, axis=1
        )
        adt_base["Decimo13"] = adt_base.apply(
            lambda x: x["ValorLiquido"] if x.get("Lancamento") == "13" else 0, axis=1
        )
        adt_base["Ferias"] = adt_base.apply(
            lambda x: x["ValorLiquido"] if x.get("Lancamento") == "Ferias" else 0, axis=1
        )
    
    adt_base["Rescisao"] = 0
    adt_base["MultaFGTS"] = 0
    adt_base["ValorExame"] = 0
    adt_base["ValorEPI"] = 0

    # ===== JUNÇÃO FINAL =====
    base = pd.concat([adm_base, dem_base, exames_base, epi_base, adt_base], ignore_index=True)
    
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
    
    # DEBUG: Mostrar colunas (remover depois)
    with st.expander("🔍 Debug - Colunas das planilhas"):
        st.write("**Demissões (colunas originais):**", list(dem.columns))
        st.write("**Amostra de dados de demissões:**")
        st.dataframe(dem.head(2))
        st.write("---")
        st.write("**Base unificada - Rescisão:**")
        st.dataframe(base[base["TipoMovimento"] == "Demissão"][["Funcionario", "CPF", "TipoMovimento", "Rescisao", "MultaFGTS"]].head(5))
    
    st.write("---")

    # ===============================
    # FILTROS NA SIDEBAR
    # ===============================
    st.sidebar.header("Filtros")

    colunas_filtro = {
        "Produtor Rural": "Produtor",
        "Filial": "Filial",
        "Unidade": "Unidade",
        "Setor": "Setor",
        "Mês": "Mes",
        "Tipo de movimento": "TipoMovimento",
    }

    filtros = {}

    for label, coluna in colunas_filtro.items():
        if coluna in base.columns:
            valores_unicos = base[coluna].dropna().unique()
            valores_unicos = sorted([str(v) for v in valores_unicos])
            selecao = st.sidebar.multiselect(label, valores_unicos)
            filtros[coluna] = selecao

    # ===============================
    # APLICAR FILTROS
    # ===============================
    filtro = base.copy()

    for coluna, valores in filtros.items():
        if valores:
            filtro = filtro[filtro[coluna].astype(str).isin([str(v) for v in valores])]

    # ===============================
    # CARDS RESUMO
    # ===============================
    col1, col2, col3, col4, col5, col6 = st.columns(6)

    col1.metric("Admissões", filtro[filtro["TipoMovimento"] == "Admissão"].shape[0])
    col2.metric("Demissões", filtro[filtro["TipoMovimento"] == "Demissão"].shape[0])
    col3.metric("Rescisão R$", f"{filtro['Rescisao'].sum():,.2f}")
    col4.metric("Multa FGTS R$", f"{filtro['MultaFGTS'].sum():,.2f}")
    col5.metric("Exames R$", f"{filtro['ValorExame'].sum():,.2f}")
    col6.metric("EPI/Uniformes R$", f"{filtro['ValorEPI'].sum():,.2f}")

    col7, col8, col9 = st.columns(3)
    col7.metric("Férias R$", f"{filtro['Ferias'].sum():,.2f}")
    col8.metric("ADT 13º R$", f"{filtro['ADT13'].sum():,.2f}")
    col9.metric("13º Salário R$", f"{filtro['Decimo13'].sum():,.2f}")

    # ===============================
    # GRÁFICO POR PRODUTOR
    # ===============================
    st.write("---")
    st.subheader("Custos por Produtor Rural")

    if not filtro.empty and "Produtor" in filtro.columns:
        graf = (
            filtro.groupby("Produtor")[["Rescisao", "ValorExame", "ValorEPI",
                                        "ADT13", "Decimo13", "Ferias", "MultaFGTS"]]
            .sum()
            .reset_index()
        )
        
        # Remove produtores sem nome
        graf = graf[graf["Produtor"].notna() & (graf["Produtor"] != "0") & (graf["Produtor"] != 0)]

        # Prepara dados para gráfico
        graf_melted = graf.melt(
            id_vars=["Produtor"],
            value_vars=["Rescisao", "ValorExame", "ValorEPI", "ADT13", "Decimo13", "Ferias", "MultaFGTS"],
            var_name="Tipo",
            value_name="Valor"
        )
        
        # Renomeia para melhor visualização
        nomes_display = {
            "Rescisao": "Rescisão Líquida",
            "ValorExame": "Exames",
            "ValorEPI": "EPI/Uniformes",
            "ADT13": "ADT 13º",
            "Decimo13": "13º Salário",
            "Ferias": "Férias",
            "MultaFGTS": "Multa FGTS"
        }
        graf_melted["Tipo"] = graf_melted["Tipo"].map(nomes_display)

        fig = px.bar(
            graf_melted,
            x="Produtor",
            y="Valor",
            color="Tipo",
            title="Custos por Produtor Rural",
            height=500,
            barmode="group"
        )
        
        fig.update_layout(
            xaxis_title="Produtor Rural",
            yaxis_title="Valor (R$)",
            legend_title="Tipo de Custo"
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Nenhum dado encontrado para os filtros selecionados.")

    # ===============================
    # TABELA DETALHADA COM EXPORTAÇÃO
    # ===============================
    st.write("---")
    st.subheader("Detalhamento dos registros")
    
    # Botões de exportação
    col_exp1, col_exp2, col_exp3 = st.columns([1, 1, 4])
    
    with col_exp1:
        # Exportar para CSV
        csv = filtro.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 Exportar CSV",
            data=csv,
            file_name=f'relatorio_rh_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
            mime='text/csv',
        )
    
    with col_exp2:
        # Exportar para Excel
        from io import BytesIO
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            filtro.to_excel(writer, sheet_name='Dados', index=False)
            writer.close()
        
        st.download_button(
            label="📥 Exportar Excel",
            data=buffer.getvalue(),
            file_name=f'relatorio_rh_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    # Mostra a tabela
    st.dataframe(filtro, use_container_width=True, height=400)
    
    # Estatísticas adicionais
    st.write("---")
    st.subheader("Resumo Estatístico")
    
    col_stat1, col_stat2, col_stat3 = st.columns(3)
    
    with col_stat1:
        st.metric("Total de Registros", len(filtro))
        st.metric("Total Geral", f"R$ {(filtro['Rescisao'].sum() + filtro['ValorExame'].sum() + filtro['ValorEPI'].sum() + filtro['ADT13'].sum() + filtro['Decimo13'].sum() + filtro['Ferias'].sum() + filtro['MultaFGTS'].sum()):,.2f}")
    
    with col_stat2:
        if len(filtro) > 0:
            st.metric("Produtores Únicos", filtro["Produtor"].nunique())
            st.metric("Filiais Únicas", filtro["Filial"].nunique())
    
    with col_stat3:
        if "Mes" in filtro.columns:
            st.metric("Meses no Período", filtro["Mes"].nunique())

except FileNotFoundError as e:
    st.error(f"❌ Erro ao carregar arquivos: {e}")
    st.info("Verifique se todos os arquivos Excel estão na pasta 'dados/'")
except Exception as e:
    st.error(f"❌ Erro no sistema: {str(e)}")
    st.info("Entre em contato com o suporte técnico.")
    
    if st.checkbox("Mostrar detalhes do erro (debug)"):
        st.exception(e)
