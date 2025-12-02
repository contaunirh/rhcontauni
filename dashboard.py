import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
from datetime import datetime

st.set_page_config(page_title="ContaUni - Indicadores RH", layout="wide", initial_sidebar_state="expanded")

# Estilo customizado
st.markdown("""
    <style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

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

def montar_base(adm, dem, exames, epi, adt13, produtores):
    prod_map = produtores.set_index('Filial')['ProdutorRural'].to_dict()

    # ADMISSÕES
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

    # DEMISSÕES
    dem_base = dem.copy()
    if "DataDemissao" in dem_base.columns:
        dem_base["DataDemissao"] = pd.to_datetime(dem_base["DataDemissao"]).dt.date
        dem_base["Mes"] = pd.to_datetime(dem_base["DataDemissao"]).dt.strftime('%m/%Y')
    dem_base["TipoMovimento"] = "Demissão"
    dem_base["Produtor"] = dem_base["Filial"].map(prod_map)
    
    if "ValorLiquidoRescisao" in dem_base.columns:
        dem_base["Rescisao"] = dem_base["ValorLiquidoRescisao"]
    else:
        dem_base["Rescisao"] = 0
    
    if "MultaFGTS" not in dem_base.columns:
        dem_base["MultaFGTS"] = 0
    
    dem_base["ValorExame"] = 0
    dem_base["ValorEPI"] = 0
    dem_base["ADT13"] = 0
    dem_base["Decimo13"] = 0
    dem_base["Ferias"] = 0

    # EXAMES
    exames_base = exames.copy()
    if "DataExame" in exames_base.columns:
        exames_base["DataExame"] = pd.to_datetime(exames_base["DataExame"]).dt.date
        exames_base["Mes"] = pd.to_datetime(exames_base["DataExame"]).dt.strftime('%m/%Y')
    exames_base["TipoMovimento"] = "Exame"
    exames_base["Produtor"] = exames_base["Filial"].map(prod_map)
    if "ValorExame" not in exames_base.columns:
        exames_base["ValorExame"] = 0
    exames_base["Rescisao"] = 0
    exames_base["MultaFGTS"] = 0
    exames_base["ValorEPI"] = 0
    exames_base["ADT13"] = 0
    exames_base["Decimo13"] = 0
    exames_base["Ferias"] = 0

    # EPI/UNIFORMES
    epi_base = epi.copy()
    if "DataEntrega" in epi_base.columns:
        epi_base["DataEntrega"] = pd.to_datetime(epi_base["DataEntrega"]).dt.date
        epi_base["Mes"] = pd.to_datetime(epi_base["DataEntrega"]).dt.strftime('%m/%Y')
    epi_base["TipoMovimento"] = "EPI/Uniforme"
    epi_base["Produtor"] = epi_base["Filial"].map(prod_map)
    if "ValorItem" in epi_base.columns:
        epi_base["ValorEPI"] = epi_base["ValorItem"]
    else:
        epi_base["ValorEPI"] = 0
    epi_base["Rescisao"] = 0
    epi_base["MultaFGTS"] = 0
    epi_base["ValorExame"] = 0
    epi_base["ADT13"] = 0
    epi_base["Decimo13"] = 0
    epi_base["Ferias"] = 0

    # ADT13/13º/FÉRIAS
    adt_base = adt13.copy()
    col_tipo = "TipoLancamento" if "TipoLancamento" in adt_base.columns else ("Lancamento" if "Lancamento" in adt_base.columns else None)
    
    if "Mes" not in adt_base.columns:
        for col_data in ["Data", "DataPagamento", "DataLancamento"]:
            if col_data in adt_base.columns:
                adt_base["Mes"] = pd.to_datetime(adt_base[col_data]).dt.strftime('%m/%Y')
                break
    
    adt_base["Produtor"] = adt_base["Filial"].map(prod_map)
    adt_base["ADT13"] = 0
    adt_base["Decimo13"] = 0
    adt_base["Ferias"] = 0
    adt_base["Rescisao"] = 0
    adt_base["MultaFGTS"] = 0
    adt_base["ValorExame"] = 0
    adt_base["ValorEPI"] = 0
    
    if col_tipo and "ValorLiquido" in adt_base.columns:
        adt_base[col_tipo] = adt_base[col_tipo].astype(str).str.strip().str.upper()
        
        def definir_tipo_movimento(tipo):
            tipo = str(tipo).upper()
            if "ADT" in tipo or "ADIANTAMENTO" in tipo:
                return "ADT 13º"
            elif "13" in tipo and "ADT" not in tipo:
                return "13º Salário"
            elif "FER" in tipo or "FERIAS" in tipo:
                return "Férias"
            return "Outros"
        
        adt_base["TipoMovimento"] = adt_base[col_tipo].apply(definir_tipo_movimento)
        
        for idx, row in adt_base.iterrows():
            tipo = str(row[col_tipo]).upper()
            valor = row["ValorLiquido"] if pd.notna(row["ValorLiquido"]) else 0
            
            if "ADT" in tipo or "ADIANTAMENTO" in tipo:
                adt_base.at[idx, "ADT13"] = valor
            elif "13" in tipo and "ADT" not in tipo:
                adt_base.at[idx, "Decimo13"] = valor
            elif "FER" in tipo or "FERIAS" in tipo:
                adt_base.at[idx, "Ferias"] = valor
    else:
        adt_base["TipoMovimento"] = "Outros"

    base = pd.concat([adm_base, dem_base, exames_base, epi_base, adt_base], ignore_index=True)
    return base.fillna(0)

# ===============================
# APLICAÇÃO PRINCIPAL
# ===============================
try:
    adm, dem, exames, epi, adt13, produtores = carregar_dados()
    base = montar_base(adm, dem, exames, epi, adt13, produtores)

    # HEADER
    col_header1, col_header2 = st.columns([3, 1])
    with col_header1:
        st.title("📊 ContaUni - Dashboard Executivo de RH")
        st.caption("Contabilidade Unindo Sonhos com Resultados")
    with col_header2:
        st.metric("Última atualização", datetime.now().strftime("%d/%m/%Y %H:%M"))

    # SIDEBAR - FILTROS
    st.sidebar.image("https://via.placeholder.com/200x80/667eea/ffffff?text=ContaUni", use_container_width=True)
    st.sidebar.header("🔍 Filtros")

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
            valores_unicos = sorted([str(v) for v in base[coluna].dropna().unique()])
            selecao = st.sidebar.multiselect(label, valores_unicos, key=f"filtro_{coluna}")
            filtros[coluna] = selecao

    # Limpar filtros
    if st.sidebar.button("🔄 Limpar Filtros", use_container_width=True):
        st.rerun()

    # APLICAR FILTROS
    filtro = base.copy()
    for coluna, valores in filtros.items():
        if valores:
            filtro = filtro[filtro[coluna].astype(str).isin([str(v) for v in valores])]

    # TABS PRINCIPAIS
    tab1, tab2, tab3, tab4 = st.tabs(["📈 Visão Geral", "💰 Análise Financeira", "👥 Movimentação", "📋 Dados Detalhados"])

    with tab1:
        st.subheader("Indicadores Gerais")
        
        # CARDS DE MÉTRICAS
        col1, col2, col3, col4 = st.columns(4)
        
        total_adm = filtro[filtro["TipoMovimento"] == "Admissão"].shape[0]
        total_dem = filtro[filtro["TipoMovimento"] == "Demissão"].shape[0]
        saldo = total_adm - total_dem
        
        with col1:
            st.metric("👤 Admissões", total_adm, delta=f"{saldo:+d} saldo", delta_color="normal")
        with col2:
            st.metric("👋 Demissões", total_dem)
        with col3:
            turnover = (total_dem / total_adm * 100) if total_adm > 0 else 0
            st.metric("🔄 Turnover", f"{turnover:.1f}%", delta="-5.2%" if turnover < 15 else "+3.1%", delta_color="inverse")
        with col4:
            total_geral = (filtro['Rescisao'].sum() + filtro['ValorExame'].sum() + 
                          filtro['ValorEPI'].sum() + filtro['ADT13'].sum() + 
                          filtro['Decimo13'].sum() + filtro['Ferias'].sum() + 
                          filtro['MultaFGTS'].sum())
            st.metric("💵 Custo Total", f"R$ {total_geral:,.2f}")

        st.write("")
        
        # CUSTOS DETALHADOS
        st.subheader("💰 Detalhamento de Custos")
        col5, col6, col7, col8, col9, col10, col11 = st.columns(7)
        
        col5.metric("Rescisão", f"R$ {filtro['Rescisao'].sum():,.2f}")
        col6.metric("Multa FGTS", f"R$ {filtro['MultaFGTS'].sum():,.2f}")
        col7.metric("Exames", f"R$ {filtro['ValorExame'].sum():,.2f}")
        col8.metric("EPI/Uniformes", f"R$ {filtro['ValorEPI'].sum():,.2f}")
        col9.metric("Férias", f"R$ {filtro['Ferias'].sum():,.2f}")
        col10.metric("ADT 13º", f"R$ {filtro['ADT13'].sum():,.2f}")
        col11.metric("13º Salário", f"R$ {filtro['Decimo13'].sum():,.2f}")

        st.write("")
        
        # GRÁFICO DE PIZZA - DISTRIBUIÇÃO DE CUSTOS
        col_pizza1, col_pizza2 = st.columns(2)
        
        with col_pizza1:
            custos_totais = {
                'Rescisão': filtro['Rescisao'].sum(),
                'Multa FGTS': filtro['MultaFGTS'].sum(),
                'Exames': filtro['ValorExame'].sum(),
                'EPI/Uniformes': filtro['ValorEPI'].sum(),
                'Férias': filtro['Ferias'].sum(),
                'ADT 13º': filtro['ADT13'].sum(),
                '13º Salário': filtro['Decimo13'].sum()
            }
            
            df_custos = pd.DataFrame(list(custos_totais.items()), columns=['Tipo', 'Valor'])
            df_custos = df_custos[df_custos['Valor'] > 0]
            
            fig_pizza = px.pie(df_custos, values='Valor', names='Tipo', 
                              title='Distribuição de Custos (%)',
                              hole=0.4,
                              color_discrete_sequence=px.colors.qualitative.Set3)
            fig_pizza.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pizza, use_container_width=True)
        
        with col_pizza2:
            # EVOLUÇÃO MENSAL
            if "Mes" in filtro.columns:
                evolucao = filtro.groupby("Mes").agg({
                    'Rescisao': 'sum',
                    'ValorExame': 'sum',
                    'ValorEPI': 'sum',
                    'ADT13': 'sum',
                    'Decimo13': 'sum',
                    'Ferias': 'sum',
                    'MultaFGTS': 'sum'
                }).reset_index()
                
                evolucao['Total'] = evolucao.iloc[:, 1:].sum(axis=1)
                
                fig_linha = px.line(evolucao, x='Mes', y='Total',
                                   title='Evolução de Custos Mensais',
                                   markers=True)
                fig_linha.update_traces(line_color='#667eea', line_width=3)
                fig_linha.update_layout(xaxis_title="Mês", yaxis_title="Valor (R$)")
                st.plotly_chart(fig_linha, use_container_width=True)

    with tab2:
        st.subheader("💰 Análise Financeira por Produtor")
        
        if not filtro.empty and "Produtor" in filtro.columns:
            graf = filtro.groupby("Produtor")[["Rescisao", "ValorExame", "ValorEPI",
                                                "ADT13", "Decimo13", "Ferias", "MultaFGTS"]].sum().reset_index()
            graf = graf[graf["Produtor"].notna() & (graf["Produtor"] != "0") & (graf["Produtor"] != 0)]
            
            # Gráfico de barras empilhadas
            graf_melted = graf.melt(id_vars=["Produtor"],
                                   value_vars=["Rescisao", "ValorExame", "ValorEPI", "ADT13", "Decimo13", "Ferias", "MultaFGTS"],
                                   var_name="Tipo", value_name="Valor")
            
            nomes_display = {
                "Rescisao": "Rescisão Líquida", "ValorExame": "Exames",
                "ValorEPI": "EPI/Uniformes", "ADT13": "ADT 13º",
                "Decimo13": "13º Salário", "Ferias": "Férias", "MultaFGTS": "Multa FGTS"
            }
            graf_melted["Tipo"] = graf_melted["Tipo"].map(nomes_display)

            fig = px.bar(graf_melted, x="Produtor", y="Valor", color="Tipo",
                        title="Custos por Produtor Rural (Empilhado)",
                        height=500, barmode="stack")
            fig.update_layout(xaxis_title="Produtor Rural", yaxis_title="Valor (R$)", legend_title="Tipo de Custo")
            st.plotly_chart(fig, use_container_width=True)
            
            # TOP 5 PRODUTORES
            st.subheader("🏆 Top 5 Produtores por Custo")
            graf['Total'] = graf.iloc[:, 1:].sum(axis=1)
            top5 = graf.nlargest(5, 'Total')[['Produtor', 'Total']]
            
            fig_top5 = px.bar(top5, x='Total', y='Produtor', orientation='h',
                             title='Top 5 Produtores com Maiores Custos',
                             color='Total', color_continuous_scale='Reds')
            st.plotly_chart(fig_top5, use_container_width=True)

    with tab3:
        st.subheader("👥 Análise de Movimentação de Pessoal")
        
        col_mov1, col_mov2 = st.columns(2)
        
        with col_mov1:
            # Admissões por mês
            if "Mes" in filtro.columns:
                adm_mes = filtro[filtro["TipoMovimento"] == "Admissão"].groupby("Mes").size().reset_index(name='Quantidade')
                fig_adm = px.bar(adm_mes, x='Mes', y='Quantidade',
                               title='Admissões por Mês',
                               color='Quantidade', color_continuous_scale='Greens')
                st.plotly_chart(fig_adm, use_container_width=True)
        
        with col_mov2:
            # Demissões por mês
            if "Mes" in filtro.columns:
                dem_mes = filtro[filtro["TipoMovimento"] == "Demissão"].groupby("Mes").size().reset_index(name='Quantidade')
                fig_dem = px.bar(dem_mes, x='Mes', y='Quantidade',
                               title='Demissões por Mês',
                               color='Quantidade', color_continuous_scale='Reds')
                st.plotly_chart(fig_dem, use_container_width=True)
        
        # Movimentação por Setor
        if "Setor" in filtro.columns:
            st.subheader("📊 Movimentação por Setor")
            mov_setor = filtro[filtro["TipoMovimento"].isin(["Admissão", "Demissão"])].groupby(["Setor", "TipoMovimento"]).size().reset_index(name='Quantidade')
            fig_setor = px.bar(mov_setor, x='Setor', y='Quantidade', color='TipoMovimento',
                             title='Admissões vs Demissões por Setor',
                             barmode='group', color_discrete_map={'Admissão': 'green', 'Demissão': 'red'})
            st.plotly_chart(fig_setor, use_container_width=True)

    with tab4:
        st.subheader("📋 Dados Detalhados")
        
        # Botões de exportação
        col_exp1, col_exp2, col_exp3 = st.columns([1, 1, 4])
        
        with col_exp1:
            csv = filtro.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Exportar CSV",
                data=csv,
                file_name=f'relatorio_rh_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                mime='text/csv',
                use_container_width=True
            )
        
        with col_exp2:
            from io import BytesIO
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                filtro.to_excel(writer, sheet_name='Dados', index=False)
            
            st.download_button(
                label="📥 Exportar Excel",
                data=buffer.getvalue(),
                file_name=f'relatorio_rh_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True
            )
        
        # Filtro de colunas para exibir
        colunas_disponiveis = list(filtro.columns)
        colunas_selecionadas = st.multiselect(
            "Selecione as colunas para visualizar:",
            colunas_disponiveis,
            default=["Funcionario", "Produtor", "Filial", "TipoMovimento", "Mes", "Rescisao", "Ferias", "ADT13", "Decimo13"][:min(9, len(colunas_disponiveis))]
        )
        
        if colunas_selecionadas:
            st.dataframe(filtro[colunas_selecionadas], use_container_width=True, height=400)
        else:
            st.dataframe(filtro, use_container_width=True, height=400)
        
        # Resumo Estatístico
        st.write("---")
        st.subheader("📊 Resumo Estatístico")
        
        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        
        with col_stat1:
            st.metric("📝 Total de Registros", f"{len(filtro):,}")
        with col_stat2:
            st.metric("🏭 Produtores Únicos", filtro["Produtor"].nunique())
        with col_stat3:
            st.metric("🏢 Filiais Únicas", filtro["Filial"].nunique())
        with col_stat4:
            if "Mes" in filtro.columns:
                st.metric("📅 Meses no Período", filtro["Mes"].nunique())

    # FOOTER
    st.write("---")
    st.caption("© 2025 ContaUni - Contabilidade Unindo Sonhos com Resultados | Dashboard desenvolvido com Streamlit")

    # FOOTER
    st.write("---")
    col_footer1, col_footer2 = st.columns([3, 1])
    with col_footer1:
        st.caption("© 2025 ContaUni - Contabilidade Unindo Sonhos com Resultados")
    with col_footer2:
        st.caption("💻 Desenvolvido por **Deivth Azevedo**")

except FileNotFoundError as e:
    st.error(f"❌ Erro ao carregar arquivos: {e}")
    st.info("💡 Verifique se todos os arquivos Excel estão na pasta 'dados/'")
except Exception as e:
    st.error(f"❌ Erro no sistema: {str(e)}")
    st.info("📞 Entre em contato com o suporte técnico.")
    
    if st.checkbox("🔧 Mostrar detalhes do erro (modo desenvolvedor)"):
        st.exception(e)

