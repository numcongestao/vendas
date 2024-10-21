import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Função para carregar dados do Excel
def load_data(file):
    xls = pd.ExcelFile(file)
    data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
    return data

# Função para calcular totais e margem ponderada corrigida
def analyze_data(df):
    total_venda = df['VENDA'].sum()
    total_custo_mercadoria = df['CUSTO MERCADORIA'].sum()
    margem_ponderada_total = df['MARGEM PONDERADA'].sum()
    
    return total_venda, total_custo_mercadoria, margem_ponderada_total

# Função para criar gráfico de barras interativo comparando vendas e custos
def plot_barras_comparativo(data, selected_sheets):
    vendas = []
    custos = []
    sheets = []
    
    for sheet in selected_sheets:
        df = data[sheet]
        total_venda, total_custo_mercadoria, _ = analyze_data(df)
        vendas.append(total_venda)
        custos.append(total_custo_mercadoria)
        sheets.append(sheet)
    
    # Gráfico interativo de barras usando Plotly
    fig = go.Figure(data=[
        go.Bar(name='Vendas', x=sheets, y=vendas, marker_color='blue'),
        go.Bar(name='Custos', x=sheets, y=custos, marker_color='red')
    ])
    
    fig.update_layout(
        title="Comparação de Vendas e Custos por Mês",
        xaxis_title="Meses",
        yaxis_title="Valores (R$)",
        barmode='group',
        template='plotly_dark'
    )
    
    st.plotly_chart(fig)

# Função para criar gráfico de linha comparando vendas e custos
def plot_linhas_comparativo(data, selected_sheets):
    vendas = []
    custos = []
    sheets = []
    
    for sheet in selected_sheets:
        df = data[sheet]
        total_venda, total_custo_mercadoria, _ = analyze_data(df)
        vendas.append(total_venda)
        custos.append(total_custo_mercadoria)
        sheets.append(sheet)
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=sheets, y=vendas, mode='lines+markers', name='Vendas', line=dict(color='blue')))
    fig.add_trace(go.Scatter(x=sheets, y=custos, mode='lines+markers', name='Custos', line=dict(color='red')))

    fig.update_layout(
        title="Vendas vs Custos - Linha",
        xaxis_title="Meses",
        yaxis_title="Valores (R$)",
        template='plotly_dark'
    )

    st.plotly_chart(fig)

# Função para gráfico de barras da margem ponderada
def plot_margem_ponderada(data, selected_sheets):
    margens = []
    sheets = []
    
    for sheet in selected_sheets:
        df = data[sheet]
        _, _, margem_ponderada_total = analyze_data(df)
        margens.append(margem_ponderada_total)
        sheets.append(sheet)
    
    fig = px.bar(
        x=sheets, 
        y=margens, 
        labels={'x': 'Meses', 'y': 'Margem Ponderada (%)'}, 
        title="Margem Ponderada por Mês", 
        template='plotly_dark', 
        color=margens, 
        color_continuous_scale='viridis'
    )
    
    st.plotly_chart(fig)

# Função para criar gráfico de comparação de margem ponderada e margem por produto
def plot_margens_por_produto(data, selected_sheets):
    # DataFrame vazio para acumular dados
    all_products = pd.DataFrame()

    for sheet in selected_sheets:
        df = data[sheet]
        
        # Verificar se as colunas necessárias estão presentes
        if 'PRODUTOS' in df.columns and 'MARGEM' in df.columns and 'MARGEM PONDERADA' in df.columns:
            df['Mês'] = sheet  # Adiciona uma coluna com o nome do mês
            all_products = pd.concat([all_products, df[['PRODUTOS', 'MARGEM', 'MARGEM PONDERADA', 'Mês']]], ignore_index=True)
        else:
            st.warning(f"As colunas necessárias não foram encontradas na planilha '{sheet}'. As colunas disponíveis são: {df.columns.tolist()}")

    if not all_products.empty:
        # Criar um gráfico de linhas com cores diferentes para cada mês
        fig = px.line(
            all_products.melt(id_vars=["PRODUTOS", "Mês"], value_vars=["MARGEM", "MARGEM PONDERADA"]),
            x="PRODUTOS",
            y="value",
            color="Mês",  # Usar "Mês" para diferenciar as cores
            line_group="variable",  # Agrupar por tipo de margem
            markers=True,
            labels={"value": "Margem (%)", "variable": "Tipo de Margem"},
            title="Comparação de Margens por Produto",
            template='plotly_dark',
            color_discrete_sequence=px.colors.qualitative.Set1  # Paleta de cores para os meses
        )
        
        st.plotly_chart(fig)




# Streamlit App
def main():
    st.title('Análise de Vendas por Mês com Gráficos Interativos')

    # Upload do arquivo Excel
    uploaded_file = st.file_uploader('Escolha um arquivo Excel com dados de vendas', type=['xlsx'])
    
    if uploaded_file:
        # Carregar dados de todas as sheets
        data = load_data(uploaded_file)
        
        # Exibir opções para selecionar as sheets (meses)
        sheets = list(data.keys())
        selected_sheets = st.multiselect('Selecione os meses que deseja comparar', sheets, default=sheets[:2])

        if selected_sheets:
            # Exibir análises por mês
            st.subheader('Totais e Margens por Mês Selecionado')
            for sheet in selected_sheets:
                df = data[sheet]
                total_venda, total_custo_mercadoria, margem_ponderada = analyze_data(df)
                
                st.markdown(f"### Dados para o mês: {sheet}")
                st.write(f"**Total de Vendas**: R$ {total_venda:,.2f}")
                st.write(f"**Custo Total da Mercadoria Vendida**: R$ {total_custo_mercadoria:,.2f}")
                st.write(f"**Margem Ponderada Total**: {margem_ponderada:.2f}%")
                st.write(df)  # Exibir a planilha completa

            # Gráficos comparativos
            st.subheader('Gráficos Comparativos')

            # Gráfico interativo de barras comparando vendas e custos
            plot_barras_comparativo(data, selected_sheets)

            # Gráfico de linha comparando vendas e custos ao longo dos meses
            plot_linhas_comparativo(data, selected_sheets)

            # Gráfico de barras interativo para margem ponderada
            plot_margem_ponderada(data, selected_sheets)

            # Gráfico comparando margens ponderada e margem por produto
            st.subheader('Comparação de Margens por Produto')
            plot_margens_por_produto(data, selected_sheets)

if __name__ == '__main__':
    main()
