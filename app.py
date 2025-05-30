import streamlit as st
import pandas as pd

st.title("Controlo Financeiro com Excel")

uploaded_file = st.file_uploader("Carrega o teu ficheiro Excel", type=["xlsx"])

if uploaded_file:
    st.success("Ficheiro carregado!")

    excel_data = pd.read_excel(uploaded_file, sheet_name=None)
    all_vendas = []
    all_despesas = []

    for sheet_name, df in excel_data.items():
        st.subheader(f"Folha: {sheet_name}")
        df.columns = df.columns.astype(str)
        df = df.dropna(how='all').reset_index(drop=True)

        if 'data' in df.values.astype(str).flatten():
            try:
                data_col = df.apply(lambda row: row.astype(str).str.contains('data', case=False).any(), axis=1).idxmax()
                descricao_col = df.apply(lambda row: row.astype(str).str.contains('descricao|cliente', case=False).any(), axis=1).idxmax()
                liquido_col = df.apply(lambda row: row.astype(str).str.contains('liquido|valor', case=False).any(), axis=1).idxmax()
                imposto_col = df.apply(lambda row: row.astype(str).str.contains('imposto|iva', case=False).any(), axis=1).idxmax()

                dados_validos = df.iloc[data_col+1:].copy()
                dados_validos.columns = df.iloc[data_col]
                dados_validos = dados_validos[[col for col in dados_validos.columns if col is not None and 'Unnamed' not in str(col)]]

                dados_validos['mes'] = pd.to_datetime(dados_validos['data'], errors='coerce').dt.month_name()
                dados_validos['ano'] = pd.to_datetime(dados_validos['data'], errors='coerce').dt.year

                vendas = dados_validos[dados_validos['tipo'] == 'Venda'].copy()
                despesas = dados_validos[dados_validos['tipo'] == 'Despesa'].copy()

                all_vendas.append(vendas)
                all_despesas.append(despesas)

                st.write("Exemplo de dados processados:")
                st.dataframe(dados_validos.head())

            except Exception as e:
                st.warning(f"Erro ao processar a folha '{sheet_name}': {e}")

    if all_vendas or all_despesas:
        st.header("Resumo Geral")

        if all_vendas:
            vendas_df = pd.concat(all_vendas, ignore_index=True)
            total_vendas = vendas_df['liquido FAC'].sum()
            total_iva_vendas = vendas_df['imposto'].sum()
        else:
            total_vendas = 0
            total_iva_vendas = 0

        if all_despesas:
            despesas_df = pd.concat(all_despesas, ignore_index=True)
            total_despesas = despesas_df['liquido FAC'].sum()
            total_iva_compras = despesas_df['imposto'].sum()
        else:
            total_despesas = 0
            total_iva_compras = 0

        iva_saldo = total_iva_vendas - total_iva_compras

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Vendas", f"{total_vendas:.2f} €")
        col2.metric("Total Despesas", f"{total_despesas:.2f} €")
        col3.metric("Saldo IVA (Vendas - Compras)", f"{iva_saldo:.2f} €")

        if all_vendas:
            vendas_mensal = vendas_df.groupby(['ano', 'mes'])['liquido FAC'].sum().reset_index()
            st.subheader("Vendas Mensais")
            st.bar_chart(vendas_mensal.set_index('mes')['liquido FAC'])

        if st.button("Exportar dados para CSV"):
            if all_vendas:
                vendas_df.to_csv("vendas_exportadas.csv", index=False)
                st.success("Vendas exportadas como 'vendas_exportadas.csv'")
            if all_despesas:
                despesas_df.to_csv("despesas_exportadas.csv", index=False)
                st.success("Despesas exportadas como 'despesas_exportadas.csv'")
    else:
        st.info("Nenhum dado válido encontrado no Excel.")
else:
    st.info("Por favor, carrega o teu ficheiro Excel.")