import streamlit as st
import pandas as pd

st.set_page_config(page_title="Auditoria de Preços", layout="wide", page_icon="🤖")
st.title("🤖 Robô de Auditoria de Preços")
st.write("Insira os arquivos abaixo para cruzar o faturamento com a tabela oficial.")

# Área de Upload
col1, col2 = st.columns(2)
with col1:
    arquivo_fat = st.file_uploader("📥 CSV do Faturamento (Looker)", type=['csv'])
with col2:
    arquivo_tab = st.file_uploader("📥 CSV da Tabela de Preços", type=['csv'])

if arquivo_fat and arquivo_tab:
    with st.spinner("Analisando dados..."):
        try:
            # Lendo os arquivos
            df_fat = pd.read_csv(arquivo_fat)
            df_tab = pd.read_csv(arquivo_tab)

            # Convertendo datas
            df_fat['emissaomovdate'] = pd.to_datetime(df_fat['emissaomovdate'], format='%d/%m/%Y', errors='coerce')
            df_tab['vigencia_inicio'] = pd.to_datetime(df_tab['vigencia_inicio'], errors='coerce')
            df_tab['vigencia_fim'] = pd.to_datetime(df_tab['vigencia_fim'], errors='coerce')

            # Calculando preço da caixa no faturamento
            df_fat['Preco_Faturado_CX'] = df_fat['totalfinanceiro'] / df_fat['QTD caixa']

            # Cruzando os dados
            df_cruzado = pd.merge(
                df_fat, df_tab,
                left_on=['recurso.variedade.c', 'cliente.classe.un', 'recurso.modelocaixa'],
                right_on=['VARIEDADE', 'TABELA', 'CAIXA'],
                how='left'
            )

            # Filtrando por data de vigência
            df_valido = df_cruzado[
                (df_cruzado['emissaomovdate'] >= df_cruzado['vigencia_inicio']) &
                (df_cruzado['emissaomovdate'] <= df_cruzado['vigencia_fim'])
            ]

            # Encontrando divergências
            df_valido['Preco_Faturado_CX'] = df_valido['Preco_Faturado_CX'].round(2)
            df_valido['Preço Final CX'] = df_valido['Preço Final CX'].round(2)
            divergencias = df_valido[df_valido['Preco_Faturado_CX'] != df_valido['Preço Final CX']]

            st.divider()
            
            # Exibindo Resultados
            if not divergencias.empty:
                st.error(f"⚠️ Foram encontradas {len(divergencias)} divergências de preço!")
                
                colunas = ['emissaomovdate', 'cliente.c', 'recurso.variedade.c', 'cliente.classe.un', 'recurso.modelocaixa', 'Preco_Faturado_CX', 'Preço Final CX']
                st.dataframe(divergencias[colunas], use_container_width=True)
                
                # Botão de Download
                csv = divergencias.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Baixar Relatório Completo",
                    data=csv,
                    file_name='divergencias_precos.csv',
                    mime='text/csv'
                )
            else:
                st.success("✅ Tudo certo! Nenhum faturamento divergiu da tabela de preços.")

        except Exception as e:
            st.error(f"Ocorreu um erro ao processar os arquivos. Verifique se o formato está correto. Detalhe técnico: {e}")
