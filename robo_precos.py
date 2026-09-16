import streamlit as st
import pandas as pd
import urllib.parse

st.set_page_config(page_title="Auditoria de Preços", layout="wide", page_icon="🤖")
st.title("🤖 Robô de Auditoria de Preços")
st.write("Insira os arquivos abaixo para cruzar o faturamento com a tabela oficial.")

col1, col2 = st.columns(2)
with col1:
    arquivo_fat = st.file_uploader("📥 CSV do Faturamento (Looker)", type=['csv'])
with col2:
    arquivo_tab = st.file_uploader("📥 CSV da Tabela de Preços", type=['csv'])

if arquivo_fat and arquivo_tab:
    with st.spinner("Analisando dados..."):
        try:
            df_fat = pd.read_csv(arquivo_fat)
            df_tab = pd.read_csv(arquivo_tab)

            # --- NOVA FUNÇÃO DE LIMPEZA DE NÚMEROS ---
            # Remove R$, espaços e ajusta vírgulas para que o Python entenda como matemática
            def limpar_numero(val):
                if pd.api.types.is_number(val):
                    return val
                val = str(val).strip().replace('R$', '').replace(' ', '')
                if ',' in val:
                    val = val.replace('.', '').replace(',', '.')
                return pd.to_numeric(val, errors='coerce')

            # Aplica a limpeza nas colunas de valor antes da divisão
            df_fat['totalfinanceiro'] = df_fat['totalfinanceiro'].apply(limpar_numero)
            df_fat['QTD caixa'] = df_fat['QTD caixa'].apply(limpar_numero)
            
            if 'Preço Final CX' in df_tab.columns:
                df_tab['Preço Final CX'] = df_tab['Preço Final CX'].apply(limpar_numero)
            # ----------------------------------------

            df_fat['emissaomovdate'] = pd.to_datetime(df_fat['emissaomovdate'], format='%d/%m/%Y', errors='coerce')
            df_tab['vigencia_inicio'] = pd.to_datetime(df_tab['vigencia_inicio'], errors='coerce')
            df_tab['vigencia_fim'] = pd.to_datetime(df_tab['vigencia_fim'], errors='coerce')

            # Agora a divisão vai funcionar perfeitamente
            df_fat['Preco_Faturado_CX'] = df_fat['totalfinanceiro'] / df_fat['QTD caixa']

            df_cruzado = pd.merge(
                df_fat, df_tab,
                left_on=['recurso.variedade.c', 'cliente.classe.un', 'recurso.modelocaixa'],
                right_on=['VARIEDADE', 'TABELA', 'CAIXA'],
                how='left'
            )

            df_valido = df_cruzado[
                (df_cruzado['emissaomovdate'] >= df_cruzado['vigencia_inicio']) &
                (df_cruzado['emissaomovdate'] <= df_cruzado['vigencia_fim'])
            ]

            df_valido['Preco_Faturado_CX'] = df_valido['Preco_Faturado_CX'].round(2)
            df_valido['Preço Final CX'] = df_valido['Preço Final CX'].round(2)
            divergencias = df_valido[df_valido['Preco_Faturado_CX'] != df_valido['Preço Final CX']]

            st.divider()
            
            if not divergencias.empty:
                st.error(f"⚠️ Foram encontradas {len(divergencias)} divergências de preço!")
                
                colunas = ['emissaomovdate', 'cliente.c', 'recurso.variedade.c', 'cliente.classe.un', 'recurso.modelocaixa', 'Preco_Faturado_CX', 'Preço Final CX']
                st.dataframe(divergencias[colunas], use_container_width=True)
                
                st.subheader("✉️ Notificar Responsáveis")
                
                resumo_itens = ""
                for index, row in divergencias.head(5).iterrows():
                    resumo_itens += f"• Cliente: {row['cliente.c']} | Variedade: {row['recurso.variedade.c']} | Faturado: R$ {row['Preco_Faturado_CX']} | Tabela: R$ {row['Preço Final CX']}\n"
                
                if len(divergencias) > 5:
                    resumo_itens += f"\n... e mais {len(divergencias) - 5} item(ns). O detalhamento completo está no arquivo CSV em anexo.\n"

                modelo_email = f"""Olá equipe Comercial / Faturamento,

O Robô de Auditoria identificou {len(divergencias)} lançamento(s) faturado(s) com preço divergente da Tabela Oficial.

Resumo das divergências:
{resumo_itens}
Por favor, verifiquem se houve alguma exceção ou desconto aprovado para estes casos, ou se é necessário realizar um ajuste.

Atenciosamente,
Auditoria de Preços"""
                
                st.text_area("📋 Modelo de E-mail gerado (Revise ou copie se necessário):", value=modelo_email, height=280)
                
                assunto = urllib.parse.quote("⚠️ Alerta: Divergência de Preços (Faturamento x Tabela)")
                corpo = urllib.parse.quote(modelo_email)
                email_destino = "comercial@suaempresa.com.br;faturamento@suaempresa.com.br"
                link_mailto = f"mailto:{email_destino}?subject={assunto}&body={corpo}"
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    csv = divergencias.to_csv(index=False).encode('utf-8')
                    st.download_button("📥 1. Baixar Relatório (Para Anexar)", data=csv, file_name='divergencias_precos.csv', mime='text/csv', use_container_width=True)
                with col_btn2:
                    st.link_button("📧 2. Abrir no E-mail (Outlook/Gmail)", link_mailto, use_container_width=True)
            else:
                st.success("✅ Tudo certo! Nenhum faturamento divergiu da tabela de preços.")

        except Exception as e:
            st.error(f"Ocorreu um erro ao processar os arquivos. Detalhe técnico: {e}")
