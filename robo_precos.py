import streamlit as st
import pandas as pd
import urllib.parse

st.set_page_config(page_title="Auditoria de Preços", layout="wide", page_icon="🤖")
st.title("🤖 Robô de Auditoria de Preços")
st.write("Insira os arquivos abaixo para cruzar o faturamento com a tabela oficial.")

col1, col2 = st.columns(2)
with col1:
    arquivo_fat = st.file_uploader("📥 CSV do Faturamento", type=['csv'])
with col2:
    arquivo_tab = st.file_uploader("📥 CSV da Tabela de Preços", type=['csv'])

if arquivo_fat and arquivo_tab:
    with st.spinner("Analisando dados..."):
        try:
            df_fat = pd.read_csv(arquivo_fat)
            df_tab = pd.read_csv(arquivo_tab)

            # --- LIMPEZA DE TEXTO PARA CRUZAMENTO ---
            df_fat['recurso.variedade.c'] = df_fat['recurso.variedade.c'].astype(str).str.strip()
            df_fat['recurso.classemarca.n'] = df_fat['recurso.classemarca.n'].astype(str).str.strip()
            df_fat['Tabela'] = df_fat['Tabela'].astype(str).str.strip()
            
            df_tab['Variedade.c'] = df_tab['Variedade.c'].astype(str).str.strip()
            df_tab['marca.n'] = df_tab['marca.n'].astype(str).str.strip()
            df_tab['TABELA'] = df_tab['TABELA'].astype(str).str.strip()

            # --- LIMPEZA DE NÚMEROS E PESOS ---
            def limpar_numero(val):
                if pd.api.types.is_number(val):
                    return val
                val = str(val).strip().replace('R$', '').replace(' ', '')
                if ',' in val:
                    val = val.replace('.', '').replace(',', '.')
                return pd.to_numeric(val, errors='coerce')
            
            # Limpa o "Kg" da tabela de preços para cruzar com o número do faturamento
            def limpar_peso(val):
                val = str(val).upper().replace('KG', '').strip()
                if ',' in val:
                    val = val.replace(',', '.')
                return pd.to_numeric(val, errors='coerce')

            df_tab['Peso_Tabela_Num'] = df_tab['PESO CX'].apply(limpar_peso)
            df_fat['Peso_Fat_Num'] = pd.to_numeric(df_fat['Peso Caixa'], errors='coerce')

            df_fat['Preço Caixa'] = df_fat['Preço Caixa'].apply(limpar_numero)
            df_tab['Preço Final CX'] = df_tab['Preço Final CX'].apply(limpar_numero)

            # --- TRATAMENTO DE DATAS ---
            df_fat['emissaomovdate'] = pd.to_datetime(df_fat['emissaomovdate'], format='%d/%m/%Y', errors='coerce')
            df_tab['vigencia_inicio'] = pd.to_datetime(df_tab['vigencia_inicio'], format='%d/%m/%Y', errors='coerce')
            df_tab['vigencia_fim'] = pd.to_datetime(df_tab['vigencia_fim'], format='%d/%m/%Y', errors='coerce')

            # --- CRUZAMENTO (MERGE) - AGORA USANDO O PESO! ---
            df_cruzado = pd.merge(
                df_fat, df_tab,
                left_on=['recurso.variedade.c', 'recurso.classemarca.n', 'Tabela', 'Peso_Fat_Num'],
                right_on=['Variedade.c', 'marca.n', 'TABELA', 'Peso_Tabela_Num'],
                how='left'
            )

            # --- DIAGNÓSTICO ---
            with st.expander("🛠️ Ver Diagnóstico do Robô (Clique para expandir)"):
                st.write(f"**1.** Total de linhas no Faturamento original: `{len(df_fat)}`")
                sem_tabela = df_cruzado[df_cruzado['Preço Final CX'].isna()]
                st.write(f"**2.** Linhas onde o robô não encontrou Tabela Correspondente: `{len(sem_tabela)}`")
                if not sem_tabela.empty:
                    st.warning("Verifique se estas combinações existem na sua Tabela de Preços (Variedade, Marca, Tabela e Peso):")
                    st.dataframe(sem_tabela[['recurso.variedade.c', 'recurso.classemarca.n', 'Tabela', 'Peso_Fat_Num']].drop_duplicates())

            # --- FILTROS DE VALIDAÇÃO (Data + Tipo) ---
            df_cruzado['Tipo_Num'] = pd.to_numeric(df_cruzado['Tipo'], errors='coerce')
            
            df_valido = df_cruzado[
                (df_cruzado['emissaomovdate'] >= df_cruzado['vigencia_inicio']) &
                (df_cruzado['emissaomovdate'] <= df_cruzado['vigencia_fim']) &
                (df_cruzado['Tipo_Num'] >= df_cruzado['Tipomin']) &
                (df_cruzado['Tipo_Num'] <= df_cruzado['Tipomax'])
            ].copy()

            # --- BUSCA DE DIVERGÊNCIAS ---
            df_valido = df_valido.dropna(subset=['Preço Final CX', 'Preço Caixa'])
            df_valido['Preço Caixa'] = df_valido['Preço Caixa'].round(2)
            df_valido['Preço Final CX'] = df_valido['Preço Final CX'].round(2)
            
            divergencias = df_valido[df_valido['Preço Caixa'] != df_valido['Preço Final CX']]

            st.divider()
            
            if not divergencias.empty:
                st.error(f"⚠️ Foram encontradas {len(divergencias)} divergências de preço!")
                
                colunas_exibicao = ['emissaomovdate', 'cliente.c', 'recurso.variedade.c', 'Tabela', 'Peso_Fat_Num', 'Tipo', 'Preço Caixa', 'Preço Final CX']
                st.dataframe(divergencias[colunas_exibicao], use_container_width=True)
                
                st.subheader("✉️ Notificar Responsáveis")
                
                resumo_itens = ""
                for index, row in divergencias.head(5).iterrows():
                    resumo_itens += f"• Cliente: {row['cliente.c']} | {row['recurso.variedade.c']} (Tipo {row['Tipo']}) | Faturado: R$ {row['Preço Caixa']} | Tabela: R$ {row['Preço Final CX']}\n"
                
                if len(divergencias) > 5:
                    resumo_itens += f"\n... e mais {len(divergencias) - 5} item(ns). O detalhamento está em anexo.\n"

                modelo_email = f"""Olá equipe Comercial / Faturamento,

O Robô de Auditoria identificou {len(divergencias)} lançamento(s) faturado(s) com preço divergente da Tabela Oficial.

Resumo:
{resumo_itens}
Por favor, verifiquem se há desconto aprovado para estes casos.

Atenciosamente,
Auditoria de Preços"""
                
                st.text_area("📋 Modelo de E-mail gerado:", value=modelo_email, height=280)
                
                assunto = urllib.parse.quote("⚠️ Alerta: Divergência de Preços (Faturamento x Tabela)")
                corpo = urllib.parse.quote(modelo_email)
                email_destino = "comercial@suaempresa.com.br;faturamento@suaempresa.com.br"
                link_mailto = f"mailto:{email_destino}?subject={assunto}&body={corpo}"
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    csv_export = divergencias.to_csv(index=False).encode('utf-8')
                    st.download_button("📥 1. Baixar Relatório", data=csv_export, file_name='divergencias_precos.csv', mime='text/csv', use_container_width=True)
                with col_btn2:
                    st.link_button("📧 2. Abrir no E-mail", link_mailto, use_container_width=True)
            else:
                st.success("✅ Tudo certo! Nenhuma divergência válida encontrada.")

        except Exception as e:
            st.error(f"Ocorreu um erro ao processar os arquivos. Detalhe técnico: {e}")
