# -*- coding: utf-8 -*-
"""
Robô de Auditoria de Preços — Itaueira
Cruza o faturamento com a Tabela de Preços oficial e aponta divergências.

Execução:
    pip install streamlit pandas numpy xlsxwriter
    streamlit run auditoria_precos.py
"""

import io
import json
import re
import unicodedata
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st

# =============================================================================
# CONFIGURAÇÃO DE NEGÓCIO
# =============================================================================

REGIOES_UF = {
    "NE": ["CE", "BA", "PE", "PB", "RN", "AL", "SE", "MA", "PI"],
    "NO": ["AM", "PA", "AP", "RR", "RO", "AC", "TO"],
    "SUL": ["RS", "SC", "PR"],
    "SE_CO": ["SP", "MG", "ES", "RJ", "GO", "DF", "MT", "MS"],
}
UF_PARA_REGIAO = {uf: reg for reg, ufs in REGIOES_UF.items() for uf in ufs}

# Famílias de produto: cada família usa uma nomenclatura diferente na coluna TABELA
FAMILIA_PIMENTAO = ["BAM", "BVM", "BLR", "BMS"]

# De/Para direto — independe da UF do cliente
DEPARA_DIRETO = {
    "CFI Mercado Interno Desconto 7%": "6. DescFinanc 7%",
    "CFI Mercado Interno Desconto 7% PR e RJ": "7. DescFinanc 7% PR+RJ",
    "CFI Mercado Interno Desconto 5%": "5. DescFinanc 5%",
    "CFI Mercado Interno Especial": "1.1 Merc Interno Esp",
    "CFI Mercado Interno GPA": "9. GPA",
    "CFI Mercado Interno SUL": "3. Merc Int SUL",
    "CFI Mercado Interno NO e RJ": "2. Merc Int (NOR, RJ, MT, MS)",
}

# Regra por região — usada quando o De/Para direto não se aplica
# (venda FOB / Mercado Interno: o preço segue a região do cliente)
REGRA_REGIAO_PADRAO = {
    "MELAO": {
        "NE": "4. Merc Int Sem Frete",
        "NO": "2. Merc Int (NOR, RJ, MT, MS)",
        "SUL": "3. Merc Int SUL",
        "SE_CO": "1. Merc Interno",
    },
    "PIMENTAO": {
        "NE": "1. NO/NE FOB",
        "NO": "1. NO/NE FOB",
        "SUL": "2. SU/SE/CO CIF SP",
        "SE_CO": "2. SU/SE/CO CIF SP",
    },
}

# RJ, MT e MS usam a tabela "NOR, RJ, MT, MS" mesmo estando no Sudeste/Centro-Oeste
UF_EXCECAO_NOR = ["RJ", "MT", "MS"]

# Tabelas do faturamento que não são venda a cliente — não entram na auditoria
TABELAS_FORA_ESCOPO_PADRAO = ["Estabelecimentos ITR"]

COLS_FAT = {
    "data": "emissaomovdate",
    "tabela": "Tabela",
    "uf": "cliente.uf.c",
    "grupo": "cliente.grupo.c",
    "cliente": "cliente.c",
    "produto": "recurso.n",
    "variedade": "recurso.variedade.c",
    "marca": "recurso.classemarca.n",
    "tipo": "Tipo",
    "peso": "Peso Caixa",
    "qtd": "QTD caixa",
    "preco": "Preço Caixa",
}
COLS_TAB = {
    "ini": "vigencia_inicio",
    "fim": "vigencia_fim",
    "variedade_label": "VARIEDADE",
    "tipo_label": "TIPO",
    "peso": "PESO CX",
    "tabela": "TABELA",
    "preco": "Preço Final CX",
    "variedade": "Variedade.c",
    "marca": "marca.n",
    "tipomin": "Tipomin",
    "tipomax": "Tipomax",
}

# =============================================================================
# PARSING
# =============================================================================


def ler_csv(arquivo) -> pd.DataFrame:
    """Lê CSV tolerando separador, encoding e BOM variados."""
    bruto = arquivo.getvalue() if hasattr(arquivo, "getvalue") else arquivo.read()
    ultimo_erro = None
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        for sep in (",", ";", "\t"):
            try:
                df = pd.read_csv(io.BytesIO(bruto), dtype=str, sep=sep, encoding=enc)
                if df.shape[1] > 1:
                    df.columns = [str(c).strip().lstrip("﻿") for c in df.columns]
                    return df
            except Exception as e:  # noqa: BLE001
                ultimo_erro = e
    raise ValueError(f"Não consegui ler o arquivo. Último erro: {ultimo_erro}")


def para_numero(valor):
    """Converte texto pt-BR em número. Trata 'R$ 1.234,56', '10 Kg', '4,5', '1.234'."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return np.nan
    if isinstance(valor, (int, float, np.number)):
        return float(valor)

    s = str(valor).strip()
    if not s or s.lower() in {"nan", "none", "-", "--"}:
        return np.nan

    s = s.replace("\xa0", " ")
    s = re.sub(r"(?i)\bkgs?\b", "", s)        # "10 Kg" -> "10"
    s = re.sub(r"(?i)r\$", "", s)
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s:
        return np.nan

    tem_virgula, tem_ponto = "," in s, "." in s
    if tem_virgula and tem_ponto:
        # o separador decimal é o que aparece por último
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") \
            else s.replace(",", "")
    elif tem_virgula:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")                # 1.234.567 -> milhar
    elif tem_ponto:
        inteiro, _, decimal = s.partition(".")
        if len(decimal) == 3 and len(inteiro) <= 3 and inteiro.lstrip("-").isdigit():
            s = inteiro + decimal             # 1.234 -> milhar, não 1,234
    return pd.to_numeric(s, errors="coerce")


def para_data(serie: pd.Series) -> pd.Series:
    """Converte datas aceitando dd/mm/aaaa e ISO."""
    d = pd.to_datetime(serie, format="%d/%m/%Y", errors="coerce")
    falta = d.isna() & serie.notna()
    if falta.any():
        d.loc[falta] = pd.to_datetime(serie[falta], errors="coerce", dayfirst=True)
    return d.dt.normalize()


def normalizar(valor) -> str:
    """Chave de cruzamento: sem acento, sem espaço duplo, maiúscula."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return ""
    s = unicodedata.normalize("NFKD", str(valor))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().upper()


def extrair_codigos_variedade(codigo, rotulo) -> list:
    """'BAM/BAM/BLR' + '8. Pimentão BVM/BAM/BLR' -> ['BAM', 'BLR', 'BVM']"""
    achados = {p.strip().upper() for p in str(codigo).split("/") if p.strip()}
    achados |= set(re.findall(r"\b[A-Z]{2,4}\b", normalizar(rotulo)))
    achados = {c for c in achados if c.isalpha() and 2 <= len(c) <= 4}
    return sorted(achados) or [str(codigo).strip().upper()]


def conferir_colunas(df: pd.DataFrame, mapa: dict, nome: str):
    faltam = [c for c in mapa.values() if c not in df.columns]
    if faltam:
        raise KeyError(
            f"O arquivo de {nome} não tem as colunas: {', '.join(faltam)}.\n"
            f"Colunas encontradas: {', '.join(map(str, df.columns))}"
        )


# =============================================================================
# PREPARO DAS BASES
# =============================================================================


@st.cache_data(show_spinner=False)
def preparar_faturamento(bruto: bytes) -> pd.DataFrame:
    df = ler_csv(io.BytesIO(bruto))
    conferir_colunas(df, COLS_FAT, "faturamento")
    c = COLS_FAT
    out = pd.DataFrame(index=df.index)
    out["data"] = para_data(df[c["data"]])
    out["tabela_fat"] = df[c["tabela"]].astype(str).str.strip()
    out["uf"] = df[c["uf"]].astype(str).str.strip().str.upper()
    out["grupo"] = df[c["grupo"]].astype(str).str.strip()
    out["cliente"] = df[c["cliente"]].astype(str).str.strip()
    out["produto"] = df[c["produto"]].astype(str).str.strip()
    out["variedade"] = df[c["variedade"]].astype(str).str.strip().str.upper()
    out["marca"] = df[c["marca"]].astype(str).str.strip()
    out["marca_key"] = out["marca"].map(normalizar)
    out["tipo_txt"] = df[c["tipo"]].astype(str).str.strip()
    out["tipo"] = pd.to_numeric(out["tipo_txt"], errors="coerce")
    out["peso"] = df[c["peso"]].map(para_numero).round(4)
    out["qtd"] = df[c["qtd"]].map(para_numero)
    out["preco_fat"] = df[c["preco"]].map(para_numero)
    out["linha_csv"] = df.index + 2
    return out.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def preparar_tabela(bruto: bytes) -> pd.DataFrame:
    df = ler_csv(io.BytesIO(bruto))
    conferir_colunas(df, COLS_TAB, "tabela de preços")
    c = COLS_TAB
    out = pd.DataFrame(index=df.index)
    out["vig_ini"] = para_data(df[c["ini"]])
    out["vig_fim"] = para_data(df[c["fim"]])
    out["tabela_of"] = df[c["tabela"]].astype(str).str.strip()
    out["variedade_raw"] = df[c["variedade"]].astype(str).str.strip()
    out["variedade_label"] = df[c["variedade_label"]].astype(str).str.strip()
    out["marca"] = df[c["marca"]].astype(str).str.strip()
    out["marca_key"] = out["marca"].map(normalizar)
    out["tipo_label"] = df[c["tipo_label"]].astype(str).str.strip()
    out["peso"] = df[c["peso"]].map(para_numero).round(4)
    out["preco_tab"] = df[c["preco"]].map(para_numero)
    out["tipo_min"] = df[c["tipomin"]].map(para_numero)
    out["tipo_max"] = df[c["tipomax"]].map(para_numero)

    # uma variedade composta vira várias linhas, uma por código
    out["variedade"] = out.apply(
        lambda r: extrair_codigos_variedade(r["variedade_raw"], r["variedade_label"]),
        axis=1,
    )
    return out.explode("variedade").reset_index(drop=True)


# =============================================================================
# REGRA DE PRECIFICAÇÃO
# =============================================================================


def montar_regra(fat: pd.DataFrame, depara: dict, regra_regiao: dict,
                 familia_pimentao: list, fora_escopo: list) -> pd.DataFrame:
    """Define, para cada linha faturada, qual linha da tabela oficial é a correta."""
    fat = fat.copy()
    fat["familia"] = np.where(fat["variedade"].isin(familia_pimentao), "PIMENTAO", "MELAO")
    fat["regiao"] = fat["uf"].map(UF_PARA_REGIAO)
    fat.loc[fat["uf"].isin(UF_EXCECAO_NOR), "regiao"] = "NO"

    depara_key = {normalizar(k): v for k, v in depara.items()}
    fora_key = {normalizar(x) for x in fora_escopo}

    def alvo(r):
        chave = normalizar(r["tabela_fat"])
        if chave in fora_key:
            return None, "Fora de escopo (não é venda a cliente)"
        if chave in depara_key:
            destino = depara_key[chave]
            # a família Pimentão não tem as tabelas regionais de melão
            if r["familia"] == "PIMENTAO" and destino.startswith(("2. Merc", "3. Merc")):
                destino = regra_regiao["PIMENTAO"].get(r["regiao"])
            return destino, "De/Para direto"
        if pd.isna(r["regiao"]):
            return None, f"UF '{r['uf']}' não classificada em nenhuma região"
        return regra_regiao[r["familia"]].get(r["regiao"]), "Regra por região"

    res = fat.apply(alvo, axis=1, result_type="expand")
    fat["tabela_alvo"] = res[0]
    fat["origem_regra"] = res[1]
    return fat


def auditar(fat: pd.DataFrame, tab: pd.DataFrame,
            tol_reais: float, tol_pct: float) -> pd.DataFrame:
    """LEFT join: nenhuma linha some — cada uma recebe um status."""
    auditavel = fat[fat["tabela_alvo"].notna()].copy()
    auditavel["_id"] = np.arange(len(auditavel))

    chaves_e = ["variedade", "marca_key", "tabela_alvo", "peso"]
    chaves_d = ["variedade", "marca_key", "tabela_of", "peso"]
    m = auditavel.merge(
        tab, left_on=chaves_e, right_on=chaves_d, how="left", suffixes=("", "_tab")
    )

    tem_candidato = m["preco_tab"].notna()
    na_vigencia = (m["data"] >= m["vig_ini"]) & (m["data"] <= m["vig_fim"])
    sem_faixa = m["tipo_min"].isna() | m["tipo_max"].isna()
    na_faixa = sem_faixa | ((m["tipo"] >= m["tipo_min"]) & (m["tipo"] <= m["tipo_max"]))
    valido = tem_candidato & na_vigencia & na_faixa

    # 1 preço por linha faturada; se nada validou, mantém a linha sem preço
    bons = m[valido].drop_duplicates(subset="_id", keep="first")
    orfas = m[~m["_id"].isin(bons["_id"])].drop_duplicates(subset="_id", keep="first")
    for col in ["preco_tab", "vig_ini", "vig_fim", "tipo_label", "tipo_min", "tipo_max"]:
        orfas[col] = np.nan

    r = pd.concat([bons, orfas], ignore_index=True).sort_values("_id")
    r["preco_fat"] = r["preco_fat"].round(2)
    r["preco_tab"] = r["preco_tab"].round(2)
    r["diferenca"] = (r["preco_fat"] - r["preco_tab"]).round(2)
    r["dif_pct"] = np.where(
        r["preco_tab"].fillna(0) != 0, r["diferenca"] / r["preco_tab"] * 100, np.nan
    ).round(2)
    r["impacto_rs"] = (r["diferenca"] * r["qtd"]).round(2)

    dentro = (r["diferenca"].abs() <= tol_reais) | (r["dif_pct"].abs() <= tol_pct)
    r["status"] = np.select(
        [r["preco_tab"].isna(), dentro, r["diferenca"] > 0, r["diferenca"] < 0],
        ["SEM PREÇO NA TABELA", "OK", "FATURADO A MAIOR", "FATURADO A MENOR"],
        default="OK",
    )

    nao_auditadas = fat[fat["tabela_alvo"].isna()].copy()
    for col in ["preco_tab", "diferenca", "dif_pct", "impacto_rs"]:
        nao_auditadas[col] = np.nan
    nao_auditadas["status"] = "NÃO AUDITADO"
    return pd.concat([r, nao_auditadas], ignore_index=True)


def identificar_origem_do_preco(res: pd.DataFrame, tab: pd.DataFrame,
                                tol_reais: float) -> pd.Series:
    """Para cada linha, diz a qual coluna da tabela o preço FATURADO corresponde.

    Distingue 'aplicaram a tabela errada' (bate com outra coluna) de
    'preço fora de qualquer tabela' (não bate com nenhuma) — o primeiro caso
    costuma ser erro de cadastro/rota, o segundo é desconto comercial.
    """
    cand = res.reset_index(drop=True).assign(_lin=lambda d: d.index).merge(
        tab, left_on=["variedade", "marca_key", "peso"],
        right_on=["variedade", "marca_key", "peso"], how="left", suffixes=("", "_c"),
    )
    cand = cand[(cand["data"] >= cand["vig_ini_c"]) & (cand["data"] <= cand["vig_fim_c"])]
    sem_faixa = cand["tipo_min_c"].isna() | cand["tipo_max_c"].isna()
    cand = cand[sem_faixa | ((cand["tipo"] >= cand["tipo_min_c"])
                             & (cand["tipo"] <= cand["tipo_max_c"]))]
    bate = cand[(cand["preco_fat"] - cand["preco_tab_c"]).abs() <= tol_reais]
    achado = bate.groupby("_lin")["tabela_of_c"].apply(
        lambda s: " / ".join(sorted(set(s)))
    )
    return achado.reindex(range(len(res))).fillna("— nenhuma tabela vigente —")


# =============================================================================
# SAÍDAS
# =============================================================================

COLS_SAIDA = [
    "data", "cliente", "grupo", "uf", "produto", "variedade", "marca",
    "tipo_txt", "peso", "qtd", "tabela_fat", "tabela_alvo", "tipo_label",
    "preco_fat", "preco_tab", "diferenca", "dif_pct", "impacto_rs", "status",
    "confere_com",
]
ROTULOS = {
    "data": "Emissão", "cliente": "Cliente", "grupo": "Grupo", "uf": "UF",
    "produto": "Produto", "variedade": "Variedade", "marca": "Marca",
    "tipo_txt": "Tipo", "peso": "Peso Cx", "qtd": "Qtd Cx",
    "tabela_fat": "Tabela (faturamento)", "tabela_alvo": "Tabela oficial aplicada",
    "tipo_label": "Faixa da tabela", "preco_fat": "Preço faturado",
    "preco_tab": "Preço tabela", "diferenca": "Diferença R$",
    "dif_pct": "Diferença %", "impacto_rs": "Impacto R$", "status": "Status",
    "confere_com": "Preço faturado confere com",
}


def formatar(df: pd.DataFrame) -> pd.DataFrame:
    out = df.reindex(columns=[c for c in COLS_SAIDA if c in df.columns]).copy()
    if "data" in out:
        out["data"] = pd.to_datetime(out["data"]).dt.strftime("%d/%m/%Y")
    return out.rename(columns=ROTULOS)


def gerar_excel(divergencias, sem_preco, nao_auditado, resumo_cliente,
                resumo_motivo, completo) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xl:
        abas = [
            ("Divergências", formatar(divergencias)),
            ("Sem preço na tabela", formatar(sem_preco)),
            ("Não auditado", formatar(nao_auditado)),
            ("Resumo por cliente", resumo_cliente),
            ("Resumo por motivo", resumo_motivo),
            ("Base completa", formatar(completo)),
        ]
        livro = xl.book
        cab = livro.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "white",
                                "border": 1, "align": "center", "valign": "vcenter"})
        moeda = livro.add_format({"num_format": "#,##0.00"})
        for nome, dados in abas:
            if dados is None or dados.empty:
                dados = pd.DataFrame({"Sem registros": []})
            dados.to_excel(xl, sheet_name=nome[:31], index=False, startrow=1, header=False)
            ws = xl.sheets[nome[:31]]
            for i, col in enumerate(dados.columns):
                ws.write(0, i, str(col), cab)
                maior = dados[col].astype(str).str.len().max()
                maior = 10 if pd.isna(maior) else int(maior)
                largura = min(max(len(str(col)) + 2, maior + 2), 45)
                fmt = moeda if dados[col].dtype.kind in "fc" else None
                ws.set_column(i, i, largura, fmt)
            ws.freeze_panes(1, 0)
            if len(dados):
                ws.autofilter(0, 0, len(dados), len(dados.columns) - 1)
    return buf.getvalue()


def brl(valor) -> str:
    """Formata no padrão pt-BR: 1.234.567,89"""
    if valor is None or pd.isna(valor):
        return "0,00"
    return f"{valor:,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def gerar_email(divergencias: pd.DataFrame, periodo: str, destinatarios: str) -> str:
    n = len(divergencias)
    impacto = divergencias["impacto_rs"].sum(skipna=True)
    a_menor = divergencias[divergencias["status"] == "FATURADO A MENOR"]
    a_maior = divergencias[divergencias["status"] == "FATURADO A MAIOR"]

    linhas = []
    top = divergencias.reindex(
        divergencias["impacto_rs"].abs().sort_values(ascending=False).index
    ).head(10)
    for _, r in top.iterrows():
        linhas.append(
            f"• {r['cliente']} ({r['uf']}) | {r['variedade']} {r['marca']} "
            f"Tipo {r['tipo_txt']} cx {r['peso']:g}kg | "
            f"Faturado R$ {brl(r['preco_fat'])} x Tabela R$ {brl(r['preco_tab'])} "
            f"({r['tabela_alvo']}) | Dif. R$ {brl(r['diferenca'])} "
            f"| Impacto R$ {brl(r['impacto_rs'])}"
        )
    corpo_itens = "\n".join(linhas)
    if n > len(top):
        corpo_itens += f"\n\n... e mais {n - len(top)} lançamento(s). O detalhamento está em anexo."

    return f"""Prezados,

Pedimos a sua verificação para os preços com as diferenças apresentadas no quadro abaixo. Solicitamos os seus comentários.

Período auditado: {periodo}
Lançamentos com divergência: {n}
Faturados a MENOR que a tabela: {len(a_menor)} (impacto R$ {brl(a_menor['impacto_rs'].sum(skipna=True))})
Faturados a MAIOR que a tabela: {len(a_maior)} (impacto R$ {brl(a_maior['impacto_rs'].sum(skipna=True))})
Impacto líquido: R$ {brl(impacto)}

Principais itens (por impacto):
{corpo_itens}

Favor informar se há desconto ou condição comercial aprovada para estes casos.

Atenciosamente,
Auditoria de Preços
Para: {destinatarios}
"""


# =============================================================================
# INTERFACE
# =============================================================================

st.set_page_config(page_title="Auditoria de Preços", layout="wide", page_icon="🤖")
st.title("🤖 Robô de Auditoria de Preços")
st.caption("Cruza o faturamento com a Tabela de Preços oficial e aponta as divergências.")

with st.sidebar:
    st.header("⚙️ Parâmetros")
    tol_reais = st.number_input("Tolerância em R$ por caixa", 0.0, 50.0, 0.01, 0.01,
                                help="Diferenças até este valor são tratadas como arredondamento.")
    tol_pct = st.number_input("Tolerância em %", 0.0, 20.0, 0.0, 0.1)
    st.divider()
    fora_escopo = st.multiselect(
        "Tabelas fora do escopo",
        options=TABELAS_FORA_ESCOPO_PADRAO + ["Clientes Diversos"],
        default=TABELAS_FORA_ESCOPO_PADRAO,
        help="Estabelecimentos ITR é filial de faturamento — transferência interna, não venda.",
    )
    familia_pimentao = st.multiselect(
        "Variedades da família Pimentão", options=FAMILIA_PIMENTAO, default=FAMILIA_PIMENTAO,
        help="Esta família usa a nomenclatura FOB / CIF SP na tabela oficial.",
    )
    st.divider()
    destinatarios = st.text_input("Destinatários do e-mail",
                                  "com.inteligencia.de.mercado@itaueira.com; dir.comercial@itaueira.com")

col1, col2 = st.columns(2)
with col1:
    arq_fat = st.file_uploader("📥 CSV do Faturamento", type=["csv"])
with col2:
    arq_tab = st.file_uploader("📥 CSV da Tabela de Preços", type=["csv"])

if not (arq_fat and arq_tab):
    st.info("Carregue os dois arquivos para iniciar a auditoria.")
    st.stop()

try:
    fat = preparar_faturamento(arq_fat.getvalue())
    tab = preparar_tabela(arq_tab.getvalue())
except Exception as erro:  # noqa: BLE001
    st.error(f"Não consegui preparar os arquivos.\n\n**{type(erro).__name__}:** {erro}")
    st.stop()

# --- De/Para editável -------------------------------------------------------
with st.expander("🔗 Regra de precificação (De/Para) — clique para ajustar", expanded=False):
    st.markdown(
        "**Como o robô escolhe a linha da tabela:** primeiro procura no De/Para direto; "
        "se não achar, aplica a regra por região da UF do cliente."
    )
    opcoes_tabela = sorted(tab["tabela_of"].dropna().unique().tolist())
    tabelas_fat = sorted(fat["tabela_fat"].dropna().unique().tolist())

    base_depara = pd.DataFrame(
        {"Tabela no faturamento": tabelas_fat,
         "Tabela oficial": [DEPARA_DIRETO.get(t, "— usar regra por região —") for t in tabelas_fat]}
    )
    ed_depara = st.data_editor(
        base_depara, hide_index=True, use_container_width=True, key="ed_depara",
        column_config={"Tabela no faturamento": st.column_config.TextColumn(disabled=True),
                       "Tabela oficial": st.column_config.SelectboxColumn(
                           options=["— usar regra por região —"] + opcoes_tabela, required=True)},
    )
    depara = {r["Tabela no faturamento"]: r["Tabela oficial"] for _, r in ed_depara.iterrows()
              if r["Tabela oficial"] != "— usar regra por região —"}

    st.markdown("**Regra por região** (aplicada quando o De/Para acima está em *regra por região*)")
    base_regiao = pd.DataFrame(
        [{"Família": fam, "Região": reg, "Tabela oficial": destino}
         for fam, mapa in REGRA_REGIAO_PADRAO.items() for reg, destino in mapa.items()]
    )
    ed_regiao = st.data_editor(
        base_regiao, hide_index=True, use_container_width=True, key="ed_regiao",
        column_config={"Família": st.column_config.TextColumn(disabled=True),
                       "Região": st.column_config.TextColumn(disabled=True),
                       "Tabela oficial": st.column_config.SelectboxColumn(
                           options=opcoes_tabela, required=True)},
    )
    regra_regiao = {"MELAO": {}, "PIMENTAO": {}}
    for _, r in ed_regiao.iterrows():
        regra_regiao[r["Família"]][r["Região"]] = r["Tabela oficial"]

    st.download_button(
        "💾 Salvar esta configuração (JSON)",
        json.dumps({"depara": depara, "regra_regiao": regra_regiao,
                    "fora_escopo": fora_escopo, "familia_pimentao": familia_pimentao},
                   ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="config_auditoria_precos.json", mime="application/json",
    )

# --- Execução ---------------------------------------------------------------
with st.spinner("Cruzando faturamento x tabela oficial..."):
    fat_regra = montar_regra(fat, depara, regra_regiao, familia_pimentao, fora_escopo)
    res = auditar(fat_regra, tab, tol_reais, tol_pct).reset_index(drop=True)
    res["confere_com"] = identificar_origem_do_preco(res, tab, max(tol_reais, 0.01))

divergencias = res[res["status"].isin(["FATURADO A MAIOR", "FATURADO A MENOR"])].copy()
sem_preco = res[res["status"] == "SEM PREÇO NA TABELA"].copy()
nao_auditado = res[res["status"] == "NÃO AUDITADO"].copy()
conferidas = res[res["status"] != "NÃO AUDITADO"]

datas = fat["data"].dropna()
periodo = (f"{datas.min():%d/%m/%Y} a {datas.max():%d/%m/%Y}" if len(datas) else "n/d")

st.divider()
st.subheader(f"Resultado — período {periodo}")
k = st.columns(6)
k[0].metric("Linhas no faturamento", f"{len(res):,}".replace(",", "."))
k[1].metric("Conferidas", f"{len(conferidas):,}".replace(",", "."))
k[2].metric("✅ Conformes", f"{(res['status'] == 'OK').sum():,}".replace(",", "."))
k[3].metric("⚠️ Divergentes", f"{len(divergencias):,}".replace(",", "."),
            delta=f"{len(divergencias) / max(len(conferidas), 1) * 100:.1f}% do conferido",
            delta_color="inverse")
k[4].metric("❓ Sem preço na tabela", f"{len(sem_preco):,}".replace(",", "."))
k[5].metric("💰 Impacto líquido",
            f"R$ {brl(divergencias['impacto_rs'].sum(skipna=True))}")

if len(divergencias) == 0 and len(sem_preco) == 0:
    st.success("✅ Nenhuma divergência encontrada no período.")

aba1, aba2, aba3, aba4, aba5 = st.tabs(
    ["⚠️ Divergências", "📊 Resumos", "❓ Sem preço", "🚫 Não auditado", "🛠️ Diagnóstico"]
)

with aba1:
    if divergencias.empty:
        st.success("Nenhuma divergência acima da tolerância.")
    else:
        filtro = st.multiselect("Filtrar status", ["FATURADO A MENOR", "FATURADO A MAIOR"],
                                default=["FATURADO A MENOR", "FATURADO A MAIOR"])
        vis = divergencias[divergencias["status"].isin(filtro)]
        vis = vis.reindex(vis["impacto_rs"].abs().sort_values(ascending=False).index)
        st.dataframe(formatar(vis), use_container_width=True, hide_index=True, height=430)

resumo_cliente = pd.DataFrame()
resumo_motivo = pd.DataFrame()
with aba2:
    if divergencias.empty:
        st.info("Sem divergências para resumir.")
    else:
        resumo_cliente = (
            divergencias.groupby(["cliente", "uf"], as_index=False)
            .agg(Lançamentos=("status", "size"), Caixas=("qtd", "sum"),
                 Impacto_RS=("impacto_rs", "sum"))
            .sort_values("Impacto_RS", key=abs, ascending=False)
            .rename(columns={"cliente": "Cliente", "uf": "UF", "Impacto_RS": "Impacto R$"})
        )
        resumo_motivo = (
            divergencias.groupby(["tabela_fat", "tabela_alvo", "status"], as_index=False)
            .agg(Lançamentos=("status", "size"), Impacto_RS=("impacto_rs", "sum"))
            .sort_values("Impacto_RS", key=abs, ascending=False)
            .rename(columns={"tabela_fat": "Tabela (faturamento)",
                             "tabela_alvo": "Tabela oficial aplicada",
                             "status": "Status", "Impacto_RS": "Impacto R$"})
        )
        e1, e2 = st.columns(2)
        with e1:
            st.markdown("**Por cliente**")
            st.dataframe(resumo_cliente, use_container_width=True, hide_index=True, height=320)
        with e2:
            st.markdown("**Por tabela aplicada**")
            st.dataframe(resumo_motivo, use_container_width=True, hide_index=True, height=320)
        st.markdown("**Por variedade e marca**")
        st.dataframe(
            divergencias.groupby(["variedade", "marca"], as_index=False)
            .agg(Lançamentos=("status", "size"), Impacto_RS=("impacto_rs", "sum"))
            .sort_values("Impacto_RS", key=abs, ascending=False)
            .rename(columns={"variedade": "Variedade", "marca": "Marca",
                             "Impacto_RS": "Impacto R$"}),
            use_container_width=True, hide_index=True,
        )

with aba3:
    st.caption("Combinações faturadas que não existem na tabela oficial vigente. "
               "Não são divergência de preço — são lacuna de cadastro.")
    if sem_preco.empty:
        st.success("Todas as combinações faturadas têm preço na tabela.")
    else:
        st.dataframe(
            sem_preco.groupby(["variedade", "marca", "peso", "tipo_txt", "tabela_alvo"],
                              as_index=False).agg(Lançamentos=("status", "size"),
                                                  Caixas=("qtd", "sum")),
            use_container_width=True, hide_index=True,
        )

with aba4:
    st.caption("Linhas deliberadamente fora da auditoria.")
    if nao_auditado.empty:
        st.info("Nenhuma linha fora do escopo.")
    else:
        st.dataframe(
            nao_auditado.groupby(["tabela_fat", "origem_regra"], as_index=False)
            .agg(Lançamentos=("status", "size"), Caixas=("qtd", "sum"))
            .rename(columns={"tabela_fat": "Tabela (faturamento)", "origem_regra": "Motivo"}),
            use_container_width=True, hide_index=True,
        )

with aba5:
    st.markdown("**Cobertura do cruzamento**")
    st.dataframe(res["status"].value_counts().rename_axis("Status").reset_index(name="Linhas"),
                 use_container_width=True, hide_index=True)
    d1, d2 = st.columns(2)
    with d1:
        st.markdown("**Vigências na tabela oficial**")
        st.dataframe(
            tab.groupby(["vig_ini", "vig_fim"], as_index=False).size()
            .assign(vig_ini=lambda d: d["vig_ini"].dt.strftime("%d/%m/%Y"),
                    vig_fim=lambda d: d["vig_fim"].dt.strftime("%d/%m/%Y"))
            .rename(columns={"vig_ini": "Início", "vig_fim": "Fim", "size": "Linhas"}),
            use_container_width=True, hide_index=True,
        )
        sem_vig = fat[~fat["data"].apply(
            lambda d: ((tab["vig_ini"] <= d) & (tab["vig_fim"] >= d)).any())]
        if len(sem_vig):
            st.warning(f"{len(sem_vig)} linha(s) faturadas fora de qualquer vigência: "
                       + ", ".join(sorted(sem_vig['data'].dt.strftime('%d/%m/%Y').unique())))
    with d2:
        st.markdown("**Valores que não viraram número**")
        problemas = pd.DataFrame({
            "Campo": ["Preço faturado", "Peso caixa", "Qtd caixas", "Data emissão",
                      "Preço tabela", "Vigência início"],
            "Nulos": [fat["preco_fat"].isna().sum(), fat["peso"].isna().sum(),
                      fat["qtd"].isna().sum(), fat["data"].isna().sum(),
                      tab["preco_tab"].isna().sum(), tab["vig_ini"].isna().sum()],
        })
        st.dataframe(problemas, use_container_width=True, hide_index=True)
        st.caption("Tipos não numéricos no faturamento (ex.: 'Vários', '250g') são "
                   "cruzados por peso da caixa, pois a tabela não define faixa para eles.")

# --- Entregáveis ------------------------------------------------------------
st.divider()
st.subheader("📤 Entregáveis")

texto_email = gerar_email(divergencias, periodo, destinatarios) if not divergencias.empty \
    else "Nenhuma divergência encontrada no período."
st.text_area("✉️ E-mail pronto para envio", texto_email, height=300)

b1, b2 = st.columns(2)
with b1:
    st.download_button(
        "📊 Baixar Excel completo",
        gerar_excel(divergencias, sem_preco, nao_auditado, resumo_cliente, resumo_motivo, res),
        file_name=f"auditoria_precos_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )
with b2:
    st.download_button("📝 Baixar texto do e-mail", texto_email.encode("utf-8"),
                       file_name=f"email_auditoria_{date.today():%Y%m%d}.txt",
                       mime="text/plain", use_container_width=True)
