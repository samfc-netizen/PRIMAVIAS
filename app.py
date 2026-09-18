# app.py
import re
from io import BytesIO
from datetime import date

import pandas as pd
import streamlit as st


# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
st.set_page_config(
    page_title="Títulos em Aberto - Primavias",
    page_icon="📄",
    layout="wide"
)


# =========================
# MAPA CLIENTE -> UNIDADE
# =========================
MAPA_UNIDADES = {
    "80005051-PRIMAVIA COMERCIO DE AUTOMOVEIS LTDA": "RENAULT UNAÍ 17730943000172",
    "80005052-PRIMAVIA COMERCIO DE VEICULOS LTDA": "NISSAN UNAÍ 17168524000199",
    "00003335-PRIMAVIA VEICULOS LTDA": "PARACATU 71145668000256",
    "00003362-PRIMAVIA VEICULOS LTDA": "FIAT UNAÍ 71145668000175",
}


# =========================
# FUNÇÕES AUXILIARES
# =========================
def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def parse_moeda(valor):
    """
    Converte valores como:
    '1.234,56' -> 1234.56
    1234.56 -> 1234.56
    """
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    texto = texto.replace("R$", "").replace(" ", "")

    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    try:
        return float(texto)
    except Exception:
        return 0.0


def formatar_moeda_br(valor):
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def extrair_rps(duplicata):
    """
    Exemplo:
    000000001026A-NS -> RPS = 1026
    Regra:
    - pega a parte antes do hífen
    - remove tudo que não for número
    - converte para inteiro para tirar zeros à esquerda
    """
    if pd.isna(duplicata):
        return None

    texto = str(duplicata).strip()
    parte_antes_hifen = texto.split("-")[0]
    numeros = re.sub(r"\D", "", parte_antes_hifen)

    if not numeros:
        return None

    try:
        return int(numeros)
    except Exception:
        return None


def calcular_ns(rps):
    if pd.isna(rps) or rps is None:
        return None
    try:
        return int(rps) - 2
    except Exception:
        return None


def encontrar_coluna(df, nome_base):
    """
    Procura coluna de forma robusta.
    Ex.: VENCTO ou VENCTO.1
    """
    cols = [c for c in df.columns if str(c).strip().upper().startswith(nome_base.upper())]
    return cols[0] if cols else None


@st.cache_data(show_spinner=False)
def processar_planilha(arquivo):
    nome_arquivo = getattr(arquivo, "name", "").lower()

    if nome_arquivo.endswith(".csv"):
        # CSV exportado pelo sistema possui linhas de cabeçalho/relatório
        # antes da tabela. Localizamos automaticamente a linha que começa
        # com CLIENTE e usamos essa linha como cabeçalho.
        bruto = arquivo.getvalue()

        texto = None
        encoding_usado = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                texto = bruto.decode(encoding)
                encoding_usado = encoding
                break
            except UnicodeDecodeError:
                continue

        if texto is None:
            raise ValueError("Não foi possível identificar a codificação do CSV.")

        linhas = texto.splitlines()
        linha_cabecalho = None

        for i, linha in enumerate(linhas):
            if linha.strip().upper().startswith("CLIENTE;"):
                linha_cabecalho = i
                break

        if linha_cabecalho is None:
            raise ValueError(
                "Cabeçalho do CSV não encontrado. Era esperada uma linha iniciando por CLIENTE."
            )

        arquivo.seek(0)
        df = pd.read_csv(
            arquivo,
            sep=";",
            encoding=encoding_usado,
            skiprows=linha_cabecalho,
            dtype=str,
            engine="python",
        )

        # O relatório CSV pode trazer colunas vazias/repetidas entre CLIENTE
        # e VENCTO. O pandas renomeia essas colunas automaticamente.
        aba = "CSV"

    else:
        xls = pd.ExcelFile(arquivo)
        aba = xls.sheet_names[0]
        df = pd.read_excel(arquivo, sheet_name=aba)

    col_vencto = encontrar_coluna(df, "VENCTO")
    col_cliente = encontrar_coluna(df, "CLIENTE")
    col_duplicata = encontrar_coluna(df, "DUPLICATA")
    col_dta_cad = encontrar_coluna(df, "DTA.CAD")
    col_valor = encontrar_coluna(df, "V.ORIGI")

    colunas_obrigatorias = {
        "VENCTO": col_vencto,
        "CLIENTE": col_cliente,
        "DUPLICATA": col_duplicata,
        "DTA.CAD": col_dta_cad,
        "V.ORIGI": col_valor,
    }

    faltantes = [nome for nome, col in colunas_obrigatorias.items() if col is None]
    if faltantes:
        raise ValueError(f"Colunas obrigatórias não encontradas: {', '.join(faltantes)}")

    base = df[[col_dta_cad, col_vencto, col_cliente, col_duplicata, col_valor]].copy()

    base.columns = [
        "Data cad",
        "Data vencimento",
        "Cliente",
        "Duplicata",
        "Valor",
    ]

    base["Cliente"] = base["Cliente"].apply(normalizar_texto)
    base["Duplicata"] = base["Duplicata"].apply(normalizar_texto)
    base["Valor"] = base["Valor"].apply(parse_moeda)

    base["Data cad"] = pd.to_datetime(base["Data cad"], errors="coerce")
    base["Data vencimento"] = pd.to_datetime(base["Data vencimento"], errors="coerce")

    base = base[
        (base["Cliente"].astype(str).str.strip() != "") &
        (base["Duplicata"].astype(str).str.strip() != "")
    ].copy()

    base["Unidade"] = base["Cliente"].map(MAPA_UNIDADES).fillna("Não mapeada")
    base["RPS"] = base["Duplicata"].apply(extrair_rps)
    base["NS"] = base["RPS"].apply(calcular_ns)

    base = base.sort_values(
        ["Data cad", "Data vencimento", "Cliente", "Duplicata"],
        ascending=True
    ).reset_index(drop=True)

    return base, aba, df


def para_excel(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Titulos_em_Aberto")
    output.seek(0)
    return output.getvalue()


# =========================
# TÍTULO
# =========================
st.title("📄 Títulos em Aberto - Primavias")
st.caption("Leitura da planilha com correlação de cliente, unidade, duplicata, RPS, NS, valor e filtros por unidade e datas.")


# =========================
# UPLOAD
# =========================
arquivo = st.file_uploader(
    "Insira a planilha Excel ou CSV",
    type=["xlsx", "xls", "csv"],
    help="Envie a planilha Excel ou o CSV exportado pelo sistema para leitura automática."
)

if not arquivo:
    st.info("Envie uma planilha para iniciar a análise.")
    st.stop()


# =========================
# PROCESSAMENTO
# =========================
try:
    base, nome_aba, df_original = processar_planilha(arquivo)
except Exception as e:
    st.error(f"Erro ao processar a planilha: {e}")
    st.stop()


# =========================
# SIDEBAR - FILTROS
# =========================
st.sidebar.header("Filtros")

unidades_disponiveis = sorted(base["Unidade"].dropna().unique().tolist())

unidade_selecionada = st.sidebar.multiselect(
    "Unidade",
    options=unidades_disponiveis,
    default=unidades_disponiveis
)

data_min = base["Data vencimento"].min()
data_max = base["Data vencimento"].max()

if pd.isna(data_min) or pd.isna(data_max):
    data_inicial = date.today()
    data_final = date.today()
else:
    data_inicial = data_min.date()
    data_final = data_max.date()

periodo = st.sidebar.date_input(
    "Período de vencimento",
    value=(data_inicial, data_final),
    min_value=data_inicial,
    max_value=data_final
)

filtrado = base.copy()

if unidade_selecionada:
    filtrado = filtrado[filtrado["Unidade"].isin(unidade_selecionada)].copy()

if isinstance(periodo, tuple) and len(periodo) == 2:
    dt_ini, dt_fim = periodo
    filtrado = filtrado[
        (filtrado["Data vencimento"].dt.date >= dt_ini) &
        (filtrado["Data vencimento"].dt.date <= dt_fim)
    ].copy()


# =========================
# CARDS DE RESUMO
# =========================
col1, col2, col3, col4 = st.columns(4)

col1.metric("Qtd. títulos", f"{len(filtrado):,}".replace(",", "."))
col2.metric("Valor total", formatar_moeda_br(filtrado["Valor"].sum()))
col3.metric("Qtd. clientes", f"{filtrado['Cliente'].nunique():,}".replace(",", "."))
col4.metric("Qtd. unidades", f"{filtrado['Unidade'].nunique():,}".replace(",", "."))


# =========================
# INFORMAÇÕES DA LEITURA
# =========================
with st.expander("Informações da leitura", expanded=False):
    c1, c2, c3 = st.columns(3)
    c1.write(f"**Aba lida:** {nome_aba}")
    c2.write(f"**Linhas originais:** {len(df_original):,}".replace(",", "."))
    c3.write(f"**Linhas válidas após tratamento:** {len(base):,}".replace(",", "."))


# =========================
# TABELA PRINCIPAL
# =========================
st.subheader("Lista de títulos em aberto")

tabela = filtrado[[
    "Data cad",
    "Data vencimento",
    "Cliente",
    "Unidade",
    "Duplicata",
    "RPS",
    "NS",
    "Valor"
]].copy()

tabela_exibicao = tabela.copy()
tabela_exibicao["Data cad"] = tabela_exibicao["Data cad"].dt.strftime("%d/%m/%Y")
tabela_exibicao["Data vencimento"] = tabela_exibicao["Data vencimento"].dt.strftime("%d/%m/%Y")
tabela_exibicao["Valor"] = tabela_exibicao["Valor"].apply(formatar_moeda_br)

st.dataframe(
    tabela_exibicao,
    use_container_width=True,
    hide_index=True
)


# =========================
# TABELA RESUMIDA POR NS
# =========================
st.subheader("Resumo por NS e Unidade")

resumo_ns = (
    filtrado.groupby(["NS", "Unidade"], dropna=False)
    .agg(
        Valor=("Valor", "sum")
    )
    .reset_index()
    .sort_values(by=["NS", "Unidade"], ascending=True)
)

resumo_ns_exib = resumo_ns.copy()
resumo_ns_exib["Valor"] = resumo_ns_exib["Valor"].apply(formatar_moeda_br)

st.dataframe(
    resumo_ns_exib,
    use_container_width=True,
    hide_index=True
)


# =========================
# RESUMO POR CLIENTE
# =========================
st.subheader("Resumo por cliente")

resumo_cliente = (
    filtrado.groupby(["Cliente", "Unidade"], dropna=False)
    .agg(
        Qtd_Titulos=("Duplicata", "count"),
        Valor_Total=("Valor", "sum"),
        Menor_Vencimento=("Data vencimento", "min"),
        Maior_Vencimento=("Data vencimento", "max"),
    )
    .reset_index()
)

resumo_cliente_exib = resumo_cliente.copy()
resumo_cliente_exib["Menor_Vencimento"] = resumo_cliente_exib["Menor_Vencimento"].dt.strftime("%d/%m/%Y")
resumo_cliente_exib["Maior_Vencimento"] = resumo_cliente_exib["Maior_Vencimento"].dt.strftime("%d/%m/%Y")
resumo_cliente_exib["Valor_Total"] = resumo_cliente_exib["Valor_Total"].apply(formatar_moeda_br)

st.dataframe(
    resumo_cliente_exib,
    use_container_width=True,
    hide_index=True
)


# =========================
# EXPORTAÇÃO
# =========================
st.subheader("Exportar")

col_exp1, col_exp2 = st.columns(2)

export_df = tabela.copy()
export_df["Data cad"] = export_df["Data cad"].dt.strftime("%d/%m/%Y")
export_df["Data vencimento"] = export_df["Data vencimento"].dt.strftime("%d/%m/%Y")

with col_exp1:
    st.download_button(
        label="📥 Baixar Excel",
        data=para_excel(export_df),
        file_name="titulos_em_aberto_primavias.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_exp2:
    st.download_button(
        label="📥 Baixar CSV",
        data=export_df.to_csv(index=False, sep=";", encoding="utf-8-sig"),
        file_name="titulos_em_aberto_primavias.csv",
        mime="text/csv"
    )