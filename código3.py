import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Comparador de Bases", page_icon="🔍", layout="wide")

st.title("🔍 Comparador de Bases — Serviços Divergentes")
st.markdown("Carregue as bases **Cbill** e **Oper** para identificar serviços divergentes por data limite.")

# ─── mapeamento fixo de colunas por sistema ──────────────────────────────────
COLUNAS_CBILL = {
    "servico": "Serviço",
    "data":    "Prazo de execução",
    "tipo":    "Tipo Serviço",
}

COLUNAS_OPER = {
    "servico": "Numero",
    "data":    "Data/Hora Limite",
    "tipo":    "Subtipo",
}

# Aliases aceitos para cada coluna canônica
ALIASES_COLUNAS = {
    "Data/Hora Limite":  ["Data/Hora Limite", "Data Limite", "Data limite", "DataHora Limite"],
    "Numero":            ["Numero", "Número", "numero", "número"],
    "Subtipo":           ["Subtipo", "subtipo"],
    "Tipo Serviço":      ["Tipo Serviço", "Tipo Servico", "Tipo de Serviço", "Tipo de Servico"],
    "Prazo de execução": ["Prazo de execução", "Prazo de Execução", "Prazo execução"],
    "Serviço":           ["Serviço", "Servico", "serviço"],
}

TIPO_EXCLUIDO_OPER = "RESTABELECIMENTO FORNEC. NORMAL"

# ─── helpers ─────────────────────────────────────────────────────────────────

def ler_excel(arquivo) -> pd.DataFrame:
    nome = arquivo.name.lower()
    conteudo = arquivo.read()
    arquivo.seek(0)

    amostra = conteudo[:10]
    if amostra.startswith(b"<") or amostra.startswith(b"\xef\xbb\xbf<"):
        for header_row in range(0, 8):
            for flavor in ("lxml", "html5lib"):
                try:
                    todas = pd.read_html(BytesIO(conteudo), header=header_row, flavor=flavor)
                except Exception:
                    continue
                for df in todas:
                    df.columns = [str(c).strip() for c in df.columns]
                    colunas_validas = [
                        c for c in df.columns
                        if c and not c.startswith("Unnamed") and c.lower() != "nan"
                    ]
                    if len(colunas_validas) >= 3:
                        df = df.dropna(how="all").reset_index(drop=True)
                        return df
        return pd.DataFrame()

    if nome.endswith(".xls"):
        return pd.read_excel(BytesIO(conteudo), engine="xlrd")
    else:
        return pd.read_excel(BytesIO(conteudo), engine="openpyxl")


def resolver_coluna(df: pd.DataFrame, col_canonica: str) -> str | None:
    aliases = ALIASES_COLUNAS.get(col_canonica, [col_canonica])
    for alias in aliases:
        if alias in df.columns:
            return alias
    col_lower = col_canonica.lower()
    for c in df.columns:
        if c.lower() == col_lower:
            return c
    return None


def resolver_colunas(df: pd.DataFrame, mapa: dict) -> dict:
    return {
        chave: resolver_coluna(df, col_canonica) or col_canonica
        for chave, col_canonica in mapa.items()
    }


def validar_colunas(df: pd.DataFrame, cols: dict, nome_arquivo: str) -> bool:
    ausentes = [v for v in cols.values() if v not in df.columns]
    if ausentes:
        st.error(
            f"❌ **{nome_arquivo}** — colunas não encontradas: `{'`, `'.join(ausentes)}`\n\n"
            f"Colunas disponíveis: `{'`, `'.join(df.columns.tolist())}`"
        )
        return False
    return True


def normalizar_datas(df: pd.DataFrame, col: str) -> pd.DataFrame:
    df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
    return df


def exportar_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Divergentes")
        ws = writer.sheets["Divergentes"]
        for col_cells in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = max(max_len + 4, 14)
    return buf.getvalue()

# ─── sidebar ─────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Configuração")
    data_filtro = st.date_input("Data limite a comparar", value=datetime.today())
    st.markdown("---")
    st.markdown("**Colunas utilizadas:**")
    st.markdown("🔵 **Cbill:** `Serviço` · `Prazo de execução` · `Tipo Serviço`")
    st.markdown("🟠 **Oper:** `Numero` · `Data/Hora Limite` · `Subtipo`")
    st.markdown("---")
    st.markdown("**Filtro automático Oper:**")
    st.markdown(f"🚫 Excluídos: `{TIPO_EXCLUIDO_OPER}`")
    st.markdown("---")
    st.info(
        "💡 **Padrão de nomes esperado:**\n"
        "`base_DD.MM_Cbill.xlsx`\n"
        "`base_DD.MM_oper.xls` _(Comercial)_\n"
        "`base_gd_DD.MM_oper.xls` _(GD)_"
    )

# ─── upload ───────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Base Cbill")
    arquivo_cbill = st.file_uploader("Selecione a base Cbill", type=["xlsx", "xls"], key="cbill")

with col2:
    st.subheader("📂 Base Oper")
    arquivo_oper_com = st.file_uploader("Base Oper — Comercial", type=["xlsx", "xls"], key="oper_com")
    arquivo_oper_gd  = st.file_uploader("Base Oper — GD (opcional)", type=["xlsx", "xls"], key="oper_gd")

# ─── processamento ────────────────────────────────────────────────────────────

if arquivo_cbill and arquivo_oper_com:
    try:
        with st.spinner("Carregando bases..."):
            df_cbill    = ler_excel(arquivo_cbill)
            df_oper_com = ler_excel(arquivo_oper_com)
            df_oper_gd  = ler_excel(arquivo_oper_gd) if arquivo_oper_gd else pd.DataFrame()

        # ── resolve colunas ──
        cols_cbill    = resolver_colunas(df_cbill,    COLUNAS_CBILL)
        cols_oper_com = resolver_colunas(df_oper_com, COLUNAS_OPER)

        if not validar_colunas(df_cbill, cols_cbill, arquivo_cbill.name):
            st.stop()
        if not validar_colunas(df_oper_com, cols_oper_com, arquivo_oper_com.name):
            st.stop()

        # ── normaliza datas ──
        df_cbill    = normalizar_datas(df_cbill,    cols_cbill["data"])
        df_oper_com = normalizar_datas(df_oper_com, cols_oper_com["data"])

        # ── processa e une GD se enviada ──
        if not df_oper_gd.empty:
            cols_oper_gd = resolver_colunas(df_oper_gd, COLUNAS_OPER)
            if not validar_colunas(df_oper_gd, cols_oper_gd, arquivo_oper_gd.name):
                st.stop()
            df_oper_gd = normalizar_datas(df_oper_gd, cols_oper_gd["data"])

            # Padroniza nomes de colunas da GD para bater com Comercial antes de unir
            df_oper_gd = df_oper_gd.rename(columns={
                cols_oper_gd["servico"]: cols_oper_com["servico"],
                cols_oper_gd["data"]:    cols_oper_com["data"],
                cols_oper_gd["tipo"]:    cols_oper_com["tipo"],
            })

            df_oper_full = pd.concat([df_oper_com, df_oper_gd], ignore_index=True)
        else:
            df_oper_full = df_oper_com.copy()

        # ── filtra pela data ──
        data_alvo  = pd.Timestamp(data_filtro).date()
        base_cbill = df_cbill[df_cbill[cols_cbill["data"]].dt.date == data_alvo].copy()
        base_oper  = df_oper_full[df_oper_full[cols_oper_com["data"]].dt.date == data_alvo].copy()

        # ── remove RESTABELECIMENTO apenas do Oper ──
        col_tipo_oper = cols_oper_com["tipo"]
        base_oper = base_oper[
            base_oper[col_tipo_oper].astype(str).str.strip().str.upper() != TIPO_EXCLUIDO_OPER.upper()
        ].copy()

        total_cbill = len(base_cbill)
        total_oper  = len(base_oper)

        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        m1.metric("Serviços Cbill", total_cbill)
        m2.metric("Serviços Oper (sem Restab.)", total_oper)
        m3.metric("Diferença", abs(total_cbill - total_oper))

        # ── identifica divergências ──
        col_srv_cbill  = cols_cbill["servico"]
        col_srv_oper   = cols_oper_com["servico"]
        col_tipo_cbill = cols_cbill["tipo"]

        srvs_cbill = set(base_cbill[col_srv_cbill].astype(str).str.strip())
        srvs_oper  = set(base_oper[col_srv_oper].astype(str).str.strip())

        apenas_cbill = srvs_cbill - srvs_oper
        apenas_oper  = srvs_oper  - srvs_cbill

        registros = []

        for srv in sorted(apenas_cbill):
            linhas = base_cbill[base_cbill[col_srv_cbill].astype(str).str.strip() == srv]
            for _, row in linhas.iterrows():
                registros.append({
                    "servico":        row[col_srv_cbill],
                    "tipo_servico":   str(row.get(col_tipo_cbill, "")).strip(),
                    "data_limite":    row[cols_cbill["data"]].date() if pd.notna(row[cols_cbill["data"]]) else data_filtro,
                    "ausente_em":     "Oper",
                    "sistema_origem": "Cbill",
                })

        for srv in sorted(apenas_oper):
            linhas = base_oper[base_oper[col_srv_oper].astype(str).str.strip() == srv]
            for _, row in linhas.iterrows():
                registros.append({
                    "servico":        row[col_srv_oper],
                    "tipo_servico":   str(row.get(col_tipo_oper, "")).strip(),
                    "data_limite":    row[cols_oper_com["data"]].date() if pd.notna(row[cols_oper_com["data"]]) else data_filtro,
                    "ausente_em":     "Cbill",
                    "sistema_origem": "Oper",
                })

        df_resultado = pd.DataFrame(registros)

        st.markdown("---")

        if df_resultado.empty:
            st.success("✅ Nenhuma divergência encontrada! As bases estão alinhadas para esta data.")
        else:
            st.warning(f"⚠️ **{len(df_resultado)} serviço(s) divergente(s)** em {data_filtro.strftime('%d/%m/%Y')}")

            tab1, tab2, tab3 = st.tabs([
                f"📋 Todos ({len(df_resultado)})",
                f"🔵 Só no Cbill ({len(apenas_cbill)})",
                f"🟠 Só no Oper ({len(apenas_oper)})",
            ])

            with tab1:
                st.dataframe(df_resultado, use_container_width=True)

            with tab2:
                sub = df_resultado[df_resultado["sistema_origem"] == "Cbill"]
                if not sub.empty:
                    st.dataframe(sub, use_container_width=True)
                else:
                    st.info("Sem exclusivos.")

            with tab3:
                sub = df_resultado[df_resultado["sistema_origem"] == "Oper"]
                if not sub.empty:
                    st.dataframe(sub, use_container_width=True)
                else:
                    st.info("Sem exclusivos.")

            st.markdown("---")
            nome_saida = f"divergentes_{data_filtro.strftime('%d.%m')}_Cbill_vs_Oper.xlsx"
            st.download_button(
                label="⬇️ Baixar planilha de divergentes (.xlsx)",
                data=exportar_excel(df_resultado),
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with st.expander(f"👁️ Ver base Cbill filtrada ({total_cbill} registros)"):
            st.dataframe(base_cbill, use_container_width=True)
        with st.expander(f"👁️ Ver base Oper filtrada ({total_oper} registros, sem Restabelecimento)"):
            st.dataframe(base_oper, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")
        st.exception(e)

else:
    st.info("👆 Faça upload da base Cbill e ao menos da base Oper Comercial para iniciar a comparação.")
