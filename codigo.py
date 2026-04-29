import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Comparador de Bases", page_icon="🔍", layout="wide")

st.title("🔍 Comparador de Bases — Serviços Divergentes")
st.markdown("Carregue as bases **Cbill** e **Oper** para identificar serviços divergentes por data limite.")

# ─── mapeamento fixo de colunas por sistema ──────────────────────────────────
COLUNAS_POR_SISTEMA = {
    "cbill": ("Serviço",  "Prazo de execução"),
    "oper":  ("Numero",   "Data/Hora Limite"),
}

# ─── helpers ─────────────────────────────────────────────────────────────────

def extrair_sistema(nome_arquivo: str) -> str:
    nome = nome_arquivo.replace(".xlsx", "").replace(".xls", "")
    partes = nome.split("_")
    return partes[-1] if len(partes) >= 3 else nome


def ler_excel(arquivo) -> pd.DataFrame:
    nome = arquivo.name.lower()
    conteudo = arquivo.read()
    arquivo.seek(0)

    # Detecta HTML disfarçado de xls (comum em sistemas legados como Oper)
    amostra = conteudo[:10]
    if amostra.startswith(b"<") or amostra.startswith(b"\xef\xbb\xbf<"):
        # Tenta diferentes combinações de tabela + linha de header
        for header_row in range(0, 5):
            try:
                todas = pd.read_html(BytesIO(conteudo), header=header_row, flavor="lxml")
            except Exception:
                try:
                    todas = pd.read_html(BytesIO(conteudo), header=header_row, flavor="html5lib")
                except Exception:
                    continue

            for df in todas:
                colunas = [str(c).strip() for c in df.columns]
                # Retorna a primeira tabela que tenha ao menos 3 colunas com nome real
                if sum(1 for c in colunas if not c.startswith("Unnamed") and c != "nan") >= 3:
                    df.columns = [str(c).strip() for c in df.columns]
                    return df

        return pd.DataFrame()

    if nome.endswith(".xls"):
        return pd.read_excel(BytesIO(conteudo), engine="xlrd")
    else:
        return pd.read_excel(BytesIO(conteudo), engine="openpyxl")


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
    st.markdown("**Colunas utilizadas por sistema:**")
    st.markdown("🔵 **Cbill:** `Serviço` · `Prazo de execução`")
    st.markdown("🟠 **Oper:** `Numero` · `Data/Hora Limite`")
    st.markdown("---")
    st.info("💡 **Padrão de nome esperado:**\n`base_DD.MM_Cbill.xlsx`\n`base_DD.MM_oper.xls`")

# ─── upload ───────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)
with col1:
    st.subheader("📂 Base Cbill")
    arquivo1 = st.file_uploader("Selecione a base Cbill", type=["xlsx", "xls"], key="base1")
with col2:
    st.subheader("📂 Base Oper")
    arquivo2 = st.file_uploader("Selecione a base Oper", type=["xlsx", "xls"], key="base2")

# ─── processamento ────────────────────────────────────────────────────────────

if arquivo1 and arquivo2:
    try:
        sistema1 = extrair_sistema(arquivo1.name)
        sistema2 = extrair_sistema(arquivo2.name)

        cfg1 = COLUNAS_POR_SISTEMA.get(sistema1.lower())
        cfg2 = COLUNAS_POR_SISTEMA.get(sistema2.lower())

        if not cfg1:
            st.error(f"❌ Sistema **{sistema1}** não reconhecido. O arquivo deve ter `_Cbill` ou `_oper` no nome.")
            st.stop()
        if not cfg2:
            st.error(f"❌ Sistema **{sistema2}** não reconhecido. O arquivo deve ter `_Cbill` ou `_oper` no nome.")
            st.stop()

        col_srv1, col_dt1 = cfg1
        col_srv2, col_dt2 = cfg2

        with st.spinner("Carregando bases..."):
            df1 = ler_excel(arquivo1)
            df2 = ler_excel(arquivo2)

        # Valida colunas
        for sistema, df, arq, cols in [
            (sistema1, df1, arquivo1.name, [col_srv1, col_dt1]),
            (sistema2, df2, arquivo2.name, [col_srv2, col_dt2]),
        ]:
            ausentes = [c for c in cols if c not in df.columns]
            if ausentes:
                st.error(
                    f"❌ **{arq}** — colunas não encontradas: `{'`, `'.join(ausentes)}`\n\n"
                    f"Colunas disponíveis: `{'`, `'.join(df.columns.tolist())}`"
                )
                st.stop()

        df1 = normalizar_datas(df1, col_dt1)
        df2 = normalizar_datas(df2, col_dt2)

        data_alvo = pd.Timestamp(data_filtro).date()
        base1 = df1[df1[col_dt1].dt.date == data_alvo].copy()
        base2 = df2[df2[col_dt2].dt.date == data_alvo].copy()

        total1, total2 = len(base1), len(base2)

        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        m1.metric(f"Serviços em {sistema1}", total1)
        m2.metric(f"Serviços em {sistema2}", total2)
        m3.metric("Diferença", abs(total1 - total2))

        srvs1 = set(base1[col_srv1].astype(str).str.strip())
        srvs2 = set(base2[col_srv2].astype(str).str.strip())

        apenas_em_1 = srvs1 - srvs2
        apenas_em_2 = srvs2 - srvs1

        registros = []

        for srv in sorted(apenas_em_1):
            linhas = base1[base1[col_srv1].astype(str).str.strip() == srv]
            for _, row in linhas.iterrows():
                registros.append({
                    "servico":        row[col_srv1],
                    "data_limite":    row[col_dt1].date() if pd.notna(row[col_dt1]) else data_filtro,
                    "sistema_origem": sistema1,
                    "presente_em":    sistema1,
                    "ausente_em":     sistema2,
                })

        for srv in sorted(apenas_em_2):
            linhas = base2[base2[col_srv2].astype(str).str.strip() == srv]
            for _, row in linhas.iterrows():
                registros.append({
                    "servico":        row[col_srv2],
                    "data_limite":    row[col_dt2].date() if pd.notna(row[col_dt2]) else data_filtro,
                    "sistema_origem": sistema2,
                    "presente_em":    sistema2,
                    "ausente_em":     sistema1,
                })

        df_resultado = pd.DataFrame(registros)

        st.markdown("---")

        if df_resultado.empty:
            st.success("✅ Nenhuma divergência encontrada! As bases estão alinhadas para esta data.")
        else:
            st.warning(f"⚠️ **{len(df_resultado)} serviço(s) divergente(s)** em {data_filtro.strftime('%d/%m/%Y')}")

            tab1, tab2, tab3 = st.tabs([
                f"📋 Todos ({len(df_resultado)})",
                f"🔵 Só em {sistema1} ({len(apenas_em_1)})",
                f"🟠 Só em {sistema2} ({len(apenas_em_2)})",
            ])
            with tab1:
                st.dataframe(df_resultado, use_container_width=True)
            with tab2:
                sub = df_resultado[df_resultado["presente_em"] == sistema1]
                st.dataframe(sub, use_container_width=True) if not sub.empty else st.info("Sem exclusivos.")
            with tab3:
                sub = df_resultado[df_resultado["presente_em"] == sistema2]
                st.dataframe(sub, use_container_width=True) if not sub.empty else st.info("Sem exclusivos.")

            st.markdown("---")
            nome_saida = f"divergentes_{data_filtro.strftime('%d.%m')}_{sistema1}_vs_{sistema2}.xlsx"
            st.download_button(
                label="⬇️ Baixar planilha de divergentes (.xlsx)",
                data=exportar_excel(df_resultado),
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with st.expander(f"👁️ Ver base {sistema1} filtrada ({total1} registros)"):
            st.dataframe(base1, use_container_width=True)
        with st.expander(f"👁️ Ver base {sistema2} filtrada ({total2} registros)"):
            st.dataframe(base2, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")
        st.exception(e)

else:
    st.info("👆 Faça upload das duas bases acima para iniciar a comparação.")
