import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Comparador de Bases",
    page_icon="🔍",
    layout="wide"
)

st.title("🔍 Comparador de Bases — Serviços Divergentes")
st.markdown("Carregue duas bases Excel e identifique os serviços divergentes por data limite.")

# ─── helpers ────────────────────────────────────────────────────────────────

def extrair_sistema(nome_arquivo: str) -> str:
    """Extrai o nome do sistema a partir do nome do arquivo.
    Exemplo: 'base_29.04_Cbill.xlsx' → 'Cbill'
    """
    nome = nome_arquivo.replace(".xlsx", "").replace(".xls", "")
    partes = nome.split("_")
    if len(partes) >= 3:
        return partes[-1]
    return nome


def normalizar_datas(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """Tenta converter a coluna de data para datetime."""
    try:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
    except Exception:
        pass
    return df


def exportar_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Divergentes")
        ws = writer.sheets["Divergentes"]
        # Ajusta largura das colunas
        for col_cells in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = max(max_len + 4, 14)
    return buf.getvalue()

# ─── sidebar — configuração ─────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Configuração")
    col_servico = st.text_input("Nome da coluna de Serviço", value="servico",
                                help="Nome exato da coluna que identifica o serviço nas planilhas")
    col_data    = st.text_input("Nome da coluna de Data Limite", value="data_limite",
                                help="Nome exato da coluna de data limite")
    data_filtro = st.date_input("Data limite a comparar", value=datetime.today(),
                                help="Filtra apenas os serviços desta data")
    st.markdown("---")
    st.info("💡 **Padrão de nome esperado:**\n`base_DD.MM_Sistema.xlsx`\n\nExemplo:\n`base_29.04_Cbill.xlsx`")

# ─── upload das bases ────────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.subheader("📂 Base 1")
    arquivo1 = st.file_uploader("Selecione a primeira base", type=["xlsx", "xls"], key="base1")

with col2:
    st.subheader("📂 Base 2")
    arquivo2 = st.file_uploader("Selecione a segunda base", type=["xlsx", "xls"], key="base2")

# ─── processamento ───────────────────────────────────────────────────────────

if arquivo1 and arquivo2:
    try:
        sistema1 = extrair_sistema(arquivo1.name)
        sistema2 = extrair_sistema(arquivo2.name)

        df1 = pd.read_excel(arquivo1)
        df2 = pd.read_excel(arquivo2)

        # Valida colunas
        colunas_necessarias = [col_servico, col_data]
        for nome, df, arq in [(sistema1, df1, arquivo1.name), (sistema2, df2, arquivo2.name)]:
            ausentes = [c for c in colunas_necessarias if c not in df.columns]
            if ausentes:
                st.error(f"❌ Arquivo **{arq}** não tem as colunas: `{'`, `'.join(ausentes)}`\n\n"
                         f"Colunas encontradas: `{'`, `'.join(df.columns.tolist())}`")
                st.stop()

        # Normaliza datas
        df1 = normalizar_datas(df1, col_data)
        df2 = normalizar_datas(df2, col_data)

        # Filtra pela data escolhida
        data_alvo = pd.Timestamp(data_filtro)
        mask1 = df1[col_data].dt.date == data_alvo.date()
        mask2 = df2[col_data].dt.date == data_alvo.date()

        base1_filtrada = df1[mask1].copy()
        base2_filtrada = df2[mask2].copy()

        total1 = len(base1_filtrada)
        total2 = len(base2_filtrada)

        # Métricas rápidas
        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        m1.metric(f"Serviços em {sistema1}", total1)
        m2.metric(f"Serviços em {sistema2}", total2)
        m3.metric("Diferença", abs(total1 - total2))

        # Identifica divergentes (serviços que existem em uma base mas não na outra)
        servicos1 = set(base1_filtrada[col_servico].astype(str).str.strip())
        servicos2 = set(base2_filtrada[col_servico].astype(str).str.strip())

        apenas_em_1 = servicos1 - servicos2   # presentes só na base 1
        apenas_em_2 = servicos2 - servicos1   # presentes só na base 2

        registros_divergentes = []

        for srv in sorted(apenas_em_1):
            linhas = base1_filtrada[base1_filtrada[col_servico].astype(str).str.strip() == srv]
            for _, row in linhas.iterrows():
                registros_divergentes.append({
                    "servico": row[col_servico],
                    "data_limite": row[col_data].date() if pd.notna(row[col_data]) else data_filtro,
                    "sistema_origem": sistema1,
                    "presente_em": sistema1,
                    "ausente_em": sistema2,
                })

        for srv in sorted(apenas_em_2):
            linhas = base2_filtrada[base2_filtrada[col_servico].astype(str).str.strip() == srv]
            for _, row in linhas.iterrows():
                registros_divergentes.append({
                    "servico": row[col_servico],
                    "data_limite": row[col_data].date() if pd.notna(row[col_data]) else data_filtro,
                    "sistema_origem": sistema2,
                    "presente_em": sistema2,
                    "ausente_em": sistema1,
                })

        df_resultado = pd.DataFrame(registros_divergentes)

        # ─── exibição dos resultados ─────────────────────────────────────────

        st.markdown("---")

        if df_resultado.empty:
            st.success("✅ Nenhuma divergência encontrada! As bases estão alinhadas para esta data.")
        else:
            st.warning(f"⚠️ **{len(df_resultado)} serviço(s) divergente(s)** encontrado(s) em {data_filtro.strftime('%d/%m/%Y')}")

            tab1, tab2, tab3 = st.tabs([
                f"📋 Todos ({len(df_resultado)})",
                f"🔵 Só em {sistema1} ({len(apenas_em_1)})",
                f"🟠 Só em {sistema2} ({len(apenas_em_2)})"
            ])

            with tab1:
                st.dataframe(df_resultado, use_container_width=True)

            with tab2:
                df_tab1 = df_resultado[df_resultado["presente_em"] == sistema1]
                if df_tab1.empty:
                    st.info("Nenhum serviço exclusivo desta base.")
                else:
                    st.dataframe(df_tab1, use_container_width=True)

            with tab3:
                df_tab2 = df_resultado[df_resultado["presente_em"] == sistema2]
                if df_tab2.empty:
                    st.info("Nenhum serviço exclusivo desta base.")
                else:
                    st.dataframe(df_tab2, use_container_width=True)

            # Botão de download
            st.markdown("---")
            nome_saida = f"divergentes_{data_filtro.strftime('%d.%m')}_{sistema1}_vs_{sistema2}.xlsx"
            st.download_button(
                label="⬇️ Baixar planilha de divergentes (.xlsx)",
                data=exportar_excel(df_resultado),
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # Preview das bases filtradas (expansível)
        with st.expander(f"👁️ Ver base {sistema1} filtrada ({total1} registros)"):
            st.dataframe(base1_filtrada, use_container_width=True)

        with st.expander(f"👁️ Ver base {sistema2} filtrada ({total2} registros)"):
            st.dataframe(base2_filtrada, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")
        st.exception(e)

else:
    st.info("👆 Faça upload das duas bases acima para iniciar a comparação.")
