import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date, time
import io
import json
import random

st.set_page_config(
    page_title="Gestão de Turno · Real-Time",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .kpi-card {
        background: #1e2235; border: 1px solid #3d4270;
        border-radius: 12px; padding: 20px 24px;
        text-align: center; margin-bottom: 8px;
    }
    .kpi-value { font-size: 2.2rem; font-weight: 800; color: #7c83ff; }
    .kpi-label { font-size: 0.75rem; color: #8b92b8; text-transform: uppercase; letter-spacing: 0.08em; }
    .kpi-card.green .kpi-value { color: #4ade80; }
    .kpi-card.yellow .kpi-value { color: #facc15; }
    .kpi-card.blue .kpi-value { color: #60a5fa; }
    .prefix-ok { background:#0f2918; border:1px solid #22c55e; border-radius:8px; padding:6px 12px; margin:3px 0; font-size:0.88rem; color:#4ade80; }
    .prefix-pending { background:#1f1a0f; border:1px solid #854d0e; border-radius:8px; padding:6px 12px; margin:3px 0; font-size:0.88rem; color:#fbbf24; }
    .prefix-interval { background:#1a1028; border:1px solid #7c3aed; border-radius:8px; padding:6px 12px; margin:3px 0; font-size:0.88rem; color:#c084fc; }
    .section-header { font-size:0.72rem; font-weight:700; text-transform:uppercase; letter-spacing:0.12em; color:#5c6394; margin:16px 0 8px 0; padding-bottom:4px; border-bottom:1px solid #2a2d45; }
</style>
""", unsafe_allow_html=True)

# ── PREFIXOS ──────────────────────────────────────────
PREFIXOS = {
    "GOOC": ["GOOC-01","GOOC-02","GOOC-03","GOOC-04","GOOC-05",
             "GOOC-06","GOOC-07","GOOC-08","GOOC-09","GOOC-10"],
    "GOOH": ["GOOH-01","GOOH-02","GOOH-03","GOOH-04",
             "GOOH-05","GOOH-06","GOOH-07","GOOH-08"],
    "GOOK": ["GOOK-01","GOOK-02","GOOK-03","GOOK-04","GOOK-05","GOOK-06"],
    "GOOL": ["GOOL-01","GOOL-02","GOOL-03","GOOL-04",
             "GOOL-05","GOOL-06","GOOL-07"],
}

# ── DADOS DEMO ────────────────────────────────────────
def gerar_dados_demo():
    hoje = date.today()
    rows = []
    todos = [p for g in PREFIXOS.values() for p in g]
    abertos = random.sample(todos, k=int(len(todos) * 0.75))
    status_opts = ["Ativo","Intervalo","Ativo","Ativo","Ativo"]
    atv_opts = ["Turno Normal","Hora Extra","Turno Normal","Patrulhamento","Turno Normal"]
    mesas = ["Mesa 01","Mesa 02","Mesa 03","Mesa 04","Mesa 05"]
    viaturas = ["VTR-001","VTR-002","VTR-003","VTR-004","VTR-005","VTR-006"]
    for p in abertos:
        h = datetime.combine(hoje, time(hour=random.randint(6,14),
                                        minute=random.choice([0,15,30,45])))
        st_ = random.choice(status_opts)
        rows.append({
            "Prefixo": p,
            "Início de Turno": h,
            "Status Comercial": st_,
            "Intervalo": h.strftime("%H:%M") if st_ == "Intervalo" else "",
            "Tipo de Atividade": random.choice(atv_opts),
            "Mesa": random.choice(mesas),
            "Viatura": random.choice(viaturas),
            "Operador": f"Op. {random.randint(100,999)}",
        })
    return pd.DataFrame(rows)

# ── SESSION STATE ─────────────────────────────────────
if "df_principal" not in st.session_state:
    st.session_state.df_principal = gerar_dados_demo()
if "intervalos_info" not in st.session_state:
    st.session_state.intervalos_info = {}

# ── HELPERS ───────────────────────────────────────────
def get_df_hoje():
    df = st.session_state.df_principal.copy()
    df["Início de Turno"] = pd.to_datetime(df["Início de Turno"])
    return df[df["Início de Turno"].dt.date == date.today()].copy()

def is_intervalo(row):
    return bool(row.get("Intervalo","")) or str(row.get("Status Comercial","")) == "Intervalo"

def tempo_turno(inicio):
    if pd.isnull(inicio): return 0
    return round((datetime.now() - pd.to_datetime(inicio)).total_seconds() / 3600, 2)

# ── SIDEBAR ───────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🛡️ Gestão de Turno")
    st.caption(f"📅 {date.today().strftime('%d/%m/%Y')} · ⏱ {datetime.now().strftime('%H:%M')}")
    st.divider()
    aba = st.radio("Nav", ["📊 O Dia","⏸ Intervalos","✏️ Atribuição","⚙️ Configurações"],
                   label_visibility="collapsed")
    st.divider()
    df_s = get_df_hoje()
    ativos_s = set(df_s["Prefixo"])
    todos_flat = [p for g in PREFIXOS.values() for p in g]
    st.markdown(f"**Online:** {len(df_s)} equipes")
    st.markdown(f"**Em Intervalo:** {len(df_s[df_s.apply(is_intervalo,axis=1)])}")
    st.markdown(f"**Pendentes:** {sum(1 for p in todos_flat if p not in ativos_s)}")

# ═════════════════════════════════════════════════════
# ABA 1 — O DIA
# ═════════════════════════════════════════════════════
if aba == "📊 O Dia":
    st.markdown("## 📊 Monitoramento do Dia")
    df = get_df_hoje()
    todos_flat = [p for g in PREFIXOS.values() for p in g]
    ativos = set(df["Prefixo"])
    em_int = df[df.apply(is_intervalo, axis=1)]
    media = round(df["Início de Turno"].apply(tempo_turno).mean(), 1) if len(df) else 0

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(f'<div class="kpi-card green"><div class="kpi-value">{len(df)}</div><div class="kpi-label">Equipes Online</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card yellow"><div class="kpi-value">{len(em_int)}</div><div class="kpi-label">Em Intervalo</div></div>', unsafe_allow_html=True)
    with c3:
        pend = sum(1 for p in todos_flat if p not in ativos)
        st.markdown(f'<div class="kpi-card"><div class="kpi-value">{pend}</div><div class="kpi-label">Pendentes</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card blue"><div class="kpi-value">{media}h</div><div class="kpi-label">Média Turno</div></div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📋 Status dos Prefixos")
    cols = st.columns(4)
    for i, (grupo, lista) in enumerate(PREFIXOS.items()):
        with cols[i]:
            st.markdown(f'<div class="section-header">{grupo}</div>', unsafe_allow_html=True)
            for p in lista:
                if p in ativos:
                    row = df[df["Prefixo"]==p].iloc[0]
                    h = pd.to_datetime(row["Início de Turno"]).strftime("%H:%M")
                    cls = "prefix-interval" if is_intervalo(row) else "prefix-ok"
                    ico = "⏸" if is_intervalo(row) else "✅"
                    st.markdown(f'<div class="{cls}">{ico} <b>{p}</b> <small>{h}</small></div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="prefix-pending">⭕ <b>{p}</b> <small>Pendente</small></div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📈 Gráficos")
    g1, g2 = st.columns(2)
    if not df.empty:
        with g1:
            df2 = df.copy()
            df2["Hora"] = pd.to_datetime(df2["Início de Turno"]).dt.hour
            vol = df2.groupby("Hora").size().reset_index(name="Aberturas")
            fig = px.bar(vol, x="Hora", y="Aberturas", title="Aberturas por Hora",
                         color="Aberturas", color_continuous_scale="blues", template="plotly_dark")
            fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        with g2:
            sc = df["Status Comercial"].value_counts().reset_index()
            sc.columns = ["Status","Total"]
            fig2 = px.pie(sc, values="Total", names="Status", title="Distribuição de Status",
                          template="plotly_dark", color_discrete_sequence=px.colors.sequential.Purples_r)
            fig2.update_layout(paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig2, use_container_width=True)
        if "Tipo de Atividade" in df.columns:
            ac = df["Tipo de Atividade"].value_counts().reset_index()
            ac.columns = ["Atividade","Total"]
            fig3 = px.bar(ac, x="Total", y="Atividade", orientation="h",
                          title="Tipo de Atividade", template="plotly_dark",
                          color="Total", color_continuous_scale="purples")
            fig3.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig3, use_container_width=True)

# ═════════════════════════════════════════════════════
# ABA 2 — INTERVALOS
# ═════════════════════════════════════════════════════
elif aba == "⏸ Intervalos":
    st.markdown("## ⏸ Edição de Status · Intervalos")
    df = get_df_hoje()
    em_int = df[df.apply(is_intervalo, axis=1)]
    if em_int.empty:
        st.info("✅ Nenhuma equipe em intervalo no momento.")
    else:
        st.caption(f"{len(em_int)} equipe(s) em intervalo")
        for _, row in em_int.iterrows():
            p = row["Prefixo"]
            info = st.session_state.intervalos_info.get(p, {})
            st.markdown(f"**⏸ {p}** — Início: {pd.to_datetime(row['Início de Turno']).strftime('%H:%M')}")
            ca, cb, cc = st.columns([2,1,1])
            with ca:
                motivo = st.text_input("Motivo", value=info.get("motivo",""), key=f"m_{p}", placeholder="Ex: Refeição")
            with cb:
                fim = st.time_input("Fim", value=info.get("fim", datetime.now().time()), key=f"f_{p}")
            with cc:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Salvar", key=f"s_{p}"):
                    inicio_dt = pd.to_datetime(row["Início de Turno"])
                    fim_dt = datetime.combine(date.today(), fim)
                    dur = round((fim_dt - inicio_dt).total_seconds() / 60, 1)
                    st.session_state.intervalos_info[p] = {"motivo": motivo, "fim": fim, "duracao_min": dur}
                    st.success(f"✅ {p} — {dur} min registrados")
                    st.rerun()
            if p in st.session_state.intervalos_info:
                sv = st.session_state.intervalos_info[p]
                st.caption(f"✔ Motivo: {sv.get('motivo','—')} · Duração: {sv.get('duracao_min','?')} min")
            st.divider()

# ═════════════════════════════════════════════════════
# ABA 3 — ATRIBUIÇÃO
# ═════════════════════════════════════════════════════
elif aba == "✏️ Atribuição":
    st.markdown("## ✏️ Atribuição e Modificação")
    df = get_df_hoje()
    if df.empty:
        st.warning("Sem dados para hoje.")
    else:
        f1, f2 = st.columns([1,2])
        with f1:
            gf = st.selectbox("Grupo", ["Todos"] + list(PREFIXOS.keys()))
        with f2:
            tf = st.text_input("Buscar prefixo", placeholder="Ex: GOOC-01")
        de = df.copy()
        if gf != "Todos": de = de[de["Prefixo"].isin(PREFIXOS[gf])]
        if tf: de = de[de["Prefixo"].str.contains(tf.upper(), na=False)]

        ATIVS = ["Turno Normal","Hora Extra","Patrulhamento","Cobertura","Treinamento"]
        MESAS = ["Mesa 01","Mesa 02","Mesa 03","Mesa 04","Mesa 05"]
        STATUS = ["Ativo","Intervalo","Encerrado","Afastado"]

        edited = st.data_editor(
            de[["Prefixo","Operador","Tipo de Atividade","Mesa","Viatura","Status Comercial"]],
            use_container_width=True, num_rows="fixed",
            column_config={
                "Prefixo": st.column_config.TextColumn(disabled=True),
                "Operador": st.column_config.TextColumn(disabled=True),
                "Tipo de Atividade": st.column_config.SelectboxColumn(options=ATIVS),
                "Mesa": st.column_config.SelectboxColumn(options=MESAS),
                "Viatura": st.column_config.TextColumn(),
                "Status Comercial": st.column_config.SelectboxColumn(options=STATUS),
            }, key="editor_atrib"
        )
        if st.button("✅ Confirmar Modificações", type="primary"):
            dp = st.session_state.df_principal.copy()
            dp["Início de Turno"] = pd.to_datetime(dp["Início de Turno"])
            for _, r in edited.iterrows():
                m = dp["Prefixo"] == r["Prefixo"]
                for c in ["Tipo de Atividade","Mesa","Viatura","Status Comercial"]:
                    if c in r.index and c in dp.columns:
                        dp.loc[m, c] = r[c]
            st.session_state.df_principal = dp
            st.success("✅ Modificações salvas!")
            st.rerun()

        with st.expander("🔧 Modificação Individual"):
            ps = st.selectbox("Prefixo", df["Prefixo"].tolist())
            rs = df[df["Prefixo"]==ps].iloc[0]
            c1,c2,c3 = st.columns(3)
            with c1:
                idx = ATIVS.index(rs.get("Tipo de Atividade","Turno Normal")) if rs.get("Tipo de Atividade") in ATIVS else 0
                na = st.selectbox("Atividade", ATIVS, index=idx)
            with c2:
                nm = st.selectbox("Mesa", MESAS)
            with c3:
                nv = st.text_input("Viatura", value=str(rs.get("Viatura","")))
            if st.button("💾 Confirmar Individual"):
                dp = st.session_state.df_principal.copy()
                m = dp["Prefixo"] == ps
                dp.loc[m,"Tipo de Atividade"] = na
                dp.loc[m,"Mesa"] = nm
                dp.loc[m,"Viatura"] = nv
                st.session_state.df_principal = dp
                st.success(f"✅ {ps} atualizado!")
                st.rerun()

# ═════════════════════════════════════════════════════
# ABA 4 — CONFIGURAÇÕES
# ═════════════════════════════════════════════════════
elif aba == "⚙️ Configurações":
    st.markdown("## ⚙️ Gestão e Configurações")
    t1, t2, t3 = st.tabs(["📁 Carregar Arquivo","📤 Exportar","📖 Manual"])

    with t1:
        st.info("CSV deve ter: Prefixo · Início de Turno · Status Comercial · Intervalo · Tipo de Atividade · Mesa · Viatura · Operador")
        up = st.file_uploader("Selecione o CSV do dia", type=["csv"])
        if up:
            try:
                dn = pd.read_csv(up)
                dn["Início de Turno"] = pd.to_datetime(dn["Início de Turno"])
                st.session_state.df_principal = dn
                st.success(f"✅ {len(dn)} registros importados!")
                st.dataframe(dn.head(), use_container_width=True)
            except Exception as e:
                st.error(f"Erro: {e}")
        if st.button("🔄 Recarregar dados demo"):
            st.session_state.df_principal = gerar_dados_demo()
            st.session_state.intervalos_info = {}
            st.success("✅ Dados demo recarregados!")
            st.rerun()

    with t2:
        df = get_df_hoje()
        if df.empty:
            st.warning("Sem dados.")
        else:
            de = df.copy()
            de["Motivo Intervalo"] = de["Prefixo"].map(lambda p: st.session_state.intervalos_info.get(p,{}).get("motivo",""))
            de["Fim Intervalo"] = de["Prefixo"].map(lambda p: str(st.session_state.intervalos_info.get(p,{}).get("fim","")))
            de["Duração (min)"] = de["Prefixo"].map(lambda p: st.session_state.intervalos_info.get(p,{}).get("duracao_min",""))
            de["Tempo Turno (h)"] = de["Início de Turno"].apply(tempo_turno)
            st.dataframe(de, use_container_width=True)
            buf = io.StringIO()
            de.to_csv(buf, index=False, encoding="utf-8-sig")
            st.download_button("⬇️ Baixar CSV", data=buf.getvalue().encode("utf-8-sig"),
                               file_name=f"relatorio_{date.today():%Y%m%d}.csv", mime="text/csv")
            todos_flat = [p for g in PREFIXOS.values() for p in g]
            ativos = set(df["Prefixo"])
            resumo = {
                "Data": str(date.today()),
                "Online": len(df),
                "Pendentes": len([p for p in todos_flat if p not in ativos]),
                "Intervalos registrados": len(st.session_state.intervalos_info),
                "Média turno (h)": round(de["Tempo Turno (h)"].mean(), 2),
            }
            for k,v in resumo.items(): st.markdown(f"**{k}:** {v}")
            st.download_button("⬇️ Baixar JSON", data=json.dumps(resumo, ensure_ascii=False, indent=2),
                               file_name=f"resumo_{date.today():%Y%m%d}.json", mime="application/json")

    with t3:
        st.markdown("""
**📊 O Dia** — KPIs + status visual de todos os prefixos + gráficos.
- ✅ Verde = turno aberto | ⭕ Amarelo = pendente | ⏸ Roxo = intervalo

**⏸ Intervalos** — Preencha motivo e fim para calcular duração automática.

**✏️ Atribuição** — Edite na tabela ou via formulário individual. Confirme para aplicar.

**⚙️ Configurações** — Upload do CSV do dia + exportação do relatório final.

---
**Formato CSV:** Prefixo · Início de Turno (YYYY-MM-DD HH:MM) · Status Comercial · Intervalo · Tipo de Atividade · Mesa · Viatura · Operador
        """)
