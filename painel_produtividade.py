"""
Painel de Produtividade — Ativos e MP
=====================================
Dashboard executivo (Streamlit) para acompanhar a produtividade da equipe de
Ativos e Reagendamentos: quantas e quais ordens cada pessoa CRIA, REAGENDA e
CANCELA — no mês e no dia.

Como rodar:
    streamlit run painel_produtividade.py

Lê a planilha 'dados_produtividade.xlsx' (gerada por extrair_produtividade.py).
A regra de negócio (o que conta como criar/reagendar/cancelar e como cada ação
é atribuída à pessoa) fica em logica.py — inclusive os campos configuráveis.
"""

import os
import io
import base64
import textwrap
import datetime as dt

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

import logica as L

# ==========================================================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================================================
st.set_page_config(
    page_title="Produtividade Ativos e MP",
    layout="wide",
)

ARQUIVO_DADOS = "dados_produtividade.xlsx"
LOGO = "logo.png"

# Cores dos tipos de evento (dentro da paleta corporativa) — usadas nos KPIs
COR_EVENTO = {
    "Criada":     "#1E5FCC",   # azul — abrir ordem
    "Reagendada": "#0E9F6E",   # verde — remarcar
    "Cancelada":  "#F59E0B",   # âmbar — cancelar (ação tratada, não "erro")
}

# Paleta 100% azul para os GRÁFICOS (três tons distintos da identidade)
COR_EVENTO_AZUL = {
    "Criada":     "#0A2A66",   # navy
    "Reagendada": "#2F6BD8",   # azul
    "Cancelada":  "#93B4E8",   # azul claro
}
AZUL_SEG = {"Instalação (Vendas)": "#1E5FCC", "Diversos": "#A9C4EE"}

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}


# ==========================================================================
# 2. TEMA VISUAL CORPORATIVO PREMIUM (CSS)
# ==========================================================================
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --navy:      #0A2A66;
    --blue:      #1E5FCC;
    --blue2:     #3B82F6;
    --blue-soft: #EAF1FB;
    --ink:       #16233F;
    --muted:     #647393;
    --line:      #E4EBF6;
    --bg:        #F4F7FD;
    --card:      #FFFFFF;
    --pos:       #0E9F6E;
    --neg:       #E02424;
    --amber:     #F59E0B;
}

html { font-size: 13px; }
html, body, .stApp, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    color: var(--ink);
}

.stApp, [data-testid="stAppViewContainer"] { background: var(--bg); }

header[data-testid="stHeader"] { background: transparent; height: 0; }
#MainMenu, footer { visibility: hidden; }
[data-testid="stToolbar"] { right: 12px; }

.block-container { padding-top: 2.2rem; padding-bottom: 2.5rem; max-width: 1500px; }

h1, h2, h3 {
    color: var(--navy) !important;
    font-weight: 800 !important;
    letter-spacing: -0.01em !important;
}
h4, h5, h6 { color: var(--ink) !important; font-weight: 700 !important; }

/* ---- Sidebar ---- */
[data-testid="stSidebar"] {
    background: var(--card);
    border-right: 1px solid var(--line);
}
[data-testid="stSidebar"] .block-container { padding-top: 1.2rem; }
[data-testid="stSidebar"] [data-testid="stImage"] { text-align: center; }
[data-testid="stSidebar"] [data-testid="stImage"] img { margin: 0 auto; display: block; }

/* ---- Cabeçalho do app ---- */
.app-header {
    display: flex; align-items: flex-end; justify-content: space-between;
    border-bottom: 1px solid var(--line); padding-bottom: 14px; margin-bottom: 22px;
}
.app-header .title { font-size: 1.9rem; font-weight: 800; color: var(--navy); letter-spacing: -0.02em; }
.app-header .subtitle { font-size: .92rem; color: var(--muted); margin-top: 2px; }
.app-header .period {
    font-size: .8rem; font-weight: 700; color: var(--blue);
    background: var(--blue-soft); padding: 6px 14px; border-radius: 999px; white-space: nowrap;
}

/* ---- KPI cards ---- */
.kpi-card {
    position: relative; background: var(--card); border: 1px solid var(--line);
    border-radius: 14px; padding: 18px 18px 16px 18px; overflow: hidden;
    box-shadow: 0 1px 2px rgba(16,40,90,.04), 0 8px 20px rgba(16,40,90,.05);
    transition: transform .18s ease, box-shadow .18s ease; height: 100%; min-height: 122px;
}
.kpi-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 2px 4px rgba(16,40,90,.06), 0 16px 34px rgba(16,40,90,.10);
}
.kpi-bar {
    position: absolute; top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, var(--navy), var(--blue2));
}
.kpi-title {
    font-size: .72rem; font-weight: 700; letter-spacing: .06em; text-transform: uppercase;
    color: var(--muted); margin-bottom: 6px; line-height: 1.25; min-height: 1.9rem;
}
.kpi-value { font-size: 2.05rem; font-weight: 800; line-height: 1.05; color: var(--navy); }
.kpi-sub { font-size: .76rem; color: var(--muted); margin-top: 6px; }

.kpi-hero {
    background: linear-gradient(180deg, #0A2A66 0%, #163D8C 100%);
    border: none; color: #fff;
}
.kpi-hero .kpi-title { color: #AFC4EC; }
.kpi-hero .kpi-value { color: #fff; }
.kpi-hero .kpi-sub { color: #C9D8F4; }

/* ---- Tabs ---- */
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid var(--line); }
.stTabs [data-baseweb="tab"] {
    font-weight: 700; color: var(--muted); padding: 8px 16px; border-radius: 8px 8px 0 0;
}
.stTabs [aria-selected="true"] {
    color: var(--navy) !important; background: var(--blue-soft);
    border-bottom: 2px solid var(--blue) !important;
}

/* ---- Botões ---- */
.stButton > button, .stDownloadButton > button {
    background: var(--navy); color: #fff; border: none; border-radius: 9px;
    font-weight: 700; padding: 8px 18px; transition: background .15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover { background: var(--blue); color:#fff; }

/* ---- Tabelas / dataframes ---- */
[data-testid="stDataFrame"] {
    border: 1px solid var(--line); border-radius: 12px; overflow: hidden;
    box-shadow: 0 1px 2px rgba(16,40,90,.04), 0 8px 20px rgba(16,40,90,.04);
}

/* ---- Seção de bloco ---- */
.section-label {
    font-size: .78rem; font-weight: 700; letter-spacing: .05em; text-transform: uppercase;
    color: var(--muted); margin: 6px 0 2px 0;
}
hr { border-color: var(--line); }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==========================================================================
# 3. HELPERS DE UI
# ==========================================================================
def fmt(n) -> str:
    try:
        return f"{int(n):,}".replace(",", ".")
    except Exception:
        return str(n)


def card_kpi(titulo, valor, sub="", cor="var(--navy)", hero=False, top_bar=True):
    if hero:
        html = (f'<div class="kpi-card kpi-hero">'
                f'<div class="kpi-title">{titulo}</div>'
                f'<div class="kpi-value">{valor}</div>'
                f'<div class="kpi-sub">{sub}</div></div>')
    else:
        bar = '<div class="kpi-bar"></div>' if top_bar else ''
        html = (f'<div class="kpi-card">{bar}'
                f'<div class="kpi-title">{titulo}</div>'
                f'<div class="kpi-value" style="color:{cor}">{valor}</div>'
                f'<div class="kpi-sub">{sub}</div></div>')
    st.markdown(html, unsafe_allow_html=True)


def rotulo_mes(ano_mes: str) -> str:
    try:
        ano, mes = ano_mes.split("-")
        return f"{MESES_PT[int(mes)]} de {ano}"
    except Exception:
        return ano_mes


def estilo_fig(fig, altura=340, legenda=True):
    fig.update_layout(
        height=altura,
        margin=dict(l=10, r=10, t=30, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", size=12, color="#16233F"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    title="", font=dict(size=11)) if legenda else dict(),
        hoverlabel=dict(font_size=12, font_family="Inter"),
    )
    fig.update_xaxes(showgrid=False, linecolor="#E4EBF6", tickfont=dict(size=11))
    fig.update_yaxes(showgrid=True, gridcolor="#EEF2FA", zeroline=False, tickfont=dict(size=11))
    return fig


# ==========================================================================
# 4. CARGA DOS DADOS
# ==========================================================================
@st.cache_data(show_spinner="Carregando dados...")
def carregar(caminho: str, mtime: float):
    df = L.carregar_dados(caminho)
    ev = L.construir_eventos(df)
    return df, ev


if not os.path.exists(ARQUIVO_DADOS):
    st.markdown(
        '<div class="app-header"><div><div class="title">Produtividade Ativos e MP</div>'
        '<div class="subtitle">Equipe de Ativos e Reagendamentos</div></div></div>',
        unsafe_allow_html=True)
    st.warning(
        f"Não encontrei **{ARQUIVO_DADOS}** nesta pasta.\n\n"
        "Rode primeiro a extração para gerar a planilha de dados:\n\n"
        "```\npython extrair_produtividade.py\n```\n\n"
        "Depois recarregue este painel.")
    st.stop()

_mtime = os.path.getmtime(ARQUIVO_DADOS)
df, ev = carregar(ARQUIVO_DADOS, _mtime)


# ==========================================================================
# 5. SIDEBAR — LOGO + FILTROS
# ==========================================================================
with st.sidebar:
    if os.path.exists(LOGO):
        try:
            with open(LOGO, "rb") as _fh:
                _b64 = base64.b64encode(_fh.read()).decode()
            _ext = (os.path.splitext(LOGO)[1].lower().lstrip(".") or "png")
            _mime = ("image/svg+xml" if _ext == "svg"
                     else f"image/{'jpeg' if _ext in ('jpg', 'jpeg') else _ext}")
            st.markdown(
                f'<div style="text-align:center; margin:2px 0 6px 0;">'
                f'<img src="data:{_mime};base64,{_b64}" width="150" '
                f'style="display:inline-block; max-width:80%; height:auto;"></div>',
                unsafe_allow_html=True)
        except Exception:
            st.image(LOGO, width=150)
    atualizado = dt.datetime.fromtimestamp(_mtime).strftime("%d/%m/%Y às %H:%M")
    st.markdown(
        f'<div style="text-align:center; margin-top:8px; margin-bottom:6px;">'
        f'<div style="font-size:1.06rem; font-weight:800; color:var(--navy); '
        f'letter-spacing:-.01em;">Produtividade Ativos e MP</div>'
        f'<div style="font-size:.75rem; color:var(--muted); margin-top:4px;">'
        f'Atualizado em {atualizado}</div></div>',
        unsafe_allow_html=True)
    if os.path.exists("MODO_DEMO.txt"):
        st.warning("⚠️ **Dados de demonstração.**\nRode `extrair_produtividade.py` "
                   "para carregar os dados reais.")
    st.divider()

    meses = L.meses_disponiveis(ev)
    if not meses:
        st.error("Nenhum evento encontrado nos dados.")
        st.stop()

    mes_atual = dt.date.today().strftime("%Y-%m")
    idx_default = meses.index(mes_atual) if mes_atual in meses else len(meses) - 1

    ano_mes = st.selectbox(
        "📅 Mês de referência",
        options=meses, index=idx_default, format_func=rotulo_mes,
    )

    ops_disp = [o for o in L.ORDEM_OPERACOES]
    operacoes = st.multiselect("🏷️ Operação", options=ops_disp, default=[],
                               placeholder="Todas as operações")

    pessoas_disp = list(L.EQUIPE.keys())
    pessoas = st.multiselect("👤 Pessoa", options=pessoas_disp, default=[],
                             placeholder="Toda a equipe")

    incluir_outros = st.toggle(
        "Incluir ações fora da equipe", value=False,
        help="Mostra também ações atribuídas a usuários que não estão na lista da equipe.")


# Aplicar filtros
evf = L.filtrar(
    ev, ano_mes=ano_mes,
    operacoes=operacoes if operacoes else None,
    pessoas=pessoas if pessoas else None,
    apenas_equipe=not incluir_outros,
)


# ==========================================================================
# 6. CABEÇALHO
# ==========================================================================
st.markdown(
    f'<div class="app-header">'
    f'<div><div class="title">Produtividade Ativos e MP</div>'
    f'<div class="subtitle">Equipe de Ativos e Reagendamentos</div></div>'
    f'<div class="period">{rotulo_mes(ano_mes)}</div>'
    f'</div>',
    unsafe_allow_html=True)

k = L.kpis(evf)

# ==========================================================================
# 7. ABAS
# ==========================================================================
aba_geral, aba_pessoa, aba_ordens = st.tabs(
    ["Visão Geral", "Por Pessoa", "Ordens (detalhe)"])


# --------------------------------------------------------------------------
# ABA 1 — VISÃO GERAL
# --------------------------------------------------------------------------
with aba_geral:
    c0, c1, c2, c3, c4, c5 = st.columns([1.35, 1, 1, 1, 1, 1])
    with c0:
        card_kpi("Total de ações", fmt(k["Total"]),
                 f'{k["Pessoas_ativas"]} pessoas ativas', hero=True)
    with c1:
        card_kpi("Ordens criadas", fmt(k["Criada"]), "abertura de OS", cor=COR_EVENTO["Criada"])
    with c2:
        card_kpi("Reagendamentos", fmt(k["Reagendada"]), "novos agendamentos", cor=COR_EVENTO["Reagendada"])
    with c3:
        card_kpi("Cancelamentos", fmt(k["Cancelada"]), "cancelados ativamente", cor=COR_EVENTO["Cancelada"])
    with c4:
        card_kpi("Canceladas na criação", fmt(k["Cancelada_criacao"]),
                 "abertas e já canceladas", cor="var(--muted)")
    with c5:
        media_dia = k["Total"] / max(evf["Dia"].nunique(), 1) if not evf.empty else 0
        card_kpi("Média diária", fmt(round(media_dia)), "por dia com atividade")

    st.markdown("<br>", unsafe_allow_html=True)

    # Prepara os dois blocos e usa uma altura comum (baseada na tabela) p/ alinhar
    op = L.resumo_por_operacao(evf)
    seg = L.reagendamentos_por_segmento(evf)

    if seg.empty or seg[["Instalação (Vendas)", "Diversos"]].sum().sum() == 0:
        seg_show = None
        altura = 320
    else:
        seg_show = seg.sort_values("Total", ascending=False).copy()
        seg_tot = pd.DataFrame([{
            "Pessoa": "TOTAL",
            "Instalação (Vendas)": seg_show["Instalação (Vendas)"].sum(),
            "Diversos": seg_show["Diversos"].sum(),
            "Total": seg_show["Total"].sum(),
        }])
        seg_show = pd.concat([seg_show, seg_tot], ignore_index=True)[
            ["Pessoa", "Instalação (Vendas)", "Diversos", "Total"]]
        altura = (len(seg_show) + 1) * 35 + 8

    colB1, colB2 = st.columns(2, vertical_alignment="top")

    # --- Ações por operação (barras verticais empilhadas) ---
    with colB1:
        st.markdown('<div class="section-label">Ações por operação</div>', unsafe_allow_html=True)
        if op.empty or op["Total"].sum() == 0:
            st.info("Sem dados.")
        else:
            op_plot = op.sort_values("Total", ascending=False)
            x_labels = ["<br>".join(textwrap.wrap(o, 14)) for o in op_plot["Operacao"]]
            total_op = op_plot["Total"]
            txt_cor = {"Criada": "#FFFFFF", "Reagendada": "#FFFFFF", "Cancelada": "#0A2A66"}
            fig = go.Figure()
            for tipo in L.TIPOS_EVENTO:
                fig.add_trace(go.Bar(
                    x=x_labels, y=op_plot[tipo], name=tipo,
                    marker_color=COR_EVENTO_AZUL[tipo],
                    text=[str(int(v)) if v > 0 else "" for v in op_plot[tipo]],
                    textposition="inside", insidetextanchor="middle", textangle=0,
                    textfont=dict(color=txt_cor[tipo], size=11),
                    hovertemplate="%{x}<br>" + tipo + ": %{y}<extra></extra>",
                ))
            fig.add_trace(go.Scatter(
                x=x_labels, y=total_op, mode="text",
                text=[f"<b>{int(v)}</b>" for v in total_op],
                textposition="top center", textfont=dict(color="#0A2A66", size=12),
                showlegend=False, hoverinfo="skip", cliponaxis=False,
            ))
            fig = estilo_fig(fig, altura=altura)
            fig.update_layout(barmode="stack", bargap=0.35,
                              uniformtext_minsize=8, uniformtext_mode="hide",
                              margin=dict(l=10, r=10, t=30, b=64))
            fig.update_yaxes(range=[0, total_op.max() * 1.16])
            fig.update_xaxes(tickfont=dict(size=10))
            st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})

    # --- Reagendamento: Vendas (Instalação) x Diversos (tabela) ---
    with colB2:
        st.markdown('<div class="section-label">Reagendamentos: Vendas × Diversos</div>',
                    unsafe_allow_html=True)
        if seg_show is None:
            st.info("Sem reagendamentos no período.")
        else:
            st.dataframe(
                seg_show, width="stretch", hide_index=True,
                column_config={
                    "Instalação (Vendas)": st.column_config.NumberColumn(
                        "Vendas (Instalação)", format="%d",
                        help="Reagendamentos de OS de Instalação"),
                    "Diversos": st.column_config.NumberColumn("Diversos", format="%d"),
                    "Total": st.column_config.NumberColumn("Total", format="%d"),
                },
            )

    st.divider()

    # --- Tabela resumo por pessoa ---
    st.markdown('<div class="section-label">Quadro resumo por pessoa</div>', unsafe_allow_html=True)
    rp = L.resumo_por_pessoa(evf)
    if not rp.empty:
        canc_cri = L.canceladas_na_criacao_por_pessoa(evf)
        rp = rp.copy()
        rp["Cancelada_criacao"] = rp["Pessoa"].map(canc_cri).fillna(0).astype(int)
        total_row = pd.DataFrame([{
            "Pessoa": "TOTAL", "Operacao": "",
            "Criada": rp["Criada"].sum(), "Reagendada": rp["Reagendada"].sum(),
            "Cancelada": rp["Cancelada"].sum(),
            "Cancelada_criacao": rp["Cancelada_criacao"].sum(), "Total": rp["Total"].sum(),
        }])
        rp_show = pd.concat([rp, total_row], ignore_index=True)
        rp_show = rp_show.rename(columns={"Operacao": "Operação"})
        rp_show = rp_show[["Pessoa", "Operação", "Criada", "Reagendada",
                           "Cancelada", "Cancelada_criacao", "Total"]]
        st.dataframe(
            rp_show, width="stretch", hide_index=True,
            column_config={
                "Criada": st.column_config.NumberColumn("Criadas", format="%d"),
                "Reagendada": st.column_config.NumberColumn("Reagendadas", format="%d"),
                "Cancelada": st.column_config.NumberColumn("Canceladas", format="%d"),
                "Cancelada_criacao": st.column_config.NumberColumn(
                    "Cancel. na criação", format="%d",
                    help="Ordens abertas e já canceladas no mesmo momento "
                         "(não entram no Total)."),
                "Total": st.column_config.NumberColumn(
                    "Total", format="%d",
                    help="Criadas + Reagendadas + Canceladas "
                         "(não inclui as canceladas na criação)."),
            },
        )

    with st.expander("ℹ️  Como os números são calculados"):
        st.markdown(
            "- **Ordem criada** — abertura de uma OS. Conta 1 por ordem, atribuída a quem "
            "abriu o caso.\n"
            "- **Reagendamento** — cada novo agendamento de uma ordem já existente (nova WOLI). "
            "Atribuído a quem fez o novo agendamento.\n"
            "- **Cancelamento** — cada ordem **cancelada ativamente** (ex.: o cliente pede para "
            "cancelar no momento do reagendamento). Atribuído a quem registrou o cancelamento.\n"
            "- **Cancelada na criação** — quando a ordem é **aberta e cancelada no mesmo momento** "
            "(o item original já nasce cancelado). É contabilizada à parte: não é criação nem "
            "cancelamento pró-ativo e **não entra no total de ações**.\n"
            "- **Reagendamento de Vendas** = OS de **Instalação**; os demais entram em **Diversos**.\n"
            "- As datas seguem o fuso de São Paulo, e o mês é fechado pelo dia de cada ação."
        )


# --------------------------------------------------------------------------
# ABA 2 — POR PESSOA (com detalhe por dia)
# --------------------------------------------------------------------------
with aba_pessoa:
    esq, dir_ = st.columns([1, 2.4])

    with esq:
        st.markdown('<div class="section-label">Filtro</div>', unsafe_allow_html=True)
        pessoas_no_mes = sorted(evf["Pessoa"].unique().tolist()) if not evf.empty else []
        pessoa_sel = st.selectbox("Pessoa", options=["(todas)"] + pessoas_no_mes)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-label">Detalhe por dia</div>', unsafe_allow_html=True)
        ver_dia = st.toggle("Filtrar um dia específico", value=False)
        dia_sel = None
        if ver_dia and not evf.empty:
            dias = sorted(evf["Dia"].unique())
            hoje = dt.date.today()
            default_dia = hoje if hoje in dias else dias[-1]
            dia_sel = st.date_input("Dia", value=default_dia,
                                    min_value=min(dias), max_value=max(dias))

    # Base filtrada por pessoa/dia
    base = evf.copy()
    if pessoa_sel != "(todas)":
        base = base[base["Pessoa"] == pessoa_sel]
    if dia_sel is not None:
        base = base[base["Dia"] == dia_sel]

    with dir_:
        titulo_ctx = pessoa_sel if pessoa_sel != "(todas)" else "Toda a equipe"
        if dia_sel is not None:
            titulo_ctx += f" · {dia_sel.strftime('%d/%m/%Y')}"
        st.markdown(f"#### {titulo_ctx}")

        kk = L.kpis(base)
        m1, m2, m3, m4, m5 = st.columns(5)
        with m1: card_kpi("Criadas", fmt(kk["Criada"]), cor=COR_EVENTO["Criada"])
        with m2: card_kpi("Reagendadas", fmt(kk["Reagendada"]), cor=COR_EVENTO["Reagendada"])
        with m3: card_kpi("Canceladas", fmt(kk["Cancelada"]), cor=COR_EVENTO["Cancelada"])
        with m4: card_kpi("Na criação", fmt(kk["Cancelada_criacao"]), cor="var(--muted)")
        with m5: card_kpi("Total", fmt(kk["Total"]))

    st.markdown("<br>", unsafe_allow_html=True)

    if dia_sel is None and not base.empty:
        st.markdown('<div class="section-label">Ações por dia</div>', unsafe_allow_html=True)
        diar = L.serie_diaria(base)
        if not diar.empty:
            dias_dt = pd.to_datetime(diar["Dia"])
            total_dia = diar[L.TIPOS_EVENTO].sum(axis=1)
            txt_cor = {"Criada": "#FFFFFF", "Reagendada": "#FFFFFF", "Cancelada": "#0A2A66"}
            fig = go.Figure()
            for tipo in L.TIPOS_EVENTO:
                fig.add_trace(go.Bar(
                    x=dias_dt, y=diar[tipo], name=tipo,
                    marker_color=COR_EVENTO_AZUL[tipo],
                    text=[str(int(v)) if v > 0 else "" for v in diar[tipo]],
                    textposition="inside", insidetextanchor="middle", textangle=0,
                    textfont=dict(color=txt_cor[tipo], size=10),
                    hovertemplate="%{x|%d/%m}<br>" + tipo + ": %{y}<extra></extra>",
                ))
            fig.add_trace(go.Scatter(
                x=dias_dt, y=total_dia, mode="text",
                text=[f"<b>{int(v)}</b>" if v > 0 else "" for v in total_dia],
                textposition="top center", textfont=dict(color="#0A2A66", size=11),
                showlegend=False, hoverinfo="skip", cliponaxis=False,
            ))
            fig.update_layout(barmode="stack", bargap=0.30,
                              uniformtext_minsize=8, uniformtext_mode="hide")
            fig.update_yaxes(range=[0, max(total_dia.max() * 1.16, 1)])
            st.plotly_chart(estilo_fig(fig, altura=360), width="stretch",
                            config={"displayModeBar": False})

    # --- Quais ordens (drill-down) ---
    st.markdown('<div class="section-label">Quais ordens</div>', unsafe_allow_html=True)
    if base.empty:
        st.info("Nenhuma ação para o filtro selecionado.")
    else:
        tabela = base.sort_values("Data_Evento", ascending=False).copy()
        tabela["Data"] = pd.to_datetime(tabela["Data_Evento"]).dt.strftime("%d/%m/%Y %H:%M")
        cols = ["Data", "Tipo_Evento", "Pessoa", "Operacao", "Numero_OS", "LineItemNumber",
                "CaseNumber", "Cliente", "Segmento_Servico", "WO_TipoServico",
                "WOLI_Status", "Nome_Tecnico", "Motivo_Insucesso"]
        tabela = tabela[cols].rename(columns={
            "Tipo_Evento": "Ação", "Operacao": "Operação", "Numero_OS": "Nº OS",
            "LineItemNumber": "Item", "CaseNumber": "Caso", "Segmento_Servico": "Segmento",
            "WO_TipoServico": "Tipo de Serviço", "WOLI_Status": "Status atual",
            "Nome_Tecnico": "Técnico", "Motivo_Insucesso": "Motivo",
        })
        st.dataframe(tabela, width="stretch", hide_index=True, height=430)
        st.caption(f"{len(tabela)} ação(ões). Clique no cabeçalho para ordenar; use a busca (ícone 🔍) para filtrar.")


# --------------------------------------------------------------------------
# ABA 3 — ORDENS (DETALHE + DOWNLOAD)
# --------------------------------------------------------------------------
with aba_ordens:
    f1, f2, f3 = st.columns(3)
    with f1:
        tipos = st.multiselect("Ação", options=L.TIPOS_EVENTO + [L.CANCELADA_CRIACAO],
                               default=L.TIPOS_EVENTO + [L.CANCELADA_CRIACAO])
    with f2:
        segs = st.multiselect("Segmento", options=["Instalação (Vendas)", "Diversos"],
                              default=["Instalação (Vendas)", "Diversos"])
    with f3:
        busca = st.text_input("Buscar (Nº OS, caso, cliente)", "")

    base = evf.copy()
    if tipos:
        base = base[base["Tipo_Evento"].isin(tipos)]
    if segs:
        base = base[base["Segmento_Servico"].isin(segs)]
    if busca.strip():
        q = busca.strip().lower()
        mask = (
            base["Numero_OS"].astype(str).str.lower().str.contains(q, na=False)
            | base["CaseNumber"].astype(str).str.lower().str.contains(q, na=False)
            | base["Cliente"].astype(str).str.lower().str.contains(q, na=False)
        )
        base = base[mask]

    st.markdown(f'<div class="section-label">{len(base)} ação(ões) encontradas</div>',
                unsafe_allow_html=True)

    if not base.empty:
        tabela = base.sort_values("Data_Evento", ascending=False).copy()
        tabela["Data"] = pd.to_datetime(tabela["Data_Evento"]).dt.strftime("%d/%m/%Y %H:%M")
        cols = ["Data", "Tipo_Evento", "Pessoa", "Operacao", "Numero_OS", "LineItemNumber",
                "CaseNumber", "Cliente", "CodigoItem", "Segmento_Servico", "WO_TipoServico",
                "WOLI_Status", "WO_Status", "Nome_Tecnico", "Motivo_Insucesso"]
        tabela_show = tabela[cols].rename(columns={
            "Tipo_Evento": "Ação", "Operacao": "Operação", "Numero_OS": "Nº OS",
            "LineItemNumber": "Item", "CaseNumber": "Caso", "CodigoItem": "Cód. Item",
            "Segmento_Servico": "Segmento", "WO_TipoServico": "Tipo de Serviço",
            "WOLI_Status": "Status WOLI", "WO_Status": "Status OS",
            "Nome_Tecnico": "Técnico", "Motivo_Insucesso": "Motivo",
        })
        st.dataframe(tabela_show, width="stretch", hide_index=True, height=460)

        # Download em Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            tabela_show.to_excel(writer, index=False, sheet_name="Ações")
        buffer.seek(0)
        st.download_button(
            "⬇️  Baixar em Excel",
            data=buffer,
            file_name=f"produtividade_{ano_mes}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Nenhuma ação para os filtros selecionados.")
