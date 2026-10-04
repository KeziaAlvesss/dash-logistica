"""
Dashboard de Carregamento x Meta Semanal  |  Colchões BonSono

Como rodar:
    pip install -r requirements.txt
    streamlit run dashboard_carregamento.py

Arquivos ao lado do script (opcionais, mas recomendados):
    logo_bonsono.png         -> logo exibido no cabeçalho (troque pelo original em alta resolução)
    .streamlit/config.toml   -> tema (cores) do Streamlit

Formato esperado da planilha:
  - Linha de cabeçalho começando com "PLACA", "MOTORISTA" e as datas do dia a dia
  - Linha "TOTAL ..." encerrando cada bloco (ex.: TOTAL COMERCIAL, TOTAL LOJAS)
  - O rótulo "META / CAMINHÃO" com o valor ao lado
O primeiro bloco é a frota com meta; os demais viram seções só com o carregamento.
"""
import io
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

NAVY = "#003399"
CORES = {"Meta batida": "#2E9E5B", "No ritmo": "#2A6FDB", "Atenção": "#F2A33A", "Crítico": "#D64545"}
FUNDO = {"Meta batida": "#D9F2E3", "No ritmo": "#DCE7FB", "Atenção": "#FDEBCB", "Crítico": "#F9D9D9"}
EMOJI = {"Meta batida": "🟢", "No ritmo": "🔵", "Atenção": "🟠", "Crítico": "🔴"}
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]

CSS = """
<style>
.stApp{background:#F8F9FA;}
[data-testid="stSidebar"]{background:#F0F2F6;}
[data-testid="stSidebar"] h3{color:#003399;font-weight:700;margin-top:.6rem;}
.block-container{padding-top:2.4rem;max-width:1500px;}
.titulo{color:#003399;font-weight:800;font-size:2.7rem;line-height:1.1;margin:0;}
.subtitulo{color:#374151;font-size:1.02rem;margin-top:.7rem;}
.pill{display:inline-block;background:#E4EBFA;color:#003399;border-radius:999px;
      padding:4px 14px;font-size:.82rem;font-weight:600;margin:.9rem .4rem 0 0;}
.pill.warn{background:#FDEBCB;color:#8A5A00;}
.sec{color:#003399;font-weight:700;font-size:1.7rem;line-height:1.2;margin:2.2rem 0 .9rem;}
.kpi{background:#fff;border-radius:12px;box-shadow:0 2px 10px rgba(0,30,90,.08);
     border-top:3px solid #003399;padding:20px 10px 12px;text-align:center;
     height:132px;box-sizing:border-box;overflow:hidden;}
.kpi-v{color:#003399;font-size:clamp(.95rem,1.15vw,1.25rem);font-weight:800;line-height:1.2;
       white-space:nowrap;}
.kpi-l{color:#6B7280;font-size:.9rem;margin-top:4px;white-space:nowrap;}
.kpi-s{font-size:.78rem;font-weight:600;margin-top:8px;white-space:nowrap;}
.kpi-bar{height:6px;background:#E5E9F2;border-radius:3px;margin:9px 14px 0;overflow:hidden;}
.kpi-bar>div{height:100%;background:linear-gradient(90deg,#003399,#2A6FDB);border-radius:3px;}
div[data-testid="stPlotlyChart"]{background:#fff;border-radius:12px;
     box-shadow:0 2px 10px rgba(0,30,90,.08);padding:6px 6px 0;}
div[data-testid="stDataFrame"]{background:#fff;border-radius:12px;
     box-shadow:0 2px 10px rgba(0,30,90,.08);padding:6px;}
[data-testid="stImage"] img{background:#fff;border-radius:14px;padding:6px;
     box-shadow:0 2px 10px rgba(0,30,90,.08);}
.stDownloadButton button{border-radius:8px;border:1px solid #003399;color:#003399;font-weight:600;}
</style>
"""

BASE = dict(
    template="plotly_white", paper_bgcolor="#fff", plot_bgcolor="#fff",
    font=dict(family="Source Sans Pro, Segoe UI, sans-serif", size=12, color="#1F2937"),
    hoverlabel=dict(bgcolor="white", font_size=12),
)


# ---------------------------------------------------------------- formatação
def brl(v, casas=2):
    if v is None or pd.isna(v):
        return "–"
    s = f"{abs(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if v < 0 else ''}R$ {s}"


def mil(v):
    if v is None or pd.isna(v):
        return ""
    return f"{v / 1000:,.1f}k".replace(",", "X").replace(".", ",").replace("X", ".")


def rotulo_data(d):
    return f"{DIAS_SEMANA[d.weekday()]} {d:%d/%m}"


# ------------------------------------------------------------------- leitura
def _norm(x):
    return " ".join(str(x).split()) if pd.notna(x) else ""


@st.cache_data(show_spinner=False)
def ler_planilha(conteudo: bytes):
    """Lê a planilha e devolve (meta_encontrada, {nome_bloco: (df, datas)})."""
    raw = pd.read_excel(io.BytesIO(conteudo), header=None, sheet_name=0)
    col0 = raw.iloc[:, 0].map(_norm).str.upper()

    meta = None
    for r in range(len(raw)):
        for c in range(raw.shape[1]):
            v = raw.iat[r, c]
            if isinstance(v, str) and len(v) < 30 and "META" in v.upper() and "CAMINH" in v.upper():
                for c2 in range(c + 1, raw.shape[1]):
                    n = pd.to_numeric(raw.iat[r, c2], errors="coerce")
                    if pd.notna(n):
                        meta = float(n)
                        break
            if meta:
                break
        if meta:
            break

    blocos = {}
    for h in [i for i, v in enumerate(col0) if v == "PLACA"]:
        datas_col = {c: pd.Timestamp(v) for c, v in raw.iloc[h].items()
                     if isinstance(v, (pd.Timestamp, datetime))}
        if not datas_col:
            continue
        nome, linhas = None, []
        for r in range(h + 1, len(raw)):
            rot = col0.iat[r]
            if rot.startswith("TOTAL"):
                nome = rot.replace("TOTAL", "").strip().title() or "Geral"
                break
            if rot == "PLACA":
                break
            if rot == "" and _norm(raw.iat[r, 1]) == "":
                continue
            linhas.append(r)
        nome = nome or f"Bloco {len(blocos) + 1}"

        regs = []
        for r in linhas:
            reg = {"Placa": col0.iat[r], "Motorista": _norm(raw.iat[r, 1]) or "(sem motorista)"}
            for c, d in datas_col.items():
                reg[d] = pd.to_numeric(raw.iat[r, c], errors="coerce")
            regs.append(reg)
        if regs:
            blocos[nome] = (pd.DataFrame(regs), sorted(datas_col.values()))
    return meta, blocos


# ------------------------------------------------------------------- cálculo
def calcular(df, datas, meta, dias_semana, metodo, limiar):
    d = df.copy()
    v = d[datas]
    d["Acumulado"] = v.sum(axis=1, skipna=True)
    d["Dias lançados"] = v.notna().sum(axis=1)

    idx = [i for i, dt in enumerate(datas) if v[dt].notna().any()]
    dias_decorridos = (max(idx) + 1) if idx else 0

    if metodo == "Dias lançados de cada motorista":
        base = d["Dias lançados"].astype(float)
    else:
        base = pd.Series(float(dias_decorridos), index=d.index)
    prev = (d["Acumulado"] / base.replace(0, np.nan) * dias_semana).fillna(0)
    d["Previsão"] = np.maximum(prev, d["Acumulado"])

    d["Meta"] = meta
    d["% Meta"] = d["Acumulado"] / meta * 100
    d["% Previsão"] = d["Previsão"] / meta * 100
    d["Falta / Excede"] = d["Acumulado"] - meta
    d["Status"] = np.select(
        [d["Acumulado"] >= meta, d["Previsão"] >= meta, d["Previsão"] >= limiar * meta],
        ["Meta batida", "No ritmo", "Atenção"], default="Crítico")
    d["Rótulo"] = d["Motorista"].where(d["Motorista"] != "(sem motorista)", d["Placa"])
    return d, dias_decorridos


# ------------------------------------------------------------------- gráficos
def _base(fig, titulo, altura, top=75, **kw):
    fig.update_layout(
        **BASE, height=altura,
        title=dict(text=f"<b>{titulo}</b>", x=0.01, xanchor="left", y=0.97, font=dict(size=14)),
        margin=dict(l=10, r=20, t=top, b=25), **kw)
    return fig


def grafico_ranking(d, altura):
    d = d.sort_values("% Meta")
    fig = go.Figure()
    fig.add_bar(
        y=d["Rótulo"], x=d["% Meta"], orientation="h", name="Acumulado", showlegend=False,
        marker_color=[CORES[s] for s in d["Status"]],
        text=[f"{p:.0f}%" for p in d["% Meta"]], textposition="outside", cliponaxis=False,
        customdata=np.stack([d["Acumulado"].map(brl), d["Previsão"].map(brl), d["Placa"]], axis=-1),
        hovertemplate="<b>%{y}</b> (%{customdata[2]})<br>Acumulado: %{customdata[0]}"
                      "<br>% da meta: %{x:.1f}%<br>Previsão sexta: %{customdata[1]}<extra></extra>")
    fig.add_scatter(
        y=d["Rótulo"], x=d["% Previsão"], mode="markers", name="Previsão até sexta",
        marker=dict(symbol="diamond", size=9, color="#1F2937", line=dict(width=1, color="white")),
        hovertemplate="Previsão: %{x:.0f}% da meta<extra></extra>")
    fig.add_vline(x=100, line_dash="dash", line_color="#6B7280",
                  annotation_text="Meta", annotation_position="top")
    _base(fig, "Acumulado × Meta por Motorista  (◆ previsão até sexta)", altura,
          xaxis=dict(title="% da meta semanal", gridcolor="#EEF1F6", zeroline=False,
                     range=[0, max(130, d["% Previsão"].max() * 1.12)]),
          yaxis=dict(showgrid=False, automargin=True), bargap=0.28,
          legend=dict(orientation="h", y=1.0, yanchor="bottom", x=0.0))
    return fig


def grafico_status(d, altura):
    cont = d["Status"].value_counts().reindex(list(CORES)).dropna()
    fig = go.Figure(go.Pie(
        labels=cont.index, values=cont.values, hole=0.58, sort=False,
        marker=dict(colors=[CORES[s] for s in cont.index], line=dict(color="white", width=2)),
        textinfo="value", textfont=dict(size=14, color="white"),
        hovertemplate="%{label}: %{value} caminhões (%{percent})<extra></extra>"))
    fig.add_annotation(text=f"<b style='font-size:28px;color:{NAVY}'>{len(d)}</b><br>caminhões",
                       showarrow=False, font=dict(size=13, color="#6B7280"))
    _base(fig, "Situação da Frota", altura, legend=dict(orientation="h", y=-0.02, x=0.5, xanchor="center"))
    return fig


def grafico_diario(d, datas, meta_total, dias_semana, altura, titulo="Carregamento Diário da Frota"):
    n = len(d)
    somas = [d[dt].sum() if d[dt].notna().any() else None for dt in datas]
    lanc = [int(d[dt].notna().sum()) for dt in datas]
    fig = go.Figure()
    fig.add_bar(
        x=[rotulo_data(dt) for dt in datas], y=somas, name="Carregado no dia",
        marker_color=[NAVY if q == n else "#8FB0EA" for q in lanc],
        text=[mil(s) if s is not None else "" for s in somas], textposition="outside",
        customdata=lanc,
        hovertemplate=f"%{{x}}<br>Carregado: R$ %{{y:,.2f}}<br>Lançamentos: %{{customdata}}/{n}"
                      "<extra></extra>")
    if meta_total:
        fig.add_hline(y=meta_total / dias_semana, line_dash="dash", line_color="#D64545",
                      annotation_text="Meta diária linear", annotation_position="top right",
                      annotation_font_color="#D64545")
    topo = max([s for s in somas if s] + [meta_total / dias_semana if meta_total else 0]) * 1.2
    _base(fig, titulo + "  (barra clara = dia parcial)", altura, showlegend=False,
          yaxis=dict(title="R$", gridcolor="#EEF1F6", range=[0, topo]), xaxis=dict(showgrid=False))
    return fig


def grafico_acumulado(d, datas, meta_total, dias_semana, dias_dec, previsao, altura):
    rot = [rotulo_data(dt) for dt in datas]
    diarios = [d[dt].sum() for dt in datas[:dias_dec]]
    acum = list(np.cumsum(diarios))
    trilha = [meta_total * min(i + 1, dias_semana) / dias_semana for i in range(len(datas))]
    fig = go.Figure()
    fig.add_scatter(x=rot, y=trilha, mode="lines", name="Trajetória da meta",
                    line=dict(color="#9CA3AF", dash="dash", width=2),
                    hovertemplate="Meta até aqui: R$ %{y:,.0f}<extra></extra>")
    fig.add_scatter(x=rot[:dias_dec], y=acum, mode="lines+markers", name="Realizado acumulado",
                    line=dict(color=NAVY, width=3), marker=dict(size=8),
                    fill="tozeroy", fillcolor="rgba(0,51,153,0.08)",
                    hovertemplate="Realizado: R$ %{y:,.0f}<extra></extra>")
    if 0 < dias_dec < len(datas):
        fig.add_scatter(x=[rot[dias_dec - 1], rot[-1]], y=[acum[-1], previsao], mode="lines+markers",
                        name="Previsão até sexta", line=dict(color="#2A6FDB", dash="dot", width=3),
                        marker=dict(size=[0, 10], symbol="diamond"),
                        hovertemplate="Previsão: R$ %{y:,.0f}<extra></extra>")
    _base(fig, "Acumulado da Semana × Trajetória da Meta", altura,
          yaxis=dict(title="R$", gridcolor="#EEF1F6", range=[0, max(previsao, meta_total) * 1.12]),
          xaxis=dict(showgrid=False), legend=dict(orientation="h", y=1.0, yanchor="bottom", x=0))
    return fig


def grafico_heatmap(d, datas):
    d = d.sort_values("Acumulado", ascending=False)
    z = d[datas].to_numpy(dtype=float)
    texto = [[mil(x) if pd.notna(x) else "" for x in linha] for linha in z]
    fig = go.Figure(go.Heatmap(
        z=z, x=[rotulo_data(dt) for dt in datas], y=d["Rótulo"], text=texto, texttemplate="%{text}",
        colorscale="Blues", zmin=0, hoverongaps=False, xgap=2, ygap=2,
        colorbar=dict(title="R$", thickness=14, len=0.9),
        hovertemplate="%{y}<br>%{x}: R$ %{z:,.2f}<extra></extra>"))
    _base(fig, "Carregamento por Motorista e Dia  (R$ mil · vazio = não lançado)",
          max(380, 30 * len(d) + 110), top=100,
          yaxis=dict(autorange="reversed", automargin=True, showgrid=False),
          xaxis=dict(side="top", showgrid=False))
    return fig


# ------------------------------------------------------------------- tabela
def tabela_estilizada(d, datas):
    cols_dias = {dt: rotulo_data(dt) for dt in datas}
    t = d.rename(columns=cols_dias)[
        ["Placa", "Motorista", *cols_dias.values(), "Acumulado", "% Meta", "Meta",
         "Falta / Excede", "Previsão", "% Previsão", "Status"]].reset_index(drop=True)
    status = list(t["Status"])
    t["Status"] = [f"{EMOJI[s]} {s}" for s in status]
    money = [*cols_dias.values(), "Acumulado", "Meta", "Falta / Excede", "Previsão"]
    sty = (t.style.format({**{c: brl for c in money}, "% Meta": "{:.1f}%", "% Previsão": "{:.1f}%"}, na_rep="–")
           .apply(lambda s: [f"background-color: {FUNDO[x]}; font-weight: 600" for x in status],
                  subset=["% Meta", "Status"]))
    return t, sty


def para_excel(t):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        t.to_excel(w, index=False, sheet_name="Carregamento")
    return buf.getvalue()


# --------------------------------------------------------------------- blocos
def secao(emoji, texto):
    st.markdown(f"<div class='sec'>{emoji} {texto}</div>", unsafe_allow_html=True)


def kpi(valor, rotulo, sub="", cor="#6B7280", barra=None):
    b = f"<div class='kpi-bar'><div style='width:{min(max(barra, 0), 100):.0f}%'></div></div>" if barra is not None else ""
    s = f"<div class='kpi-s' style='color:{cor}'>{sub}</div>" if sub else ""
    return (f"<div class='kpi'><div class='kpi-v'>{valor}</div><div class='kpi-l'>{rotulo}</div>{b}{s}</div>")


def cabecalho(pills=""):
    logo = next((p for p in (Path(__file__).parent / n for n in ("logo_bonsono.png", "logo.png")) if p.exists()), None)
    c1, c2 = st.columns([1, 4], vertical_alignment="center")
    if logo:
        with c1:
            st.image(str(logo), width=190)
    with c2:
        st.markdown(
            "<div class='titulo'>🚚 Dashboard de Carregamento x Meta</div>"
            "<div class='subtitulo'>Colchões BonSono | Acompanhamento Semanal de Carregamento por Motorista</div>"
            f"{pills}", unsafe_allow_html=True)


# ------------------------------------------------------------------------ app
def main():
    st.set_page_config(page_title="Carregamento x Meta | BonSono", page_icon="🚚", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### 📁 Upload de Dados")
        arq = st.file_uploader("Faça upload da planilha de carregamento (Excel)", type=["xlsx", "xlsm"])
        if arq is not None:
            st.success("Arquivo carregado com sucesso!")

    if arq is None:
        cabecalho()
        st.info("⬅️ Faça o upload da planilha na barra lateral para montar o dashboard.")
        return

    try:
        meta_arq, blocos = ler_planilha(arq.getvalue())
    except Exception as e:  # noqa: BLE001
        cabecalho()
        st.error(f"Não consegui ler a planilha: {e}")
        return
    if not blocos:
        cabecalho()
        st.error("Não encontrei nenhuma tabela com cabeçalho 'PLACA' e datas. Confira o formato da planilha.")
        return

    nomes = list(blocos)
    df_frota, datas = blocos[nomes[0]]

    with st.sidebar:
        st.markdown("### 🔍 Filtros")
        box_mot, box_sit = st.container(), st.container()
        st.markdown("### ⚙️ Parâmetros")
        meta = st.number_input("Meta semanal por caminhão (R$)", min_value=0.0, step=500.0,
                               value=float(meta_arq or 87500.0), format="%.2f")
        dias_semana = int(st.number_input("Dias úteis da semana", 1, 7, len(datas)))
        metodo = st.radio(
            "Previsão até sexta", ["Dias lançados de cada motorista", "Dias decorridos da semana (todos)"],
            help="• Dias lançados: acumulado ÷ nº de dias com lançamento do próprio motorista × dias úteis. "
                 "Célula em branco = ainda não lançado; 0 = lançado sem carga.\n\n"
                 "• Dias decorridos: acumulado ÷ último dia com dado na frota × dias úteis.")
        limiar = st.slider("Limite 'Atenção' (previsão ≥ % da meta)", 50, 100, 80, 5) / 100

    d_all, dias_dec = calcular(df_frota, datas, meta, dias_semana, metodo, limiar)
    with box_mot:
        sel_mot = st.multiselect("Motorista", sorted(d_all["Rótulo"]), placeholder="Todos")
    with box_sit:
        sel_sit = st.multiselect("Situação", list(CORES), placeholder="Todas")

    d = d_all[d_all["Rótulo"].isin(sel_mot or d_all["Rótulo"]) & d_all["Status"].isin(sel_sit or list(CORES))]

    ultimo = datas[dias_dec - 1] if dias_dec else None
    pills = f"<span class='pill'>📅 Semana de {datas[0]:%d/%m} a {datas[-1]:%d/%m/%Y}</span>"
    if ultimo is not None:
        pills += f"<span class='pill'>Dados até {rotulo_data(ultimo)}</span>"
        parcial = int(df_frota[ultimo].notna().sum())
        if parcial < len(df_frota):
            pills += (f"<span class='pill warn'>⚠️ {rotulo_data(ultimo)}: apenas {parcial} de "
                      f"{len(df_frota)} caminhões lançados</span>")
    cabecalho(pills)

    if d.empty:
        st.warning("Nenhum caminhão corresponde aos filtros selecionados.")
        return

    n = len(d)
    meta_total = meta * n
    realizado, previsao = d["Acumulado"].sum(), d["Previsão"].sum()
    falta = meta_total - realizado
    pct = realizado / meta_total * 100 if meta_total else 0
    dif = previsao - meta_total

    secao("📈", "Métricas Gerais")
    k = st.columns(6)
    cards = [
        kpi(brl(realizado), "Realizado Acumulado"),
        kpi(brl(meta_total), "Meta da Frota", f"{n} × {brl(meta, 0)}"),
        kpi(f"{pct:.1f}%", "Meta Atingida", barra=pct),
        kpi(brl(abs(falta)), "Falta para a Meta" if falta >= 0 else "Excedente"),
        kpi(brl(previsao), "Previsão até Sexta", f"{'▲ +' if dif >= 0 else '▼ '}{brl(dif)} vs meta",
            "#2E9E5B" if dif >= 0 else "#D64545"),
        kpi(f"{int((d['Status'] == 'Meta batida').sum())}/{n}", "Caminhões na Meta",
            f"{int((d['Status'] == 'No ritmo').sum())} a caminho"),
    ]
    for col, html in zip(k, cards):
        col.markdown(html, unsafe_allow_html=True)

    altura1 = max(430, 30 * n + 130)
    secao("🎯", "Atingimento da Meta por Motorista")
    c1, c2 = st.columns([3, 2])
    c1.plotly_chart(grafico_ranking(d, altura1), width="stretch", theme=None)
    c2.plotly_chart(grafico_status(d, altura1), width="stretch", theme=None)

    secao("📅", "Evolução Diária")
    c1, c2 = st.columns(2)
    c1.plotly_chart(grafico_diario(d, datas, meta_total, dias_semana, 400), width="stretch", theme=None)
    c2.plotly_chart(grafico_acumulado(d, datas, meta_total, dias_semana, dias_dec, previsao, 400),
                    width="stretch", theme=None)

    secao("🔥", "Mapa de Calor: Motorista × Dia")
    st.plotly_chart(grafico_heatmap(d, datas), width="stretch", theme=None)

    secao("📋", "Detalhamento por Motorista")
    ordenado = d.sort_values("Acumulado", ascending=False)
    t, sty = tabela_estilizada(ordenado, datas)
    st.dataframe(sty, width="stretch", hide_index=True)
    st.download_button("⬇️ Baixar tabela (Excel)", para_excel(t), "carregamento_vs_meta.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    for nome in nomes[1:]:
        dfb, datas_b = blocos[nome]
        tot = dfb[datas_b].sum(axis=1, skipna=True)
        secao("🏪", nome)
        c0, c1, c2 = st.columns([1, 2, 3])
        c0.markdown(kpi(brl(tot.sum()), f"Total {nome}"), unsafe_allow_html=True)
        c1.plotly_chart(grafico_diario(dfb, datas_b, 0, len(datas_b), 360, f"Carregamento Diário · {nome}"),
                        width="stretch", theme=None)
        vis = dfb.assign(Acumulado=tot).rename(columns={dt: rotulo_data(dt) for dt in datas_b})
        c2.dataframe(vis.style.format({c: brl for c in vis.columns[2:]}, na_rep="–"),
                     width="stretch", hide_index=True)


if __name__ == "__main__":
    main()