import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os

# ══════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Painel de Pendências",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ══════════════════════════════════════════════════════════════════════
# CSS PERSONALIZADO
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
    * { font-family: 'DM Sans', sans-serif; }
    .main { background-color: #0e1117; }
    .metric-card {
        background: linear-gradient(135deg, #1a1f2e 0%, #252b3b 100%);
        border-radius: 16px; padding: 24px;
        border: 1px solid rgba(255,255,255,0.06);
        text-align: center; transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-2px); }
    .metric-value { font-size: 2.2rem; font-weight: 700; margin: 4px 0; line-height: 1.1; }
    .metric-label { font-size: 0.8rem; color: #8892a4; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 600; }
    .color-orange { color: #ff8c42; }
    .color-red { color: #ff4d6a; }
    .color-blue { color: #4dabf7; }
    .color-green { color: #51cf66; }
    .color-purple { color: #b197fc; }
    .header-bar {
        background: linear-gradient(135deg, #1a5e2a 0%, #2e7d32 50%, #43a047 100%);
        padding: 28px 36px; border-radius: 20px; margin-bottom: 28px;
        border: 1px solid rgba(255,255,255,0.1);
    }
    .header-bar h1 { color: white; font-size: 1.8rem; font-weight: 700; margin: 0; }
    .header-bar p { color: rgba(255,255,255,0.75); font-size: 0.95rem; margin: 6px 0 0 0; }
    .section-title {
        font-size: 1.15rem; font-weight: 700; color: #e0e0e0;
        margin: 32px 0 16px 0; padding-bottom: 8px;
        border-bottom: 2px solid #2e7d32; display: inline-block;
    }
    div[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }
    .stSelectbox > div > div { border-radius: 10px !important; }
    .stTextArea > div > div { border-radius: 10px !important; }
    .stDateInput > div > div { border-radius: 10px !important; }
    div[data-testid="stSidebar"] { background: linear-gradient(180deg, #111827 0%, #1a1f2e 100%); }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# FUNÇÕES AUXILIARES
# ══════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════
# FUNÇÕES AUXILIARES
# ══════════════════════════════════════════════════════════════════════
import gspread

OPCOES_JUSTIFICATIVA = [
    "— Selecione —",
    "Aguardando NF do fornecedor",
    "NF com divergência de valores",
    "Pedido em análise de aprovação",
    "Fornecedor sem previsão de entrega",
    "Aguardando liberação financeira",
    "Produto em falta no fornecedor",
    "Erro no pedido - refazendo",
    "Em negociação de preço",
    "Documentação pendente",
    "Outro (detalhar na observação)"
]

SHEET_ID = "1XABBxLxziTZpMPCOzeoYWhsxDiS-F5sfZnTS3lraa-o"
COLS_JUST = ["ID", "Justificativa", "Observacao", "Prazo_Resolucao", "Data_Preenchimento", "Responsavel"]

@st.cache_resource
def get_gsheet_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    return gspread.service_account_from_dict(creds_dict, scopes=scopes)

def get_worksheet():
    client = get_gsheet_client()
    sh = client.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet("Página1")
    except:
        try:
            ws = sh.worksheet("Sheet1")
        except:
            ws = sh.sheet1
    # Garante cabeçalho na primeira vez
    try:
        first = ws.acell('A1').value
        if not first or first != 'ID':
            ws.clear()
            ws.append_row(COLS_JUST)
    except:
        pass
    return ws

def load_justificativas():
    try:
        ws = get_worksheet()
        data = ws.get_all_records()
        if data:
            return pd.DataFrame(data, dtype=str)
        return pd.DataFrame(columns=COLS_JUST)
    except Exception as e:
        st.warning(f"⚠️ Não foi possível carregar justificativas: {e}")
        return pd.DataFrame(columns=COLS_JUST)

def save_justificativa(row_id, justificativa, observacao, prazo, responsavel=""):
    try:
        ws = get_worksheet()
        data = ws.get_all_records()
        df_just = pd.DataFrame(data, dtype=str) if data else pd.DataFrame(columns=COLS_JUST)
        new_row = {
            "ID": str(row_id),
            "Justificativa": justificativa,
            "Observacao": observacao,
            "Prazo_Resolucao": str(prazo) if prazo else "",
            "Data_Preenchimento": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Responsavel": responsavel
        }
        if len(df_just) > 0 and str(row_id) in df_just['ID'].values:
            cell = ws.find(str(row_id))
            ws.update(f'A{cell.row}:F{cell.row}', [[new_row[c] for c in COLS_JUST]])
        else:
            ws.append_row([new_row[c] for c in COLS_JUST])
        return True
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {e}")
        return False



def parse_valor(v):
    if pd.isna(v) or str(v).strip() in ["", "—", "-"]:
        return 0.0
    s = str(v).replace("R$", "").replace("r$", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def load_data(file):
    df_raw = pd.read_excel(file, header=None, dtype=object)
    df_raw = df_raw.map(lambda x: str(x) if not (isinstance(x, float) and pd.isna(x)) else "")
    header_row = None
    for i in range(min(10, len(df_raw))):
        row_values = df_raw.iloc[i].tolist()
        row_text = " ".join([str(v).lower() for v in row_values if v is not None])
        if "fornecedor" in row_text and "comprador" in row_text:
            header_row = i
            break
        elif "fornecedor" in row_text or "comprador" in row_text:
            header_row = i
            break
    if header_row is not None:
        new_headers = df_raw.iloc[header_row].tolist()
        df = df_raw.iloc[header_row + 1:].reset_index(drop=True)
        df.columns = [str(c).strip() if c not in ("", "nan") else f"Col_{i}" for i, c in enumerate(new_headers)]
    else:
        df = pd.read_excel(file, dtype=str)
        df = df.map(lambda x: str(x) if not (isinstance(x, float) and pd.isna(x)) else "")
    df.columns = [str(c).strip() for c in df.columns]
    col_map = {}
    for c in df.columns:
        cl = c.strip().lower()
        if cl == "tipo": col_map[c] = "Tipo"
        elif "fornecedor" in cl: col_map[c] = "Fornecedor"
        elif "valor" in cl: col_map[c] = "Valor"
        elif "vencimento" in cl: col_map[c] = "Vencimento"
        elif "comprador" in cl: col_map[c] = "Comprador"
        elif "solicitante" in cl: col_map[c] = "Solicitante"
        elif "status sf" in cl: col_map[c] = "Status"
        elif cl == "status manual": col_map[c] = "Status Manual"
        elif "filial" in cl: col_map[c] = "Filial"
        elif cl == "dias": col_map[c] = "Dias"
        elif "pc" in cl and ("nº" in cl or "n°" in cl or "num" in cl or cl.startswith("nº")): col_map[c] = "Nº PC"
        elif cl == "nº pc" or cl == "n° pc": col_map[c] = "Nº PC"
        elif "controle" in cl: col_map[c] = "Controle"
        elif "entrega" in cl: col_map[c] = "Dt Entrega PC"
        elif "emiss" in cl: col_map[c] = "Dt Emissão"
        elif "chave" in cl: col_map[c] = "Chave Sefaz"
        elif "nota" in cl and "nº" in cl.lower(): col_map[c] = "Nº Nota"
    df = df.rename(columns=col_map)
    if "Valor" in df.columns:
        df["Valor"] = df["Valor"].apply(parse_valor)
    if "Vencimento" in df.columns:
        df["Vencimento"] = pd.to_datetime(df["Vencimento"], dayfirst=True, errors="coerce")
    if "Dias" in df.columns:
        df["Dias"] = pd.to_numeric(df["Dias"].replace("—", ""), errors="coerce").fillna(0).astype(int)
    key_cols = [c for c in ["Fornecedor", "Comprador", "Valor"] if c in df.columns]
    if key_cols:
        df = df.dropna(subset=key_cols, how="all").reset_index(drop=True)
    df["ID"] = df.index.astype(str)
    return df

def format_brl(valor):
    return f"R$ {valor:,.2f}".replace(",","X").replace(".",",").replace("X",".")

# ══════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁 Dados")
    uploaded = st.file_uploader("Carregar Excel de Pendências", type=["xlsx","xls","csv"])

    if uploaded:
        if uploaded.name.endswith('.csv'):
            df = pd.read_csv(uploaded, dtype=str)
        else:
            df = load_data(uploaded)
        st.session_state['df'] = df
        st.success(f"✅ {len(df)} registros carregados")

        with st.expander("🔧 Colunas detectadas"):
            for c in df.columns:
                sample = df[c].dropna().head(2).tolist()
                st.text(f"{c}: {sample}")

    if 'df' not in st.session_state:
        st.info("👆 Faça upload do arquivo Excel para começar")
        st.stop()

    df = st.session_state['df']

    st.markdown("---")
    st.markdown("### 🔍 Filtros")

    compradores = ["Todos"] + sorted([x for x in df['Comprador'].dropna().unique().tolist() if x not in ['—','']]) if 'Comprador' in df.columns else ["Todos"]
    sel_comprador = st.selectbox("Comprador", compradores)

    solicitantes = ["Todos"] + sorted([x for x in df['Solicitante'].dropna().unique().tolist() if x not in ['—','']]) if 'Solicitante' in df.columns else ["Todos"]
    sel_solicitante = st.selectbox("Solicitante", solicitantes)

    if 'Status' in df.columns:
        status_opts = ["Todos"] + sorted([x for x in df['Status'].dropna().unique().tolist() if x not in ['—','']])
        sel_status = st.selectbox("Status", status_opts)
    else:
        sel_status = "Todos"

    if 'Filial' in df.columns:
        filiais = ["Todas"] + sorted([x for x in df['Filial'].dropna().astype(str).unique().tolist() if x not in ['—','']])
        sel_filial = st.selectbox("Filial", filiais)
    else:
        sel_filial = "Todas"

    df_filtered = df.copy()
    if sel_comprador != "Todos":
        df_filtered = df_filtered[df_filtered['Comprador'] == sel_comprador]
    if sel_solicitante != "Todos":
        df_filtered = df_filtered[df_filtered['Solicitante'] == sel_solicitante]
    if sel_status != "Todos" and 'Status' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Status'] == sel_status]
    if sel_filial != "Todas" and 'Filial' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Filial'].astype(str) == sel_filial]

    st.markdown("---")
    st.markdown(f"**Exibindo:** {len(df_filtered)} de {len(df)} itens")

# ══════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════
hoje = datetime.now().strftime("%d/%m/%Y")
st.markdown(f"""
<div class="header-bar">
    <h1>📊 Painel de Pendências</h1>
    <p>Atualizado em {hoje} &nbsp;|&nbsp; {len(df)} itens totais &nbsp;|&nbsp; {len(df_filtered)} filtrados</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# KPIs
# ══════════════════════════════════════════════════════════════════════
total_valor = df_filtered['Valor'].sum() if 'Valor' in df_filtered.columns else 0
total_itens = len(df_filtered)

vencidas = 0
if 'Vencimento' in df_filtered.columns:
    vencidas = len(df_filtered[df_filtered['Vencimento'] < pd.Timestamp.now()])

just_df = load_justificativas()
com_justificativa = len(df_filtered[df_filtered['ID'].isin(just_df['ID'])])
sem_justificativa = total_itens - com_justificativa

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Total de Itens</div>
        <div class="metric-value color-blue">{total_itens}</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Valor Total</div>
        <div class="metric-value color-green">{format_brl(total_valor)}</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Vencidas</div>
        <div class="metric-value color-red">{vencidas}</div>
    </div>""", unsafe_allow_html=True)
with c4:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Com Justificativa</div>
        <div class="metric-value color-green">{com_justificativa}</div>
    </div>""", unsafe_allow_html=True)
with c5:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Sem Justificativa</div>
        <div class="metric-value color-orange">{sem_justificativa}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# GRÁFICOS
# ══════════════════════════════════════════════════════════════════════
COLORS = ["#4dabf7","#ff8c42","#51cf66","#ff4d6a","#b197fc","#ffd43b","#20c997","#f06595"]
PLOT_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="DM Sans", color="#c9d1d9"),
    margin=dict(l=20, r=20, t=40, b=20),
    legend=dict(bgcolor="rgba(0,0,0,0)")
)

st.markdown('<div class="section-title">📈 Visão por Comprador / Solicitante</div>', unsafe_allow_html=True)

# Linha 1: Qtd por Solicitante | Valor por Comprador
col_g1, col_g2 = st.columns(2)

with col_g1:
    if 'Solicitante' in df_filtered.columns:
        df_sol = df_filtered[df_filtered['Solicitante'].notna() & (~df_filtered['Solicitante'].isin(['—','']))].groupby('Solicitante').size().reset_index(name='Qtd').sort_values('Qtd', ascending=False)
        if len(df_sol) > 0:
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=df_sol['Solicitante'], y=df_sol['Qtd'],
                marker=dict(color=COLORS[4], cornerradius=6),
                text=df_sol['Qtd'], textposition='auto', textfont=dict(size=13, color="white")
            ))
            fig1.update_layout(**PLOT_LAYOUT, title="👤 Qtd Pendências por Solicitante", height=380,
                xaxis=dict(showgrid=False, tickangle=-30),
                yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.05)"))
            st.plotly_chart(fig1, use_container_width=True)

with col_g2:
    if 'Comprador' in df_filtered.columns and 'Valor' in df_filtered.columns:
        df_comp = df_filtered[df_filtered['Comprador'].notna() & (df_filtered['Comprador'] != '—')].groupby('Comprador').agg(
            Valor=('Valor','sum')
        ).reset_index().sort_values('Valor', ascending=True)
        if len(df_comp) > 0:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                y=df_comp['Comprador'], x=df_comp['Valor'], orientation='h',
                marker=dict(color=COLORS[2], cornerradius=6),
                text=df_comp['Valor'].apply(lambda v: format_brl(v)),
                textposition='auto', textfont=dict(size=11)
            ))
            fig2.update_layout(**PLOT_LAYOUT, title="💰 Valor Total por Comprador", height=380,
                xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.05)"),
                yaxis=dict(showgrid=False))
            st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════
# SEÇÃO: PROCESSOS VENCIDOS / ATRASADOS POR RESPONSÁVEL
# ══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🚨 Processos Vencidos / Atrasados por Responsável</div>', unsafe_allow_html=True)

if 'Vencimento' in df_filtered.columns:
    agora = pd.Timestamp.now().normalize()
    df_vencidos = df_filtered[df_filtered['Vencimento'] < agora].copy()
    df_vencidos['Dias_Atraso'] = (agora - df_vencidos['Vencimento']).dt.days

    if len(df_vencidos) == 0:
        st.success("✅ Nenhum processo vencido no filtro atual!")
    else:
        # KPIs de vencidos
        total_venc_valor = df_vencidos['Valor'].sum() if 'Valor' in df_vencidos.columns else 0
        maior_atraso = df_vencidos['Dias_Atraso'].max()
        media_atraso = df_vencidos['Dias_Atraso'].mean()

        kv1, kv2, kv3 = st.columns(3)
        with kv1:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">⚠️ Processos Atrasados</div>
                <div class="metric-value color-red">{len(df_vencidos)}</div>
            </div>""", unsafe_allow_html=True)
        with kv2:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">💸 Valor em Atraso</div>
                <div class="metric-value color-red">{format_brl(total_venc_valor)}</div>
            </div>""", unsafe_allow_html=True)
        with kv3:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">📆 Maior Atraso (dias)</div>
                <div class="metric-value color-orange">{int(maior_atraso)}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        col_vr1, col_vr2 = st.columns(2)

        # Gráfico: Qtd de vencidos por Comprador
        with col_vr1:
            if 'Comprador' in df_vencidos.columns:
                df_comp_venc = df_vencidos[
                    df_vencidos['Comprador'].notna() & (~df_vencidos['Comprador'].isin(['—','']))
                ].groupby('Comprador').agg(
                    Qtd=('Comprador','size'),
                    Valor=('Valor','sum'),
                    Atraso_Medio=('Dias_Atraso','mean')
                ).reset_index().sort_values('Qtd', ascending=False)

                if len(df_comp_venc) > 0:
                    fig_cv = go.Figure()
                    fig_cv.add_trace(go.Bar(
                        x=df_comp_venc['Comprador'],
                        y=df_comp_venc['Qtd'],
                        marker=dict(color='#ff4d6a', cornerradius=6),
                        text=df_comp_venc['Qtd'],
                        textposition='auto',
                        textfont=dict(size=13, color='white'),
                        customdata=df_comp_venc[['Valor','Atraso_Medio']],
                        hovertemplate=(
                            "<b>%{x}</b><br>"
                            "Qtd vencidos: %{y}<br>"
                            "Valor total: R$ %{customdata[0]:,.2f}<br>"
                            "Atraso médio: %{customdata[1]:.0f} dias<extra></extra>"
                        )
                    ))
                    fig_cv.update_layout(**PLOT_LAYOUT,
                        title="🔴 Processos Vencidos por Comprador",
                        height=380,
                        xaxis=dict(showgrid=False, tickangle=-30),
                        yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.05)", title="Qtd Vencidos"))
                    st.plotly_chart(fig_cv, use_container_width=True)

        # Gráfico: Qtd de vencidos por Solicitante
        with col_vr2:
            if 'Solicitante' in df_vencidos.columns:
                df_sol_venc = df_vencidos[
                    df_vencidos['Solicitante'].notna() & (~df_vencidos['Solicitante'].isin(['—','']))
                ].groupby('Solicitante').agg(
                    Qtd=('Solicitante','size'),
                    Valor=('Valor','sum'),
                    Atraso_Medio=('Dias_Atraso','mean')
                ).reset_index().sort_values('Qtd', ascending=True)

                if len(df_sol_venc) > 0:
                    fig_sv = go.Figure()
                    fig_sv.add_trace(go.Bar(
                        y=df_sol_venc['Solicitante'],
                        x=df_sol_venc['Qtd'],
                        orientation='h',
                        marker=dict(color='#ff8c42', cornerradius=6),
                        text=df_sol_venc['Qtd'],
                        textposition='auto',
                        textfont=dict(size=13, color='white'),
                        customdata=df_sol_venc[['Valor','Atraso_Medio']],
                        hovertemplate=(
                            "<b>%{y}</b><br>"
                            "Qtd vencidos: %{x}<br>"
                            "Valor total: R$ %{customdata[0]:,.2f}<br>"
                            "Atraso médio: %{customdata[1]:.0f} dias<extra></extra>"
                        )
                    ))
                    fig_sv.update_layout(**PLOT_LAYOUT,
                        title="🟠 Processos Vencidos por Solicitante",
                        height=380,
                        xaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.05)", title="Qtd Vencidos"),
                        yaxis=dict(showgrid=False))
                    st.plotly_chart(fig_sv, use_container_width=True)

        # Tabela de vencidos em evidência
        st.markdown('<div class="section-title">📋 Detalhe dos Processos Vencidos</div>', unsafe_allow_html=True)
        just_df_venc = load_justificativas()
        df_venc_display = df_vencidos.merge(just_df_venc, on='ID', how='left')
        cols_venc = [c for c in ['Comprador','Solicitante','Fornecedor','Valor','Vencimento','Dias_Atraso','Filial','Nº PC','Justificativa','Observacao'] if c in df_venc_display.columns]

        def highlight_atraso(row):
            dias = row.get('Dias_Atraso', 0)
            if dias > 30:
                return ['background-color: rgba(255,77,106,0.18)'] * len(row)
            elif dias > 7:
                return ['background-color: rgba(255,140,66,0.15)'] * len(row)
            else:
                return ['background-color: rgba(255,212,59,0.10)'] * len(row)

        styled_venc = df_venc_display[cols_venc].sort_values('Dias_Atraso', ascending=False).style\
            .apply(highlight_atraso, axis=1)\
            .format({'Valor': lambda x: format_brl(x) if pd.notna(x) and isinstance(x, (int,float)) else x})

        st.dataframe(styled_venc, use_container_width=True, height=350)

else:
    st.info("Coluna 'Vencimento' não encontrada nos dados.")

# ══════════════════════════════════════════════════════════════════════
# GRÁFICO DE BOLHAS: Comprador × Solicitante
# ══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🔵 Processos Pendentes: Comprador × Solicitante</div>', unsafe_allow_html=True)

if 'Comprador' in df_filtered.columns and 'Solicitante' in df_filtered.columns:
    df_cross = df_filtered[
        df_filtered['Comprador'].notna() & (~df_filtered['Comprador'].isin(['—',''])) &
        df_filtered['Solicitante'].notna() & (~df_filtered['Solicitante'].isin(['—','']))
    ].copy()

    if len(df_cross) > 0:
        df_grp = df_cross.groupby(['Comprador', 'Solicitante']).size().reset_index(name='Qtd')
        fig_bubble = go.Figure()
        for idx, sol in enumerate(df_grp['Solicitante'].unique()):
            df_s = df_grp[df_grp['Solicitante'] == sol]
            fig_bubble.add_trace(go.Scatter(
                x=df_s['Comprador'],
                y=[sol] * len(df_s),
                mode='markers+text',
                name=sol,
                marker=dict(
                    size=df_s['Qtd'] * 8,
                    color=COLORS[idx % len(COLORS)],
                    opacity=0.85,
                    line=dict(width=1, color='white')
                ),
                text=df_s['Qtd'],
                textposition='middle center',
                textfont=dict(size=12, color='white', family='DM Sans'),
            ))
        fig_bubble.update_layout(**PLOT_LAYOUT,
            title="🔵 Bolhas: Qtd Processos por Comprador × Solicitante",
            height=420,
            xaxis=dict(tickangle=-30, showgrid=True, gridcolor="rgba(255,255,255,0.06)"),
            yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.06)"))
        st.plotly_chart(fig_bubble, use_container_width=True)
    else:
        st.info("Sem dados suficientes de Comprador e Solicitante para gerar o cruzamento.")

# ══════════════════════════════════════════════════════════════════════
# TABELA + JUSTIFICATIVAS
# ══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📝 Pendências e Justificativas</div>', unsafe_allow_html=True)
just_df = load_justificativas()
df_display = df_filtered.merge(just_df, on='ID', how='left')

show_cols = [c for c in ['Fornecedor','Valor','Vencimento','Status','Comprador','Solicitante',
                         'Filial','Nº PC','Justificativa','Observacao','Prazo_Resolucao'] if c in df_display.columns]

st.dataframe(
    df_display[show_cols].style.format({
        'Valor': lambda x: format_brl(x) if pd.notna(x) and isinstance(x, (int,float)) else x,
    }),
    use_container_width=True, height=400
)

# ══════════════════════════════════════════════════════════════════════
# FORMULÁRIO DE JUSTIFICATIVA
# ══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">✏️ Preencher Justificativa</div>', unsafe_allow_html=True)
col_form1, col_form2 = st.columns([1, 2])

with col_form1:
    responsavel = st.text_input("Seu nome (quem está preenchendo)", placeholder="Ex: João Silva")
    st.session_state['responsavel'] = responsavel
    opcoes_pend = df_filtered.apply(
        lambda r: f"#{r['ID']} - {str(r.get('Fornecedor','?'))[:40]} - {format_brl(r.get('Valor',0))}",
        axis=1).tolist()
    sel_pendencia = st.selectbox("Selecione a Pendência", ["— Selecione —"] + opcoes_pend)
    sel_justificativa = st.selectbox("Motivo da Pendência", OPCOES_JUSTIFICATIVA)
    prazo_resolucao = st.date_input("Prazo previsto de resolução", value=None, min_value=date.today())

with col_form2:
    observacao = st.text_area("Observação adicional", height=130,
                              placeholder="Descreva detalhes adicionais sobre esta pendência...")
    if st.button("💾  Salvar Justificativa", type="primary", use_container_width=True):
        if sel_pendencia == "— Selecione —":
            st.error("⚠️ Selecione uma pendência!")
        elif sel_justificativa == "— Selecione —":
            st.error("⚠️ Selecione um motivo!")
        else:
            row_id = sel_pendencia.split(" - ")[0].replace("#","")
            save_justificativa(row_id, sel_justificativa, observacao, prazo_resolucao, responsavel=st.session_state.get('responsavel',''))
            st.success("✅ Justificativa salva com sucesso!")
            st.rerun()

# ══════════════════════════════════════════════════════════════════════
# EXPORTAR
# ══════════════════════════════════════════════════════════════════════
st.markdown("---")
col_exp1, col_exp2, _ = st.columns([1,1,2])
with col_exp1:
    csv_data = df_display[show_cols].to_csv(index=False).encode('utf-8-sig')
    st.download_button("📥 Exportar Dados (CSV)", csv_data, "pendencias_com_justificativas.csv", "text/csv")
with col_exp2:
    just_df_exp = load_justificativas()
    if len(just_df_exp) > 0:
        just_csv = just_df_exp.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Exportar Justificativas", just_csv, "justificativas.csv", "text/csv")
