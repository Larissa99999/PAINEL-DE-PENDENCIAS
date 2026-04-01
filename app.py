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
COLS_JUST = ["ID", "Nº_PC", "Fornecedor", "Comprador", "Solicitante", "Filial", "Valor", "Vencimento", "Dias_Atraso", "Justificativa", "Observacao", "Prazo_Resolucao", "Data_Preenchimento", "Responsavel", "Status_Resolucao"]

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

@st.cache_data(ttl=30)
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

def save_justificativa(row_id, justificativa, observacao, prazo, responsavel="", df_ref=None):
    try:
        ws = get_worksheet()
        data = ws.get_all_records()
        df_just = pd.DataFrame(data, dtype=str) if data else pd.DataFrame(columns=COLS_JUST)

        # Busca dados da pendência para enriquecer o histórico
        row_data = {}
        if df_ref is not None and len(df_ref) > 0:
            match = df_ref[df_ref['ID'] == str(row_id)]
            if len(match) > 0:
                r = match.iloc[0]
                agora = pd.Timestamp.now().normalize()
                venc = r.get('Vencimento', '')
                dias_atraso = int((agora - venc).days) if pd.notna(venc) and isinstance(venc, pd.Timestamp) else 0
                row_data = {
                    "Nº_PC": str(r.get('Nº PC', '')),
                    "Fornecedor": str(r.get('Fornecedor', '')),
                    "Comprador": str(r.get('Comprador', '')),
                    "Solicitante": str(r.get('Solicitante', '')),
                    "Filial": str(r.get('Filial', '')),
                    "Valor": str(r.get('Valor', '')),
                    "Vencimento": str(venc)[:10] if venc else '',
                    "Dias_Atraso": str(dias_atraso),
                }

        new_row = {
            "ID": str(row_id),
            "Nº_PC": row_data.get('Nº_PC', ''),
            "Fornecedor": row_data.get('Fornecedor', ''),
            "Comprador": row_data.get('Comprador', ''),
            "Solicitante": row_data.get('Solicitante', ''),
            "Filial": row_data.get('Filial', ''),
            "Valor": row_data.get('Valor', ''),
            "Vencimento": row_data.get('Vencimento', ''),
            "Dias_Atraso": row_data.get('Dias_Atraso', ''),
            "Justificativa": justificativa,
            "Observacao": observacao,
            "Prazo_Resolucao": str(prazo) if prazo else "",
            "Data_Preenchimento": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Responsavel": responsavel,
            "Status_Resolucao": "Pendente"
        }
        if len(df_just) > 0 and str(row_id) in df_just['ID'].values:
            cell = ws.find(str(row_id))
            ws.update(f'A{cell.row}:O{cell.row}', [[new_row[c] for c in COLS_JUST]])
        else:
            ws.append_row([new_row[c] for c in COLS_JUST])
        load_justificativas.clear()
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
EXCEL_PATH = "PENDENCIAS TOTVS BIOFLOR.xlsx"

with st.sidebar:
    st.markdown("### 📁 Dados")

    # Carrega automaticamente do repositório
    if 'df' not in st.session_state:
        try:
            df = load_data(EXCEL_PATH)
            st.session_state['df'] = df
        except Exception as e:
            st.error(f"Erro ao carregar planilha: {e}")
            st.stop()

    # Permite também upload manual para atualizar
    uploaded = st.file_uploader("🔄 Atualizar Excel (opcional)", type=["xlsx","xls","csv"])
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
        st.info("Carregando dados...")
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

    # Filtro por vencimento
    st.markdown("#### 📅 Vencimento")
    if 'Vencimento' in df.columns:
        venc_min = df['Vencimento'].min()
        venc_max = df['Vencimento'].max()
        if pd.notna(venc_min) and pd.notna(venc_max):
            data_de = st.date_input("De", value=venc_min.date(), key="venc_de")
            data_ate = st.date_input("Até", value=venc_max.date(), key="venc_ate")
        else:
            data_de = None
            data_ate = None
    else:
        data_de = None
        data_ate = None

    df_filtered = df.copy()
    if sel_comprador != "Todos":
        df_filtered = df_filtered[df_filtered['Comprador'] == sel_comprador]
    if sel_solicitante != "Todos":
        df_filtered = df_filtered[df_filtered['Solicitante'] == sel_solicitante]
    if sel_status != "Todos" and 'Status' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Status'] == sel_status]
    if sel_filial != "Todas" and 'Filial' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Filial'].astype(str) == sel_filial]
    if data_de and data_ate and 'Vencimento' in df_filtered.columns:
        df_filtered = df_filtered[
            (df_filtered['Vencimento'].dt.date >= data_de) &
            (df_filtered['Vencimento'].dt.date <= data_ate)
        ]

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
agora_now = pd.Timestamp.now()
total_valor = df_filtered['Valor'].sum() if 'Valor' in df_filtered.columns else 0
total_itens = len(df_filtered)

vencidas = 0
valor_vencido = 0
maior_atraso_kpi = 0
if 'Vencimento' in df_filtered.columns:
    df_venc_kpi = df_filtered[df_filtered['Vencimento'] < agora_now]
    vencidas = len(df_venc_kpi)
    valor_vencido = df_venc_kpi['Valor'].sum() if 'Valor' in df_venc_kpi.columns else 0
    if vencidas > 0:
        maior_atraso_kpi = int((agora_now.normalize() - df_venc_kpi['Vencimento']).dt.days.max())

entrega_enc = 0
if 'Dt Entrega PC' in df_filtered.columns:
    df_filtered['Dt Entrega PC'] = pd.to_datetime(df_filtered['Dt Entrega PC'], dayfirst=True, errors='coerce')
    entrega_enc = len(df_filtered[df_filtered['Dt Entrega PC'] < agora_now])

just_df = load_justificativas()
com_justificativa = len(df_filtered[df_filtered['ID'].isin(just_df['ID'])])
sem_justificativa = total_itens - com_justificativa
pct_just = int(com_justificativa / total_itens * 100) if total_itens > 0 else 0
pct_vencidos = int(vencidas / total_itens * 100) if total_itens > 0 else 0

# ── Linha 1: os 3 mais críticos em destaque ──
st.markdown("""<style>
.big-card {
    background: linear-gradient(135deg, #1a1f2e 0%, #252b3b 100%);
    border-radius: 16px; padding: 28px 20px;
    border: 1px solid rgba(255,255,255,0.06);
    text-align: center;
}
.big-card .label { font-size: 0.75rem; color: #8892a4; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 600; }
.big-card .value { font-size: 2.6rem; font-weight: 700; line-height: 1.15; margin: 6px 0 2px 0; }
.big-card .sub { font-size: 0.78rem; color: #8892a4; }
.big-card.critical { border: 1.5px solid rgba(255,77,106,0.4); background: linear-gradient(135deg, #2a1520 0%, #2e1a26 100%); }
.big-card.warning { border: 1.5px solid rgba(177,151,252,0.35); background: linear-gradient(135deg, #1e1a2e 0%, #231f35 100%); }
</style>""", unsafe_allow_html=True)

r1c1, r1c2, r1c3 = st.columns(3)
with r1c1:
    st.markdown(f"""<div class="big-card critical">
        <div class="label">🚨 Processos Vencidos</div>
        <div class="value" style="color:#ff4d6a">{vencidas} <span style="font-size:1.1rem">({pct_vencidos}%)</span></div>
        <div class="sub">Valor: {format_brl(valor_vencido)} &nbsp;|&nbsp; Maior atraso: {maior_atraso_kpi} dias</div>
    </div>""", unsafe_allow_html=True)
with r1c2:
    st.markdown(f"""<div class="big-card warning">
        <div class="label">📦 Prazo de Entrega Encerrado</div>
        <div class="value" style="color:#b197fc">{entrega_enc}</div>
        <div class="sub">de {total_itens} pendências no filtro atual</div>
    </div>""", unsafe_allow_html=True)
with r1c3:
    cor_just = '#ff4d6a' if pct_just < 30 else '#ff8c42' if pct_just < 60 else '#51cf66'
    st.markdown(f"""<div class="big-card">
        <div class="label">✅ Justificativas Preenchidas</div>
        <div class="value" style="color:{cor_just}">{pct_just}%</div>
        <div class="sub">{com_justificativa} de {total_itens} &nbsp;|&nbsp; {sem_justificativa} sem justificativa</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Linha 2: resumo geral ──
r2c1, r2c2, r2c3, r2c4 = st.columns(4)
with r2c1:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Total de Pendências</div>
        <div class="metric-value color-blue">{total_itens}</div>
    </div>""", unsafe_allow_html=True)
with r2c2:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Valor Total</div>
        <div class="metric-value color-green">{format_brl(total_valor)}</div>
    </div>""", unsafe_allow_html=True)
with r2c3:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Valor Vencido</div>
        <div class="metric-value color-red">{format_brl(valor_vencido)}</div>
    </div>""", unsafe_allow_html=True)
with r2c4:
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

agora_kpi = pd.Timestamp.now().normalize()

col_g1, col_g2 = st.columns(2)

with col_g1:
    if 'Solicitante' in df_filtered.columns and 'Vencimento' in df_filtered.columns:
        base = df_filtered[df_filtered['Solicitante'].notna() & (~df_filtered['Solicitante'].isin(['—','']))]
        df_sol = base.groupby('Solicitante').size().reset_index(name='Qtd')
        df_sol_venc = base[base['Vencimento'] < agora_kpi].groupby('Solicitante').size().reset_index(name='Vencidos')
        df_sol = df_sol.merge(df_sol_venc, on='Solicitante', how='left').fillna(0)
        df_sol['Vencidos'] = df_sol['Vencidos'].astype(int)
        df_sol['Pct_Vencido'] = (df_sol['Vencidos'] / df_sol['Qtd'] * 100).round(0).astype(int)
        if 'Dt Entrega PC' in df_filtered.columns:
            df_sol_ent = base[base['Dt Entrega PC'] < agora_kpi].groupby('Solicitante').size().reset_index(name='Entrega_Enc')
            df_sol = df_sol.merge(df_sol_ent, on='Solicitante', how='left').fillna(0)
            df_sol['Entrega_Enc'] = df_sol['Entrega_Enc'].astype(int)
        else:
            df_sol['Entrega_Enc'] = 0
        df_sol['Pct_just'] = 0
        if len(just_df) > 0:
            ids_just = set(just_df['ID'].values)
            df_sol['Com_Just'] = base.groupby('Solicitante').apply(lambda x: x['ID'].isin(ids_just).sum()).values
            df_sol['Pct_just'] = (df_sol['Com_Just'] / df_sol['Qtd'] * 100).round(0).astype(int)
        df_sol = df_sol.sort_values('Qtd', ascending=False)
        if len(df_sol) > 0:
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=df_sol['Solicitante'], y=df_sol['Entrega_Enc'],
                name='Prazo Entrega Enc. 📦',
                marker=dict(color='#b197fc'),
                textfont=dict(size=11, color='white')
            ))
            fig1.add_trace(go.Bar(
                x=df_sol['Solicitante'], y=df_sol['Vencidos'],
                name='Vencidos 🚨',
                marker=dict(color='#ff4d6a'),
                text=[f"{v} ({p}%)" for v, p in zip(df_sol['Vencidos'], df_sol['Pct_Vencido'])],
                textposition='inside',
                textfont=dict(size=11, color='white', family='DM Sans')
            ))
            fig1.update_layout(**PLOT_LAYOUT)
            fig1.update_layout(
                barmode='stack',
                title="👤 Por Solicitante — Entrega Enc. + Vencidos",
                height=400,
                xaxis=dict(showgrid=False, tickangle=-30),
                yaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.05)'),
                legend=dict(orientation='h', y=1.1)
            )
            st.plotly_chart(fig1, use_container_width=True)

with col_g2:
    if 'Comprador' in df_filtered.columns and 'Valor' in df_filtered.columns and 'Vencimento' in df_filtered.columns:
        base2 = df_filtered[df_filtered['Comprador'].notna() & (~df_filtered['Comprador'].isin(['—','']))]
        df_comp = base2.groupby('Comprador').agg(Valor=('Valor','sum')).reset_index()
        df_comp_venc = base2[base2['Vencimento'] < agora_kpi].groupby('Comprador').agg(Valor_Vencido=('Valor','sum')).reset_index()
        df_comp = df_comp.merge(df_comp_venc, on='Comprador', how='left').fillna(0)
        if 'Dt Entrega PC' in df_filtered.columns:
            df_comp_ent = base2[base2['Dt Entrega PC'] < agora_kpi].groupby('Comprador').agg(Valor_Entrega=('Valor','sum')).reset_index()
            df_comp = df_comp.merge(df_comp_ent, on='Comprador', how='left').fillna(0)
        else:
            df_comp['Valor_Entrega'] = 0
        df_comp['Pct_Vencido'] = (df_comp['Valor_Vencido'] / df_comp['Valor'] * 100).round(0).astype(int)
        df_comp = df_comp.sort_values('Valor', ascending=True)
        if len(df_comp) > 0:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                y=df_comp['Comprador'], x=df_comp['Valor_Entrega'],
                name='Prazo Entrega Enc. 📦',
                orientation='h',
                marker=dict(color='#b197fc'),
                textfont=dict(size=11)
            ))
            fig2.add_trace(go.Bar(
                y=df_comp['Comprador'], x=df_comp['Valor_Vencido'],
                name='Vencido 🚨',
                orientation='h',
                marker=dict(color='#ff4d6a'),
                text=[f"{p}% — {format_brl(v)}" for v, p in zip(df_comp['Valor_Vencido'], df_comp['Pct_Vencido'])],
                textposition='inside',
                textfont=dict(size=10, color='white')
            ))
            fig2.update_layout(**PLOT_LAYOUT)
            fig2.update_layout(
                barmode='stack',
                title="💰 Valor por Comprador — Entrega Enc. + Vencido",
                height=400,
                xaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.05)'),
                yaxis=dict(showgrid=False),
                legend=dict(orientation='h', y=1.1)
            )
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
# TABELA + JUSTIFICATIVAS
# ══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📝 Pendências e Justificativas</div>', unsafe_allow_html=True)
just_df = load_justificativas()
df_display = df_filtered.merge(just_df, on='ID', how='left')

show_cols = [c for c in ['Fornecedor','Valor','Vencimento','Status','Comprador','Solicitante',
                         'Filial','Nº PC','Nº Nota','Dt Emissão','Chave Sefaz','Justificativa','Observacao','Prazo_Resolucao'] if c in df_display.columns]

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
            save_justificativa(row_id, sel_justificativa, observacao, prazo_resolucao, responsavel=st.session_state.get('responsavel',''), df_ref=df_filtered)
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
