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
    "Divergência entre NF x PC",
    "Pedido em análise de aprovação",
    "Fornecedor sem previsão de entrega",
    "Eliminar resíduo",
    "Outro (detalhar na observação)"
]

SHEET_ID = "1XABBxLxziTZpMPCOzeoYWhsxDiS-F5sfZnTS3lraa-o"
COLS_JUST = ["ID", "Comprador", "Solicitante", "Fornecedor", "Filial", "Nº_PC", "Nº_Nota", "Dt_Entrega", "Vencimento", "Valor", "Dias_Atraso", "Justificativa", "Prazo_Resolucao", "Observacao", "Responsavel", "Data_Preenchimento", "Status_Resolucao"]

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
    try:
        first = ws.acell('A1').value
        if not first or first != 'ID':
            ws.clear()
            ws.append_row(COLS_JUST)
    except:
        pass
    return ws

def load_sheet_safe(ws):
    """Load records handling duplicate headers safely."""
    try:
        data = ws.get_all_records(expected_headers=COLS_JUST)
        return data
    except Exception:
        # Fallback: read raw and use first row as header
        try:
            all_vals = ws.get_all_values()
            if len(all_vals) <= 1:
                return []
            headers = all_vals[0]
            rows = []
            for row in all_vals[1:]:
                d = {}
                for i, h in enumerate(COLS_JUST):
                    d[h] = row[i] if i < len(row) else ''
                rows.append(d)
            return rows
        except:
            return []

@st.cache_data(ttl=30)
@st.cache_data(ttl=30)
def load_justificativas():
    try:
        ws = get_worksheet()
        data = load_sheet_safe(ws)
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
                    "Comprador": str(r.get('Comprador', '')),
                    "Solicitante": str(r.get('Solicitante', '')),
                    "Fornecedor": str(r.get('Fornecedor', '')),
                    "Filial": str(r.get('Filial', '')),
                    "Nº_PC": str(r.get('Nº PC', '')),
                    "Nº_Nota": str(r.get('Nº Nota', '')),
                    "Dt_Entrega": str(r.get('Dt Entrega PC', ''))[:10] if r.get('Dt Entrega PC', '') else '',
                    "Vencimento": str(venc)[:10] if venc else '',
                    "Valor": str(r.get('Valor', '')),
                    "Dias_Atraso": str(dias_atraso),
                }

        new_row = {
            "ID": str(row_id),
            "Comprador": row_data.get('Comprador', ''),
            "Solicitante": row_data.get('Solicitante', ''),
            "Fornecedor": row_data.get('Fornecedor', ''),
            "Filial": row_data.get('Filial', ''),
            "Nº_PC": row_data.get('Nº_PC', ''),
            "Nº_Nota": row_data.get('Nº_Nota', ''),
            "Dt_Entrega": row_data.get('Dt_Entrega', ''),
            "Vencimento": row_data.get('Vencimento', ''),
            "Valor": row_data.get('Valor', ''),
            "Dias_Atraso": row_data.get('Dias_Atraso', ''),
            "Justificativa": justificativa,
            "Prazo_Resolucao": str(prazo) if prazo else "",
            "Observacao": observacao,
            "Responsavel": responsavel,
            "Data_Preenchimento": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Status_Resolucao": "Pendente"
        }
        # Use Nº_PC as match key if available, otherwise ID
        match_key = new_row.get('Nº_PC', '') or str(row_id)
        matched = False
        if len(df_just) > 0 and 'Nº_PC' in df_just.columns:
            match_rows = df_just[df_just['Nº_PC'] == match_key]
            if len(match_rows) > 0:
                idx_row = match_rows.index[0] + 2  # +2 for header and 1-based
                ws.update(f'A{idx_row}:Q{idx_row}', [[new_row[c] for c in COLS_JUST]])
                matched = True
        if not matched:
            ws.append_row([new_row[c] for c in COLS_JUST])
        load_justificativas.clear()
        return True
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {e}")
        return False



def parse_valor(v):
    if pd.isna(v) or str(v).strip() in ["", "—", "-", "nan"]:
        return 0.0
    s = str(v).replace("R$", "").replace("r$", "").replace("R$","").strip()
    s = s.replace(" ", "")
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
        # Clean 'VENCIDA Xd' prefix - extract just the date part
        df["Vencimento"] = df["Vencimento"].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')[0].fillna(df["Vencimento"])
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
EXCEL_PATH = "PENDENCIAS_TOTVS_BIOFLOR.xlsx"  # Update if filename changes

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
    
    with st.expander("🔧 Debug colunas detectadas"):
        _df_tmp = st.session_state['df']
        st.write("Colunas:", list(_df_tmp.columns))
        if 'Solicitante' in _df_tmp.columns:
            st.write("Solicitantes únicos:", _df_tmp['Solicitante'].dropna().unique().tolist()[:10])
        if 'Comprador' in _df_tmp.columns:
            st.write("Compradores únicos:", _df_tmp['Comprador'].dropna().unique().tolist()[:10])

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

    if 'Controle' in df.columns:
        sel_aprovacao = st.selectbox("Aprovação", ["Todos", "B — Em aprovação", "L — Aprovado"])
    else:
        sel_aprovacao = "Todos"

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

    # Filtro por data de emissão
    st.markdown("#### 🧾 Dt Emissão")
    if 'Dt Emissão' in df.columns:
        df['Dt Emissão'] = pd.to_datetime(df['Dt Emissão'], dayfirst=True, errors='coerce')
        emiss_min = df['Dt Emissão'].min()
        emiss_max = df['Dt Emissão'].max()
        if pd.notna(emiss_min) and pd.notna(emiss_max):
            emiss_de = st.date_input("De", value=emiss_min.date(), key="emiss_de")
            emiss_ate = st.date_input("Até", value=emiss_max.date(), key="emiss_ate")
        else:
            emiss_de = None
            emiss_ate = None
    else:
        emiss_de = None
        emiss_ate = None

    df_filtered = df.copy()
    if sel_comprador != "Todos":
        df_filtered = df_filtered[df_filtered['Comprador'] == sel_comprador]
    if sel_solicitante != "Todos":
        df_filtered = df_filtered[df_filtered['Solicitante'] == sel_solicitante]
    if sel_status != "Todos" and 'Status' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Status'] == sel_status]
    if sel_filial != "Todas" and 'Filial' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Filial'].astype(str) == sel_filial]
    if sel_aprovacao != "Todos" and 'Controle' in df_filtered.columns:
        cod = sel_aprovacao.split(" ")[0]
        df_filtered = df_filtered[df_filtered['Controle'].astype(str).str.upper().str.startswith(cod)]
    if data_de and data_ate and 'Vencimento' in df_filtered.columns:
        df_filtered = df_filtered[
            (df_filtered['Vencimento'].dt.date >= data_de) &
            (df_filtered['Vencimento'].dt.date <= data_ate)
        ]
    if emiss_de and emiss_ate and 'Dt Emissão' in df_filtered.columns:
        df_filtered['Dt Emissão'] = pd.to_datetime(df_filtered['Dt Emissão'], dayfirst=True, errors='coerce')
        df_filtered = df_filtered[
            (df_filtered['Dt Emissão'].dt.date >= emiss_de) &
            (df_filtered['Dt Emissão'].dt.date <= emiss_ate)
        ]

    st.markdown("---")
    st.markdown(f"**Exibindo:** {len(df_filtered)} de {len(df)} itens")
    if st.button("🔄 Atualizar Justificativas"):
        load_justificativas.clear()
        st.rerun()

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
if len(just_df) > 0 and 'Nº_PC' in just_df.columns and 'Nº PC' in df_filtered.columns:
    pcs_com_just = set(just_df['Nº_PC'].astype(str).values)
    com_justificativa = len(df_filtered[df_filtered['Nº PC'].astype(str).isin(pcs_com_just)])
else:
    com_justificativa = 0
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
        <div class="value" style="color:{cor_just}">{com_justificativa} <span style="font-size:1rem;color:#8892a4">de {total_itens}</span></div>
        <div class="sub">{pct_just}% concluído &nbsp;|&nbsp; {sem_justificativa} pendentes</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Linha 2: resumo geral ──
em_aprovacao = 0
aprovados = 0
if 'Controle' in df_filtered.columns:
    em_aprovacao = len(df_filtered[df_filtered['Controle'].astype(str).str.upper().str.startswith('B')])
    aprovados = len(df_filtered[df_filtered['Controle'].astype(str).str.upper().str.startswith('L')])

r2c1, r2c2, r2c3, r2c4, r2c5, r2c6 = st.columns(6)
with r2c1:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Total de Pendências</div>
        <div class="metric-value color-blue">{total_itens}</div>
    </div>""", unsafe_allow_html=True)
with r2c2:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Valor Total</div>
        <div class="metric-value color-green" style="font-size:1.4rem">{format_brl(total_valor)}</div>
    </div>""", unsafe_allow_html=True)
with r2c3:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Valor Vencido</div>
        <div class="metric-value color-red" style="font-size:1.4rem">{format_brl(valor_vencido)}</div>
    </div>""", unsafe_allow_html=True)
with r2c4:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">Sem Justificativa</div>
        <div class="metric-value color-orange">{sem_justificativa}</div>
    </div>""", unsafe_allow_html=True)
with r2c5:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">⏳ Em Aprovação (B)</div>
        <div class="metric-value color-orange">{em_aprovacao}</div>
    </div>""", unsafe_allow_html=True)
with r2c6:
    st.markdown(f"""<div class="metric-card">
        <div class="metric-label">✅ Aprovados (L)</div>
        <div class="metric-value color-green">{aprovados}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

COLORS = ["#4dabf7","#ff8c42","#51cf66","#ff4d6a","#b197fc","#ffd43b","#20c997","#f06595"]
PLOT_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="DM Sans", color="#c9d1d9"),
    margin=dict(l=20, r=20, t=40, b=20),
    legend=dict(bgcolor="rgba(0,0,0,0)")
)

# ══════════════════════════════════════════════════════════════════════
# KPIs + PAINEL: NOTAS PENDENTES DE LANÇAMENTO (SF1)
# ══════════════════════════════════════════════════════════════════════
if 'Status' in df_filtered.columns:
    df_sf1 = df_filtered[df_filtered['Status'].notna() & (df_filtered['Status'].astype(str) != '—')].copy()
    tem_pc = df_sf1['Nº PC'].notna() & (~df_sf1['Nº PC'].isin(['', '—', 'nan'])) if 'Nº PC' in df_sf1.columns else pd.Series([False]*len(df_sf1))
    df_com_pc  = df_sf1[tem_pc]
    df_sem_pc  = df_sf1[~tem_pc]
    qtd_sf1    = len(df_sf1)
    qtd_com_pc = len(df_com_pc)
    qtd_sem_pc = len(df_sem_pc)
    valor_sf1    = df_sf1['Valor'].sum() if 'Valor' in df_sf1.columns else 0
    valor_sem_pc = df_sem_pc['Valor'].sum() if 'Valor' in df_sem_pc.columns else 0

    if qtd_sf1 > 0:
        st.markdown('<div class="section-title">⚠️ Notas Pendentes de Lançamento (SF1)</div>', unsafe_allow_html=True)

        # Tabela 1: Sem PC — mais crítica
        st.markdown('<div class="section-title">🚨 Notas Sem PC e Sem Lançamento — Crítico</div>', unsafe_allow_html=True)
        cols_sem_pc = [c for c in ['Comprador','Solicitante','Fornecedor','Filial','Nº PC','Nº Nota','Controle','Dt Emissão','Dt Entrega PC','Vencimento','Valor','Justificativa','Prazo_Resolucao','Responsavel'] if c in df_sem_pc.columns]
        if len(df_sem_pc) > 0 and len(cols_sem_pc) > 0:
            st.dataframe(
                df_sem_pc[cols_sem_pc].sort_values('Valor', ascending=False).style
                    .set_properties(**{'background-color': 'rgba(255,77,106,0.12)'})
                    .format({'Valor': lambda x: format_brl(x) if pd.notna(x) and isinstance(x, (int,float)) else x}),
                use_container_width=True, height=350
            )
        else:
            st.success('✅ Nenhuma nota sem PC encontrada!')

        # Tabela 2: Com PC mas sem lançamento
        st.markdown('<div class="section-title">📋 Notas Com PC e Sem Lançamento — Atenção</div>', unsafe_allow_html=True)
        cols_com_pc = [c for c in ['Comprador','Solicitante','Fornecedor','Filial','Nº PC','Nº Nota','Controle','Dt Emissão','Dt Entrega PC','Vencimento','Valor','Justificativa','Prazo_Resolucao','Responsavel'] if c in df_com_pc.columns]
        if len(df_com_pc) > 0 and len(cols_com_pc) > 0:
            st.dataframe(
                df_com_pc[cols_com_pc].sort_values('Valor', ascending=False).style
                    .set_properties(**{'background-color': 'rgba(177,151,252,0.08)'})
                    .format({'Valor': lambda x: format_brl(x) if pd.notna(x) and isinstance(x, (int,float)) else x}),
                use_container_width=True, height=350
            )
        else:
            st.success('✅ Nenhuma nota com PC pendente encontrada!')

st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# GRÁFICOS
# ══════════════════════════════════════════════════════════════════════


st.markdown('<div class="section-title">📈 Visão por Comprador / Solicitante</div>', unsafe_allow_html=True)

agora_kpi = pd.Timestamp.now().normalize()

col_g1, col_g2 = st.columns(2)

with col_g1:
    if 'Solicitante' in df_filtered.columns:
        df_filtered['Solicitante'] = df_filtered['Solicitante'].astype(str).str.strip()
        base = df_filtered[df_filtered['Solicitante'].notna() & (~df_filtered['Solicitante'].isin(['—','','nan']))]
        df_sol = base.groupby('Solicitante').size().reset_index(name='Qtd')
        # Vencidos
        if 'Vencimento' in df_filtered.columns:
            df_sol_venc = base[base['Vencimento'] < agora_kpi].groupby('Solicitante').size().reset_index(name='Vencidos')
            df_sol = df_sol.merge(df_sol_venc, on='Solicitante', how='left').fillna(0)
        else:
            df_sol['Vencidos'] = 0
        df_sol['Vencidos'] = df_sol['Vencidos'].astype(int)
        df_sol['Pct_Vencido'] = (df_sol['Vencidos'] / df_sol['Qtd'] * 100).round(0).astype(int)
        # Entrega encerrada
        if 'Dt Entrega PC' in df_filtered.columns:
            df_sol_ent = base[base['Dt Entrega PC'] < agora_kpi].groupby('Solicitante').size().reset_index(name='Entrega_Enc')
            df_sol = df_sol.merge(df_sol_ent, on='Solicitante', how='left').fillna(0)
            df_sol['Entrega_Enc'] = df_sol['Entrega_Enc'].astype(int)
        else:
            df_sol['Entrega_Enc'] = 0
        # Justificativas por Nº PC
        df_sol['Com_Just_Qtd'] = 0
        if len(just_df) > 0 and 'Nº_PC' in just_df.columns and 'Nº PC' in base.columns:
            pcs_just = set(just_df['Nº_PC'].values)
            for idx2, row2 in df_sol.iterrows():
                sol_name = row2['Solicitante']
                sol_pcs = base[base['Solicitante'] == sol_name]['Nº PC'].astype(str).values
                df_sol.at[idx2, 'Com_Just_Qtd'] = sum(1 for pc in sol_pcs if pc in pcs_just)
        df_sol = df_sol.sort_values('Qtd', ascending=True)
        if len(df_sol) > 0:
            # Calcula justificativas por solicitante
            # Com_Just_Qtd already calculated above

            # Sort ascending para horizontal ficar do maior para o menor no topo
            df_sol = df_sol.sort_values('Qtd', ascending=True)

            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                y=df_sol['Solicitante'], x=df_sol['Com_Just_Qtd'],
                name='Com justificativa',
                orientation='h',
                marker=dict(color='#51cf66'),
            ))
            fig1.add_trace(go.Bar(
                y=df_sol['Solicitante'], x=df_sol['Entrega_Enc'],
                name='Entrega enc.',
                orientation='h',
                marker=dict(color='#b197fc'),
            ))
            fig1.add_trace(go.Bar(
                y=df_sol['Solicitante'], x=df_sol['Vencidos'],
                name='Vencidos',
                orientation='h',
                marker=dict(color='#ff4d6a'),
                text=[f"{v} ({p}%)" for v, p in zip(df_sol['Vencidos'], df_sol['Pct_Vencido'])],
                textposition='inside',
                textfont=dict(size=11, color='white')
            ))
            fig1.update_layout(**PLOT_LAYOUT)
            fig1.update_layout(
                barmode='stack',
                title="👤 Por Solicitante",
                height=max(300, len(df_sol) * 60 + 80),
                xaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.05)', title='Qtd'),
                yaxis=dict(showgrid=False, automargin=True),
                margin=dict(l=10, r=20, t=40, b=90),
                legend=dict(orientation='h', y=-0.28, font=dict(size=11), x=0)
            )
            st.plotly_chart(fig1, use_container_width=True)

with col_g2:
    if 'Comprador' in df_filtered.columns and 'Valor' in df_filtered.columns:
        df_filtered['Comprador'] = df_filtered['Comprador'].astype(str).str.strip()
        base2 = df_filtered[df_filtered['Comprador'].notna() & (~df_filtered['Comprador'].isin(['—','','nan']))]
        df_comp = base2.groupby('Comprador').agg(Valor=('Valor','sum')).reset_index()
        if 'Vencimento' in base2.columns:
            df_comp_venc = base2[base2['Vencimento'] < agora_kpi].groupby('Comprador').agg(Valor_Vencido=('Valor','sum')).reset_index()
            df_comp = df_comp.merge(df_comp_venc, on='Comprador', how='left').fillna(0)
        else:
            df_comp['Valor_Vencido'] = 0
        if 'Dt Entrega PC' in df_filtered.columns:
            df_comp_ent = base2[base2['Dt Entrega PC'] < agora_kpi].groupby('Comprador').agg(Valor_Entrega=('Valor','sum')).reset_index()
            df_comp = df_comp.merge(df_comp_ent, on='Comprador', how='left').fillna(0)
        else:
            df_comp['Valor_Entrega'] = 0
        df_comp['Pct_Vencido'] = (df_comp['Valor_Vencido'] / df_comp['Valor'] * 100).round(0).astype(int)
        df_comp = df_comp.sort_values('Valor', ascending=True)
        if len(df_comp) > 0:
            fig2 = go.Figure()
            # Calcula valor com justificativa por comprador
            if len(just_df) > 0:
                ids_just2 = set(just_df['ID'].values)
                df_comp_just = base2[base2['ID'].isin(ids_just2)].groupby('Comprador').agg(Valor_Just=('Valor','sum')).reset_index()
                df_comp = df_comp.merge(df_comp_just, on='Comprador', how='left').fillna(0)
            else:
                df_comp['Valor_Just'] = 0

            fig2.add_trace(go.Bar(
                y=df_comp['Comprador'], x=df_comp['Valor_Just'],
                name='Com justificativa ✅',
                orientation='h',
                marker=dict(color='#51cf66'),
                textfont=dict(size=10, color='white')
            ))
            fig2.add_trace(go.Bar(
                y=df_comp['Comprador'], x=df_comp['Valor_Entrega'],
                name='Entrega enc. 📦',
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
                title="💰 Valor por Comprador",
                height=max(300, len(df_comp) * 50 + 100),
                xaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.05)'),
                yaxis=dict(showgrid=False, automargin=True),
                margin=dict(l=10, r=20, t=40, b=90),
                legend=dict(orientation='h', y=-0.22, font=dict(size=11), x=0)
            )
            st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════
# SEÇÃO: PROCESSOS VENCIDOS / ATRASADOS POR RESPONSÁVEL
# ══════════════════════════════════════════════════════════════════════

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

        # Tabela de vencidos em evidência
        st.markdown('<div class="section-title">📋 Detalhe dos Processos Vencidos</div>', unsafe_allow_html=True)
        just_df_venc = load_justificativas()
        # Rename just_df cols to match df columns for display
        just_rename = {'Nº_PC':'Nº PC','Nº_Nota':'Nº Nota','Dt_Entrega':'Dt Entrega PC'}
        if len(just_df_venc) > 0 and 'Nº_PC' in just_df_venc.columns:
            just_venc_cols = just_df_venc[['Nº_PC','Justificativa','Prazo_Resolucao','Observacao','Responsavel']].copy()
            just_venc_cols = just_venc_cols.rename(columns={'Nº_PC': 'Nº PC'})
            df_venc_display = df_vencidos.merge(just_venc_cols, on='Nº PC', how='left') if 'Nº PC' in df_vencidos.columns else df_vencidos.copy()
        else:
            df_venc_display = df_vencidos.copy()
        for col in ['Justificativa','Prazo_Resolucao','Observacao','Responsavel']:
            if col in df_venc_display.columns:
                df_venc_display[col] = df_venc_display[col].fillna('').replace('None', '')
        cols_venc = [c for c in ['Comprador','Solicitante','Fornecedor','Filial','Nº PC','Nº Nota','Controle','Dt Emissão','Dt Entrega PC','Vencimento','Dias_Atraso','Valor','Justificativa','Prazo_Resolucao','Responsavel'] if c in df_venc_display.columns]

        def highlight_atraso(row):
            dias = row.get('Dias_Atraso', 0)
            if dias > 30:
                return ['background-color: rgba(255,77,106,0.18)'] * len(row)
            elif dias > 7:
                return ['background-color: rgba(255,140,66,0.15)'] * len(row)
            else:
                return ['background-color: rgba(255,212,59,0.10)'] * len(row)

        _sort_col = 'Dias_Atraso' if 'Dias_Atraso' in df_venc_display.columns else cols_venc[0]
        styled_venc = df_venc_display[cols_venc].sort_values(_sort_col, ascending=False).style\
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
if len(just_df) > 0 and 'Nº_PC' in just_df.columns:
    just_cols_main = just_df[['Nº_PC','Justificativa','Prazo_Resolucao','Observacao','Responsavel']].copy()
    just_cols_main = just_cols_main.rename(columns={'Nº_PC': 'Nº PC'})
    df_display = df_filtered.merge(just_cols_main, on='Nº PC', how='left') if 'Nº PC' in df_filtered.columns else df_filtered.copy()
else:
    df_display = df_filtered.copy()
for col in ['Justificativa','Prazo_Resolucao','Observacao','Responsavel']:
    if col in df_display.columns:
        df_display[col] = df_display[col].fillna('').replace('None', '')

show_cols = [c for c in ['Comprador','Solicitante','Fornecedor','Filial','Nº PC','Nº Nota','Controle','Dt Emissão','Dt Entrega PC','Vencimento','Valor','Status','Justificativa','Prazo_Resolucao','Responsavel','Observacao'] if c in df_display.columns]

# Clean None values in display
df_display_clean = df_display[show_cols].copy()
for col in df_display_clean.columns:
    df_display_clean[col] = df_display_clean[col].replace({'None': '', 'nan': '', 'NaT': ''}).fillna('')
st.dataframe(
    df_display_clean.style.format({
        'Valor': lambda x: format_brl(x) if pd.notna(x) and isinstance(x, (int,float)) else x,
        'Dt Emissão': lambda x: str(x)[:10] if x and str(x) not in ['', 'None', 'nan', 'NaT'] else '',
        'Dt Entrega PC': lambda x: str(x)[:10] if x and str(x) not in ['', 'None', 'nan', 'NaT'] else '',
        'Vencimento': lambda x: str(x)[:10] if x and str(x) not in ['', 'None', 'nan', 'NaT'] else '',
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
    def label_pend(r):
        pc = str(r.get('Nº PC', '')).strip()
        nota = str(r.get('Nº Nota', '')).strip()
        forn = str(r.get('Fornecedor', '?'))[:35].strip()
        pc_str = f"PC:{pc}" if pc and pc not in ['—', '', 'nan'] else ''
        nota_str = f"NF:{nota}" if nota and nota not in ['—', '', 'nan'] else ''
        ref = ' | '.join(filter(None, [pc_str, nota_str]))
        return f"{ref} — {forn}" if ref else f"#{r['ID']} — {forn}"
    opcoes_pend = df_filtered.apply(label_pend, axis=1).tolist()
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
