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
        border-radius: 16px; padding: 16px 10px;
        border: 1px solid rgba(255,255,255,0.06);
        text-align: center; transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-2px); }
    .metric-value { font-size: 1.6rem; font-weight: 700; margin: 4px 0; line-height: 1.2; word-break: break-word; }
    .metric-label { font-size: 0.7rem; color: #8892a4; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; word-break: break-word; }
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
    "Pedido Mãe",
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
        # Only add header if sheet is completely empty - NEVER clear existing data
        if not first or str(first).strip() == '':
            ws.append_row(COLS_JUST)
        # If header exists but is wrong format, just leave it - do not clear
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

                # Formata datas em DD/MM/YYYY
                def _fmt_br(x):
                    if x is None or str(x).strip() in ['', 'nan', 'NaT', 'None']:
                        return ''
                    if isinstance(x, (pd.Timestamp, datetime, date)):
                        try: return x.strftime('%d/%m/%Y')
                        except: return ''
                    d = pd.to_datetime(x, dayfirst=True, errors='coerce')
                    return d.strftime('%d/%m/%Y') if pd.notna(d) else str(x)[:10]

                row_data = {
                    "Comprador": str(r.get('Comprador', '')),
                    "Solicitante": str(r.get('Solicitante', '')),
                    "Fornecedor": str(r.get('Fornecedor', '')),
                    "Filial": str(r.get('Filial', '')),
                    "Nº_PC": str(r.get('Nº PC', '')),
                    "Nº_Nota": str(r.get('Nº Nota', '')),
                    "Dt_Entrega": _fmt_br(r.get('Dt Entrega PC', '')),
                    "Vencimento": _fmt_br(venc),
                    "Valor": str(r.get('Valor', '')),
                    "Dias_Atraso": str(dias_atraso),
                }

        # Formata prazo de resolução em DD/MM/YYYY
        prazo_fmt = ''
        if prazo:
            if isinstance(prazo, (pd.Timestamp, datetime, date)):
                try: prazo_fmt = prazo.strftime('%d/%m/%Y')
                except: prazo_fmt = str(prazo)
            else:
                d = pd.to_datetime(prazo, dayfirst=True, errors='coerce')
                prazo_fmt = d.strftime('%d/%m/%Y') if pd.notna(d) else str(prazo)

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
            "Prazo_Resolucao": prazo_fmt,
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



def parse_data(v):
    """Converte qualquer formato para datetime: texto BR ('DD/MM/AAAA'), ISO ('YYYY-MM-DD'), datetime, date ou Timestamp."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, pd.Timestamp):
        return v
    if isinstance(v, datetime):
        return pd.Timestamp(v)
    if isinstance(v, date):
        return pd.Timestamp(v)
    s = str(v).strip()
    if s in ["", "—", "-", "nan", "NaT", "None"]:
        return pd.NaT
    # Remove prefixos tipo "VENCIDA 3d" ou "3 dias" mantendo só a data BR
    import re
    m = re.search(r'(\d{2}/\d{2}/\d{4})', s)
    if m:
        # Formato brasileiro DD/MM/YYYY explícito
        return pd.to_datetime(m.group(1), format='%d/%m/%Y', errors='coerce')
    # Formato ISO YYYY-MM-DD (vindo de Excel datetime convertido para string)
    m_iso = re.search(r'(\d{4}-\d{2}-\d{2})', s)
    if m_iso:
        return pd.to_datetime(m_iso.group(1), format='%Y-%m-%d', errors='coerce')
    # Fallback: tenta parse genérico com dayfirst
    return pd.to_datetime(s, dayfirst=True, errors='coerce')

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
    # Conversão robusta de TODAS as colunas de data (aceita texto OU objeto date)
    for date_col in ["Vencimento", "Dt Entrega PC", "Dt Emissão"]:
        if date_col in df.columns:
            df[date_col] = df[date_col].apply(parse_data)
    if "Dias" in df.columns:
        df["Dias"] = pd.to_numeric(df["Dias"].replace("—", ""), errors="coerce").fillna(0).astype(int)
    key_cols = [c for c in ["Fornecedor", "Comprador", "Valor"] if c in df.columns]
    if key_cols:
        df = df.dropna(subset=key_cols, how="all").reset_index(drop=True)
    df["ID"] = df.index.astype(str)

    # Normaliza Comprador e Solicitante vazios
    PENDENTE_RESP = "⚠️ Pendente Identificação"
    VAZIOS = ['', '—', 'nan', 'None', 'NaN']

    if 'Comprador' in df.columns:
        df['Comprador'] = df['Comprador'].apply(
            lambda x: str(x).strip() if str(x).strip() not in VAZIOS else '')
    if 'Solicitante' in df.columns:
        df['Solicitante'] = df['Solicitante'].apply(
            lambda x: str(x).strip() if str(x).strip() not in VAZIOS else '')

    # Se Solicitante vazio mas Comprador preenchido → usa Comprador
    if 'Solicitante' in df.columns and 'Comprador' in df.columns:
        mask_sol_vazio = df['Solicitante'] == ''
        df.loc[mask_sol_vazio & (df['Comprador'] != ''), 'Solicitante'] = df.loc[mask_sol_vazio & (df['Comprador'] != ''), 'Comprador']

        # Se Comprador vazio mas Solicitante preenchido → usa Solicitante
        mask_comp_vazio = df['Comprador'] == ''
        df.loc[mask_comp_vazio & (df['Solicitante'] != ''), 'Comprador'] = df.loc[mask_comp_vazio & (df['Solicitante'] != ''), 'Solicitante']

        # Se ambos vazios → marca como pendente de identificação
        mask_ambos = (df['Comprador'] == '') & (df['Solicitante'] == '')
        df.loc[mask_ambos, 'Comprador'] = PENDENTE_RESP
        df.loc[mask_ambos, 'Solicitante'] = PENDENTE_RESP
    elif 'Comprador' in df.columns:
        df.loc[df['Comprador'] == '', 'Comprador'] = PENDENTE_RESP
    elif 'Solicitante' in df.columns:
        df.loc[df['Solicitante'] == '', 'Solicitante'] = PENDENTE_RESP

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

    # ── Calcula situação de cada processo ──
    agora_sit = pd.Timestamp.now()
    just_df_sit = load_justificativas()
    pcs_just_sit = set()
    if len(just_df_sit) > 0 and 'Nº_PC' in just_df_sit.columns:
        pcs_just_sit = set(just_df_sit['Nº_PC'].astype(str).str.strip().str.lstrip('0').values) - {'','nan','None','—'}

    def calc_situacao(row):
        # Se está pendente de identificação de responsável, essa é a situação prioritária
        comp = str(row.get('Comprador', '')).strip()
        sol = str(row.get('Solicitante', '')).strip()
        if '⚠️' in comp or '⚠️' in sol or 'Pendente Identificação' in comp or 'Pendente Identificação' in sol:
            return 'Pendente Identificação Responsável'

        venc = parse_data(row.get('Vencimento', pd.NaT))
        pc = str(row.get('Nº PC', '')).strip().lstrip('0')
        tem_just = pc in pcs_just_sit and pc != ''
        if pd.notna(venc) and venc < agora_sit:
            if tem_just:
                return 'Vencido c/ Justificativa'
            return 'Vencido s/ Justificativa'
        if 'Dt Entrega PC' in row.index:
            entr = parse_data(row.get('Dt Entrega PC', pd.NaT))
            if pd.notna(entr) and entr < agora_sit:
                return 'Entrega Encerrada'
        if tem_just:
            return 'Em Dia (Justificado)'
        return 'Pendente'

    df['Situação'] = df.apply(calc_situacao, axis=1)

    st.markdown("---")
    st.markdown("### 🔍 Filtros")

    compradores = ["Todos"] + sorted([x for x in df['Comprador'].dropna().unique().tolist() if x not in ['—','']]) if 'Comprador' in df.columns else ["Todos"]
    sel_comprador = st.selectbox("Comprador", compradores)

    solicitantes = ["Todos"] + sorted([x for x in df['Solicitante'].dropna().unique().tolist() if x not in ['—','']]) if 'Solicitante' in df.columns else ["Todos"]
    sel_solicitante = st.selectbox("Solicitante", solicitantes)

    sit_opts = ["Todas"] + [
        'Pendente Identificação Responsável',
        'Vencido s/ Justificativa',
        'Vencido c/ Justificativa',
        'Entrega Encerrada',
        'Em Dia (Justificado)',
        'Pendente'
    ]
    sel_situacao = st.selectbox("🏷️ Situação", sit_opts)

    if 'Filial' in df.columns:
        filiais = ["Todas"] + sorted([x for x in df['Filial'].dropna().astype(str).unique().tolist() if x not in ['—','']])
        sel_filial = st.selectbox("Filial", filiais)
    else:
        sel_filial = "Todas"

    if 'Controle' in df.columns:
        sel_aprovacao = st.selectbox("Aprovação", ["Todos", "B — Em aprovação", "L — Aprovado"])
    else:
        sel_aprovacao = "Todos"

    data_de = None
    data_ate = None

    # Filtro por data de emissão
    st.markdown("#### 🧾 Dt Emissão")
    if 'Dt Emissão' in df.columns:
        df['Dt Emissão'] = df['Dt Emissão'].apply(parse_data)
        emiss_validas = df['Dt Emissão'].dropna()
        emiss_validas = emiss_validas[emiss_validas.dt.year <= pd.Timestamp.now().year + 1]
        if len(emiss_validas) > 0:
            emiss_min = emiss_validas.min()
            emiss_max = emiss_validas.max()
            from datetime import date as date_type
            emiss_de = st.date_input("De", value=emiss_min.date(), min_value=emiss_min.date(), max_value=emiss_max.date(), key="emiss_de", format="DD/MM/YYYY")
            emiss_ate = st.date_input("Até", value=emiss_max.date(), min_value=emiss_min.date(), max_value=emiss_max.date(), key="emiss_ate", format="DD/MM/YYYY")
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
    if sel_situacao != "Todas" and 'Situação' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Situação'] == sel_situacao]
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
        df_filtered['Dt Emissão'] = df_filtered['Dt Emissão'].apply(parse_data)
        emiss_de_ts = pd.Timestamp(emiss_de)
        emiss_ate_ts = pd.Timestamp(emiss_ate)
        mask_emiss = (
            df_filtered['Dt Emissão'].isna() |
            (
                (df_filtered['Dt Emissão'] >= emiss_de_ts) &
                (df_filtered['Dt Emissão'] <= emiss_ate_ts)
            )
        )
        df_filtered = df_filtered[mask_emiss]

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
    # Garante tipo datetime mesmo se vier de atualização
    df_filtered['Vencimento'] = df_filtered['Vencimento'].apply(parse_data)
    venc_series = df_filtered['Vencimento']
    df_venc_kpi = df_filtered[venc_series < agora_now].copy()
    vencidas = len(df_venc_kpi)
    valor_vencido = df_venc_kpi['Valor'].sum() if 'Valor' in df_venc_kpi.columns else 0
    if vencidas > 0:
        maior_atraso_kpi = int((agora_now.normalize() - df_venc_kpi['Vencimento']).dt.days.max())

entrega_enc = 0
if 'Dt Entrega PC' in df_filtered.columns:
    df_filtered['Dt Entrega PC'] = df_filtered['Dt Entrega PC'].apply(parse_data)
    entrega_enc = len(df_filtered[df_filtered['Dt Entrega PC'] < agora_now])

just_df = load_justificativas()
if len(just_df) > 0 and 'Nº_PC' in just_df.columns and 'Nº PC' in df_filtered.columns:
    pcs_com_just = set(just_df['Nº_PC'].astype(str).str.strip().str.lstrip('0').values) - {'', 'nan', 'None', '—'}
    df_pcs = df_filtered['Nº PC'].astype(str).str.strip().str.lstrip('0')
    com_justificativa = len(df_filtered[df_pcs.isin(pcs_com_just)]) if pcs_com_just else 0
else:
    com_justificativa = 0
sem_justificativa = total_itens - com_justificativa
# ── Helpers de formatação ──
def fmt_pct(num, den):
    if den == 0 or num == 0: return '0%'
    p = num / den * 100
    return f'{p:.1f}%' if p < 1 else f'{round(p)}%'

def fmt_sub_pct(num, den, label='do total'):
    p = fmt_pct(num, den)
    return f'{p} {label}' if num > 0 else 'nenhum registro'

pct_vencidos_str = fmt_pct(vencidas, total_itens)
pct_entrega_str  = fmt_pct(entrega_enc, total_itens)
pct_just_str     = fmt_pct(com_justificativa, total_itens)

# Situação calculada
n_venc_sem_just = len(df_filtered[df_filtered['Situação'] == 'Vencido s/ Justificativa']) if 'Situação' in df_filtered.columns else 0
n_venc_com_just = len(df_filtered[df_filtered['Situação'] == 'Vencido c/ Justificativa']) if 'Situação' in df_filtered.columns else 0
n_em_dia        = len(df_filtered[df_filtered['Situação'] == 'Em Dia (Justificado)'])      if 'Situação' in df_filtered.columns else 0
n_entrega_enc   = len(df_filtered[df_filtered['Situação'] == 'Entrega Encerrada'])          if 'Situação' in df_filtered.columns else 0
val_venc_sem    = df_filtered[df_filtered['Situação'] == 'Vencido s/ Justificativa']['Valor'].sum() if 'Situação' in df_filtered.columns and 'Valor' in df_filtered.columns else 0
val_venc_com    = df_filtered[df_filtered['Situação'] == 'Vencido c/ Justificativa']['Valor'].sum() if 'Situação' in df_filtered.columns and 'Valor' in df_filtered.columns else 0

# ── Linha 1: KPIs críticos ──
st.markdown("""<style>
.big-card {
    background: linear-gradient(135deg, #1a1f2e 0%, #252b3b 100%);
    border-radius: 16px; padding: 24px 20px;
    border: 1px solid rgba(255,255,255,0.06);
    text-align: center; height: 100%;
}
.big-card .label { font-size: 0.72rem; color: #8892a4; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 700; }
.big-card .value { font-size: 2.4rem; font-weight: 700; line-height: 1.2; margin: 8px 0 4px 0; }
.big-card .pct   { font-size: 1rem; font-weight: 500; opacity: 0.75; margin-left: 6px; }
.big-card .sub   { font-size: 0.78rem; color: #8892a4; margin-top: 4px; }
.big-card.critical { border: 1.5px solid rgba(255,77,106,0.5); background: linear-gradient(135deg, #2a1520, #3a1a28); }
.big-card.warning  { border: 1.5px solid rgba(177,151,252,0.4); background: linear-gradient(135deg, #1e1a2e, #261f3a); }
.big-card.success  { border: 1.5px solid rgba(81,207,102,0.4); background: linear-gradient(135deg, #0f2218, #152e1f); }
.big-card.neutral  { border: 1.5px solid rgba(255,255,255,0.08); }
</style>""", unsafe_allow_html=True)

# ── Justificativas vencidas (prazo de resolução já passou) ──
n_just_vencida = 0
val_just_vencida = 0
if len(just_df) > 0 and 'Nº_PC' in just_df.columns and 'Prazo_Resolucao' in just_df.columns and 'Nº PC' in df_filtered.columns:
    hoje_ts = pd.Timestamp.now().normalize()
    just_df_copy = just_df.copy()
    just_df_copy['_prazo'] = just_df_copy['Prazo_Resolucao'].apply(parse_data)
    just_df_copy['_pc_norm'] = just_df_copy['Nº_PC'].astype(str).str.strip().str.lstrip('0')
    mask_just_venc = (just_df_copy['_prazo'].notna()) & (just_df_copy['_prazo'] < hoje_ts)
    pcs_just_vencidas = set(just_df_copy[mask_just_venc]['_pc_norm'].values) - {'', 'nan', 'None', '—'}

    if pcs_just_vencidas:
        df_pcs_filt = df_filtered['Nº PC'].astype(str).str.strip().str.lstrip('0')
        df_just_vencidas = df_filtered[df_pcs_filt.isin(pcs_just_vencidas)]
        n_just_vencida = len(df_just_vencidas)
        val_just_vencida = df_just_vencidas['Valor'].sum() if 'Valor' in df_just_vencidas.columns else 0

# ── Dados adicionais para governança ──
em_aprovacao = 0
if 'Controle' in df_filtered.columns:
    em_aprovacao = len(df_filtered[df_filtered['Controle'].astype(str).str.upper().str.startswith('B')])

valor_em_aprovacao = 0
if 'Controle' in df_filtered.columns and 'Valor' in df_filtered.columns and 'Vencimento' in df_filtered.columns:
    agora_aprov = pd.Timestamp.now()
    mask_aprov = (df_filtered['Controle'].astype(str).str.upper().str.startswith('B')) & (df_filtered['Vencimento'] < agora_aprov)
    valor_em_aprovacao = df_filtered[mask_aprov]['Valor'].sum()

n_pend_ident = 0
valor_pend_ident = 0
if 'Situação' in df_filtered.columns:
    df_pend_ident = df_filtered[df_filtered['Situação'] == 'Pendente Identificação Responsável']
    n_pend_ident = len(df_pend_ident)
    if 'Valor' in df_pend_ident.columns:
        valor_pend_ident = df_pend_ident['Valor'].sum()

# ══════════════════════════════════════════════════════════════════════
# LINHA 1: Visão geral
# ══════════════════════════════════════════════════════════════════════
r1c1, r1c2 = st.columns(2)

with r1c1:
    st.markdown(f"""<div class="big-card neutral">
        <div class="label">🧾 Pendências Totais</div>
        <div class="value" style="color:#e0e0e0">{total_itens}</div>
        <div class="sub" style="font-size:0.9rem;margin-top:8px">
            💰 Valor total: <strong style="color:#51cf66;font-size:1.05rem">{format_brl(total_valor)}</strong><br>
            <span style="opacity:0.7">Entregas em atraso + processos vencidos</span>
        </div>
    </div>""", unsafe_allow_html=True)

with r1c2:
    st.markdown(f"""<div class="big-card critical">
        <div class="label">🔴 Vencidas Sem Justificativa</div>
        <div class="value" style="color:#ff4d6a">{n_venc_sem_just}
            <span class="pct" style="color:#ff8080">({fmt_pct(n_venc_sem_just, total_itens)})</span>
        </div>
        <div class="sub" style="font-size:0.9rem;margin-top:8px">
            🔥 Impacto financeiro vencido: <strong style="color:#ff4d6a;font-size:1.05rem">{format_brl(val_venc_sem)}</strong><br>
            ⏱️ Maior atraso: <strong>{maior_atraso_kpi} dias</strong>
        </div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# LINHA 2: Justificativas + Governança
# ══════════════════════════════════════════════════════════════════════
total_com_just = n_venc_com_just + n_em_dia
val_com_just = val_venc_com + (df_filtered[df_filtered['Situação'] == 'Em Dia (Justificado)']['Valor'].sum() if 'Situação' in df_filtered.columns and 'Valor' in df_filtered.columns else 0)

r2c1, r2c2 = st.columns(2)

with r2c1:
    st.markdown(f"""<div class="big-card success">
        <div class="label">✅ Com Justificativa</div>
        <div class="value" style="color:#51cf66">{total_com_just}
            <span class="pct" style="color:#80e89a">({fmt_pct(total_com_just, total_itens)})</span>
        </div>
        <div class="sub" style="font-size:0.9rem;margin-top:8px">
            💰 Valor: <strong style="color:#51cf66;font-size:1.05rem">{format_brl(val_com_just)}</strong><br>
            <span style="color:#ffd43b;font-weight:600">👀 Importante manter monitoramento</span>
        </div>
    </div>""", unsafe_allow_html=True)

with r2c2:
    st.markdown(f"""<div class="big-card" style="border:1.5px solid #4dabf7;background:linear-gradient(135deg,#0f1a2e,#152038)">
        <div class="label" style="color:#4dabf7">🛡️ Governança</div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-top:10px;text-align:left;padding:0 4px">
            <div>
                <div style="font-size:0.7rem;color:#8892a4;text-transform:uppercase;letter-spacing:1px">⚠️ Sem responsável</div>
                <div style="font-size:1.6rem;font-weight:700;color:#ffd43b;line-height:1.1;margin-top:4px">{n_pend_ident}</div>
                <div style="font-size:0.72rem;color:#ffd43b;opacity:0.7">{format_brl(valor_pend_ident)}</div>
            </div>
            <div>
                <div style="font-size:0.7rem;color:#8892a4;text-transform:uppercase;letter-spacing:1px">🕓 Em aprovação</div>
                <div style="font-size:1.6rem;font-weight:700;color:#4dabf7;line-height:1.1;margin-top:4px">{em_aprovacao}</div>
                <div style="font-size:0.72rem;color:#4dabf7;opacity:0.7">{fmt_pct(em_aprovacao, total_itens)} do total</div>
            </div>
            <div>
                <div style="font-size:0.7rem;color:#8892a4;text-transform:uppercase;letter-spacing:1px">💰 Vencido em aprov.</div>
                <div style="font-size:1.2rem;font-weight:700;color:#ff8c42;line-height:1.1;margin-top:8px">{format_brl(valor_em_aprovacao)}</div>
                <div style="font-size:0.72rem;color:#ff8c42;opacity:0.7">aguardando aprovação</div>
            </div>
        </div>
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
st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# GRÁFICOS
# ══════════════════════════════════════════════════════════════════════


st.markdown('<div class="section-title">📈 Visão por Comprador / Solicitante</div>', unsafe_allow_html=True)


agora_kpi = pd.Timestamp.now().normalize()

col_g1, col_g2 = st.columns(2)

def get_cor_barra(pct_vencido, pct_just, pct_entrega):
    """Define cor da barra por prioridade"""
    if pct_vencido >= 80: return '#ff4d6a'
    if pct_entrega >= 50: return '#b197fc'
    if pct_just >= 50:    return '#51cf66'
    if pct_vencido > 0:   return '#ff8c42'
    return '#4dabf7'

# ══════════════════════════════════════════════════════════════════════
# Função auxiliar para montar dados segmentados (vencido/entrega/just/normal)
# ══════════════════════════════════════════════════════════════════════
def montar_segmentos(base, group_col, agg_col=None):
    """
    Monta dataframe com segmentos mutuamente exclusivos.
    Se agg_col for None: agrega por quantidade.
    Se agg_col='Valor': agrega por soma do valor.
    Prioridade: Com Just > Vencido > Entrega enc. > Normal
    (justificativa tira status de vencido)
    """
    pcs_j = set(just_df['Nº_PC'].astype(str).str.strip().str.lstrip('0').values) - {'','nan','None','—'} if len(just_df)>0 and 'Nº_PC' in just_df.columns else set()

    grupos = base[group_col].dropna().unique()
    rows = []
    for g in grupos:
        rows_g = base[base[group_col] == g]
        # Justificativa tem prioridade máxima
        just_m = rows_g['Nº PC'].astype(str).str.strip().str.lstrip('0').isin(pcs_j) if 'Nº PC' in rows_g.columns and pcs_j else pd.Series([False]*len(rows_g), index=rows_g.index)
        venc_m = (rows_g['Vencimento'] < agora_kpi) & ~just_m if 'Vencimento' in rows_g.columns else pd.Series([False]*len(rows_g), index=rows_g.index)
        entr_m = (rows_g['Dt Entrega PC'] < agora_kpi) & ~just_m & ~venc_m if 'Dt Entrega PC' in rows_g.columns else pd.Series([False]*len(rows_g), index=rows_g.index)
        norm_m = ~just_m & ~venc_m & ~entr_m

        if agg_col and agg_col in rows_g.columns:
            rows.append({
                group_col: g,
                'Total': rows_g[agg_col].sum(),
                'Seg_Vencido': rows_g[venc_m][agg_col].sum(),
                'Seg_Entrega': rows_g[entr_m][agg_col].sum(),
                'Seg_Just':    rows_g[just_m][agg_col].sum(),
                'Seg_Normal':  rows_g[norm_m][agg_col].sum(),
            })
        else:
            rows.append({
                group_col: g,
                'Total': len(rows_g),
                'Seg_Vencido': int(venc_m.sum()),
                'Seg_Entrega': int(entr_m.sum()),
                'Seg_Just':    int(just_m.sum()),
                'Seg_Normal':  int(norm_m.sum()),
            })
    return pd.DataFrame(rows).sort_values('Total', ascending=True)

def criar_fig_segmentado(df_seg, group_col, titulo, eh_valor=False):
    """Cria gráfico de barras horizontais segmentado com total visível fora da barra."""
    fig = go.Figure()

    # Segmentos coloridos (sem texto interno para evitar ilegibilidade)
    for seg, cor, nome in [
        ('Seg_Normal',  '#4a5568', 'Pendente normal'),
        ('Seg_Just',    '#51cf66', 'Com justificativa'),
        ('Seg_Entrega', '#b197fc', 'Entrega enc.'),
        ('Seg_Vencido', '#ff4d6a', 'Vencido'),
    ]:
        if eh_valor:
            hover = f'<b>%{{y}}</b><br>{nome}: R$ %{{x:,.2f}}<extra></extra>'
        else:
            hover = f'<b>%{{y}}</b><br>{nome}: %{{x}}<extra></extra>'
        fig.add_trace(go.Bar(
            y=df_seg[group_col], x=df_seg[seg], name=nome,
            orientation='h', marker=dict(color=cor),
            hovertemplate=hover
        ))

    # Anotações com TOTAL por fora da barra (à direita)
    if eh_valor:
        textos_totais = [format_brl(v) for v in df_seg['Total']]
    else:
        textos_totais = [f"{int(v)}" for v in df_seg['Total']]

    # Adiciona texto do total no final de cada barra (posição x = Total, com padding)
    x_max = df_seg['Total'].max() if len(df_seg) > 0 else 0
    for i, (categoria, total, txt) in enumerate(zip(df_seg[group_col], df_seg['Total'], textos_totais)):
        fig.add_annotation(
            x=total, y=categoria,
            text=f"<b>{txt}</b>",
            showarrow=False,
            xanchor='left', yanchor='middle',
            xshift=8,
            font=dict(size=12, color='#e0e0e0')
        )

    fig.update_layout(**PLOT_LAYOUT)
    fig.update_layout(
        barmode='stack', title=titulo,
        height=max(280, len(df_seg) * 55 + 60),
        xaxis=dict(
            showgrid=True,
            gridcolor='rgba(255,255,255,0.05)',
            # Adiciona espaço à direita para mostrar o total
            range=[0, x_max * 1.25] if x_max > 0 else None
        ),
        yaxis=dict(showgrid=False, automargin=True),
        margin=dict(l=10, r=40, t=40, b=80),
        legend=dict(orientation='h', y=-0.22, font=dict(size=11))
    )
    return fig

# ── LINHA 1: SOLICITANTE ──
if 'Solicitante' in df_filtered.columns:
    df_filtered['Solicitante'] = df_filtered['Solicitante'].astype(str).str.strip()
    base_sol = df_filtered[df_filtered['Solicitante'].notna() & (~df_filtered['Solicitante'].isin(['—','','nan','None']))]

    if len(base_sol) > 0:
        with col_g1:
            df_sol_qtd = montar_segmentos(base_sol, 'Solicitante', agg_col=None)
            fig_sol_qtd = criar_fig_segmentado(df_sol_qtd, 'Solicitante', '👤 Solicitante — Qtd de Processos', eh_valor=False)
            st.plotly_chart(fig_sol_qtd, use_container_width=True)

        with col_g2:
            if 'Valor' in base_sol.columns:
                df_sol_val = montar_segmentos(base_sol, 'Solicitante', agg_col='Valor')
                fig_sol_val = criar_fig_segmentado(df_sol_val, 'Solicitante', '💰 Solicitante — Valor Total', eh_valor=True)
                st.plotly_chart(fig_sol_val, use_container_width=True)

# ── LINHA 2: COMPRADOR ──
col_g3, col_g4 = st.columns(2)

if 'Comprador' in df_filtered.columns:
    df_filtered['Comprador'] = df_filtered['Comprador'].astype(str).str.strip()
    base_comp = df_filtered[df_filtered['Comprador'].notna() & (~df_filtered['Comprador'].isin(['—','','nan','None']))]

    if len(base_comp) > 0:
        with col_g3:
            df_comp_qtd = montar_segmentos(base_comp, 'Comprador', agg_col=None)
            fig_comp_qtd = criar_fig_segmentado(df_comp_qtd, 'Comprador', '🛒 Comprador — Qtd de Processos', eh_valor=False)
            st.plotly_chart(fig_comp_qtd, use_container_width=True)

        with col_g4:
            if 'Valor' in base_comp.columns:
                df_comp_val = montar_segmentos(base_comp, 'Comprador', agg_col='Valor')
                fig_comp_val = criar_fig_segmentado(df_comp_val, 'Comprador', '💰 Comprador — Valor Total', eh_valor=True)
                st.plotly_chart(fig_comp_val, use_container_width=True)



# ══════════════════════════════════════════════════════════════════════
# TABELAS SEPARADAS POR PRIORIDADE
# ══════════════════════════════════════════════════════════════════════
just_df = load_justificativas()
if len(just_df) > 0 and 'Nº_PC' in just_df.columns:
    just_cols_main = just_df[['Nº_PC','Justificativa','Prazo_Resolucao','Observacao','Responsavel']].copy()
    just_cols_main = just_cols_main.rename(columns={'Nº_PC': 'Nº PC'})
    just_cols_main['Nº PC'] = just_cols_main['Nº PC'].astype(str).str.strip().str.lstrip('0')
    df_filtered_merge = df_filtered.copy()
    if 'Nº PC' in df_filtered_merge.columns:
        df_filtered_merge['Nº PC_norm'] = df_filtered_merge['Nº PC'].astype(str).str.strip().str.lstrip('0')
        just_cols_main = just_cols_main.rename(columns={'Nº PC': 'Nº PC_norm'})
        df_display = df_filtered_merge.merge(just_cols_main, on='Nº PC_norm', how='left').drop(columns=['Nº PC_norm'])
    else:
        df_display = df_filtered.copy()
else:
    df_display = df_filtered.copy()

for col in ['Justificativa','Prazo_Resolucao','Observacao','Responsavel']:
    if col in df_display.columns:
        df_display[col] = df_display[col].fillna('').replace({'None':'','nan':'','NaT':''})

show_cols = [c for c in ['Comprador','Solicitante','Filial','Fornecedor','Nº PC','Nº Nota','Controle','Situação','Dt Emissão','Dt Entrega PC','Vencimento','Valor','Justificativa','Prazo_Resolucao','Responsavel'] if c in df_display.columns]

def fmt_data(x):
    """Exibe data no formato DD/MM/AAAA independentemente do formato de entrada."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    s = str(x).strip()
    if s in ['', 'None', 'nan', 'NaT', '—']:
        return ''
    # Se for Timestamp/datetime/date
    if isinstance(x, (pd.Timestamp, datetime, date)):
        try:
            return x.strftime('%d/%m/%Y')
        except:
            return ''
    # Tenta parsear texto
    try:
        d = parse_data(x)
        if pd.notna(d):
            return d.strftime('%d/%m/%Y')
    except:
        pass
    return s[:10]

FMT = {
    'Valor': lambda x: format_brl(x) if pd.notna(x) and isinstance(x, (int,float)) else x,
    'Dt Emissão': fmt_data,
    'Dt Entrega PC': fmt_data,
    'Vencimento': fmt_data,
    'Prazo_Resolucao': fmt_data,
}

def clean_table(df, cols):
    d = df[[c for c in cols if c in df.columns]].copy()
    for c in d.columns:
        d[c] = d[c].replace({'None':'','nan':'','NaT':''}).fillna('')
    return d

agora_tab = pd.Timestamp.now()
em_10_dias = agora_tab + pd.Timedelta(days=10)

# Tabela 1a: Já vencidos
if 'Vencimento' in df_display.columns:
    venc_s = df_display['Vencimento'].apply(parse_data)
    mask_vencido = venc_s < agora_tab
    mask_a_vencer = (venc_s >= agora_tab) & (venc_s <= em_10_dias)
    df_t1a = df_display[mask_vencido].copy()
    df_t1a['_dias'] = (agora_tab - df_t1a['Vencimento'].apply(parse_data)).dt.days.fillna(0).astype(int)
    df_t1b = df_display[mask_a_vencer].copy()
    df_t1b['_dias_restantes'] = (df_t1b['Vencimento'].apply(parse_data) - agora_tab).dt.days.fillna(0).astype(int)
    ids_t1 = set(df_t1a.index) | set(df_t1b.index)
else:
    df_t1a = pd.DataFrame()
    df_t1b = pd.DataFrame()
    ids_t1 = set()

# Tabela 1a: Já vencidos
n_t1a = len(df_t1a)
st.markdown(f"""<div style='background:linear-gradient(135deg,#2a0e0e,#3a1010);border:1.5px solid #ff4d6a;border-radius:12px;padding:16px 20px;margin:24px 0 4px 0'>
    <span style='color:#ff4d6a;font-size:1.1rem;font-weight:700'>🔴 JÁ VENCIDOS — Ação Imediata Necessária</span>
    <span style='color:#f08090;font-size:0.85rem;margin-left:12px'>({n_t1a} registros)</span><br>
    <span style='color:#f08090;font-size:0.8rem'>Processos que já passaram do prazo de vencimento. Ordenados pelo maior atraso.</span>
</div>""", unsafe_allow_html=True)

if n_t1a > 0:
    t1a_clean = clean_table(df_t1a.sort_values('_dias', ascending=False), show_cols)
    st.dataframe(t1a_clean.style.set_properties(**{'background-color':'rgba(255,77,106,0.15)'}).format(FMT),
        use_container_width=True, height=350)
else:
    st.success("✅ Nenhum processo vencido!")

# Tabela 1b: Vencendo em 10 dias
n_t1b = len(df_t1b)
st.markdown(f"""<div style='background:linear-gradient(135deg,#1e1000,#2e1800);border:1.5px solid #ffd43b;border-radius:12px;padding:16px 20px;margin:24px 0 4px 0'>
    <span style='color:#ffd43b;font-size:1.1rem;font-weight:700'>🟡 ATENÇÃO — Vencendo nos Próximos 10 Dias</span>
    <span style='color:#ffe066;font-size:0.85rem;margin-left:12px'>({n_t1b} registros)</span><br>
    <span style='color:#ffe066;font-size:0.8rem'>Ainda dentro do prazo, mas vencerão em breve. Antecipe a ação para evitar atraso.</span>
</div>""", unsafe_allow_html=True)

if n_t1b > 0:
    t1b_clean = clean_table(df_t1b.sort_values('_dias_restantes', ascending=True), show_cols)
    st.dataframe(t1b_clean.style.set_properties(**{'background-color':'rgba(255,212,59,0.10)'}).format(FMT),
        use_container_width=True, height=300)
else:
    st.success("✅ Nenhum processo vencendo nos próximos 10 dias!")

# Tabela 2: Notas sem PC identificado
if 'Nº PC' in df_display.columns:
    sem_pc = df_display['Nº PC'].astype(str).str.strip().isin(['','—','nan','None'])
    df_t2 = df_display[sem_pc & ~df_display.index.isin(ids_t1)].copy()
else:
    df_t2 = pd.DataFrame()
ids_t2 = set(df_t2.index)
n_t2 = len(df_t2)

st.markdown(f"""<div style='background:linear-gradient(135deg,#1e1208,#2e1c10);border:1.5px solid #ff8c42;border-radius:12px;padding:16px 20px;margin:24px 0 4px 0'>
    <span style='color:#ff8c42;font-size:1.1rem;font-weight:700'>🟠 PENDENTE DE VÍNCULO — Notas Sem PC Identificado</span>
    <span style='color:#ffa570;font-size:0.85rem;margin-left:12px'>({n_t2} registros)</span><br>
    <span style='color:#ffa570;font-size:0.8rem'>Notas fiscais sem pedido de compra associado. Necessitam regularização urgente.</span>
</div>""", unsafe_allow_html=True)

if n_t2 > 0:
    t2_clean = clean_table(df_t2, show_cols)
    st.dataframe(t2_clean.style.set_properties(**{'background-color':'rgba(255,140,66,0.10)'}).format(FMT), use_container_width=True, height=300)
else:
    st.success("✅ Todas as notas possuem PC identificado!")

# Tabela 3: Demais pendências
df_t3 = df_display[~df_display.index.isin(ids_t1 | ids_t2)].copy()
n_t3 = len(df_t3)

st.markdown(f"""<div style='background:linear-gradient(135deg,#111827,#1a2035);border:1.5px solid #4dabf7;border-radius:12px;padding:16px 20px;margin:24px 0 4px 0'>
    <span style='color:#4dabf7;font-size:1.1rem;font-weight:700'>📋 ACOMPANHAMENTO — Demais Pendências de Atendimento</span>
    <span style='color:#8ab8d4;font-size:0.85rem;margin-left:12px'>({n_t3} registros)</span><br>
    <span style='color:#8ab8d4;font-size:0.8rem'>Processos em andamento. Monitore para evitar evolução para atraso.</span>
</div>""", unsafe_allow_html=True)

if n_t3 > 0:
    t3_clean = clean_table(df_t3, show_cols)
    st.dataframe(t3_clean.style.format(FMT), use_container_width=True, height=350)
else:
    st.info("Sem demais pendências.")

df_display_export = df_display.copy()

# ══════════════════════════════════════════════════════════════════════
# FORMULÁRIO DE JUSTIFICATIVA
# ══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">✏️ Preencher Justificativa</div>', unsafe_allow_html=True)

def label_pend(r):
    pc = str(r.get('Nº PC', '')).strip()
    nota = str(r.get('Nº Nota', '')).strip()
    forn = str(r.get('Fornecedor', '?'))[:35].strip()
    pc_str = f"PC:{pc}" if pc and pc not in ['—', '', 'nan'] else ''
    nota_str = f"NF:{nota}" if nota and nota not in ['—', '', 'nan'] else ''
    ref = ' | '.join(filter(None, [pc_str, nota_str]))
    return f"{ref} — {forn}" if ref else f"#{r['ID']} — {forn}"
opcoes_pend = df_filtered.apply(label_pend, axis=1).tolist()

with st.form("form_justificativa", clear_on_submit=True):
    col_form1, col_form2 = st.columns([1, 2])
    with col_form1:
        responsavel = st.text_input("Seu nome (quem está preenchendo)", placeholder="Ex: João Silva")
        sel_pendencia = st.selectbox("Selecione a Pendência", ["— Selecione —"] + opcoes_pend)
        sel_justificativa = st.selectbox("Motivo da Pendência", OPCOES_JUSTIFICATIVA)
        prazo_resolucao = st.date_input("📅 Prazo previsto p/ resolução/entrega *", value=None, min_value=date.today(), format="DD/MM/YYYY")
    with col_form2:
        observacao = st.text_area("Observação adicional", height=160,
                                  placeholder="Descreva detalhes adicionais sobre esta pendência...")
        submitted = st.form_submit_button("💾  Salvar Justificativa", type="primary", use_container_width=True)

    if submitted:
        if sel_pendencia == "— Selecione —":
            st.error("⚠️ Selecione uma pendência!")
        elif sel_justificativa == "— Selecione —":
            st.error("⚠️ Selecione um motivo!")
        elif prazo_resolucao is None:
            st.error("⚠️ Informe o prazo previsto para resolução/entrega!")
        else:
            idx_sel = opcoes_pend.index(sel_pendencia)
            row = df_filtered.iloc[idx_sel]
            row_id = str(row.get('ID', idx_sel))
            save_justificativa(row_id, sel_justificativa, observacao, prazo_resolucao, responsavel=responsavel, df_ref=df_filtered)
            st.success("✅ Justificativa salva com sucesso!")
            load_justificativas.clear()
            st.rerun()

# ══════════════════════════════════════════════════════════════════════
# EXPORTAR
# ══════════════════════════════════════════════════════════════════════
st.markdown("---")
col_exp1, col_exp2, _ = st.columns([1,1,2])
with col_exp1:
    exp_cols = [c for c in show_cols if c in df_display_export.columns]
    csv_data = df_display_export[exp_cols].to_csv(index=False).encode('utf-8-sig')
    st.download_button("📥 Exportar Dados (CSV)", csv_data, "pendencias_com_justificativas.csv", "text/csv")
with col_exp2:
    if len(just_df) > 0:
        just_csv = just_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Exportar Justificativas", just_csv, "justificativas.csv", "text/csv")
