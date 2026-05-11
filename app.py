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
        elif "status_lanc" in cl or "status lanc" in cl or "status lancamento" in cl: col_map[c] = "Status"
        elif cl == "status manual": col_map[c] = "Status Manual"
        elif cl == "status" and "Status" not in col_map.values(): col_map[c] = "Status"
        elif "filial" in cl and ("nf" in cl or "nota" in cl): col_map[c] = "Filial NF"
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
    # Remove colunas duplicadas, mantendo apenas a primeira ocorrência
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    if "Valor" in df.columns:
        df["Valor"] = df["Valor"].apply(parse_valor)
    # Conversão robusta de TODAS as colunas de data (aceita texto OU objeto date)
    for date_col in ["Vencimento", "Dt Entrega PC", "Dt Emissão"]:
        if date_col in df.columns:
            df[date_col] = df[date_col].apply(parse_data)
    if "Dias" in df.columns:
        df["Dias"] = pd.to_numeric(df["Dias"].replace("—", ""), errors="coerce").fillna(0).astype(int)
    
    # Normaliza textos da coluna Status
    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str).str.strip()
        df["Status"] = df["Status"].replace({
            "Pendente SF1": "PENDENTE LANÇAMENTO",
            "PENDENTE SF1": "PENDENTE LANÇAMENTO",
            "pendente sf1": "PENDENTE LANÇAMENTO",
            "Pendente sf1": "PENDENTE LANÇAMENTO",
            "Aguardando NF": "AGUARDANDO NF",
        })
    
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
    if pd.isna(valor) or not isinstance(valor, (int, float)):
        return str(valor) if valor else ''
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
        st.session_state['upload_filename'] = uploaded.name
        st.session_state['upload_datetime'] = datetime.now().strftime("%d/%m/%Y às %H:%M")
        # Persiste no Google Sheets para sobreviver a restarts
        try:
            _gc = get_gspread_client()
            _sh = _gc.open_by_key(SHEET_ID)
            try:
                _ws_meta = _sh.worksheet("_meta_upload")
            except:
                _ws_meta = _sh.add_worksheet(title="_meta_upload", rows=2, cols=3)
                _ws_meta.update('A1:C1', [['filename', 'datetime', 'registros']])
            _ws_meta.update('A2:C2', [[uploaded.name, datetime.now().strftime("%d/%m/%Y às %H:%M"), str(len(df))]])
        except:
            pass
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
    hoje_sit = agora_sit.normalize()
    just_df_sit = load_justificativas()

    def _nk(s):
        s = str(s).strip().lstrip('0')
        return s if s not in ['', 'nan', 'None', '—', 'NaN'] else ''

    def _nf(s):
        s = str(s).strip().upper()
        s = ' '.join(s.split())
        return s if s not in ['', 'NAN', 'NONE', '—'] else ''

    # Dicionários (chave, fornecedor) → prazo de resolução, para validar expiração
    pcs_just_sit   = {}   # (pc, forn) → prazo (pd.Timestamp ou None)
    notas_just_sit = {}   # (nota, forn) → prazo
    if len(just_df_sit) > 0:
        for _, r in just_df_sit.iterrows():
            forn = _nf(r.get('Fornecedor', ''))
            if not forn:
                continue
            prazo_raw = r.get('Prazo_Resolucao', '')
            prazo = parse_data(prazo_raw)
            pc_k = _nk(r.get('Nº_PC', ''))
            nota_k = _nk(r.get('Nº_Nota', ''))
            if pc_k:
                pcs_just_sit[(pc_k, forn)] = prazo
            if nota_k:
                notas_just_sit[(nota_k, forn)] = prazo

    def calc_situacao(row):
        # Se está pendente de identificação de responsável, essa é a situação prioritária
        comp = str(row.get('Comprador', '')).strip()
        sol = str(row.get('Solicitante', '')).strip()
        if '⚠️' in comp or '⚠️' in sol or 'Pendente Identificação' in comp or 'Pendente Identificação' in sol:
            return 'Pendente Identificação Responsável'

        venc = parse_data(row.get('Vencimento', pd.NaT))
        forn = _nf(row.get('Fornecedor', ''))
        pc = _nk(row.get('Nº PC', ''))
        nota = _nk(row.get('Nº Nota', ''))

        # Verifica se tem justificativa E se o prazo ainda não expirou
        tem_just = False
        just_expirada = False
        if forn:
            prazo = None
            if pc and (pc, forn) in pcs_just_sit:
                prazo = pcs_just_sit[(pc, forn)]
                tem_just = True
            elif nota and (nota, forn) in notas_just_sit:
                prazo = notas_just_sit[(nota, forn)]
                tem_just = True

            # Se o prazo passou, justificativa não vale mais
            if tem_just and pd.notna(prazo) and prazo < hoje_sit:
                just_expirada = True
                tem_just = False

        if pd.notna(venc) and venc < agora_sit:
            if tem_just:
                return 'Vencido c/ Justificativa'
            if just_expirada:
                return 'Vencido — Justificativa Expirada'
            return 'Vencido s/ Justificativa'
        if 'Dt Entrega PC' in row.index:
            entr = parse_data(row.get('Dt Entrega PC', pd.NaT))
            if pd.notna(entr) and entr < agora_sit:
                if just_expirada:
                    return 'Vencido — Justificativa Expirada'
                return 'Entrega Encerrada'
        if tem_just:
            return 'Em Dia (Justificado)'
        if just_expirada:
            return 'Vencido — Justificativa Expirada'
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
        'Vencido — Justificativa Expirada',
        'Vencido c/ Justificativa',
        'Entrega Encerrada',
        'Em Dia (Justificado)',
        'Pendente'
    ]
    sel_situacao = st.selectbox("🏷️ Situação", sit_opts)

    if 'Filial' in df.columns:
        filial_col = df['Filial']
        if isinstance(filial_col, pd.DataFrame):
            filial_col = filial_col.iloc[:, 0]
        filiais = ["Todas"] + sorted([x for x in filial_col.dropna().astype(str).unique().tolist() if x not in ['—','']])
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
import re as re_header

# Lê metadata: primeiro de session_state, fallback de Google Sheets
data_ref = st.session_state.get('upload_datetime', '')
nome_arq = st.session_state.get('upload_filename', '')
if not data_ref or not nome_arq:
    try:
        _gc_h = get_gspread_client()
        _sh_h = _gc_h.open_by_key(SHEET_ID)
        _ws_meta_h = _sh_h.worksheet("_meta_upload")
        _meta_row = _ws_meta_h.row_values(2)
        if len(_meta_row) >= 2:
            nome_arq = nome_arq or _meta_row[0]
            data_ref = data_ref or _meta_row[1]
    except:
        pass

if not data_ref:
    data_ref = "Aguardando upload"

# Tenta extrair data do nome do arquivo (ex: PENDENCIAS_28042026.xlsx → 28/04/2026)
m_data = re_header.search(r'(\d{2})(\d{2})(\d{4})', nome_arq)
if m_data:
    data_base = f"{m_data.group(1)}/{m_data.group(2)}/{m_data.group(3)}"
else:
    data_base = data_ref

st.markdown(f"""
<div class="header-bar">
    <h1>📊 Painel de Pendências</h1>
    <p>Base: {data_base} &nbsp;|&nbsp; Última atualização: {data_ref} &nbsp;|&nbsp; {len(df)} itens totais &nbsp;|&nbsp; {len(df_filtered)} filtrados</p>
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

# ── Lógica central: linha tem justificativa VÁLIDA se PC+Fornecedor OU Nota+Fornecedor bater
#    E o prazo de resolução não tiver expirado ──
def _norm_key(s):
    s = str(s).strip().lstrip('0')
    return s if s not in ['', 'nan', 'None', '—', 'NaN'] else ''

def _norm_forn(s):
    """Normaliza fornecedor: upper + strip + remove múltiplos espaços."""
    s = str(s).strip().upper()
    s = ' '.join(s.split())
    return s if s not in ['', 'NAN', 'NONE', '—'] else ''

# Dicionários (chave, fornecedor) → prazo de resolução
hoje_flag = pd.Timestamp.now().normalize()
pcs_com_just = {}
notas_com_just = {}
if len(just_df) > 0:
    for _, r in just_df.iterrows():
        forn = _norm_forn(r.get('Fornecedor', ''))
        if not forn:
            continue
        prazo = parse_data(r.get('Prazo_Resolucao', ''))
        pc_k = _norm_key(r.get('Nº_PC', ''))
        nota_k = _norm_key(r.get('Nº_Nota', ''))
        if pc_k:
            pcs_com_just[(pc_k, forn)] = prazo
        if nota_k:
            notas_com_just[(nota_k, forn)] = prazo

def linha_tem_justificativa(row):
    """Só retorna True se a justificativa existe E o prazo ainda não expirou."""
    forn = _norm_forn(row.get('Fornecedor', ''))
    if not forn:
        return False
    pc_k = _norm_key(row.get('Nº PC', ''))
    nota_k = _norm_key(row.get('Nº Nota', ''))
    prazo = None
    if pc_k and (pc_k, forn) in pcs_com_just:
        prazo = pcs_com_just[(pc_k, forn)]
    elif nota_k and (nota_k, forn) in notas_com_just:
        prazo = notas_com_just[(nota_k, forn)]
    else:
        return False
    # Se prazo não preenchido, considera válida; se preenchido, precisa estar no futuro
    if pd.notna(prazo) and prazo < hoje_flag:
        return False
    return True

# Aplica a flag no df_filtered como booleano puro
if 'Fornecedor' in df_filtered.columns and ('Nº PC' in df_filtered.columns or 'Nº Nota' in df_filtered.columns):
    df_filtered['_tem_just'] = df_filtered.apply(linha_tem_justificativa, axis=1).astype(bool)
    com_justificativa = int(df_filtered['_tem_just'].astype(int).sum())
else:
    df_filtered['_tem_just'] = False
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
# 🚨 NOTAS CRÍTICAS — Sem PC e Sem Lançamento (Pendentes de Lançamento)
# ══════════════════════════════════════════════════════════════════════
VAZIOS_PC = ['', '—', 'nan', 'None', 'NaN', '0', '000000']

def _is_pendente_sf1(row):
    """Identifica se a nota está pendente de lançamento."""
    status = str(row.get('Status', '')).strip().upper()
    return 'PENDENTE' in status

def _is_sem_pc(row):
    """Identifica se a nota NÃO tem PC vinculado."""
    pc = str(row.get('Nº PC', '')).strip().lstrip('0')
    return pc in VAZIOS_PC

# Filtra notas críticas: sem lançamento E sem PC
df_criticas = pd.DataFrame()
if 'Status' in df_filtered.columns:
    mask_pendente = df_filtered.apply(_is_pendente_sf1, axis=1)
    mask_sem_pc = df_filtered.apply(_is_sem_pc, axis=1)
    df_criticas = df_filtered[mask_pendente & mask_sem_pc].copy()

n_criticas = len(df_criticas)
val_criticas = df_criticas['Valor'].sum() if 'Valor' in df_criticas.columns and n_criticas > 0 else 0

# KPI Card de alerta
if n_criticas > 0:
    # Top 5 fornecedores com mais notas críticas
    if 'Fornecedor' in df_criticas.columns:
        top_forn = df_criticas.groupby('Fornecedor').agg(
            Qtd=('Fornecedor', 'size'),
            Valor=('Valor', 'sum') if 'Valor' in df_criticas.columns else ('Fornecedor', 'size')
        ).reset_index().sort_values('Qtd', ascending=False).head(5)
        top_forn_html = '<br>'.join([
            f"• {r['Fornecedor'][:35]} — <b>{int(r['Qtd'])} NFs</b> | {format_brl(r['Valor'])}"
            for _, r in top_forn.iterrows()
        ])
    else:
        top_forn_html = ''

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#3a0a0a,#5a1010);border:2px solid #ff4d6a;border-radius:16px;padding:24px 28px;margin-bottom:20px">
        <div style="display:flex;align-items:center;gap:16px;margin-bottom:16px">
            <div style="font-size:2.5rem">🚨</div>
            <div>
                <div style="font-size:1.3rem;font-weight:700;color:#ff4d6a">NFs Emitidas SEM Pedido de Compra e SEM Lançamento</div>
                <div style="font-size:0.85rem;color:#ff8080;margin-top:4px">Notas com status pendentes de lançamento e sem PC vinculado — requerem ação imediata</div>
            </div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;margin-bottom:16px">
            <div style="text-align:center">
                <div style="font-size:0.7rem;color:#8892a4;text-transform:uppercase;letter-spacing:1px">Quantidade</div>
                <div style="font-size:2.4rem;font-weight:700;color:#ff4d6a;line-height:1.1;margin-top:4px">{n_criticas}</div>
                <div style="font-size:0.75rem;color:#ff8080">{fmt_pct(n_criticas, total_itens)} do total</div>
            </div>
            <div style="text-align:center">
                <div style="font-size:0.7rem;color:#8892a4;text-transform:uppercase;letter-spacing:1px">Valor Total</div>
                <div style="font-size:1.8rem;font-weight:700;color:#ff4d6a;line-height:1.1;margin-top:4px">{format_brl(val_criticas)}</div>
                <div style="font-size:0.75rem;color:#ff8080">exposição financeira sem cobertura</div>
            </div>
            <div style="text-align:left">
                <div style="font-size:0.7rem;color:#8892a4;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px">Top fornecedores</div>
                <div style="font-size:0.75rem;color:#e0e0e0;line-height:1.6">{top_forn_html}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Tabela dedicada — notas críticas ordenadas por valor (maior primeiro)
    criticas_cols = [c for c in ['Tipo','Filial','Fornecedor','Filial NF','Nº Nota','Chave Sefaz','Dt Emissão','Vencimento','Valor','Comprador','Solicitante','Status','Controle'] if c in df_criticas.columns]
    df_criticas_show = df_criticas[criticas_cols].copy()
    if 'Valor' in df_criticas_show.columns:
        df_criticas_show = df_criticas_show.sort_values('Valor', ascending=False)

    _fmt_d = lambda x: parse_data(x).strftime('%d/%m/%Y') if pd.notna(parse_data(x)) else ''
    _fmt_v = lambda x: format_brl(x) if pd.notna(x) and isinstance(x, (int, float)) else str(x) if x else ''
    fmt_dict = {}
    if 'Valor' in df_criticas_show.columns:
        fmt_dict['Valor'] = _fmt_v
    if 'Dt Emissão' in df_criticas_show.columns:
        fmt_dict['Dt Emissão'] = _fmt_d
    if 'Vencimento' in df_criticas_show.columns:
        fmt_dict['Vencimento'] = _fmt_d

    st.dataframe(
        df_criticas_show.style.format(fmt_dict),
        use_container_width=True,
        height=min(400, max(200, n_criticas * 38 + 40))
    )

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
    grupos = base[group_col].dropna().unique()
    rows = []
    for g in grupos:
        rows_g = base[base[group_col] == g]
        # Justificativa tem prioridade máxima — usa flag pré-computada
        if '_tem_just' in rows_g.columns:
            just_m = rows_g['_tem_just'].fillna(False)
        else:
            just_m = pd.Series([False]*len(rows_g), index=rows_g.index)
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

def format_abrev(v):
    """Formata valor de forma abreviada: 1.234.567 → R$ 1,2 mi / 45.300 → R$ 45,3 k"""
    if pd.isna(v) or not isinstance(v, (int, float)) or v == 0:
        return 'R$ 0'
    if abs(v) >= 1_000_000:
        return f"R$ {v/1_000_000:,.1f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
    if abs(v) >= 1_000:
        return f"R$ {v/1_000:,.1f} k".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

def criar_fig_segmentado(df_seg, group_col, titulo, eh_valor=False):
    """Cria gráfico de barras horizontais segmentado com total visível fora da barra."""
    df_seg = df_seg.copy()
    df_seg['_Total_Visivel'] = df_seg['Seg_Vencido'] + df_seg['Seg_Entrega'] + df_seg['Seg_Just']
    df_seg = df_seg.sort_values('_Total_Visivel', ascending=True)
    grand_total = df_seg['_Total_Visivel'].sum()

    fig = go.Figure()

    # Segmentos coloridos
    for seg, cor, nome in [
        ('Seg_Just',    '#51cf66', 'Com justificativa'),
        ('Seg_Entrega', '#b197fc', 'Entrega enc.'),
        ('Seg_Vencido', '#ff4d6a', 'Vencido'),
    ]:
        if eh_valor:
            customdata = [[format_brl(v)] for v in df_seg[seg]]
            hover = f'<b>%{{y}}</b><br>{nome}: %{{customdata[0]}}<extra></extra>'
            fig.add_trace(go.Bar(
                y=df_seg[group_col], x=df_seg[seg], name=nome,
                orientation='h', marker=dict(color=cor),
                customdata=customdata, hovertemplate=hover
            ))
        else:
            hover = f'<b>%{{y}}</b><br>{nome}: %{{x}}<extra></extra>'
            fig.add_trace(go.Bar(
                y=df_seg[group_col], x=df_seg[seg], name=nome,
                orientation='h', marker=dict(color=cor),
                hovertemplate=hover
            ))

    # Anotações à direita da barra: valor abreviado + qtd de pendências
    x_max = df_seg['_Total_Visivel'].max() if len(df_seg) > 0 else 0

    # Precisamos da qtd de processos por grupo (mesmo nos gráficos de valor)
    # Usa o Total original (conta de linhas) do montar_segmentos
    for _, row in df_seg.iterrows():
        total_vis = row['_Total_Visivel']
        total_qtd = int(row.get('Seg_Vencido', 0) + row.get('Seg_Entrega', 0) + row.get('Seg_Just', 0)) if not eh_valor else 0

        if eh_valor:
            # Para gráfico de valor: mostra valor abreviado + qtd de processos
            # Busca qtd do grupo no df_filtered
            nome_grupo = row[group_col]
            if group_col in df_filtered.columns:
                _mask_grp = df_filtered[group_col] == nome_grupo
                _mask_vis = df_filtered['_tem_just'] | False
                if 'Vencimento' in df_filtered.columns:
                    _v_tmp = df_filtered['Vencimento'].apply(parse_data)
                    _mask_vis = _mask_vis | (_v_tmp < agora_kpi)
                if 'Dt Entrega PC' in df_filtered.columns:
                    _e_tmp = df_filtered['Dt Entrega PC'].apply(parse_data)
                    _mask_vis = _mask_vis | (_e_tmp < agora_kpi)
                qtd_proc = int((_mask_grp & _mask_vis).sum())
            else:
                qtd_proc = 0
            txt = f"<b>{format_abrev(total_vis)}</b> <span style='color:#8892a4'>| {qtd_proc} pendências</span>"
        else:
            txt = f"<b>{int(total_vis)}</b> <span style='color:#8892a4'>pendências</span>"

        fig.add_annotation(
            x=total_vis, y=row[group_col],
            text=txt, showarrow=False,
            xanchor='left', yanchor='middle',
            xshift=8, font=dict(size=11, color='#e0e0e0')
        )

    fig.update_layout(**PLOT_LAYOUT)
    fig.update_layout(
        barmode='stack', title=titulo,
        height=max(280, len(df_seg) * 55 + 60),
        xaxis=dict(
            showgrid=True,
            gridcolor='rgba(255,255,255,0.05)',
            range=[0, x_max * 1.35] if x_max > 0 else None,
            tickformat=',.0f' if not eh_valor else None,
            # Eixo X abreviado para valores
        ),
        yaxis=dict(showgrid=False, automargin=True),
        margin=dict(l=10, r=50, t=40, b=80),
        legend=dict(orientation='h', y=-0.22, font=dict(size=11))
    )

    # Formata eixo X com labels abreviados e espaçamento adequado
    if eh_valor and x_max > 0:
        import math
        # Calcula intervalos proporcionais ao range (4-5 ticks no máximo)
        magnitude = 10 ** math.floor(math.log10(max(x_max, 1)))
        step = magnitude
        if x_max / step < 3:
            step = magnitude / 2
        elif x_max / step > 6:
            step = magnitude * 2
        
        tick_vals = []
        v = 0
        while v <= x_max * 1.1:
            tick_vals.append(v)
            v += step
        
        tick_text = []
        for v in tick_vals:
            if v == 0:
                tick_text.append('0')
            elif v >= 1_000_000:
                tick_text.append(f"{v/1_000_000:,.1f} mi".replace(",", "."))
            elif v >= 1_000:
                tick_text.append(f"{v/1_000:,.0f} k".replace(",", "."))
            else:
                tick_text.append(f"{v:,.0f}".replace(",", "."))
        
        fig.update_xaxes(tickvals=tick_vals, ticktext=tick_text)

    return fig

# ── VALOR: Solicitante (esq) + Comprador (dir) — lado a lado ──
col_g1, col_g2 = st.columns(2)

if 'Solicitante' in df_filtered.columns:
    df_filtered['Solicitante'] = df_filtered['Solicitante'].astype(str).str.strip()
    base_sol = df_filtered[df_filtered['Solicitante'].notna() & (~df_filtered['Solicitante'].isin(['—','','nan','None']))]

    if len(base_sol) > 0 and 'Valor' in base_sol.columns:
        with col_g1:
            df_sol_val = montar_segmentos(base_sol, 'Solicitante', agg_col='Valor')
            fig_sol_val = criar_fig_segmentado(df_sol_val, 'Solicitante', '💰 Solicitante — Valor Total', eh_valor=True)
            st.plotly_chart(fig_sol_val, use_container_width=True)

if 'Comprador' in df_filtered.columns:
    df_filtered['Comprador'] = df_filtered['Comprador'].astype(str).str.strip()
    base_comp = df_filtered[df_filtered['Comprador'].notna() & (~df_filtered['Comprador'].isin(['—','','nan','None']))]

    if len(base_comp) > 0 and 'Valor' in base_comp.columns:
        with col_g2:
            df_comp_val = montar_segmentos(base_comp, 'Comprador', agg_col='Valor')
            fig_comp_val = criar_fig_segmentado(df_comp_val, 'Comprador', '💰 Comprador — Valor Total', eh_valor=True)
            st.plotly_chart(fig_comp_val, use_container_width=True)



# ══════════════════════════════════════════════════════════════════════
# TABELAS SEPARADAS POR PRIORIDADE
# ══════════════════════════════════════════════════════════════════════
just_df = load_justificativas()

if len(just_df) > 0 and 'Nº_PC' in just_df.columns:
    # Dois dicionários de lookup: chave composta (identificador, fornecedor)
    cols_to_bring = ['Justificativa','Prazo_Resolucao','Observacao','Responsavel']
    lookup_pc = {}
    lookup_nota = {}
    for _, r in just_df.iterrows():
        forn = _norm_forn(r.get('Fornecedor', ''))
        if not forn:
            continue
        data = {c: r.get(c, '') for c in cols_to_bring}
        pc_k = _norm_key(r.get('Nº_PC', ''))
        nota_k = _norm_key(r.get('Nº_Nota', ''))
        if pc_k:
            lookup_pc[(pc_k, forn)] = data
        if nota_k:
            lookup_nota[(nota_k, forn)] = data

    df_display = df_filtered.copy()
    # Inicializa colunas vazias
    for c in cols_to_bring:
        df_display[c] = ''

    # Faz lookup linha a linha: tenta por (PC, fornecedor), depois por (Nota, fornecedor)
    def _buscar_just(row):
        forn = _norm_forn(row.get('Fornecedor', ''))
        if not forn:
            return None
        pc_key = _norm_key(row.get('Nº PC', ''))
        nota_key = _norm_key(row.get('Nº Nota', ''))
        if pc_key and (pc_key, forn) in lookup_pc:
            return lookup_pc[(pc_key, forn)]
        if nota_key and (nota_key, forn) in lookup_nota:
            return lookup_nota[(nota_key, forn)]
        return None

    for idx, row in df_display.iterrows():
        match = _buscar_just(row)
        if match:
            for c in cols_to_bring:
                df_display.at[idx, c] = match.get(c, '')
else:
    df_display = df_filtered.copy()

for col in ['Justificativa','Prazo_Resolucao','Observacao','Responsavel']:
    if col in df_display.columns:
        df_display[col] = df_display[col].fillna('').replace({'None':'','nan':'','NaT':''})

show_cols = [c for c in ['Tipo','Comprador','Solicitante','Filial','Fornecedor','Nº PC','Filial NF','Nº Nota','Chave Sefaz','Controle','Situação','Dt Emissão','Dt Entrega PC','Vencimento','Valor','Justificativa','Observacao','Prazo_Resolucao','Responsavel'] if c in df_display.columns]

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

# ══════════════════════════════════════════════════════════════════════
# Coluna de ALERTA DE VENCIMENTO
# ══════════════════════════════════════════════════════════════════════
def calc_alerta_venc(row):
    """Cria coluna de alerta visual com base no vencimento."""
    venc = parse_data(row.get('Vencimento', pd.NaT))
    if pd.isna(venc):
        return '⚪ Sem vencimento'
    delta = (venc - agora_tab).days
    if delta < 0:
        return f'🔴 VENCIDO há {abs(delta)}d'
    elif delta == 0:
        return '🔴 VENCE HOJE'
    elif delta <= 5:
        return f'🟠 URGENTE {delta}d'
    elif delta <= 10:
        return f'🟡 PRÓXIMO {delta}d'
    else:
        return f'🟢 OK {delta}d'

df_display['⚠️ Vencimento'] = df_display.apply(calc_alerta_venc, axis=1)

# Atualiza show_cols com a nova coluna de alerta logo após Vencimento
show_cols = [c for c in ['Tipo','Comprador','Solicitante','Filial','Fornecedor','Nº PC','Filial NF','Nº Nota','Chave Sefaz','Controle','Situação','Dt Emissão','Dt Entrega PC','Vencimento','⚠️ Vencimento','Valor','Justificativa','Observacao','Prazo_Resolucao','Responsavel'] if c in df_display.columns]

# ══════════════════════════════════════════════════════════════════════
# CATEGORIZAÇÃO DAS TABELAS
# ══════════════════════════════════════════════════════════════════════

# Identifica notas com NF emitida (tem Nº Nota) mas pendentes de lançamento
def _has_nf(row):
    nota = str(row.get('Nº Nota', '')).strip().lstrip('0')
    return nota not in ['', 'nan', 'None', '—']

def _is_pendente_lanc(row):
    status = str(row.get('Status', '')).strip().upper()
    return 'PENDENTE' in status

mask_tem_nf = df_display.apply(_has_nf, axis=1)
mask_pendente = df_display.apply(_is_pendente_lanc, axis=1) if 'Status' in df_display.columns else pd.Series([False]*len(df_display), index=df_display.index)

# Categoria 1: NFs emitidas pendentes de lançamento
df_cat1 = df_display[mask_tem_nf & mask_pendente].copy()
ids_cat1 = set(df_cat1.index)

# Categoria 2: PCs vencidos (sem NF ou com NF já lançada mas vencidos)
if 'Vencimento' in df_display.columns:
    venc_s = df_display['Vencimento'].apply(parse_data)
    mask_vencido = venc_s < agora_tab
    df_cat2 = df_display[mask_vencido & ~df_display.index.isin(ids_cat1)].copy()
else:
    df_cat2 = pd.DataFrame()
ids_cat2 = set(df_cat2.index)

# Categoria 3: Entrega encerrada (Dt Entrega PC já passou)
if 'Dt Entrega PC' in df_display.columns:
    df_display['Dt Entrega PC'] = df_display['Dt Entrega PC'].apply(parse_data)
    mask_entrega_enc = (df_display['Dt Entrega PC'] < agora_tab) & (~df_display.index.isin(ids_cat1 | ids_cat2))
    df_cat3 = df_display[mask_entrega_enc].copy()
else:
    df_cat3 = pd.DataFrame()
ids_cat3 = set(df_cat3.index)

# Categoria 4: Demais pendências
df_cat4 = df_display[~df_display.index.isin(ids_cat1 | ids_cat2 | ids_cat3)].copy()

# ── Ordena cada categoria por urgência de vencimento ──
for _df in [df_cat1, df_cat2, df_cat3, df_cat4]:
    if 'Vencimento' in _df.columns and len(_df) > 0:
        _df['_venc_sort'] = _df['Vencimento'].apply(parse_data)
        _df.sort_values('_venc_sort', ascending=True, inplace=True, na_position='last')

# ══════════════════════════════════════════════════════════════════════
# TABELA 1: NFs EMITIDAS — PENDENTES DE LANÇAMENTO
# ══════════════════════════════════════════════════════════════════════
n_cat1 = len(df_cat1)
val_cat1 = df_cat1['Valor'].sum() if 'Valor' in df_cat1.columns and n_cat1 > 0 else 0

# Conta quantas vencem em 10 dias
n_venc_10d_cat1 = 0
if 'Vencimento' in df_cat1.columns and n_cat1 > 0:
    venc_cat1 = df_cat1['Vencimento'].apply(parse_data)
    n_venc_10d_cat1 = int(((venc_cat1 >= agora_tab) & (venc_cat1 <= em_10_dias)).sum())
    n_ja_venc_cat1 = int((venc_cat1 < agora_tab).sum())

st.markdown(f"""<div style='background:linear-gradient(135deg,#0f1a2e,#152038);border:2px solid #ff8c42;border-radius:12px;padding:18px 22px;margin:24px 0 4px 0'>
    <div style='display:flex;justify-content:space-between;align-items:center'>
        <div>
            <span style='color:#ff8c42;font-size:1.2rem;font-weight:700'>📄 NFs EMITIDAS — Pendentes de Lançamento</span><br>
            <span style='color:#ffa570;font-size:0.8rem'>Notas fiscais emitidas por fornecedores, pendentes de lançamento</span>
        </div>
        <div style='display:flex;gap:20px;text-align:center'>
            <div>
                <div style='font-size:1.8rem;font-weight:700;color:#ff8c42'>{n_cat1}</div>
                <div style='font-size:0.7rem;color:#ffa570'>notas</div>
            </div>
            <div>
                <div style='font-size:1rem;font-weight:700;color:#ff8c42'>{format_brl(val_cat1)}</div>
                <div style='font-size:0.7rem;color:#ffa570'>valor total</div>
            </div>
            <div>
                <div style='font-size:1.8rem;font-weight:700;color:#ff4d6a'>{n_ja_venc_cat1 if n_cat1 > 0 else 0}</div>
                <div style='font-size:0.7rem;color:#ff8080'>já vencidas</div>
            </div>
            <div>
                <div style='font-size:1.8rem;font-weight:700;color:#ffd43b'>{n_venc_10d_cat1}</div>
                <div style='font-size:0.7rem;color:#ffe066'>vencem em 10d</div>
            </div>
        </div>
    </div>
</div>""", unsafe_allow_html=True)

if n_cat1 > 0:
    t_cat1 = clean_table(df_cat1, show_cols)
    st.dataframe(
        t_cat1.style
            .set_properties(**{'background-color':'rgba(255,140,66,0.08)'})
            .format(FMT),
        use_container_width=True, height=min(500, max(250, n_cat1 * 38 + 40))
    )
else:
    st.success("✅ Todas as notas emitidas foram lançadas!")

# ══════════════════════════════════════════════════════════════════════
# TABELA 2: PCs VENCIDOS
# ══════════════════════════════════════════════════════════════════════
n_cat2 = len(df_cat2)
val_cat2 = df_cat2['Valor'].sum() if 'Valor' in df_cat2.columns and n_cat2 > 0 else 0

st.markdown(f"""<div style='background:linear-gradient(135deg,#2a0e0e,#3a1010);border:1.5px solid #ff4d6a;border-radius:12px;padding:16px 20px;margin:24px 0 4px 0'>
    <div style='display:flex;justify-content:space-between;align-items:center'>
        <div>
            <span style='color:#ff4d6a;font-size:1.1rem;font-weight:700'>🔴 VENCIDOS — Ação Imediata Necessária</span>
            <span style='color:#f08090;font-size:0.85rem;margin-left:12px'>({n_cat2} registros | {format_brl(val_cat2)})</span><br>
            <span style='color:#f08090;font-size:0.8rem'>Processos que já passaram do prazo. Ordenados pelo mais urgente.</span>
        </div>
    </div>
</div>""", unsafe_allow_html=True)

if n_cat2 > 0:
    t_cat2 = clean_table(df_cat2, show_cols)
    st.dataframe(
        t_cat2.style
            .set_properties(**{'background-color':'rgba(255,77,106,0.12)'})
            .format(FMT),
        use_container_width=True, height=min(500, max(250, n_cat2 * 38 + 40))
    )
else:
    st.success("✅ Nenhum processo vencido!")

# ══════════════════════════════════════════════════════════════════════
# TABELA 3: ENTREGA ENCERRADA
# ══════════════════════════════════════════════════════════════════════
n_cat3 = len(df_cat3)
val_cat3 = df_cat3['Valor'].sum() if 'Valor' in df_cat3.columns and n_cat3 > 0 else 0

st.markdown(f"""<div style='background:linear-gradient(135deg,#1a1030,#251840);border:1.5px solid #b197fc;border-radius:12px;padding:16px 20px;margin:24px 0 4px 0'>
    <div style='display:flex;justify-content:space-between;align-items:center'>
        <div>
            <span style='color:#b197fc;font-size:1.1rem;font-weight:700'>🟣 ENTREGA ENCERRADA — Prazo de Entrega Expirado</span>
            <span style='color:#cdb4fc;font-size:0.85rem;margin-left:12px'>({n_cat3} registros | {format_brl(val_cat3)})</span><br>
            <span style='color:#cdb4fc;font-size:0.8rem'>Processos cuja data de entrega do PC já passou. Verificar recebimento e lançamento.</span>
        </div>
    </div>
</div>""", unsafe_allow_html=True)

if n_cat3 > 0:
    t_cat3 = clean_table(df_cat3, show_cols)
    st.dataframe(
        t_cat3.style
            .set_properties(**{'background-color':'rgba(177,151,252,0.08)'})
            .format(FMT),
        use_container_width=True, height=min(500, max(250, n_cat3 * 38 + 40))
    )
else:
    st.success("✅ Nenhum processo com entrega encerrada!")

# ══════════════════════════════════════════════════════════════════════
# TABELA 4: DEMAIS PENDÊNCIAS
# ══════════════════════════════════════════════════════════════════════
n_cat4 = len(df_cat4)

st.markdown(f"""<div style='background:linear-gradient(135deg,#111827,#1a2035);border:1.5px solid #4dabf7;border-radius:12px;padding:16px 20px;margin:24px 0 4px 0'>
    <span style='color:#4dabf7;font-size:1.1rem;font-weight:700'>📋 ACOMPANHAMENTO — Demais Pendências</span>
    <span style='color:#8ab8d4;font-size:0.85rem;margin-left:12px'>({n_cat4} registros)</span><br>
    <span style='color:#8ab8d4;font-size:0.8rem'>Processos em andamento. Monitore para evitar evolução para atraso.</span>
</div>""", unsafe_allow_html=True)

if n_cat4 > 0:
    t_cat4 = clean_table(df_cat4, show_cols)
    st.dataframe(t_cat4.style.format(FMT), use_container_width=True, height=min(500, max(250, n_cat4 * 38 + 40)))
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
# 📧 ALERTAS — Enviar cobrança por responsável via Outlook Web
# ══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📧 Enviar Alertas de Pendências</div>', unsafe_allow_html=True)
st.markdown("<span style='color:#8892a4;font-size:0.85rem'>Clique no botão do responsável para abrir o Outlook Web com o email de cobrança já preenchido.</span>", unsafe_allow_html=True)

import urllib.parse as _urlparse

# Monta mapa de pendências vencidas por comprador
_agora_alerta = pd.Timestamp.now()
_df_alerta = df_display.copy()

# Identifica vencidos (vencimento ou entrega passada)
_mask_venc_alerta = pd.Series([False]*len(_df_alerta), index=_df_alerta.index)
if 'Vencimento' in _df_alerta.columns:
    _v = _df_alerta['Vencimento'].apply(parse_data)
    _mask_venc_alerta = _mask_venc_alerta | (_v < _agora_alerta)
if 'Dt Entrega PC' in _df_alerta.columns:
    _e = _df_alerta['Dt Entrega PC'].apply(parse_data)
    _mask_venc_alerta = _mask_venc_alerta | (_e < _agora_alerta)

_df_vencidos_alerta = _df_alerta[_mask_venc_alerta].copy()

# Agrupa por Comprador (responsável principal)
_col_resp = 'Comprador' if 'Comprador' in _df_vencidos_alerta.columns else 'Solicitante'
if _col_resp in _df_vencidos_alerta.columns and len(_df_vencidos_alerta) > 0:
    _responsaveis = _df_vencidos_alerta[_col_resp].dropna().unique()
    _responsaveis = [r for r in _responsaveis if str(r).strip() not in ['', '—', 'nan', 'None', '⚠️ Pendente Identificação']]
    _responsaveis = sorted(_responsaveis)

    if len(_responsaveis) > 0:
        # Input do domínio de email
        dominio_email = st.text_input("Domínio do email corporativo", value="sdflorestal.com.br", placeholder="Ex: suaempresa.com.br",
                                       help="Os emails serão montados como nome.sobrenome@dominio")

        # Link do painel
        link_painel = "https://painel-de-pendencias-jqvwy4tzlek87tqm5z9zam.streamlit.app"

        cols_alerta = st.columns(min(4, len(_responsaveis)))

        for i, resp in enumerate(_responsaveis):
            _pend_resp = _df_vencidos_alerta[_df_vencidos_alerta[_col_resp] == resp]
            _n = len(_pend_resp)
            _val = _pend_resp['Valor'].sum() if 'Valor' in _pend_resp.columns else 0

            # Separa em duas categorias: NFs pendentes de lançamento vs PCs pendentes
            # NFs emitidas pendentes de lançamento: tem Nº Nota preenchido + Status PENDENTE
            _nfs_pendentes = []
            _pcs_pendentes = []

            for _, _r in _pend_resp.iterrows():
                _forn = str(_r.get('Fornecedor', ''))[:40]
                _pc = str(_r.get('Nº PC', '')).strip()
                _nota = str(_r.get('Nº Nota', '')).strip()
                _valor = format_brl(_r.get('Valor', 0))
                _venc = fmt_data(_r.get('Vencimento', ''))
                _status = str(_r.get('Status', '')).strip().upper()
                _pc_display = _pc if _pc not in ['', '—', 'nan', 'None'] else 'Sem PC'
                _nota_display = _nota if _nota not in ['', '—', 'nan', 'None'] else '—'

                linha = f"  • {_forn} | PC: {_pc_display} | NF: {_nota_display} | Valor: {_valor} | Venc: {_venc}"

                # Se tem NF e status pendente → NF emitida pendente de lançamento
                if _nota_display != '—' and 'PENDENTE' in _status:
                    _nfs_pendentes.append(linha)
                else:
                    _pcs_pendentes.append(linha)

            # Monta corpo categorizado
            _body_parts = []

            if _nfs_pendentes:
                _val_nfs = sum([_r.get('Valor', 0) for _, _r in _pend_resp.iterrows()
                               if str(_r.get('Nº Nota', '')).strip() not in ['', '—', 'nan', 'None']
                               and 'PENDENTE' in str(_r.get('Status', '')).upper()])
                _body_parts.append(
                    f"🚨 NOTAS FISCAIS EMITIDAS — PENDENTES DE LANÇAMENTO ({len(_nfs_pendentes)})\n"
                    f"   Valor: {format_brl(_val_nfs)}\n"
                    f"   Notas emitidas por fornecedores que ainda não foram lançadas.\n"
                    f"   AÇÃO NECESSÁRIA: providenciar lançamento imediato.\n\n"
                    + '\n'.join(_nfs_pendentes)
                )

            if _pcs_pendentes:
                _body_parts.append(
                    f"📋 PEDIDOS DE COMPRA PENDENTES ({len(_pcs_pendentes)})\n"
                    f"   Processos com prazo vencido aguardando providências.\n\n"
                    + '\n'.join(_pcs_pendentes)
                )

            _lista_txt = '\n\n'.join(_body_parts) if _body_parts else '(sem pendências detalhadas)'

            _subject = f"⚠️ Pendências em atraso — {_n} processos requerem sua ação"
            _body = f"""Prezado(a) {resp.replace('.', ' ').title()},

Identificamos {_n} processo(s) sob sua responsabilidade com pendências vencidas, totalizando {format_brl(_val)}.

{'⚠️ ATENÇÃO: ' + str(len(_nfs_pendentes)) + ' nota(s) fiscal(is) emitida(s) por fornecedores ainda não foram lançadas no sistema.' if _nfs_pendentes else ''}

{_lista_txt}

Solicitamos que as pendências sejam tratadas com urgência e que a justificativa seja registrada no painel:
{link_painel}

Processos sem justificativa ou com prazo de resolução expirado permanecerão sinalizados como críticos para a gestão.

Atenciosamente,
Controle de Pendências"""

            # Monta URL do Outlook Web
            _email_dest = f"{resp}@{dominio_email}"
            _url_outlook = (
                "https://outlook.office.com/mail/deeplink/compose?"
                + _urlparse.urlencode({
                    'to': _email_dest,
                    'subject': _subject,
                    'body': _body,
                }, quote_via=_urlparse.quote)
            )

            col_idx = i % min(4, len(_responsaveis))
            with cols_alerta[col_idx]:
                st.markdown(f"""
                <a href="{_url_outlook}" target="_blank" style="text-decoration:none">
                    <div style="background:linear-gradient(135deg,#1a2035,#252b3b);border:1px solid #4dabf7;border-radius:12px;padding:14px;text-align:center;margin:6px 0;cursor:pointer;transition:transform 0.2s">
                        <div style="font-size:0.75rem;color:#8892a4;text-transform:uppercase;letter-spacing:1px">📧 Enviar para</div>
                        <div style="font-size:1rem;font-weight:600;color:#4dabf7;margin:4px 0">{resp}</div>
                        <div style="font-size:0.8rem;color:#ff4d6a;font-weight:600">{_n} pendências | {format_brl(_val)}</div>
                    </div>
                </a>
                """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
    else:
        st.success("✅ Nenhum responsável com pendências vencidas!")
else:
    st.info("Sem dados de responsável para gerar alertas.")

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
