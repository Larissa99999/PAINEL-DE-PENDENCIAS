import streamlit as st
import pandas as pd
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
# CSS EXECUTIVO — MINIMALISTA
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
    * { font-family: 'Inter', sans-serif; }
    
    /* Cards de KPI */
    .kpi-card {
        background: var(--secondary-background-color, #f8f9fa);
        border-radius: 8px; padding: 16px 20px;
    }
    .kpi-card.critical { border-left: 3px solid #E24B4A; }
    .kpi-card.warning { border-left: 3px solid #EF9F27; }
    .kpi-card.success { border-left: 3px solid #1D9E75; }
    .kpi-card.info { border-left: 3px solid #378ADD; }
    .kpi-card.neutral { border-left: 3px solid #888780; }
    .kpi-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin: 0 0 4px; font-weight: 500; }
    .kpi-value { font-size: 24px; font-weight: 600; margin: 0; line-height: 1.1; }
    .kpi-sub { font-size: 12px; color: #888; margin: 4px 0 0; }
    
    /* Resumo executivo */
    .exec-summary {
        background: var(--secondary-background-color, #f8f9fa);
        border-radius: 10px; padding: 16px 20px; margin-bottom: 20px;
    }
    
    /* Seções de tabela */
    .table-header {
        border-radius: 0; padding: 12px 16px; margin: 20px 0 4px;
    }
    .table-header.critical { border-left: 3px solid #E24B4A; background: var(--secondary-background-color, #f8f9fa); }
    .table-header.warning { border-left: 3px solid #EF9F27; background: var(--secondary-background-color, #f8f9fa); }
    .table-header.purple { border-left: 3px solid #7F77DD; background: var(--secondary-background-color, #f8f9fa); }
    .table-header.neutral { border-left: 3px solid #888; background: var(--secondary-background-color, #f8f9fa); }
    
    /* Sidebar */
    div[data-testid="stSidebar"] { background: var(--secondary-background-color, #f8f9fa); }
    
    /* Section labels */
    .section-label {
        font-size: 11px; color: #888; text-transform: uppercase;
        letter-spacing: 1.5px; margin: 28px 0 10px; font-weight: 500;
    }
    
    /* Badge de SLA */
    .sla-badge {
        padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 500; display: inline-block;
    }
    .sla-critical { background: #FCEBEB; color: #A32D2D; }
    .sla-warning { background: #FAEEDA; color: #854F0B; }
    .sla-ok { background: #EAF3DE; color: #3B6D11; }
    
    /* Barra de progresso por responsável */
    .resp-bar { background: var(--secondary-background-color, #eee); border-radius: 3px; height: 6px; overflow: hidden; }
    .resp-bar-inner { height: 100%; display: flex; }
    
    /* Email buttons */
    .email-btn {
        background: var(--secondary-background-color, #f8f9fa);
        border: 0.5px solid var(--secondary-background-color, #ddd);
        border-radius: 8px; padding: 10px; text-align: center;
        text-decoration: none; display: block; transition: opacity 0.2s;
    }
    .email-btn:hover { opacity: 0.8; }
    
    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
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
        elif "valor" in cl and "pc" in cl: col_map[c] = "Valor PC"
        elif "valor" in cl and ("nf" in cl or "nota" in cl): col_map[c] = "Valor NF"
        elif "valor" in cl and "Valor NF" not in col_map.values() and "Valor PC" not in col_map.values(): col_map[c] = "Valor"
        elif "vencimento" in cl: col_map[c] = "Vencimento"
        elif "comprador" in cl: col_map[c] = "Comprador"
        elif "solicitante" in cl: col_map[c] = "Solicitante"
        elif "status sf" in cl: col_map[c] = "Status"
        elif "status_lanc" in cl or "status lanc" in cl or "status lancamento" in cl: col_map[c] = "Status"
        elif cl == "status manual": col_map[c] = "Status Manual"
        elif cl == "status" and "Status" not in col_map.values(): col_map[c] = "Status"
        elif "filial" in cl and ("nf" in cl or "nota" in cl): col_map[c] = "Filial NF"
        elif "filial" in cl and "pc" in cl: col_map[c] = "Filial"
        elif "filial" in cl and "Filial" not in col_map.values(): col_map[c] = "Filial"
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
    
    # Unifica colunas de valor: Valor NF + Valor PC → Valor
    # Prioridade: Valor NF se preenchido, senão Valor PC
    if "Valor NF" in df.columns or "Valor PC" in df.columns:
        df["Valor NF"] = df["Valor NF"].apply(parse_valor) if "Valor NF" in df.columns else 0
        df["Valor PC"] = df["Valor PC"].apply(parse_valor) if "Valor PC" in df.columns else 0
        # Usa Valor NF quando > 0, senão Valor PC
        df["Valor"] = df.apply(lambda r: r["Valor NF"] if r.get("Valor NF", 0) > 0 else r.get("Valor PC", 0), axis=1)
    elif "Valor" in df.columns:
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
        # Tenta pegar a data de modificação do arquivo Excel
        _dt_upload = datetime.now().strftime("%d/%m/%Y às %H:%M")
        try:
            import openpyxl
            uploaded.seek(0)
            _wb_meta = openpyxl.load_workbook(uploaded, read_only=True)
            if _wb_meta.properties.modified:
                # Excel salva em UTC — converte para horário de Brasília (UTC-3)
                from datetime import timedelta
                _dt_utc = _wb_meta.properties.modified
                _dt_br = _dt_utc - timedelta(hours=3)
                _dt_upload = _dt_br.strftime("%d/%m/%Y às %H:%M")
            _wb_meta.close()
            uploaded.seek(0)
        except:
            pass
        st.session_state['upload_datetime'] = _dt_upload
        # Persiste no Google Sheets para sobreviver a restarts
        try:
            _gc = get_gsheet_client()
            _sh = _gc.open_by_key(SHEET_ID)
            try:
                _ws_meta = _sh.worksheet("_meta_upload")
            except:
                _ws_meta = _sh.add_worksheet(title="_meta_upload", rows=2, cols=3)
                _ws_meta.update('A1:C1', [['filename', 'datetime', 'registros']])
            _ws_meta.update('A2:C2', [[uploaded.name, _dt_upload, str(len(df))]])
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
# HEADER MINIMALISTA
# ══════════════════════════════════════════════════════════════════════
import re as re_header

data_ref = st.session_state.get('upload_datetime', '')
nome_arq = st.session_state.get('upload_filename', '')
if not data_ref or not nome_arq:
    try:
        _gc_h = get_gsheet_client()
        _sh_h = _gc_h.open_by_key(SHEET_ID)
        _ws_meta_h = _sh_h.worksheet("_meta_upload")
        _meta_row = _ws_meta_h.row_values(2)
        if len(_meta_row) >= 2:
            nome_arq = nome_arq or _meta_row[0]
            data_ref = data_ref or _meta_row[1]
    except:
        pass

# Fallback: se ainda não tem data, usa a data mais recente dos dados
if not data_ref or data_ref == "Aguardando upload":
    if 'Dt Emissão' in df.columns:
        _max_emissao = df['Dt Emissão'].apply(parse_data).dropna().max()
        if pd.notna(_max_emissao):
            data_ref = _max_emissao.strftime("%d/%m/%Y") + " (data dos dados)"
if not data_ref:
    data_ref = "Aguardando upload"

m_data = re_header.search(r'(\d{2})(\d{2})(\d{4})', nome_arq)
data_base = f"{m_data.group(1)}/{m_data.group(2)}/{m_data.group(3)}" if m_data else data_ref

st.markdown(f"**Painel de pendências** · Base: {data_base} · Atualizado: {data_ref} · {len(df)} itens · {len(df_filtered)} filtrados")
st.markdown("---")

# ══════════════════════════════════════════════════════════════════════
# KPIs
# ══════════════════════════════════════════════════════════════════════
agora_now = pd.Timestamp.now()

vencidas = 0; valor_vencido = 0; maior_atraso_kpi = 0
if 'Vencimento' in df_filtered.columns:
    df_filtered['Vencimento'] = df_filtered['Vencimento'].apply(parse_data)
    df_venc_kpi = df_filtered[df_filtered['Vencimento'] < agora_now].copy()
    vencidas = len(df_venc_kpi)
    valor_vencido = df_venc_kpi['Valor'].sum() if 'Valor' in df_venc_kpi.columns else 0
    if vencidas > 0:
        maior_atraso_kpi = int((agora_now.normalize() - df_venc_kpi['Vencimento']).dt.days.max())

entrega_enc = 0
if 'Dt Entrega PC' in df_filtered.columns:
    df_filtered['Dt Entrega PC'] = df_filtered['Dt Entrega PC'].apply(parse_data)
    entrega_enc = len(df_filtered[df_filtered['Dt Entrega PC'] < agora_now])

just_df = load_justificativas()

def _norm_key(s):
    s = str(s).strip().lstrip('0')
    return s if s not in ['', 'nan', 'None', '—', 'NaN'] else ''

def _norm_forn(s):
    s = str(s).strip().upper(); s = ' '.join(s.split())
    return s if s not in ['', 'NAN', 'NONE', '—'] else ''

hoje_flag = pd.Timestamp.now().normalize()
pcs_com_just = {}; notas_com_just = {}
if len(just_df) > 0:
    for _, r in just_df.iterrows():
        forn = _norm_forn(r.get('Fornecedor', ''))
        if not forn: continue
        prazo = parse_data(r.get('Prazo_Resolucao', ''))
        pc_k = _norm_key(r.get('Nº_PC', ''))
        nota_k = _norm_key(r.get('Nº_Nota', ''))
        if pc_k: pcs_com_just[(pc_k, forn)] = prazo
        if nota_k: notas_com_just[(nota_k, forn)] = prazo

def linha_tem_justificativa(row):
    forn = _norm_forn(row.get('Fornecedor', ''))
    if not forn: return False
    pc_k = _norm_key(row.get('Nº PC', ''))
    nota_k = _norm_key(row.get('Nº Nota', ''))
    prazo = None
    if pc_k and (pc_k, forn) in pcs_com_just: prazo = pcs_com_just[(pc_k, forn)]
    elif nota_k and (nota_k, forn) in notas_com_just: prazo = notas_com_just[(nota_k, forn)]
    else: return False
    if pd.notna(prazo) and prazo < hoje_flag: return False
    return True

if 'Fornecedor' in df_filtered.columns and ('Nº PC' in df_filtered.columns or 'Nº Nota' in df_filtered.columns):
    df_filtered['_tem_just'] = df_filtered.apply(linha_tem_justificativa, axis=1).astype(bool)
    com_justificativa = int(df_filtered['_tem_just'].astype(int).sum())
else:
    df_filtered['_tem_just'] = False; com_justificativa = 0

sem_justificativa = len(df_filtered) - com_justificativa

def fmt_pct(num, den):
    if den == 0 or num == 0: return '0%'
    p = num / den * 100
    return f'{p:.1f}%' if p < 1 else f'{round(p)}%'

# Situação
n_venc_sem_just = len(df_filtered[df_filtered['Situação'] == 'Vencido s/ Justificativa']) if 'Situação' in df_filtered.columns else 0
n_venc_com_just = len(df_filtered[df_filtered['Situação'] == 'Vencido c/ Justificativa']) if 'Situação' in df_filtered.columns else 0
n_em_dia = len(df_filtered[df_filtered['Situação'] == 'Em Dia (Justificado)']) if 'Situação' in df_filtered.columns else 0
n_entrega_enc = len(df_filtered[df_filtered['Situação'] == 'Entrega Encerrada']) if 'Situação' in df_filtered.columns else 0
val_venc_sem = df_filtered[df_filtered['Situação'] == 'Vencido s/ Justificativa']['Valor'].sum() if 'Situação' in df_filtered.columns and 'Valor' in df_filtered.columns else 0
val_venc_com = df_filtered[df_filtered['Situação'] == 'Vencido c/ Justificativa']['Valor'].sum() if 'Situação' in df_filtered.columns and 'Valor' in df_filtered.columns else 0

# NFs pendentes de lançamento
n_nfs_pend = 0; val_nfs_pend = 0
if 'Status' in df_filtered.columns and 'Nº Nota' in df_filtered.columns:
    _st = df_filtered['Status'].astype(str).str.upper()
    _nf_ok = ~df_filtered['Nº Nota'].astype(str).str.strip().isin(['', '—', 'nan', 'None'])
    _mask_nf_pend = _nf_ok & _st.str.contains('PENDENTE', na=False)
    n_nfs_pend = int(_mask_nf_pend.sum())
    val_nfs_pend = df_filtered[_mask_nf_pend]['Valor'].sum() if 'Valor' in df_filtered.columns else 0

# Governança
em_aprovacao = len(df_filtered[df_filtered['Controle'].astype(str).str.upper().str.startswith('B')]) if 'Controle' in df_filtered.columns else 0
n_pend_ident = len(df_filtered[df_filtered['Situação'] == 'Pendente Identificação Responsável']) if 'Situação' in df_filtered.columns else 0

total_com_just = n_venc_com_just + n_em_dia
val_com_just = val_venc_com + (df_filtered[df_filtered['Situação'] == 'Em Dia (Justificado)']['Valor'].sum() if 'Situação' in df_filtered.columns and 'Valor' in df_filtered.columns else 0)

# Recalcula totais excluindo "Demais pendências" (PCs com entrega não encerrada)
# Pre-compute para KPIs
_agora_pre = pd.Timestamp.now()
_mask_nf_pre = pd.Series([False]*len(df_filtered), index=df_filtered.index)
if 'Status' in df_filtered.columns and 'Nº Nota' in df_filtered.columns:
    _st_pre = df_filtered['Status'].astype(str).str.upper()
    _nf_pre = ~df_filtered['Nº Nota'].astype(str).str.strip().isin(['', '—', 'nan', 'None'])
    _mask_nf_pre = _nf_pre & _st_pre.str.contains('PENDENTE', na=False)

_mask_servico_pre = df_filtered['Tipo'].astype(str).str.lower().str.contains('servi', na=False) if 'Tipo' in df_filtered.columns else pd.Series([False]*len(df_filtered), index=df_filtered.index)
_mask_servico_pre = _mask_servico_pre & ~_mask_nf_pre

_mask_ativos = _mask_nf_pre | _mask_servico_pre
total_itens = int(_mask_ativos.sum())
total_valor = df_filtered[_mask_ativos]['Valor'].sum() if 'Valor' in df_filtered.columns else 0

# KPIs removidos do topo — foco direto nas tabelas

# ══════════════════════════════════════════════════════════════════════
# MERGE COM JUSTIFICATIVAS + PREPARAÇÃO DAS TABELAS
# ══════════════════════════════════════════════════════════════════════
VAZIOS_PC = ['', '—', 'nan', 'None', 'NaN', '0', '000000']

def format_abrev(v):
    if pd.isna(v) or not isinstance(v, (int, float)) or v == 0: return 'R$ 0'
    if abs(v) >= 1_000_000: return f"R$ {v/1_000_000:,.1f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
    if abs(v) >= 1_000: return f"R$ {v/1_000:,.1f} k".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")



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

show_cols = [c for c in ['Tipo','Solicitante','Comprador','Fornecedor','Nº Nota','Nº PC','Valor','Dt Emissão','Dt Entrega PC','⚠️ Prazo','Filial','STATUS APROV','Chave Sefaz','Status','Justificativa','Observacao','Prazo_Resolucao','Responsavel'] if c in df_display.columns]

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
    'Prazo_Resolucao': fmt_data,
}

def clean_table(df, cols):
    d = df[[c for c in cols if c in df.columns]].copy()
    for c in d.columns:
        d[c] = d[c].replace({'None':'','nan':'','NaT':''}).fillna('')
    return d

agora_tab = pd.Timestamp.now()
em_10_dias = agora_tab + pd.Timedelta(days=10)

# Coluna de alerta de vencimento
def calc_alerta_venc(row):
    venc = parse_data(row.get('Vencimento', pd.NaT))
    entr = parse_data(row.get('Dt Entrega PC', pd.NaT))
    # Usa vencimento se tiver, senão usa Dt Entrega PC
    ref = venc if pd.notna(venc) else entr
    if pd.isna(ref):
        return '⚪ VENCIMENTO NÃO LOCALIZADO'
    delta = (ref - agora_tab).days
    if delta < 0:
        return f'🔴 VENCIDO {abs(delta)}d'
    elif delta == 0:
        return '🔴 VENCE HOJE'
    elif delta <= 5:
        return f'🟠 URGENTE {delta}d'
    elif delta <= 10:
        return f'🟡 {delta}d'
    else:
        return f'🟢 {delta}d'

df_display['⚠️ Prazo'] = df_display.apply(calc_alerta_venc, axis=1)

# Coluna PC unificada: mostra número ou "PC NÃO LOCALIZADO"
def calc_pc_unificado(row):
    pc = str(row.get('Nº PC', '')).strip().lstrip('0')
    if pc in VAZIOS_PC:
        return '⚠️ PC NÃO LOCALIZADO'
    return str(row.get('Nº PC', ''))

df_display['Nº PC'] = df_display.apply(calc_pc_unificado, axis=1)

# Renomeia Controle para STATUS APROV
if 'Controle' in df_display.columns:
    df_display = df_display.rename(columns={'Controle': 'STATUS APROV'})

# ══════════════════════════════════════════════════════════════════════
# TABELA 1: NFs EMITIDAS — PENDENTES DE LANÇAMENTO
# ══════════════════════════════════════════════════════════════════════
def _is_pendente_lanc(row):
    status = str(row.get('Status', '')).strip().upper()
    return 'PENDENTE' in status

def _has_nf(row):
    nota = str(row.get('Nº Nota', '')).strip().lstrip('0')
    return nota not in ['', 'nan', 'None', '—']

mask_tem_nf = df_display.apply(_has_nf, axis=1)
mask_pendente = df_display.apply(_is_pendente_lanc, axis=1) if 'Status' in df_display.columns else pd.Series([False]*len(df_display), index=df_display.index)

df_nfs = df_display[mask_tem_nf & mask_pendente].copy()
if 'Valor' in df_nfs.columns:
    df_nfs = df_nfs.sort_values('Valor', ascending=False)

n_nfs = len(df_nfs)
val_nfs = df_nfs['Valor'].sum() if 'Valor' in df_nfs.columns and n_nfs > 0 else 0
n_sem_pc = int(df_nfs['Nº PC'].astype(str).str.contains('NÃO LOCALIZADO').sum()) if n_nfs > 0 else 0

# Colunas tabela NFs: ordem conforme layout
nf_cols = [c for c in ['Tipo','Filial','Dt Emissão','Nº Nota','Nº PC','Solicitante','Comprador','Fornecedor','Valor','STATUS APROV','Dt Entrega PC','⚠️ Prazo','Chave Sefaz','Justificativa','Observacao','Prazo_Resolucao','Responsavel'] if c in df_nfs.columns]

st.markdown(f"""<div class="table-header critical">
    <p style="font-size:14px;font-weight:600;color:#E24B4A;margin:0">Notas fiscais emitidas — pendentes de lançamento</p>
    <p style="font-size:12px;color:#888;margin:2px 0 0">{n_nfs} notas · {format_brl(val_nfs)} · {n_sem_pc} sem pedido de compra</p>
</div>""", unsafe_allow_html=True)

if n_nfs > 0:
    t_nfs = clean_table(df_nfs, nf_cols)
    st.dataframe(t_nfs.style.format(FMT), use_container_width=True, height=min(600, max(250, n_nfs * 36 + 40)))
else:
    st.success("Todas as notas emitidas foram lançadas.")

# ══════════════════════════════════════════════════════════════════════
# TABELA 2: PCs DE SERVIÇO — PENDENTES DE ATENDIMENTO
# ══════════════════════════════════════════════════════════════════════
# PCs de serviço = Tipo contém "serviço" ou "servico" ou "PC serviço"
# Exclui os que já estão na tabela de NFs
mask_servico = df_display['Tipo'].astype(str).str.lower().str.contains('servi', na=False) if 'Tipo' in df_display.columns else pd.Series([False]*len(df_display), index=df_display.index)
df_pcs = df_display[mask_servico & ~df_display.index.isin(set(df_nfs.index))].copy()
if 'Valor' in df_pcs.columns:
    df_pcs = df_pcs.sort_values('Valor', ascending=False)

n_pcs = len(df_pcs)
val_pcs = df_pcs['Valor'].sum() if 'Valor' in df_pcs.columns and n_pcs > 0 else 0

# Contadores de vencimento
n_pcs_vencidos = 0; n_pcs_10d = 0
if 'Vencimento' in df_pcs.columns and n_pcs > 0:
    _v_pcs = df_pcs['Vencimento'].apply(parse_data)
    n_pcs_vencidos = int((_v_pcs < agora_tab).sum())
    n_pcs_10d = int(((_v_pcs >= agora_tab) & (_v_pcs <= em_10_dias)).sum())
if 'Dt Entrega PC' in df_pcs.columns and n_pcs > 0:
    _e_pcs = df_pcs['Dt Entrega PC'].apply(parse_data)
    n_pcs_entr_enc = int((_e_pcs < agora_tab).sum())
else:
    n_pcs_entr_enc = 0

pc_cols = [c for c in ['Tipo','Solicitante','Comprador','Fornecedor','Valor','STATUS APROV','Dt Entrega PC','⚠️ Prazo','Chave Sefaz','Justificativa','Observacao','Prazo_Resolucao','Responsavel'] if c in df_pcs.columns]

st.markdown(f"""<div class="table-header warning" style="border-left-color:#EF9F27">
    <p style="font-size:14px;font-weight:600;color:#EF9F27;margin:0">Pedidos de compra (serviço) — pendentes de atendimento</p>
    <p style="font-size:12px;color:#888;margin:2px 0 0">{n_pcs} pedidos · {format_brl(val_pcs)} · {n_pcs_vencidos} vencidos · {n_pcs_entr_enc} entrega encerrada · {n_pcs_10d} vencem em 10d</p>
</div>""", unsafe_allow_html=True)

if n_pcs > 0:
    t_pcs = clean_table(df_pcs, pc_cols)
    st.dataframe(t_pcs.style.format(FMT), use_container_width=True, height=min(600, max(250, n_pcs * 36 + 40)))
else:
    st.success("Nenhum PC de serviço pendente.")

# Tabela "Demais pendências" removida — foco apenas em processos com ação necessária

df_display_export = df_display.copy()

# ══════════════════════════════════════════════════════════════════════
# FORMULÁRIO DE JUSTIFICATIVA
# ══════════════════════════════════════════════════════════════════════
st.markdown('<p class="section-label">Preencher justificativa</p>', unsafe_allow_html=True)

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
st.markdown('<p class="section-label">Alertas por email</p>', unsafe_allow_html=True)

import urllib.parse as _urlparse

_agora_alerta = pd.Timestamp.now()
_df_alerta = df_display.copy()

_mask_venc_alerta = pd.Series([False]*len(_df_alerta), index=_df_alerta.index)
if 'Vencimento' in _df_alerta.columns:
    _v = _df_alerta['Vencimento'].apply(parse_data)
    _mask_venc_alerta = _mask_venc_alerta | (_v < _agora_alerta)
if 'Dt Entrega PC' in _df_alerta.columns:
    _e = _df_alerta['Dt Entrega PC'].apply(parse_data)
    _mask_venc_alerta = _mask_venc_alerta | (_e < _agora_alerta)

_mask_nf_pendente = pd.Series([False]*len(_df_alerta), index=_df_alerta.index)
if 'Status' in _df_alerta.columns and 'Nº Nota' in _df_alerta.columns:
    _status_upper = _df_alerta['Status'].astype(str).str.strip().str.upper()
    _nota_preenchida = ~_df_alerta['Nº Nota'].astype(str).str.strip().isin(['', '—', 'nan', 'None'])
    _mask_nf_pendente = _nota_preenchida & _status_upper.str.contains('PENDENTE', na=False)

_df_vencidos_alerta = _df_alerta[_mask_venc_alerta | _mask_nf_pendente].copy()

if '_tem_just' in _df_vencidos_alerta.columns:
    _df_vencidos_alerta = _df_vencidos_alerta[~_df_vencidos_alerta['_tem_just']].copy()

dominio_email = st.text_input("Domínio do email corporativo", value="sdflorestal.com.br",
                               help="Os emails serão montados como nome.sobrenome@dominio")
link_painel = "https://painel-de-pendencias-jqvwy4tzlek87tqm5z9zam.streamlit.app"

EMAIL_OVERRIDES = {
    "igor.costa": "igor.costa@sdbioflor.com.br",
}

def _gerar_botoes_email(df_pend, col_grupo, titulo, key_prefix):
    """Gera botões de email agrupados por coluna (Comprador ou Solicitante)."""
    if col_grupo not in df_pend.columns or len(df_pend) == 0:
        return

    _pessoas = df_pend[col_grupo].dropna().unique()
    _pessoas = sorted([p for p in _pessoas if str(p).strip() not in ['', '—', 'nan', 'None', '⚠️ Pendente Identificação']])

    if not _pessoas:
        return

    st.markdown(f"<p style='font-size:13px;font-weight:500;margin:16px 0 8px;color:#888'>{titulo}</p>", unsafe_allow_html=True)
    _cols = st.columns(min(4, len(_pessoas)))

    for i, resp in enumerate(_pessoas):
        _pend = df_pend[df_pend[col_grupo] == resp]
        _n = len(_pend)
        _val = _pend['Valor'].sum() if 'Valor' in _pend.columns else 0

        # Monta corpo do email
        _linhas = []
        for _idx, (_, _r) in enumerate(_pend.iterrows(), 1):
            _forn = str(_r.get('Fornecedor', ''))[:35]
            _pc = str(_r.get('Nº PC', '')).strip()
            _nota = str(_r.get('Nº Nota', '')).strip()
            _valor_fmt = format_brl(_r.get('Valor', 0))
            _pc_d = f"PC {_pc}" if _pc not in ['', '—', 'nan', 'None'] else '⚠️ PC NÃO LOCALIZADO'
            _nf_d = f"NF {_nota}" if _nota not in ['', '—', 'nan', 'None'] else 'Sem NF'
            _sol = str(_r.get('Solicitante', '')).strip()
            _sol = _sol if _sol not in ['', '—', 'nan', 'None', '⚠️ Pendente Identificação'] else '—'

            _linhas.append(f"{_idx}. {_forn}\n    {_nf_d} | {_pc_d} | {_valor_fmt} | Solic: {_sol}")
            if _idx >= 10:
                break

        _lista = '\n\n'.join(_linhas)
        _rest = max(0, _n - 10)
        _rodape = f"\n\n... e mais {_rest} pendência(s). Consulte o painel." if _rest > 0 else ''
        _nome = resp.replace('.', ' ').title().split()[0]

        _body = f"""{_nome},

⚠️ As notas abaixo já foram emitidas pelos fornecedores e seguem pendentes de lançamento/regularização:

{_lista}{_rodape}

🔎 Os detalhes completos, histórico das tratativas e atualização das pendências devem ser consultados e registrados diretamente no painel:

{link_painel}"""

        _email_dest = EMAIL_OVERRIDES.get(resp, f"{resp}@{dominio_email}")
        _url = "https://outlook.office.com/mail/deeplink/compose?" + _urlparse.urlencode({
            'to': _email_dest, 'subject': "Pendências Fiscais _ Atualizar painel"
        }, quote_via=_urlparse.quote)

        _bid = f"{key_prefix}_{resp.replace('.','_')}"
        col_idx = i % min(4, len(_pessoas))
        with _cols[col_idx]:
            if st.button(f"📋 Copiar — {resp}", key=_bid):
                st.code(_body, language=None)
                st.info("☝️ Copie e cole no corpo do email.")
            st.markdown(f"""<a href="{_url}" target="_blank" class="email-btn">
                <div style="font-size:12px;font-weight:500">✉️ {resp}</div>
                <div style="font-size:11px;color:#E24B4A;font-weight:500;margin-top:4px">{_n} pend. · {format_abrev(_val)}</div>
            </a>""", unsafe_allow_html=True)

# Gera botões por COMPRADOR
_gerar_botoes_email(_df_vencidos_alerta, 'Comprador', '📧 Por Comprador', 'comp')

# Gera botões por SOLICITANTE
_gerar_botoes_email(_df_vencidos_alerta, 'Solicitante', '📧 Por Solicitante', 'solic')

st.markdown("<br>", unsafe_allow_html=True)

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
