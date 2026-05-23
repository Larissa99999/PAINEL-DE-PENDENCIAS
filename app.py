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
    initial_sidebar_state="collapsed"
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
if not data_ref:
    data_ref = "Aguardando upload"

m_data = re_header.search(r'(\d{2})(\d{2})(\d{4})', nome_arq)
data_base = f"{m_data.group(1)}/{m_data.group(2)}/{m_data.group(3)}" if m_data else data_ref

st.markdown(f"**Painel de pendências** · Base: {data_base} · Atualizado: {data_ref} · {len(df)} itens · {len(df_filtered)} filtrados")

# ══════════════════════════════════════════════════════════════════════
# KPIs
# ══════════════════════════════════════════════════════════════════════
agora_now = pd.Timestamp.now()
total_valor = df_filtered['Valor'].sum() if 'Valor' in df_filtered.columns else 0
total_itens = len(df_filtered)

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

sem_justificativa = total_itens - com_justificativa

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

# ══════════════════════════════════════════════════════════════════════
# RESUMO EXECUTIVO
# ══════════════════════════════════════════════════════════════════════
_alerta_parts = []
if n_nfs_pend > 0:
    _alerta_parts.append(f"<strong style='color:#E24B4A'>{n_nfs_pend} notas fiscais</strong> emitidas aguardam lançamento ({format_brl(val_nfs_pend)})")
if n_venc_sem_just > 0:
    _alerta_parts.append(f"<strong style='color:#E24B4A'>{n_venc_sem_just} pedidos</strong> vencidos sem justificativa ({format_brl(val_venc_sem)})")
if em_aprovacao > 0:
    _alerta_parts.append(f"<strong style='color:#EF9F27'>{em_aprovacao}</strong> processos em aprovação")
_alerta_txt = ". ".join(_alerta_parts) + "." if _alerta_parts else "Sem alertas críticos no momento."

st.markdown(f"""<div class="exec-summary">
    <p style="font-size:13px;color:#888;margin:0 0 4px;font-weight:500">RESUMO EXECUTIVO</p>
    <p style="font-size:14px;margin:0;line-height:1.7">{_alerta_txt} Impacto financeiro total: <strong style="color:#E24B4A">{format_brl(total_valor)}</strong>.</p>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# KPIs EXECUTIVOS — 4 CARDS
# ══════════════════════════════════════════════════════════════════════
st.markdown('<p class="section-label">Indicadores principais</p>', unsafe_allow_html=True)
k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(f"""<div class="kpi-card neutral">
        <p class="kpi-label">Pendências totais</p>
        <p class="kpi-value">{total_itens}</p>
        <p class="kpi-sub">{format_brl(total_valor)}</p>
    </div>""", unsafe_allow_html=True)
with k2:
    st.markdown(f"""<div class="kpi-card critical">
        <p class="kpi-label">NFs sem lançamento</p>
        <p class="kpi-value" style="color:#E24B4A">{n_nfs_pend}</p>
        <p class="kpi-sub">{format_brl(val_nfs_pend)}</p>
    </div>""", unsafe_allow_html=True)
with k3:
    st.markdown(f"""<div class="kpi-card critical">
        <p class="kpi-label">Vencidas s/ justificativa</p>
        <p class="kpi-value" style="color:#E24B4A">{n_venc_sem_just}</p>
        <p class="kpi-sub">{format_brl(val_venc_sem)} · atraso máx: {maior_atraso_kpi}d</p>
    </div>""", unsafe_allow_html=True)
with k4:
    st.markdown(f"""<div class="kpi-card success">
        <p class="kpi-label">Com justificativa</p>
        <p class="kpi-value" style="color:#1D9E75">{total_com_just}</p>
        <p class="kpi-sub">{format_brl(val_com_just)}</p>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# NFs CRÍTICAS — Sem PC e Sem Lançamento
# ══════════════════════════════════════════════════════════════════════
VAZIOS_PC = ['', '—', 'nan', 'None', 'NaN', '0', '000000']

def _is_pendente_sf1(row):
    status = str(row.get('Status', '')).strip().upper()
    return 'PENDENTE' in status

def _is_sem_pc(row):
    pc = str(row.get('Nº PC', '')).strip().lstrip('0')
    return pc in VAZIOS_PC

df_criticas = pd.DataFrame()
if 'Status' in df_filtered.columns:
    mask_pendente = df_filtered.apply(_is_pendente_sf1, axis=1)
    mask_sem_pc = df_filtered.apply(_is_sem_pc, axis=1)
    df_criticas = df_filtered[mask_pendente & mask_sem_pc].copy()

n_criticas = len(df_criticas)
val_criticas = df_criticas['Valor'].sum() if 'Valor' in df_criticas.columns and n_criticas > 0 else 0

if n_criticas > 0:
    if 'Fornecedor' in df_criticas.columns:
        top_forn = df_criticas.groupby('Fornecedor').agg(
            Qtd=('Fornecedor', 'size'),
            Valor=('Valor', 'sum') if 'Valor' in df_criticas.columns else ('Fornecedor', 'size')
        ).reset_index().sort_values('Qtd', ascending=False).head(5)
        top_forn_html = ''.join([
            f"<div style='font-size:12px;color:var(--text-color,#333);padding:2px 0'>{r['Fornecedor'][:35]} — <b>{int(r['Qtd'])} NFs</b> · {format_brl(r['Valor'])}</div>"
            for _, r in top_forn.iterrows()
        ])
    else:
        top_forn_html = ''

    st.markdown(f"""<div class="table-header critical" style="margin-top:24px">
        <p style="font-size:14px;font-weight:600;color:#E24B4A;margin:0">NFs emitidas sem pedido de compra e sem lançamento</p>
        <p style="font-size:12px;color:#888;margin:2px 0 0">Notas pendentes de lançamento e sem PC vinculado — requerem ação imediata</p>
        <div style="display:grid;grid-template-columns:1fr 1fr 2fr;gap:16px;margin-top:12px">
            <div><span style="font-size:11px;color:#888;text-transform:uppercase">Quantidade</span><br><span style="font-size:20px;font-weight:600;color:#E24B4A">{n_criticas}</span></div>
            <div><span style="font-size:11px;color:#888;text-transform:uppercase">Valor total</span><br><span style="font-size:16px;font-weight:600;color:#E24B4A">{format_brl(val_criticas)}</span></div>
            <div><span style="font-size:11px;color:#888;text-transform:uppercase">Top fornecedores</span><br>{top_forn_html}</div>
        </div>
    </div>""", unsafe_allow_html=True)

    criticas_cols = [c for c in ['Tipo','Fornecedor','Nº Nota','Dt Emissão','Vencimento','Valor','Comprador','Solicitante','Status'] if c in df_criticas.columns]
    df_criticas_show = df_criticas[criticas_cols].copy()
    if 'Valor' in df_criticas_show.columns:
        df_criticas_show = df_criticas_show.sort_values('Valor', ascending=False)
    _fmt_d = lambda x: parse_data(x).strftime('%d/%m/%Y') if pd.notna(parse_data(x)) else ''
    _fmt_v = lambda x: format_brl(x) if pd.notna(x) and isinstance(x, (int, float)) else str(x) if x else ''
    fmt_dict = {}
    for c in ['Valor']: 
        if c in df_criticas_show.columns: fmt_dict[c] = _fmt_v
    for c in ['Dt Emissão', 'Vencimento']:
        if c in df_criticas_show.columns: fmt_dict[c] = _fmt_d
    st.dataframe(df_criticas_show.style.format(fmt_dict), use_container_width=True, height=min(350, max(180, n_criticas * 36 + 40)))

# ══════════════════════════════════════════════════════════════════════
# IMPACTO FINANCEIRO POR RESPONSÁVEL — Barras CSS leves
# ══════════════════════════════════════════════════════════════════════
agora_kpi = pd.Timestamp.now().normalize()

def montar_segmentos(base, group_col, agg_col=None):
    grupos = base[group_col].dropna().unique()
    rows = []
    for g in grupos:
        rows_g = base[base[group_col] == g]
        if '_tem_just' in rows_g.columns:
            just_m = rows_g['_tem_just'].fillna(False)
        else:
            just_m = pd.Series([False]*len(rows_g), index=rows_g.index)
        venc_m = (rows_g['Vencimento'] < agora_kpi) & ~just_m if 'Vencimento' in rows_g.columns else pd.Series([False]*len(rows_g), index=rows_g.index)
        entr_m = (rows_g['Dt Entrega PC'] < agora_kpi) & ~just_m & ~venc_m if 'Dt Entrega PC' in rows_g.columns else pd.Series([False]*len(rows_g), index=rows_g.index)
        norm_m = ~just_m & ~venc_m & ~entr_m
        if agg_col and agg_col in rows_g.columns:
            rows.append({group_col: g, 'Total': rows_g[agg_col].sum(), 'Qtd': len(rows_g),
                'Seg_Vencido': rows_g[venc_m][agg_col].sum(), 'Seg_Entrega': rows_g[entr_m][agg_col].sum(),
                'Seg_Just': rows_g[just_m][agg_col].sum(), 'Seg_Normal': rows_g[norm_m][agg_col].sum()})
        else:
            rows.append({group_col: g, 'Total': len(rows_g), 'Qtd': len(rows_g),
                'Seg_Vencido': int(venc_m.sum()), 'Seg_Entrega': int(entr_m.sum()),
                'Seg_Just': int(just_m.sum()), 'Seg_Normal': int(norm_m.sum())})
    return pd.DataFrame(rows).sort_values('Total', ascending=False)

def format_abrev(v):
    if pd.isna(v) or not isinstance(v, (int, float)) or v == 0: return 'R$ 0'
    if abs(v) >= 1_000_000: return f"R$ {v/1_000_000:,.1f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
    if abs(v) >= 1_000: return f"R$ {v/1_000:,.1f} k".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

def render_resp_bars(df_seg, group_col, titulo):
    """Renderiza barras CSS por responsável."""
    if len(df_seg) == 0: return
    st.markdown(f"<p style='font-size:13px;color:#888;margin:0 0 10px'>{titulo}</p>", unsafe_allow_html=True)
    max_val = df_seg['Total'].max() if df_seg['Total'].max() > 0 else 1
    for _, row in df_seg.head(8).iterrows():
        total = row['Total']
        qtd = int(row['Qtd'])
        pct_venc = (row['Seg_Vencido'] / total * 100) if total > 0 else 0
        pct_entr = (row['Seg_Entrega'] / total * 100) if total > 0 else 0
        pct_just = (row['Seg_Just'] / total * 100) if total > 0 else 0
        bar_width = max(total / max_val * 100, 2)
        st.markdown(f"""<div style="margin-bottom:8px">
            <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">
                <span>{row[group_col]}</span>
                <span style="color:#888">{format_abrev(total)} · {qtd} pend.</span>
            </div>
            <div class="resp-bar" style="width:100%">
                <div class="resp-bar-inner" style="width:{bar_width}%">
                    <div style="width:{pct_just}%;background:#1D9E75"></div>
                    <div style="width:{pct_entr}%;background:#7F77DD"></div>
                    <div style="width:{pct_venc}%;background:#E24B4A"></div>
                </div>
            </div>
        </div>""", unsafe_allow_html=True)
    # Legenda
    st.markdown("""<div style="display:flex;gap:12px;margin-top:6px;font-size:11px;color:#888">
        <span style="display:flex;align-items:center;gap:3px"><span style="width:8px;height:8px;border-radius:2px;background:#E24B4A"></span>Vencido</span>
        <span style="display:flex;align-items:center;gap:3px"><span style="width:8px;height:8px;border-radius:2px;background:#7F77DD"></span>Entrega enc.</span>
        <span style="display:flex;align-items:center;gap:3px"><span style="width:8px;height:8px;border-radius:2px;background:#1D9E75"></span>Justificado</span>
    </div>""", unsafe_allow_html=True)

st.markdown('<p class="section-label">Impacto financeiro por responsável</p>', unsafe_allow_html=True)
_gc1, _gc2 = st.columns(2)

with _gc1:
    if 'Solicitante' in df_filtered.columns and 'Valor' in df_filtered.columns:
        base_sol = df_filtered[~df_filtered['Solicitante'].astype(str).str.strip().isin(['', '—', 'nan', 'None'])]
        if len(base_sol) > 0:
            df_sol_val = montar_segmentos(base_sol, 'Solicitante', agg_col='Valor')
            render_resp_bars(df_sol_val, 'Solicitante', 'Solicitante')

with _gc2:
    if 'Comprador' in df_filtered.columns and 'Valor' in df_filtered.columns:
        base_comp = df_filtered[~df_filtered['Comprador'].astype(str).str.strip().isin(['', '—', 'nan', 'None'])]
        if len(base_comp) > 0:
            df_comp_val = montar_segmentos(base_comp, 'Comprador', agg_col='Valor')
            render_resp_bars(df_comp_val, 'Comprador', 'Comprador')

# ══════════════════════════════════════════════════════════════════════
# RANKING DE MAIORES IMPACTOS
# ══════════════════════════════════════════════════════════════════════
st.markdown('<p class="section-label">Ranking de maiores impactos</p>', unsafe_allow_html=True)

if 'Fornecedor' in df_filtered.columns and 'Valor' in df_filtered.columns:
    df_ranking = df_filtered.nlargest(10, 'Valor')[['Fornecedor', 'Valor', 'Vencimento', 'Situação', 'Comprador', 'Nº PC', 'Nº Nota']].copy()
    
    ranking_rows = ''
    for _, r in df_ranking.iterrows():
        venc = parse_data(r.get('Vencimento', pd.NaT))
        if pd.notna(venc):
            delta = (venc - agora_kpi).days
            if delta < 0:
                sla_html = f"<span class='sla-badge sla-critical'>{abs(delta)}d atraso</span>"
            elif delta <= 10:
                sla_html = f"<span class='sla-badge sla-warning'>{delta}d</span>"
            else:
                sla_html = f"<span class='sla-badge sla-ok'>{delta}d</span>"
        else:
            sla_html = "<span style='font-size:11px;color:#888'>—</span>"
        
        sit = str(r.get('Situação', ''))
        if 'Vencido' in sit:
            sit_html = "<span class='sla-badge sla-critical'>Vencido</span>"
        elif 'Justificat' in sit or 'Em Dia' in sit:
            sit_html = "<span class='sla-badge sla-ok'>Justificado</span>"
        elif 'Entrega' in sit:
            sit_html = "<span class='sla-badge sla-warning'>Entrega enc.</span>"
        elif 'PENDENTE' in sit.upper():
            sit_html = "<span class='sla-badge sla-warning'>Pend. lanç.</span>"
        else:
            sit_html = f"<span style='font-size:11px;color:#888'>{sit[:20]}</span>"
        
        ranking_rows += f"""<tr style="border-bottom:0.5px solid var(--secondary-background-color,#eee)">
            <td style="padding:8px 10px;font-size:12px">{str(r['Fornecedor'])[:35]}</td>
            <td style="padding:8px 6px;font-size:12px;font-weight:500">{format_brl(r['Valor'])}</td>
            <td style="padding:8px 6px">{sla_html}</td>
            <td style="padding:8px 6px">{sit_html}</td>
            <td style="padding:8px 6px;font-size:12px;color:#888">{str(r.get('Comprador',''))}</td>
        </tr>"""
    
    st.markdown(f"""<div style="border:0.5px solid var(--secondary-background-color,#ddd);border-radius:8px;overflow:hidden">
        <table style="width:100%;border-collapse:collapse">
            <thead><tr style="border-bottom:0.5px solid var(--secondary-background-color,#ddd)">
                <td style="padding:8px 10px;font-size:11px;color:#888;font-weight:500">Fornecedor</td>
                <td style="padding:8px 6px;font-size:11px;color:#888;font-weight:500">Valor</td>
                <td style="padding:8px 6px;font-size:11px;color:#888;font-weight:500">SLA</td>
                <td style="padding:8px 6px;font-size:11px;color:#888;font-weight:500">Status</td>
                <td style="padding:8px 6px;font-size:11px;color:#888;font-weight:500">Responsável</td>
            </tr></thead>
            <tbody>{ranking_rows}</tbody>
        </table>
    </div>""", unsafe_allow_html=True)



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

show_cols = [c for c in ['Tipo','Solicitante','Comprador','Fornecedor','Nº Nota','Nº PC','Valor','Dt Emissão','Dt Entrega PC','Vencimento','Filial','Controle','Chave Sefaz','Status','Justificativa','Observacao','Prazo_Resolucao','Responsavel'] if c in df_display.columns]

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
show_cols = [c for c in ['Tipo','Solicitante','Comprador','Fornecedor','Nº Nota','Nº PC','Valor','Dt Emissão','Dt Entrega PC','Vencimento','⚠️ Vencimento','Filial','Controle','Chave Sefaz','Status','Justificativa','Observacao','Prazo_Resolucao','Responsavel'] if c in df_display.columns]

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
n_venc_10d_cat1 = 0; n_ja_venc_cat1 = 0
if 'Vencimento' in df_cat1.columns and n_cat1 > 0:
    venc_cat1 = df_cat1['Vencimento'].apply(parse_data)
    n_venc_10d_cat1 = int(((venc_cat1 >= agora_tab) & (venc_cat1 <= em_10_dias)).sum())
    n_ja_venc_cat1 = int((venc_cat1 < agora_tab).sum())

st.markdown('<p class="section-label">Tabelas categorizadas</p>', unsafe_allow_html=True)

st.markdown(f"""<div class="table-header critical">
    <p style="font-size:14px;font-weight:500;color:#E24B4A;margin:0">NFs emitidas — pendentes de lançamento</p>
    <p style="font-size:12px;color:#888;margin:2px 0 0">{n_cat1} notas · {format_brl(val_cat1)} · {n_ja_venc_cat1} já vencidas · {n_venc_10d_cat1} vencem em 10d</p>
</div>""", unsafe_allow_html=True)

if n_cat1 > 0:
    t_cat1 = clean_table(df_cat1, show_cols)
    st.dataframe(t_cat1.style.format(FMT), use_container_width=True, height=min(450, max(220, n_cat1 * 36 + 40)))
else:
    st.success("Todas as notas emitidas foram lançadas.")

# ══════════════════════════════════════════════════════════════════════
# TABELA 2: VENCIDOS
# ══════════════════════════════════════════════════════════════════════
n_cat2 = len(df_cat2)
val_cat2 = df_cat2['Valor'].sum() if 'Valor' in df_cat2.columns and n_cat2 > 0 else 0

st.markdown(f"""<div class="table-header critical">
    <p style="font-size:14px;font-weight:500;color:#E24B4A;margin:0">Vencidos — ação imediata</p>
    <p style="font-size:12px;color:#888;margin:2px 0 0">{n_cat2} processos · {format_brl(val_cat2)}</p>
</div>""", unsafe_allow_html=True)

if n_cat2 > 0:
    t_cat2 = clean_table(df_cat2, show_cols)
    st.dataframe(t_cat2.style.format(FMT), use_container_width=True, height=min(450, max(220, n_cat2 * 36 + 40)))
else:
    st.success("Nenhum processo vencido.")

# ══════════════════════════════════════════════════════════════════════
# TABELA 3: ENTREGA ENCERRADA
# ══════════════════════════════════════════════════════════════════════
n_cat3 = len(df_cat3)
val_cat3 = df_cat3['Valor'].sum() if 'Valor' in df_cat3.columns and n_cat3 > 0 else 0

st.markdown(f"""<div class="table-header purple">
    <p style="font-size:14px;font-weight:500;color:#7F77DD;margin:0">Entrega encerrada</p>
    <p style="font-size:12px;color:#888;margin:2px 0 0">{n_cat3} processos · {format_brl(val_cat3)}</p>
</div>""", unsafe_allow_html=True)

if n_cat3 > 0:
    t_cat3 = clean_table(df_cat3, show_cols)
    st.dataframe(t_cat3.style.format(FMT), use_container_width=True, height=min(450, max(220, n_cat3 * 36 + 40)))
else:
    st.success("Nenhum processo com entrega encerrada.")

# ══════════════════════════════════════════════════════════════════════
# TABELA 4: DEMAIS PENDÊNCIAS
# ══════════════════════════════════════════════════════════════════════
n_cat4 = len(df_cat4)

st.markdown(f"""<div class="table-header neutral">
    <p style="font-size:14px;font-weight:500;margin:0">Demais pendências</p>
    <p style="font-size:12px;color:#888;margin:2px 0 0">{n_cat4} processos</p>
</div>""", unsafe_allow_html=True)

if n_cat4 > 0:
    t_cat4 = clean_table(df_cat4, show_cols)
    st.dataframe(t_cat4.style.format(FMT), use_container_width=True, height=min(450, max(220, n_cat4 * 36 + 40)))
else:
    st.info("Sem demais pendências.")

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

# Inclui também NFs pendentes de lançamento (mesmo sem vencimento)
_mask_nf_pendente = pd.Series([False]*len(_df_alerta), index=_df_alerta.index)
if 'Status' in _df_alerta.columns and 'Nº Nota' in _df_alerta.columns:
    _status_upper = _df_alerta['Status'].astype(str).str.strip().str.upper()
    _nota_preenchida = ~_df_alerta['Nº Nota'].astype(str).str.strip().isin(['', '—', 'nan', 'None'])
    _mask_nf_pendente = _nota_preenchida & _status_upper.str.contains('PENDENTE', na=False)

_df_vencidos_alerta = _df_alerta[_mask_venc_alerta | _mask_nf_pendente].copy()

# Filtra: só processos SEM justificativa válida (sem justificativa OU prazo expirado)
if '_tem_just' in _df_vencidos_alerta.columns:
    _df_vencidos_alerta = _df_vencidos_alerta[~_df_vencidos_alerta['_tem_just']].copy()

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
                _venc = _venc if _venc.strip() else 'Sem vencimento'
                _entrega = fmt_data(_r.get('Dt Entrega PC', ''))
                _entrega = _entrega if _entrega.strip() else '—'
                _status = str(_r.get('Status', '')).strip().upper()
                _pc_display = _pc if _pc not in ['', '—', 'nan', 'None'] else 'Sem PC'
                _nota_display = _nota if _nota not in ['', '—', 'nan', 'None'] else '—'

                # Identifica o outro responsável (comprador ou solicitante)
                _comprador = str(_r.get('Comprador', '')).strip()
                _solicitante = str(_r.get('Solicitante', '')).strip()
                _outro_resp = ''
                if _col_resp == 'Comprador' and _solicitante and _solicitante != resp and _solicitante not in ['', '—', 'nan', 'None', '⚠️ Pendente Identificação']:
                    _outro_resp = f" | Solicitante: {_solicitante}"
                elif _col_resp == 'Solicitante' and _comprador and _comprador != resp and _comprador not in ['', '—', 'nan', 'None', '⚠️ Pendente Identificação']:
                    _outro_resp = f" | Comprador: {_comprador}"

                linha = f"  • {_forn} | PC: {_pc_display} | NF: {_nota_display} | Valor: {_valor} | Emissão: {fmt_data(_r.get('Dt Emissão', ''))} | Prev. Entrega: {_entrega} | Venc: {_venc}{_outro_resp}"

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
                    f"   Caso a mercadoria/produto tenham sido entregues, realizar lançamento. Se não, atualizar previsão de entrega no painel de pendências.\n\n"
                    + '\n'.join(_nfs_pendentes)
                )

            if _pcs_pendentes:
                # Separa PCs vencidos dos demais
                _pcs_vencidos_email = []
                _pcs_outros_email = []
                for _, _r in _pend_resp.iterrows():
                    _nota_chk = str(_r.get('Nº Nota', '')).strip()
                    _status_chk = str(_r.get('Status', '')).upper()
                    # Já foi pra NFs pendentes? Pula
                    if _nota_chk not in ['', '—', 'nan', 'None'] and 'PENDENTE' in _status_chk:
                        continue
                    _venc_chk = parse_data(_r.get('Vencimento', pd.NaT))
                    _forn = str(_r.get('Fornecedor', ''))[:40]
                    _pc = str(_r.get('Nº PC', '')).strip()
                    _valor = format_brl(_r.get('Valor', 0))
                    _venc_fmt = fmt_data(_r.get('Vencimento', ''))
                    _venc_fmt = _venc_fmt if _venc_fmt.strip() else 'Sem vencimento'
                    _entrega_fmt = fmt_data(_r.get('Dt Entrega PC', ''))
                    _entrega_fmt = _entrega_fmt if _entrega_fmt.strip() else '—'
                    _pc_display = _pc if _pc not in ['', '—', 'nan', 'None'] else 'Sem PC'
                    _nota_display = str(_r.get('Nº Nota', '')).strip()
                    _nota_display = _nota_display if _nota_display not in ['', '—', 'nan', 'None'] else '—'

                    # Outro responsável
                    _comprador2 = str(_r.get('Comprador', '')).strip()
                    _solicitante2 = str(_r.get('Solicitante', '')).strip()
                    _outro2 = ''
                    if _col_resp == 'Comprador' and _solicitante2 and _solicitante2 != resp and _solicitante2 not in ['', '—', 'nan', 'None', '⚠️ Pendente Identificação']:
                        _outro2 = f" | Solicitante: {_solicitante2}"
                    elif _col_resp == 'Solicitante' and _comprador2 and _comprador2 != resp and _comprador2 not in ['', '—', 'nan', 'None', '⚠️ Pendente Identificação']:
                        _outro2 = f" | Comprador: {_comprador2}"

                    linha = f"  • {_forn} | PC: {_pc_display} | NF: {_nota_display} | Valor: {_valor} | Prev. Entrega: {_entrega_fmt} | Venc: {_venc_fmt}{_outro2}"

                    if pd.notna(_venc_chk) and _venc_chk < _agora_alerta:
                        _dias_atraso = (pd.Timestamp.now() - _venc_chk).days
                        linha += f" | ⚠️ {_dias_atraso} dias em atraso"
                        _pcs_vencidos_email.append(linha)
                    else:
                        _pcs_outros_email.append(linha)

                _pc_section = []
                if _pcs_vencidos_email:
                    _pc_section.append(
                        f"🔴 PEDIDOS DE COMPRA VENCIDOS ({len(_pcs_vencidos_email)})\n"
                        f"   Processos com prazo EXPIRADO — requerem ação imediata.\n\n"
                        + '\n'.join(_pcs_vencidos_email)
                    )
                if _pcs_outros_email:
                    _pc_section.append(
                        f"📋 DEMAIS PEDIDOS PENDENTES ({len(_pcs_outros_email)})\n"
                        f"   Processos com prazo próximo ou entrega encerrada.\n\n"
                        + '\n'.join(_pcs_outros_email)
                    )
                if _pc_section:
                    _body_parts.extend(_pc_section)

            _lista_txt = '\n\n'.join(_body_parts) if _body_parts else '(sem pendências detalhadas)'

            # Limita tamanho total — NFs primeiro (mais críticas), depois PCs
            MAX_ITENS_EMAIL = 15
            _total_linhas = len(_nfs_pendentes) + len(_pcs_pendentes)
            _restantes = 0

            if _total_linhas > MAX_ITENS_EMAIL:
                # Prioriza NFs pendentes de lançamento
                _nfs_limit = _nfs_pendentes[:MAX_ITENS_EMAIL]
                _sobra = MAX_ITENS_EMAIL - len(_nfs_limit)

                _body_parts_limited = []
                if _nfs_limit:
                    _val_nfs = sum([_r.get('Valor', 0) for _, _r in _pend_resp.iterrows()
                                   if str(_r.get('Nº Nota', '')).strip() not in ['', '—', 'nan', 'None']
                                   and 'PENDENTE' in str(_r.get('Status', '')).upper()])
                    _body_parts_limited.append(
                        f"🚨 NOTAS FISCAIS EMITIDAS — PENDENTES DE LANÇAMENTO ({len(_nfs_pendentes)})\n"
                        f"   Valor: {format_brl(_val_nfs)}\n"
                        f"   Caso a mercadoria/produto tenham sido entregues, realizar lançamento. Se não, atualizar previsão de entrega no painel de pendências.\n\n"
                        + '\n'.join(_nfs_limit)
                    )

                if _sobra > 0:
                    # Junta PCs vencidos + outros e limita
                    _all_pcs = []
                    if '_pcs_vencidos_email' in dir():
                        _all_pcs.extend(_pcs_vencidos_email[:_sobra])
                        _sobra2 = _sobra - len(_pcs_vencidos_email[:_sobra])
                        if _sobra2 > 0 and '_pcs_outros_email' in dir():
                            _all_pcs.extend(_pcs_outros_email[:_sobra2])
                    if _all_pcs:
                        _body_parts_limited.append(
                            f"🔴 PEDIDOS PENDENTES (mostrando {len(_all_pcs)} de {len(_pcs_pendentes)})\n\n"
                            + '\n'.join(_all_pcs)
                        )

                _restantes = _total_linhas - MAX_ITENS_EMAIL
                _lista_txt = '\n\n'.join(_body_parts_limited) if _body_parts_limited else _lista_txt

            _subject = f"⚠️ Pendências em atraso — {_n} processos requerem sua ação"
            _rodape_restantes = f"\n\n📌 + {_restantes} processo(s) não listado(s) neste email. Consulte o painel para a lista completa." if _restantes > 0 else ''
            _body = f"""Prezado(a) {resp.replace('.', ' ').title()},

Identificamos {_n} processo(s) sob sua responsabilidade com pendências vencidas, totalizando {format_brl(_val)}.

{'⚠️ ATENÇÃO: ' + str(len(_nfs_pendentes)) + ' nota(s) fiscal(is) emitida(s) por fornecedores ainda não foram lançadas no sistema.' if _nfs_pendentes else ''}

{_lista_txt}{_rodape_restantes}

Solicitamos que as pendências sejam tratadas com urgência e que a justificativa seja registrada no painel:
{link_painel}

Processos sem justificativa ou com prazo de resolução expirado permanecerão sinalizados como críticos para a gestão.

Atenciosamente,
Controle de Pendências"""

            # Monta URL do Outlook Web
            # Mapa de exceções de email (usuários com domínio diferente)
            EMAIL_OVERRIDES = {
                "igor.costa": "igor.costa@sdbioflor.com.br",
            }
            _email_dest = EMAIL_OVERRIDES.get(resp, f"{resp}@{dominio_email}")
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
                st.markdown(f"""<a href="{_url_outlook}" target="_blank" class="email-btn">
                    <div style="font-size:12px;font-weight:500">{resp}</div>
                    <div style="font-size:11px;color:#E24B4A;font-weight:500;margin-top:4px">{_n} pend. · {format_abrev(_val)}</div>
                </a>""", unsafe_allow_html=True)

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
