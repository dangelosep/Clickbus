import pandas as pd
import os
import re
import base64

pasta_atual = os.path.dirname(os.path.abspath(__file__))
caminho_excel = os.path.join(pasta_atual, 'relatorio_precos_clickbus.xlsx')
caminho_html = os.path.join(pasta_atual, 'Dashboard_Executivo.html')
caminho_logo = os.path.join(pasta_atual, 'logo_itapemirim.png')

logo_base64_src = ''
if os.path.exists(caminho_logo):
    try:
        with open(caminho_logo, 'rb') as f_logo:
            logo_base64_src = f"data:image/png;base64,{base64.b64encode(f_logo.read()).decode('utf-8')}"
    except Exception as e:
        print(f"[!] Erro ao carregar logo: {e}")

logo_img_tag = f'<img src="{logo_base64_src}" alt="Nova Itapemirim" class="header-logo">' if logo_base64_src else '<img src="logo_itapemirim.png" alt="Nova Itapemirim" class="header-logo" onerror="this.style.display=\'none\'">'

try:
    df = pd.read_excel(caminho_excel)
except FileNotFoundError:
    print(f"[ERRO] Arquivo não encontrado no caminho: {caminho_excel}")
    exit()

# ==========================================
# 1. LIMPEZA E PADRONIZAÇÃO DE DADOS
# ==========================================
if 'Status da Busca' in df.columns:
    df = df[df['Status da Busca'] == 'Sucesso'].copy()

df['Preço (R$)'] = pd.to_numeric(df['Preço (R$)'], errors='coerce').fillna(0)
df['Assentos Disponíveis'] = df['Assentos Disponíveis'].astype(str).str.extract(r'(\d+)')[0]
df['Assentos Disponíveis'] = pd.to_numeric(df['Assentos Disponíveis'], errors='coerce').fillna(0)

# ========================================================
# MAPEAMENTO DINÂMICO DE CLASSES E ORIGENS 
# ========================================================
coluna_classe_original = 'Classe'

# Detectar coluna PROCV Classes (PROCV CLASSES, PROCV CLASSE, etc.)
colunas_exatas = [c for c in df.columns if re.search(r'^PROCV\s*CLASSES?$', str(c).strip(), re.IGNORECASE)]
colunas_candidatas_classe = [
    c for c in df.columns 
    if c != coluna_classe_original 
    and ('PROCV' in str(c).upper() and ('CLASSE' in str(c).upper() or 'CLASS' in str(c).upper() or 'PADR' in str(c).upper()))
    and 'ORIGEM' not in str(c).upper() 
    and 'DESTINO' not in str(c).upper()
]

if colunas_exatas:
    df['Classe Padronizada'] = df[colunas_exatas[0]].fillna(df[coluna_classe_original]).fillna('INDEFINIDO').astype(str).str.upper().str.strip()
    print(f"[*] Coluna de Classe detectada via: {colunas_exatas[0]}")
elif colunas_candidatas_classe:
    df['Classe Padronizada'] = df[colunas_candidatas_classe[0]].fillna(df[coluna_classe_original]).fillna('INDEFINIDO').astype(str).str.upper().str.strip()
    print(f"[*] Coluna de Classe detectada via: {colunas_candidatas_classe[0]}")
else:
    df['Classe Padronizada'] = df[coluna_classe_original].fillna('INDEFINIDO').astype(str).str.upper().str.strip()
    print("[!] Coluna PROCV de Classes NÃO ENCONTRADA. Usando 'Classe' original.")

# Remove status inválidos ou strings de ausência de viagens, preservando SEMILEITO intacto
df = df[~df['Classe Padronizada'].str.contains(r'SEM\s+(?:ÔNIBUS|ONIBUS|VIAGENS|PASSAGENS)|INDEFINIDO|^NAN$', case=False, regex=True, na=False)]

col_origem_procv = [c for c in df.columns if 'PROCV' in str(c).upper() and 'ORIGEM' in str(c).upper()]
col_destino_procv = [c for c in df.columns if 'PROCV' in str(c).upper() and 'DESTINO' in str(c).upper()]

if col_origem_procv:
    df['Origem'] = df[col_origem_procv[0]].fillna(df['Origem']).astype(str).str.strip()
if col_destino_procv:
    df['Destino'] = df[col_destino_procv[0]].fillna(df['Destino']).astype(str).str.strip()

df['Origem'] = df['Origem'].astype(str).str.replace('-TODOS', '', case=False, regex=False).str.strip()
df['Destino'] = df['Destino'].astype(str).str.replace('-TODOS', '', case=False, regex=False).str.strip()

def extrair_data_hora(valor):
    texto = str(valor)
    date_match = re.search(r"['\"]date['\"]:\s*['\"]([^'\"]+)['\"]", texto)
    time_match = re.search(r"['\"]time['\"]:\s*['\"]([^'\"]+)['\"]", texto)
    if date_match and time_match:
        return f"{date_match.group(1)} {time_match.group(1)}"
    return texto

df['Saída Limpa'] = df['Saída'].apply(extrair_data_hora)
df['Horário Partida'] = pd.to_datetime(df['Saída Limpa'], errors='coerce')
df['Horário Formatado'] = df['Horário Partida'].dt.strftime('%d/%m/%Y %H:%M').fillna('')
df['Hora Apenas'] = df['Horário Partida'].dt.strftime('%H:%M').fillna('')
df['Timestamp'] = ((df['Horário Partida'] - pd.Timestamp("1970-01-01")) // pd.Timedelta('1s')).astype('int64')
df['Horário Partida'] = df['Horário Partida'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')

df = df.dropna(subset=['Origem', 'Destino', 'Viação'])
df = df[(df['Horário Partida'] != '') & (df['Preço (R$)'] > 0)]

# Remove a linha anômala isolada das 12:00 da Itapemirim em SP x RJ (glitch da 2ª raspagem)
m_anomalia_12h = (
    (df['Origem'].str.contains('SAO-PAULO|SÃO PAULO', case=False, regex=True)) & 
    (df['Destino'].str.contains('RIO DE JANEIRO|RIO-DE-JANEIRO', case=False, regex=True)) & 
    (df['Viação'].str.contains('Itapemirim', case=False)) & 
    (df['Hora Apenas'] == '12:00')
)
df = df[~m_anomalia_12h].copy()

# Deduplicação definitiva: usando Trip ID quando disponível, ou atributos textuais como fallback
colunas_dedup = ['Origem', 'Destino', 'Viação', 'Classe Padronizada', 'Horário Partida']
if 'Chegada' in df.columns:
    colunas_dedup.append('Chegada')

if 'Trip ID' in df.columns and (df['Trip ID'].notna()).any():
    m_valid = df['Trip ID'].notna() & (~df['Trip ID'].astype(str).str.upper().isin(['N/A', 'NAN', '']))
    df_com_id = df[m_valid].drop_duplicates(subset=['Trip ID'], keep='first')
    df_sem_id = df[~m_valid].drop_duplicates(subset=colunas_dedup, keep='first')
    df = pd.concat([df_com_id, df_sem_id], ignore_index=True).copy()
else:
    df = df.drop_duplicates(subset=colunas_dedup, keep='first').copy()

data_min = df['Horário Partida'].str[:10].min()
data_max = df['Horário Partida'].str[:10].max()

# Data padrão inicial: hoje se estiver nos dados raspados com volume, senão a data operacional mais recente com volume
contagem_datas = df['Horário Partida'].str[:10].value_counts()
datas_validas = sorted([d for d, c in contagem_datas.items() if c >= 50 and str(d).strip() != ''])
if not datas_validas:
    datas_validas = sorted([d for d in df['Horário Partida'].str[:10].dropna().unique() if str(d).strip() != ''])

hoje_str = pd.Timestamp.now().strftime('%Y-%m-%d')
if hoje_str in datas_validas:
    data_padrao_inicial = hoje_str
elif len(datas_validas) > 0:
    data_padrao_inicial = datas_validas[-1]
else:
    data_padrao_inicial = data_max

colunas_export = ['Origem', 'Destino', 'Viação', 'Classe Padronizada', 'Horário Partida', 'Horário Formatado', 'Hora Apenas', 'Timestamp', 'Preço (R$)', 'Assentos Disponíveis']
if 'Trip ID' in df.columns:
    colunas_export.append('Trip ID')

df_clean = df[colunas_export].copy()
json_data = df_clean.to_json(orient='records')

# ==========================================
# 2. CONSTRUÇÃO DO HTML E CSS
# ==========================================
html_parte1 = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Central de Inteligência Comercial - CIC</title>
    <script src="https://cdn.plot.ly/plotly-2.30.0.min.js"></script>
    <style>
        :root {
            --bg-body: #121212; --bg-card: #1e1e1e; --text-main: #ffffff; --text-muted: #a0a0a0;
            --border-color: #333333; --primary-yellow: #f1c40f; --input-bg: #2c2c2c;
            --alert-red: #e74c3c; --alert-green: #2ecc71; --alert-blue: #3498db; --alert-yellow: #f39c12;
        }
        body.light-mode {
            --bg-body: #f4f7f6; --bg-card: #ffffff; --text-main: #2c3e50; --text-muted: #7f8c8d;
            --border-color: #e1e8ed; --primary-yellow: #d4ac0d; --input-bg: #ffffff;
            --alert-yellow: #d4ac0d;
        }

        body { font-family: 'Segoe UI', Roboto, Helvetica, sans-serif; background-color: var(--bg-body); color: var(--text-main); margin: 0; padding: 14px 20px; transition: 0.3s; font-size: 15px; }
        
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; border-bottom: 1px solid var(--border-color); padding-bottom: 7px; }
        .header-brand { display: flex; align-items: center; gap: 14px; }
        .header-logo { height: 36px; max-height: 38px; width: auto; object-fit: contain; border-radius: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.25); }
        .header h1 { margin: 0; font-size: 23px; font-weight: 700; color: var(--primary-yellow); } 
        .theme-toggle, .nav-toggle { background: transparent; border: 1px solid var(--border-color); color: var(--text-main); padding: 7px 14px; border-radius: 5px; cursor: pointer; font-weight: bold; transition: 0.3s; font-size: 13.5px;}
        .theme-toggle:hover, .nav-toggle:hover { border-color: var(--primary-yellow); color: var(--primary-yellow); }
        
        .filter-container { background-color: var(--bg-card); padding: 10px 14px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid var(--primary-yellow); display: flex; gap: 12px; flex-wrap: wrap; box-shadow: 0 4px 6px rgba(0,0,0,0.1); align-items: flex-end; }
        .filter-group { display: flex; flex-direction: column; flex: 1; min-width: 135px; position: relative; }
        .filter-label-row { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }
        .filter-label-row label { font-size: 12px; color: var(--text-muted); text-transform: uppercase; font-weight: 700; margin: 0; letter-spacing: 0.3px; }
        .filter-clear-link { font-size: 11px; color: var(--primary-yellow); cursor: pointer; text-transform: uppercase; font-weight: 700; padding: 1.5px 5px; border-radius: 3px; background: rgba(241, 196, 15, 0.08); border: 1px solid rgba(241, 196, 15, 0.2); transition: all 0.2s; user-select: none; }
        .filter-clear-link:hover { background: rgba(241, 196, 15, 0.2); color: #ffffff; }
        .filtro-input { background-color: var(--input-bg); color: var(--text-main); border: 1px solid var(--border-color); padding: 7px 12px; border-radius: 4px; font-size: 13.5px; outline: none; width: 100%; box-sizing: border-box;}
        .filtro-input:focus { border-color: var(--primary-yellow); }
        
        .ms-button { background-color: var(--input-bg); color: var(--text-main); border: 1px solid var(--border-color); padding: 7px 12px; border-radius: 4px; font-size: 13.5px; font-weight: 600; cursor: pointer; display: flex; justify-content: space-between; align-items: center; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; user-select: none; }
        .ms-button::after { content: '▼'; font-size: 10px; margin-left: 8px; color: var(--text-muted); }
        .ms-dropdown { position: absolute; top: 100%; left: 0; right: 0; background-color: var(--input-bg); border: 1px solid var(--border-color); z-index: 1000; max-height: 350px; display: none; flex-direction: column; border-radius: 4px; margin-top: 5px; box-shadow: 0 8px 16px rgba(0,0,0,0.5); }
        .ms-dropdown.show { display: flex; }
        .ms-search-box { padding: 8px; border-bottom: 1px solid var(--border-color); position: sticky; top: 0; background-color: var(--input-bg); z-index: 2; border-radius: 4px 4px 0 0; display: flex; align-items: center; position: relative;}
        .ms-search-box input { width: 100%; padding: 6px 26px 6px 8px; box-sizing: border-box; background: var(--bg-body); border: 1px solid var(--border-color); color: var(--text-main); border-radius: 4px; outline: none; font-size: 12.5px; }
        .ms-search-box input:focus { border-color: var(--primary-yellow); }
        .ms-search-clear-btn { position: absolute; right: 14px; cursor: pointer; color: var(--text-muted); font-size: 12px; font-weight: bold; display: none; user-select: none; padding: 2px 5px; border-radius: 3px; }
        .ms-search-clear-btn:hover { color: var(--alert-red); background: rgba(255,255,255,0.1); }
        .ms-dropdown-actions { display: flex; justify-content: space-between; padding: 5px 8px; border-bottom: 1px solid rgba(255,255,255,0.05); background: rgba(0,0,0,0.15); font-size: 11px; }
        .ms-action-link { color: var(--text-muted); cursor: pointer; text-transform: uppercase; font-weight: 600; padding: 2px 4px; border-radius: 3px; transition: color 0.2s; }
        .ms-action-link:hover { color: var(--primary-yellow); }
        .ms-options { overflow-y: auto; display: flex; flex-direction: column; padding-bottom: 5px; }
        .ms-option { padding: 6px 10px; display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12.5px; transition: 0.2s;}
        .ms-option:hover { background-color: rgba(255,255,255,0.05); }
        .ms-option input { cursor: pointer; accent-color: var(--primary-yellow); width: 15px; height: 15px; margin: 0; }

        .btn-clear-all { background: rgba(255,255,255,0.05); color: var(--text-main); border: 1px solid var(--border-color); padding: 7px 12px; border-radius: 4px; font-size: 12.5px; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 5px; transition: all 0.2s ease; height: 34px; white-space: nowrap; box-sizing: border-box; }
        .btn-clear-all:hover { background: rgba(231, 76, 60, 0.15); border-color: var(--alert-red); color: var(--alert-red); }

        .kpi-section-title { font-size: 14px; color: var(--text-muted); text-transform: uppercase; margin: 18px 0 8px 0; letter-spacing: 1px; font-weight: bold; display: flex; justify-content: space-between; align-items: flex-end;}
        .kpi-container { display: flex; justify-content: space-between; flex-wrap: wrap; gap: 12px; }
        .kpi-card { background-color: var(--bg-card); padding: 14px 16px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); flex: 1; min-width: 170px; border-top: 3px solid; }
        .kpi-label { font-size: 11px; color: var(--text-muted); text-transform: uppercase; font-weight: 600; }
        .kpi-value { font-size: 24px; font-weight: bold; margin-top: 6px; color: var(--text-main); display: flex; align-items: baseline; gap: 6px; flex-wrap: wrap; }
        .kpi-value span.prefix { color: var(--primary-yellow); font-size: 16px; }
        .kpi-sub { font-size: 12px; font-weight: normal; width: 100%; display: block; margin-top: 4px; }
        
        .tabs-container { margin-top: 14px; border-bottom: 2px solid var(--border-color); display: flex; gap: 6px; flex-wrap: wrap; }
        .tab-btn { background: transparent; border: none; color: var(--text-muted); padding: 9px 16px; font-size: 13px; font-weight: bold; cursor: pointer; text-transform: uppercase; border-bottom: 3px solid transparent; transition: 0.3s; margin-bottom: -2px; }
        .tab-btn:hover { color: var(--text-main); }
        .tab-btn.active { color: var(--primary-yellow); border-bottom: 3px solid var(--primary-yellow); }
        .tab-content { display: none; margin-top: 14px; }
        .tab-content.active { display: block; }

        .chart-container { background-color: var(--bg-card); padding: 16px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); width: 100%; box-sizing: border-box; margin-bottom: 20px;}
        
        .table-wrapper { 
            background-color: var(--bg-card); 
            border-radius: 8px; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
            height: calc(100vh - 460px); 
            min-height: 350px; 
            overflow-y: auto; 
            margin-bottom: 10px; 
            box-sizing: border-box;
            border: 1px solid var(--border-color);
        }
        .data-table { width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }
        .data-table th, .data-table td { padding: 11px 14px; border-bottom: 1px solid var(--border-color); }
        .data-table th { background-color: var(--input-bg); color: var(--text-muted); font-weight: bold; position: sticky; top: 0; z-index: 10; text-transform: uppercase; font-size: 11px; }
        .data-table tr:hover { background-color: rgba(255,255,255,0.02); }
        .light-mode .data-table tr:hover { background-color: rgba(0,0,0,0.02); }
        
        .pricing-container { 
            background-color: var(--bg-card); 
            border-radius: 8px; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
            height: calc(100vh - 410px); 
            min-height: 380px; 
            overflow-y: auto; 
            box-sizing: border-box;
            border: 1px solid var(--border-color);
        }
        .pricing-table { width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }
        .pricing-table th, .pricing-table td { padding: 11px 14px; border-bottom: 1px solid var(--border-color); }
        .pricing-table th { background-color: var(--input-bg); color: var(--text-muted); font-weight: bold; position: sticky; top: 0; z-index: 10; text-transform: uppercase; font-size: 11px; }
        .pricing-table th.sortable-th, .data-table th.sortable-th { cursor: pointer; user-select: none; transition: background 0.2s, color 0.2s; }
        .pricing-table th.sortable-th:hover, .data-table th.sortable-th:hover { background-color: rgba(241, 196, 15, 0.12); color: var(--primary-yellow); }
        .pricing-table th.active-sort, .data-table th.active-sort { color: var(--primary-yellow); border-bottom: 2px solid var(--primary-yellow); }
        .pricing-table tr:hover { background-color: rgba(255,255,255,0.02); }
        .sort-arrow { display: inline-block; margin-left: 4px; font-size: 10px; opacity: 0.6; }
        /* =======================================================
           NOVO DESIGN INLINE COMPARATIVO: VISÃO POR CLASSES (ALTA DENSIDADE & BAIXA CARGA COGNITIVA)
           ======================================================= */
        .classes-header-bar { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            margin-bottom: 8px; 
            padding-bottom: 5px; 
            border-bottom: 1px solid var(--border-color); 
            flex-wrap: wrap; 
            gap: 8px; 
        }
        .classes-title-box h2 { 
            margin: 0; 
            font-size: 18px; 
            text-transform: uppercase; 
            letter-spacing: 0.5px; 
            color: var(--primary-yellow); 
        }
        .classes-subtitle { 
            font-size: 13px; 
            color: var(--text-muted); 
            display: block; 
            margin-top: 2px; 
        }
        .classes-summary-strip { 
            display: flex; 
            gap: 8px; 
            flex-wrap: wrap; 
        }
        .summary-chip { 
            background: var(--bg-card); 
            border: 1px solid var(--border-color); 
            border-radius: 5px; 
            padding: 5px 12px; 
            display: flex; 
            flex-direction: column; 
            gap: 2px; 
        }
        .summary-chip .chip-lbl { 
            font-size: 11px; 
            color: var(--text-muted); 
            text-transform: uppercase; 
            font-weight: bold; 
        }
        .summary-chip .chip-val { 
            font-size: 14px; 
            color: var(--text-main); 
        }

        /* Deck Panorâmico em Grid Responsivo em 2 Linhas */
        .classes-deck-panoramic { 
            display: flex; 
            flex-wrap: wrap; 
            gap: 12px; 
            width: 100%; 
            align-items: stretch; 
            justify-content: center; 
            box-sizing: border-box; 
        }

        .class-panoramic-card { 
            background: var(--bg-card); 
            border: 1px solid var(--border-color); 
            border-radius: 8px; 
            padding: 13px 15px; 
            display: flex; 
            flex-direction: column; 
            justify-content: space-between; 
            box-shadow: 0 4px 10px rgba(0,0,0,0.25); 
            transition: transform 0.2s ease, border-color 0.2s ease, box-shadow 0.2s ease; 
            box-sizing: border-box;
            height: calc((100vh - 240px) / 2); 
            min-height: 250px; 
        }
        .class-panoramic-card:hover { 
            transform: translateY(-2px); 
            border-color: #555555; 
            box-shadow: 0 6px 14px rgba(0,0,0,0.35); 
        }

        /* Distribuição Balanceada em 2 Linhas (4 em cima, 3 embaixo) */
        .class-panoramic-card.card-r4 { 
            flex: 1 1 calc(25% - 12px); 
            max-width: calc(25% - 9px); 
            min-width: 260px; 
        }
        .class-panoramic-card.card-r3 { 
            flex: 1 1 calc(33.333% - 12px); 
            max-width: calc(33.333% - 9px); 
            min-width: 320px; 
        }
        .class-panoramic-card.card-r2 { 
            flex: 1 1 calc(50% - 12px); 
            max-width: calc(50% - 9px); 
            min-width: 360px; 
        }
        .class-panoramic-card.card-r2-wide { 
            flex: 1 1 calc(50% - 12px); 
            max-width: 540px; 
            min-width: 360px; 
        }
        .class-panoramic-card.card-r1-solo { 
            flex: 0 1 540px; 
            width: 100%; 
            max-width: 540px; 
        }

        /* 1. Header do Card */
        .class-card-header { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            border-bottom: 1px solid var(--border-color); 
            padding-bottom: 8px; 
        }
        .class-title { 
            font-size: 17px; 
            font-weight: 800; 
            color: #ffffff; 
            text-transform: uppercase; 
            letter-spacing: 0.5px;
            white-space: nowrap; 
            overflow: hidden; 
            text-overflow: ellipsis; 
        }
        .status-tag { 
            font-size: 12px; 
            font-weight: 700; 
            padding: 3px 9px; 
            border-radius: 4px; 
            white-space: nowrap; 
            text-transform: uppercase; 
            letter-spacing: 0.4px;
            border: 1px solid var(--border-color); 
            color: var(--text-muted); 
            background: rgba(255,255,255,0.02);
            flex-shrink: 0;
        }
        .status-tag.competitive { 
            color: var(--primary-yellow); 
            border-color: rgba(241, 196, 15, 0.4); 
            background: rgba(241, 196, 15, 0.08); 
        }
        .status-tag.aggressive { 
            color: var(--alert-red); 
            border-color: rgba(231, 76, 60, 0.4); 
            background: rgba(231, 76, 60, 0.08); 
        }
        .status-tag.premium { 
            color: var(--alert-green); 
            border-color: rgba(46, 204, 113, 0.4); 
            background: rgba(46, 204, 113, 0.08); 
        }

        /* 2. Matriz de Preços (Tabela Executiva de 3 Colunas) */
        .price-table-box {
            background: rgba(0, 0, 0, 0.22);
            border: 1px solid rgba(255, 255, 255, 0.04);
            border-radius: 6px;
            padding: 8px 10px;
            display: flex;
            flex-direction: column;
            gap: 6px;
        }
        .price-table-header {
            display: grid;
            grid-template-columns: 95px 1fr 1.35fr;
            font-size: 11.5px;
            text-transform: uppercase;
            color: var(--text-muted);
            font-weight: 700;
            letter-spacing: 0.5px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.06);
            padding-bottom: 4px;
        }
        .price-table-header .th-ita {
            color: var(--primary-yellow);
            text-align: right;
        }
        .price-table-header .th-mkt {
            text-align: right;
            padding-right: 10px;
        }

        .price-row {
            display: grid;
            grid-template-columns: 95px 1fr 1.35fr;
            align-items: center;
            font-size: 15.5px;
        }
        .price-lbl {
            font-size: 12.5px;
            font-weight: 600;
            color: var(--text-muted);
            text-transform: uppercase;
        }
        .price-val-mkt {
            text-align: right;
            color: var(--text-main);
            font-weight: 600;
            padding-right: 10px;
            white-space: nowrap;
            font-size: 15px;
        }
        .price-val-ita {
            text-align: right;
            color: var(--primary-yellow);
            font-weight: 700;
            display: flex;
            justify-content: flex-end;
            align-items: center;
            gap: 5px;
            white-space: nowrap;
            font-size: 15.5px;
        }
        .price-val-ita.no-op {
            color: var(--text-muted);
            opacity: 0.5;
            font-style: italic;
        }
        .gap-badge-inline {
            font-size: 11.5px;
            font-weight: 800;
            padding: 1px 5px;
            border-radius: 3px;
            letter-spacing: 0.1px;
            white-space: nowrap;
        }
        .gap-pos { color: var(--alert-green); background: rgba(46, 204, 113, 0.14); }
        .gap-neg { color: var(--alert-red); background: rgba(231, 76, 60, 0.14); }
        .gap-neutral { color: var(--text-muted); background: rgba(255, 255, 255, 0.05); }

        /* 3. Bloco de Capacidade & Volume (2 Mini-Cards em Alto Destaque) */
        .capacity-container {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            margin: 4px 0;
        }
        .capacity-card {
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid rgba(255, 255, 255, 0.05);
            border-radius: 6px;
            padding: 8px 10px;
            display: flex;
            flex-direction: column;
            gap: 4px;
        }
        .capacity-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .capacity-title {
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            color: var(--text-muted);
            letter-spacing: 0.5px;
            display: flex;
            align-items: center;
            gap: 4px;
        }
        .capacity-share-badge {
            font-size: 13.5px;
            font-weight: 800;
            color: var(--primary-yellow);
        }
        .capacity-numbers {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            font-size: 15.5px;
            font-weight: 700;
            color: #ffffff;
        }
        .capacity-numbers .sub-total {
            font-size: 12px;
            color: var(--text-muted);
            font-weight: 500;
        }
        .capacity-track {
            background: #2a2a2a;
            border-radius: 3px;
            height: 6px;
            width: 100%;
            overflow: hidden;
            margin-top: 2px;
            border: 1px solid rgba(255, 255, 255, 0.02);
        }
        .capacity-fill {
            background: var(--primary-yellow);
            height: 100%;
            border-radius: 3px;
            transition: width 0.3s ease;
        }

        /* 4. Ghost Button Centralizado no Rodapé */
        .class-card-footer { 
            display: flex; 
            justify-content: center; 
            padding-top: 4px; 
        }
        .ghost-action-btn { 
            width: 100%; 
            background: transparent; 
            border: 1px solid var(--border-color); 
            color: var(--text-muted); 
            border-radius: 5px; 
            padding: 7px 12px; 
            font-size: 12px; 
            font-weight: 600; 
            text-transform: uppercase; 
            letter-spacing: 0.5px; 
            cursor: pointer; 
            transition: all 0.2s ease; 
            text-align: center; 
            display: flex; 
            align-items: center; 
            justify-content: center; 
            gap: 5px; 
        }
        .ghost-action-btn:hover { 
            background: var(--primary-yellow); 
            border-color: var(--primary-yellow); 
            color: #121212; 
            font-weight: 700; 
            box-shadow: 0 2px 8px rgba(241, 196, 15, 0.3); 
        }

        /* Estilo Interativo dos Cards de Resumo do Pricing (Verde e Vermelho) */
        .pricing-kpi-card {
            cursor: pointer;
            user-select: none;
            transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease, opacity 0.2s ease;
        }
        .pricing-kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 14px rgba(0,0,0,0.25);
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="header-brand">
            {LOGO_IMG_TAG}
            <h1>Central de Inteligência Comercial - CIC</h1>
        </div>
        <button class="theme-toggle" onclick="toggleTheme()">☀️ Modo Claro</button>
    </div>
    
    <div class="filter-container">
        <div class="filter-group" style="max-width: 145px;">
            <div class="filter-label-row">
                <label>Período Inicial</label>
            </div>
            <input type="date" id="filtro_data_inicio" value="{DATA_PADRAO}" class="filtro-input">
        </div>
        <div class="filter-group" style="max-width: 145px;">
            <div class="filter-label-row">
                <label>Período Final</label>
            </div>
            <input type="date" id="filtro_data_fim" value="{DATA_PADRAO}" class="filtro-input">
        </div>
        <div class="filter-group">
            <div class="filter-label-row">
                <label>Origem</label>
                <span class="filter-clear-link" id="badge-reset-origem" onclick="clearSingleFilter('origem')" style="display:none;" title="Limpar filtro de Origem">✕ Limpar</span>
            </div>
            <div class="ms-button" id="btn-origem" onclick="toggleDropdown('origem')">TODAS</div>
            <div class="ms-dropdown" id="dd-origem">
                <div class="ms-search-box">
                    <input type="text" id="search-origem" placeholder="Buscar..." oninput="filterOptionsList('origem')" onkeyup="filterOptionsList('origem')">
                    <span class="ms-search-clear-btn" id="clear-search-origem" onclick="clearSearchInput('origem')" title="Limpar busca">✕</span>
                </div>
                <div class="ms-dropdown-actions">
                    <span class="ms-action-link" onclick="clearSingleFilter('origem')">✕ Limpar Seleção</span>
                </div>
                <div class="ms-options" id="opt-origem"></div>
            </div>
        </div>
        <div class="filter-group">
            <div class="filter-label-row">
                <label>Destino</label>
                <span class="filter-clear-link" id="badge-reset-destino" onclick="clearSingleFilter('destino')" style="display:none;" title="Limpar filtro de Destino">✕ Limpar</span>
            </div>
            <div class="ms-button" id="btn-destino" onclick="toggleDropdown('destino')">TODOS</div>
            <div class="ms-dropdown" id="dd-destino">
                <div class="ms-search-box">
                    <input type="text" id="search-destino" placeholder="Buscar..." oninput="filterOptionsList('destino')" onkeyup="filterOptionsList('destino')">
                    <span class="ms-search-clear-btn" id="clear-search-destino" onclick="clearSearchInput('destino')" title="Limpar busca">✕</span>
                </div>
                <div class="ms-dropdown-actions">
                    <span class="ms-action-link" onclick="clearSingleFilter('destino')">✕ Limpar Seleção</span>
                </div>
                <div class="ms-options" id="opt-destino"></div>
            </div>
        </div>
        <div class="filter-group">
            <div class="filter-label-row">
                <label>Classe</label>
                <span class="filter-clear-link" id="badge-reset-classe" onclick="clearSingleFilter('classe')" style="display:none;" title="Limpar filtro de Classe">✕ Limpar</span>
            </div>
            <div class="ms-button" id="btn-classe" onclick="toggleDropdown('classe')">TODOS</div>
            <div class="ms-dropdown" id="dd-classe">
                <div class="ms-search-box">
                    <input type="text" id="search-classe" placeholder="Buscar..." oninput="filterOptionsList('classe')" onkeyup="filterOptionsList('classe')">
                    <span class="ms-search-clear-btn" id="clear-search-classe" onclick="clearSearchInput('classe')" title="Limpar busca">✕</span>
                </div>
                <div class="ms-dropdown-actions">
                    <span class="ms-action-link" onclick="clearSingleFilter('classe')">✕ Limpar Seleção</span>
                </div>
                <div class="ms-options" id="opt-classe"></div>
            </div>
        </div>
        <div class="filter-group">
            <div class="filter-label-row">
                <label>Viação</label>
                <span class="filter-clear-link" id="badge-reset-viacao" onclick="clearSingleFilter('viacao')" style="display:none;" title="Limpar filtro de Viação">✕ Limpar</span>
            </div>
            <div class="ms-button" id="btn-viacao" onclick="toggleDropdown('viacao')">TODOS</div>
            <div class="ms-dropdown" id="dd-viacao">
                <div class="ms-search-box">
                    <input type="text" id="search-viacao" placeholder="Buscar..." oninput="filterOptionsList('viacao')" onkeyup="filterOptionsList('viacao')">
                    <span class="ms-search-clear-btn" id="clear-search-viacao" onclick="clearSearchInput('viacao')" title="Limpar busca">✕</span>
                </div>
                <div class="ms-dropdown-actions">
                    <span class="ms-action-link" onclick="clearSingleFilter('viacao')">✕ Limpar Seleção</span>
                </div>
                <div class="ms-options" id="opt-viacao"></div>
            </div>
        </div>
        <div class="filter-group" style="min-width: 120px; flex: 0 0 auto;">
            <div class="filter-label-row"><label style="visibility: hidden;">Ações</label></div>
            <button type="button" class="btn-clear-all" onclick="resetAllFilters()" title="Limpar todos os filtros">
                🔄 Limpar Filtros
            </button>
        </div>
    </div>
    
    <div id="visao-global-kpis">
        <div class="kpi-section-title">🌍 Visão de Mercado</div>
        <div class="kpi-container">
            <div class="kpi-card" style="border-color: var(--alert-blue);">
                <div class="kpi-label">Piso Mínimo</div>
                <div class="kpi-value"><span class="prefix">R$</span> <span id="mkt-min">0,00</span></div>
            </div>
            <div class="kpi-card" style="border-color: #f1c40f;">
                <div class="kpi-label">Ticket Médio</div>
                <div class="kpi-value"><span class="prefix">R$</span> <span id="mkt-media">0,00</span></div>
            </div>
            <div class="kpi-card" style="border-color: #e74c3c;">
                <div class="kpi-label">Teto Máximo</div>
                <div class="kpi-value"><span class="prefix">R$</span> <span id="mkt-max">0,00</span></div>
            </div>
            <div class="kpi-card" style="border-color: #2ecc71;">
                <div class="kpi-label">Assentos Livres</div>
                <div class="kpi-value"><span id="mkt-assentos">0</span> <span id="mkt-assentos-sub" class="kpi-sub"></span></div>
            </div>
            <div class="kpi-card" style="border-color: #9b59b6;">
                <div class="kpi-label">Total Viagens</div>
                <div class="kpi-value"><span id="mkt-viagens">0</span> <span id="mkt-viagens-sub" class="kpi-sub"></span></div>
            </div>
        </div>

        <div class="kpi-section-title">🚌 Operação Nova Itapemirim</div>
        <div class="kpi-container">
            <div class="kpi-card" style="border-color: var(--alert-blue);">
                <div class="kpi-label">Piso Mínimo</div>
                <div class="kpi-value"><span class="prefix">R$</span> <span id="ita-min">0,00</span> <span id="ita-min-sub" class="kpi-sub"></span></div>
            </div>
            <div class="kpi-card" style="border-color: #f1c40f;">
                <div class="kpi-label">Ticket Médio</div>
                <div class="kpi-value"><span class="prefix">R$</span> <span id="ita-media">0,00</span> <span id="ita-media-sub" class="kpi-sub"></span></div>
            </div>
            <div class="kpi-card" style="border-color: #e74c3c;">
                <div class="kpi-label">Teto Máximo</div>
                <div class="kpi-value"><span class="prefix">R$</span> <span id="ita-max">0,00</span> <span id="ita-max-sub" class="kpi-sub"></span></div>
            </div>
            <div class="kpi-card" style="border-color: #2ecc71;">
                <div class="kpi-label">Assentos Livres</div>
                <div class="kpi-value"><span id="ita-assentos">0</span> <span id="ita-assentos-sub" class="kpi-sub"></span></div>
            </div>
            <div class="kpi-card" style="border-color: #9b59b6;">
                <div class="kpi-label">Total Viagens</div>
                <div class="kpi-value"><span id="ita-viagens">0</span> <span id="ita-viagens-sub" class="kpi-sub"></span></div>
            </div>
        </div>
    </div>
    
    <div class="tabs-container">
        <button class="tab-btn active" onclick="switchTab('tab-mercado')">📊 Gráficos de Mercado</button>
        <button class="tab-btn" onclick="switchTab('tab-pricing')">🎯 Pricing</button>
        <button class="tab-btn" onclick="switchTab('tab-classes')">💺 Visão por Classes</button>
        <button class="tab-btn" onclick="switchTab('tab-viagens')">🚌 Tabela de Viagens</button>
    </div>

    <div id="tab-mercado" class="tab-content active">
        <div class="chart-container" style="position: relative;">
            <div style="position: absolute; top: 15px; right: 20px; z-index: 10;">
                <button id="btn-dragmode" class="nav-toggle" onclick="toggleDragMode()">🖱️ NAVEGAÇÃO DE MAPA (PAN)</button>
            </div>
            <div id="chart-timeline" style="height: 570px;"></div>
        </div>
        <div class="chart-container" id="chart-combo" style="height: 450px;"></div>
    </div>

    <div id="tab-pricing" class="tab-content">
        <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 10px;">
            <div id="turnos-container" style="display: flex; gap: 8px; border-bottom: 1px solid var(--border-color); padding-bottom: 6px;">
                <button class="tab-btn active" id="btn-turno-TODOS" onclick="filterTurno('TODOS')" style="padding: 7px 13px; font-size: 12px;">🕒 Todos os Turnos</button>
                <button class="tab-btn" id="btn-turno-MADRUGADA" onclick="filterTurno('MADRUGADA')" style="padding: 7px 13px; font-size: 12px;">🌒 Madrugada (00h-06h)</button>
                <button class="tab-btn" id="btn-turno-MANHA" onclick="filterTurno('MANHA')" style="padding: 7px 13px; font-size: 12px;">☀️ Manhã (06h-12h)</button>
                <button class="tab-btn" id="btn-turno-TARDE" onclick="filterTurno('TARDE')" style="padding: 7px 13px; font-size: 12px;">🌤️ Tarde (12h-18h)</button>
                <button class="tab-btn" id="btn-turno-NOITE" onclick="filterTurno('NOITE')" style="padding: 7px 13px; font-size: 12px;">🌙 Noite (18h-00h)</button>
            </div>
        </div>
        
        <div class="kpi-container" id="pricing-kpis" style="margin-bottom: 10px; display: none; gap: 12px;">
            <div class="kpi-card pricing-kpi-card" id="card-pricing-win" onclick="togglePricingKpiFilter('win')" title="Clique para filtrar apenas os horários com vantagem de preço (GAP decrescente)" style="border-top: none; border-left: 5px solid var(--alert-green); padding: 12px 18px; flex: 1; min-width: 260px; box-shadow: 0 4px 6px rgba(0,0,0,0.15);">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div class="kpi-label" style="color: var(--alert-green); font-size: 12px; font-weight: 700; letter-spacing: 0.5px;">🏆 Preço Itapemirim ≥ Preço Concorrência</div>
                    <span class="kpi-filter-badge" id="badge-pricing-win" style="display: none; font-size: 10px; padding: 2px 7px; border-radius: 10px; background: var(--alert-green); color: #121212; font-weight: bold;">FILTRO ATIVO ✕</span>
                </div>
                <div class="kpi-value" style="font-size: 26px; font-weight: 800; color: var(--text-main); margin-top: 4px;">
                    <span id="kpi-pricing-win" style="color: var(--alert-green);">0</span> <span style="font-size: 13px; color: var(--text-muted); font-weight: normal;">mercados</span>
                </div>
                <div class="kpi-sub" id="kpi-pricing-win-sub" style="font-size: 12px; color: var(--text-muted); margin-top: 4px;">
                    Itapemirim com preço igual ou superior
                </div>
            </div>
            <div class="kpi-card pricing-kpi-card" id="card-pricing-loss" onclick="togglePricingKpiFilter('loss')" title="Clique para filtrar apenas os horários com concorrência acima (maior GAP de oportunidade primeiro)" style="border-top: none; border-left: 5px solid var(--alert-red); padding: 12px 18px; flex: 1; min-width: 260px; box-shadow: 0 4px 6px rgba(0,0,0,0.15);">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div class="kpi-label" style="color: var(--alert-red); font-size: 12px; font-weight: 700; letter-spacing: 0.5px;">⚠️ Preço Concorrência > Preço Itapemirim</div>
                    <span class="kpi-filter-badge" id="badge-pricing-loss" style="display: none; font-size: 10px; padding: 2px 7px; border-radius: 10px; background: var(--alert-red); color: #ffffff; font-weight: bold;">FILTRO ATIVO ✕</span>
                </div>
                <div class="kpi-value" style="font-size: 26px; font-weight: 800; color: var(--text-main); margin-top: 4px;">
                    <span id="kpi-pricing-loss" style="color: var(--alert-red);">0</span> <span style="font-size: 13px; color: var(--text-muted); font-weight: normal;">mercados</span>
                </div>
                <div class="kpi-sub" id="kpi-pricing-loss-sub" style="font-size: 12px; color: var(--text-muted); margin-top: 4px;">
                    Oportunidade tarifária / Concorrência mais cara
                </div>
            </div>
        </div>

        <div class="pricing-container">
            <table class="pricing-table">
                <thead>
                    <tr>
                        <th class="sortable-th" data-col="mercado" onclick="sortPricingTable('mercado')">Mercado <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="horarioStr" onclick="sortPricingTable('horarioStr')">Nosso Horário <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="classe" onclick="sortPricingTable('classe')">Classe <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="compCount" onclick="sortPricingTable('compCount')">Concorrentes na Janela <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="nossoPreco" onclick="sortPricingTable('nossoPreco')">Nosso Preço <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="compMax" onclick="sortPricingTable('compMax')">Preço Concorrente (Min - Max) <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="nossoAssentos" onclick="sortPricingTable('nossoAssentos')">Assentos (Nós x Eles) <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="gapRS" onclick="sortPricingTable('gapRS')">GAP Máximo (R$) <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="gapPct" onclick="sortPricingTable('gapPct')">GAP Máximo (%) <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="validation" onclick="sortPricingTable('validation')">Validação <span class="sort-arrow">↕</span></th>
                    </tr>
                </thead>
                <tbody id="pricing-tbody"></tbody>
            </table>
        </div>
    </div>

    <div id="tab-classes" class="tab-content">
        <div class="classes-header-bar">
            <div class="classes-title-box">
                <h2>💺 Matriz Estratégica por Categoria de Assento</h2>
                <span class="classes-subtitle">Comparativo de Precificação, Yield, Capacidade e Presença Competitiva</span>
            </div>
            <div class="classes-summary-strip" id="classes-summary-strip"></div>
        </div>
        <div class="classes-deck-panoramic" id="classes-grid-container"></div>
    </div>

    <div id="tab-viagens" class="tab-content">
        <div class="table-wrapper">
            <table class="data-table">
                <thead>
                    <tr>
                        <th class="sortable-th" data-col="Horário Partida" onclick="sortViagensTable('Horário Partida')">Data/Hora Partida <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="Viação" onclick="sortViagensTable('Viação')">Viação <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="Origem" onclick="sortViagensTable('Origem')">Origem <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="Destino" onclick="sortViagensTable('Destino')">Destino <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="Classe Padronizada" onclick="sortViagensTable('Classe Padronizada')">Classe <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="Preço (R$)" onclick="sortViagensTable('Preço (R$)')">Preço (R$) <span class="sort-arrow">↕</span></th>
                        <th class="sortable-th" data-col="Assentos Disponíveis" onclick="sortViagensTable('Assentos Disponíveis')">Assentos Livres <span class="sort-arrow">↕</span></th>
                    </tr>
                </thead>
                <tbody id="table-body"></tbody>
            </table>
        </div>
        <div id="viagens-pagination-bar" style="display: flex; justify-content: space-between; align-items: center; padding: 10px 18px; background-color: var(--bg-card); border: 1px solid var(--border-color); border-radius: 8px; font-size: 13px;">
            <div id="viagens-page-info" style="color: var(--text-muted);">Carregando viagens...</div>
            <div style="display: flex; gap: 8px;">
                <button id="btn-viagens-prev" onclick="changeViagensPage(-1)" class="tab-btn" style="padding: 6px 14px; font-size: 12px; border-radius: 4px; border: 1px solid var(--border-color); cursor: pointer;">◀ Anterior</button>
                <button id="btn-viagens-next" onclick="changeViagensPage(1)" class="tab-btn" style="padding: 6px 14px; font-size: 12px; border-radius: 4px; border: 1px solid var(--border-color); cursor: pointer;">Próximo ▶</button>
            </div>
        </div>
    </div>
"""
# ==========================================
# 3. LÓGICA DE JAVASCRIPT E GERAÇÃO DO ARQUIVO
# ==========================================
html_parte2 = """
    <script>
        const rawData = {DATA_PLACEHOLDER};
        let isLightMode = false;
        let isPanMode = true; 
        let itapemirimMediaGlobal = 0;
        let currentTab = 'tab-mercado';
        let currentTurno = 'TODOS';
        
        // Controle de Ordenação do Tático Pricing
        let pricingDataArray = [];
        let currentSortCol = '';
        let currentSortAsc = true;
        let activePricingKpiFilter = null; // null | 'win' | 'loss'
        
        const filtersList = [
            {id: 'origem', field: 'Origem'},
            {id: 'destino', field: 'Destino'},
            {id: 'classe', field: 'Classe Padronizada'},
            {id: 'viacao', field: 'Viação'}
        ];
        let selections = { origem: [], destino: [], classe: [], viacao: [] };

        function initDropdowns() {
            filtersList.forEach(f => {
                let distinctVals = [...new Set(rawData.map(i => i[f.field]))].sort();
                if (f.id === 'viacao') {
                    distinctVals = distinctVals.filter(v => {
                        const u = (v || '').toUpperCase();
                        return !u.includes('ITAPEMIRIM') && !u.includes('SUZANTUR');
                    });
                }
                const container = document.getElementById('opt-'+f.id);
                let html = `<label class="ms-option ms-option-container"><input type="checkbox" value="ALL" onchange="handleSelect('${f.id}', 'ALL', this)" checked> <b>TODAS AS OPÇÕES</b> <span class="opt-count" style="margin-left:auto; color:var(--text-muted); font-size:11px;"></span></label>`;
                distinctVals.forEach(val => {
                    html += `<label class="ms-option ms-option-container"><input type="checkbox" value="${val}" onchange="handleSelect('${f.id}', '${val}', this)"> <span>${val}</span> <span class="opt-count" style="margin-left:auto; color:var(--text-muted); font-size:11px;"></span></label>`;
                });
                container.innerHTML = html;
            });
            updateDropdowns(); 
        }

        function updateDropdowns() {
            const dtInicio = dtInicioInput.value;
            const dtFim = dtFimInput.value;

            // Filtra por data uma única vez para todos os 4 seletores
            const dateFiltered = (dtInicio || dtFim) ? rawData.filter(row => {
                const rowDate = row['Horário Partida'].substring(0, 10);
                if (dtInicio && rowDate < dtInicio) return false;
                if (dtFim && rowDate > dtFim) return false;
                return true;
            }) : rawData;

            filtersList.forEach(f => {
                const otherFilters = filtersList.filter(otherF => otherF.id !== f.id && selections[otherF.id].length > 0);
                
                let availableData;
                if (otherFilters.length === 0) {
                    availableData = dateFiltered;
                } else {
                    availableData = dateFiltered.filter(row => {
                        for (let i = 0; i < otherFilters.length; i++) {
                            const otherF = otherFilters[i];
                            if (otherF.id === 'viacao') {
                                const vUpper = (row['Viação'] || '').toUpperCase();
                                const isOur = vUpper.includes('ITAPEMIRIM') || vUpper.includes('SUZANTUR');
                                if (!isOur && !selections.viacao.includes(row['Viação'])) return false;
                            } else {
                                if (!selections[otherF.id].includes(row[otherF.field])) return false;
                            }
                        }
                        return true;
                    });
                }

                const counts = {};
                for (let i = 0; i < availableData.length; i++) {
                    const val = availableData[i][f.field];
                    counts[val] = (counts[val] || 0) + 1;
                }
                
                const optContainer = document.getElementById('opt-'+f.id);
                const labels = optContainer.querySelectorAll('.ms-option-container');
                let activeCount = 0;

                labels.forEach(label => {
                    const cb = label.querySelector('input');
                    const val = cb.value;
                    if (val === 'ALL') return;
                    
                    if (counts[val]) {
                        label.style.display = 'flex';
                        label.querySelector('.opt-count').innerText = `(${counts[val]})`;
                        activeCount++;
                    } else {
                        label.style.display = 'none';
                        const idx = selections[f.id].indexOf(val);
                        if(idx > -1) {
                            selections[f.id].splice(idx, 1);
                            cb.checked = false;
                        }
                    }
                });

                const allCb = optContainer.querySelector('input[value="ALL"]');
                allCb.parentNode.querySelector('.opt-count').innerText = `(${activeCount})`;
                if(selections[f.id].length === 0) allCb.checked = true;

                updateButtonText(f.id, activeCount);
            });
        }

        function toggleDropdown(id) {
            filtersList.forEach(f => { if(f.id !== id) document.getElementById('dd-'+f.id).classList.remove('show'); });
            document.getElementById('dd-'+id).classList.toggle('show');
        }

        document.addEventListener('click', function(e) {
            if (!e.target.closest('.filter-group')) {
                filtersList.forEach(f => document.getElementById('dd-'+f.id).classList.remove('show'));
            }
        });

        function filterOptionsList(id) {
            const searchInput = document.getElementById('search-' + id);
            const clearBtn = document.getElementById('clear-search-' + id);
            const term = (searchInput ? searchInput.value : '').toLowerCase().trim();
            
            if (clearBtn) {
                clearBtn.style.display = term ? 'block' : 'none';
            }
            
            const opts = document.getElementById('opt-' + id).querySelectorAll('.ms-option-container');
            opts.forEach(opt => {
                if (opt.innerText.toLowerCase().includes(term)) opt.style.display = 'flex';
                else opt.style.display = 'none';
            });
        }

        function clearSearchInput(id) {
            const searchInput = document.getElementById('search-' + id);
            if (searchInput) {
                searchInput.value = '';
                filterOptionsList(id);
            }
        }

        function clearSingleFilter(id) {
            selections[id] = [];
            const optContainer = document.getElementById('opt-' + id);
            if (optContainer) {
                optContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                    cb.checked = (cb.value === 'ALL');
                });
            }
            clearSearchInput(id);
            updateDropdowns();
            applyDataFilter();
        }

        function resetAllFilters() {
            filtersList.forEach(f => {
                selections[f.id] = [];
                clearSearchInput(f.id);
                const optContainer = document.getElementById('opt-' + f.id);
                if (optContainer) {
                    optContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                        cb.checked = (cb.value === 'ALL');
                    });
                }
            });

            setDefaultDates();
            updateDropdowns();
            applyDataFilter();
        }

        function setDefaultDates() {
            const availableDates = [...new Set(rawData.map(r => r['Horário Partida'].substring(0, 10)))].filter(Boolean).sort();
            const now = new Date();
            const localToday = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
            let targetDate = localToday;
            if (!availableDates.includes(localToday)) {
                targetDate = '{DATA_PADRAO}';
            }
            dtInicioInput.value = targetDate;
            dtFimInput.value = targetDate;
        }

        function handleSelect(filterId, value, checkboxElement) {
            if (value === 'ALL') {
                selections[filterId] = [];
                const checkboxes = document.getElementById('opt-'+filterId).querySelectorAll('input[type="checkbox"]');
                checkboxes.forEach(cb => { if(cb.value !== 'ALL') cb.checked = false; else cb.checked = true; });
            } else {
                const idx = selections[filterId].indexOf(value);
                if (idx > -1) selections[filterId].splice(idx, 1);
                else selections[filterId].push(value);
                
                document.querySelector(`#opt-${filterId} input[value="ALL"]`).checked = (selections[filterId].length === 0);
            }
            updateDropdowns();
            applyDataFilter();
        }

        function updateButtonText(filterId, activeCount) {
            const btn = document.getElementById('btn-'+filterId);
            if (selections[filterId].length === 0) btn.innerHTML = `TODAS AS OPÇÕES (${activeCount || 0})`;
            else if (selections[filterId].length === 1) btn.innerHTML = selections[filterId][0];
            else btn.innerHTML = `${selections[filterId].length} Selecionados`;

            const resetBadge = document.getElementById('badge-reset-' + filterId);
            if (resetBadge) {
                resetBadge.style.display = (selections[filterId].length > 0) ? 'inline-block' : 'none';
            }
        }
        
        function getCapacidadeEstimada(classe) {
            const c = (classe || '').toUpperCase();
            if (c.includes('CAMA DUPLO')) return 12;
            if (c.includes('CAMA INDIVIDUAL')) return 8;
            if (c.includes('CAMA')) return 15;
            if (c.includes('LEITO INDIVIDUAL')) return 12;
            if (c.includes('LEITO DUPLO')) return 24;
            if (c.includes('LEITO')) return 26;
            if (c.includes('SEMILEITO')) return 46;
            if (c.includes('EXECUTIVO')) return 46;
            return 46;
        }

        function renderCurrentTab() {
            if (currentTab === 'tab-mercado') drawCharts();
            else if (currentTab === 'tab-pricing') renderPricingTab();
            else if (currentTab === 'tab-classes') renderClassesTab(window._lastBaseMarketData || rawData);
            else if (currentTab === 'tab-viagens') drawTable();
        }

        function switchTab(tabId) {
            currentTab = tabId;
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
            const clickedBtn = document.querySelector(`.tab-btn[onclick*="${tabId}"]`);
            if (clickedBtn) clickedBtn.classList.add('active');
            document.getElementById(tabId).classList.add('active');
            
            if(tabId === 'tab-classes' || tabId === 'tab-pricing') {
                document.getElementById('visao-global-kpis').style.display = 'none';
            } else {
                document.getElementById('visao-global-kpis').style.display = 'block';
            }
            renderCurrentTab();
            setTimeout(adjustAllLayouts, 30);
        }

        function filterTurno(turno) {
            currentTurno = turno;
            document.querySelectorAll('#turnos-container .tab-btn').forEach(btn => btn.classList.remove('active'));
            document.getElementById('btn-turno-' + turno).classList.add('active');
            renderPricingTab();
            setTimeout(adjustAllLayouts, 30);
        }

        function toggleTheme() {
            isLightMode = !isLightMode;
            document.body.classList.toggle('light-mode', isLightMode);
            document.querySelector('.theme-toggle').innerHTML = isLightMode ? '🌙 Modo Escuro' : '☀️ Modo Claro';
            drawCharts();
        }

        function toggleDragMode() {
            isPanMode = !isPanMode;
            const btn = document.getElementById('btn-dragmode');
            btn.innerHTML = isPanMode ? '🖱️ NAVEGAÇÃO DE MAPA (PAN)' : '🔎 SELEÇÃO DE ÁREA (ZOOM)';
            btn.style.borderColor = isPanMode ? 'var(--border-color)' : 'var(--primary-yellow)';
            btn.style.color = isPanMode ? 'var(--text-main)' : 'var(--primary-yellow)';
            Plotly.relayout('chart-timeline', { dragmode: isPanMode ? 'pan' : 'zoom' });
        }

        function formatBR(val, isMoney) {
            if (val === undefined || val === null || isNaN(val)) return '-';
            if (isMoney) return val.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
            return val.toLocaleString('pt-BR');
        }

        const dtInicioInput = document.getElementById('filtro_data_inicio');
        const dtFimInput = document.getElementById('filtro_data_fim');

        dtInicioInput.addEventListener('change', function() {
            dtFimInput.min = this.value;
            if(dtFimInput.value && dtFimInput.value < this.value) dtFimInput.value = this.value;
            updateDropdowns();
            applyDataFilter();
        });

        dtFimInput.addEventListener('change', function() {
            dtInicioInput.max = this.value;
            if(dtInicioInput.value && dtInicioInput.value > this.value) dtInicioInput.value = this.value;
            updateDropdowns();
            applyDataFilter();
        });

        // ==========================================
        // DEFINIÇÃO RIGOROSA DOS GRUPOS DE CLASSES (PROCV BASE)
        // ==========================================
        const GRUPO_1_CLASSES = ['CAMA DUPLO', 'CAMA INDIVIDUAL', 'LEITO DUPLO'];
        const GRUPO_2_CLASSES = ['LEITO', 'SEMILEITO', 'CONEXÃO', 'CONEXAO', 'CONVENCIONAL'];

        function getClasseGrupo(cls) {
            const c = (cls || '').toUpperCase().trim();
            if (GRUPO_1_CLASSES.includes(c)) return 1;
            return 2;
        }

        // ==========================================
        // ABA PRICING: Lógica e Ordenação
        // ==========================================
        function renderPricingTab() {
            const tbody = document.getElementById('pricing-tbody');
            const kpiContainer = document.getElementById('pricing-kpis');
            tbody.innerHTML = '';
            kpiContainer.style.display = 'none';

            const dtInicio = dtInicioInput.value;
            const dtFim = dtFimInput.value;

            // 1. Filtrar base global de dados com base nos filtros ativos (Datas, Origem, Destino, Classe)
            const marketData = rawData.filter(row => {
                const rowDate = row['Horário Partida'].substring(0, 10);
                if (dtInicio && rowDate < dtInicio) return false;
                if (dtFim && rowDate > dtFim) return false;
                if (selections.origem.length > 0 && !selections.origem.includes(row['Origem'])) return false;
                if (selections.destino.length > 0 && !selections.destino.includes(row['Destino'])) return false;
                if (selections.classe.length > 0 && !selections.classe.includes(row['Classe Padronizada'])) return false;
                return true;
            });

            // 2. Viagens da Itapemirim em todo o dia (necessárias para o Nearest Neighbor Global)
            const allItaTrips = marketData.filter(row => {
                const v = row['Viação'].toUpperCase();
                return v.includes('ITAPEMIRIM') || v.includes('SUZANTUR');
            });

            // 3. Viagens dos Concorrentes
            const compTrips = marketData.filter(row => {
                const v = row['Viação'].toUpperCase();
                if (v.includes('ITAPEMIRIM') || v.includes('SUZANTUR')) return false;
                if (selections.viacao.length > 0 && !selections.viacao.includes(row['Viação'])) return false;
                return true;
            });

            // Otimização O(1): Indexação prévia das viagens da Itapemirim por mercado
            const itaByMarket = {};
            const totalViagensPorMercado = {};
            allItaTrips.forEach((ita, idx) => {
                ita._itaId = 'ita_' + idx;
                ita._assignedComps = [];
                const mKey = (ita['Origem'] || '') + '|' + (ita['Destino'] || '');
                if (!itaByMarket[mKey]) itaByMarket[mKey] = [];
                itaByMarket[mKey].push(ita);

                const orig = ita['Origem'].replace('-TODOS', '');
                const dest = ita['Destino'].replace('-TODOS', '');
                const displayKey = `${orig} x ${dest}`;
                totalViagensPorMercado[displayKey] = (totalViagensPorMercado[displayKey] || 0) + 1;
            });

            // 4. Regra de Vizinho Mais Próximo GLOBAL com Comparações Excepcionais e Sem Descarte (Otimizada com O(1) Hash Map)
            compTrips.forEach(comp => {
                const compTime = comp['Timestamp'] || Math.floor(new Date(comp['Horário Partida'].replace(' ', 'T')).getTime() / 1000);
                const compCls = (comp['Classe Padronizada'] || '').toUpperCase().trim();
                const compGroup = getClasseGrupo(compCls);

                const mKey = (comp['Origem'] || '') + '|' + (comp['Destino'] || '');
                const marketItaTrips = itaByMarket[mKey];
                if (!marketItaTrips || marketItaTrips.length === 0) return;

                // Passo 1: Busca mesma classe exata
                let candidates = marketItaTrips.filter(ita => 
                    (ita['Classe Padronizada'] || '').toUpperCase().trim() === compCls
                );

                // Passo 2: Se não houver, busca no mesmo grupo de proximidade física
                if (candidates.length === 0) {
                    candidates = marketItaTrips.filter(ita => 
                        getClasseGrupo(ita['Classe Padronizada']) === compGroup
                    );
                }

                // Passo 3: Fallback geral (qualquer classe da Itapemirim no mercado)
                if (candidates.length === 0) {
                    candidates = marketItaTrips;
                }

                let closestIta = null;
                let minDiff = Infinity;
                candidates.forEach(ita => {
                    const itaTime = ita['Timestamp'] || Math.floor(new Date(ita['Horário Partida'].replace(' ', 'T')).getTime() / 1000);
                    const diff = Math.abs(compTime - itaTime);
                    if (diff < minDiff) {
                        minDiff = diff;
                        closestIta = ita;
                    }
                });

                if (closestIta) {
                    const timeDiffHours = minDiff / 3600.0;
                    const isSameClass = compCls === (closestIta['Classe Padronizada'] || '').toUpperCase().trim();
                    const isWithin3h = timeDiffHours <= 3.0;

                    let tagValidation = 'Direto';
                    if (!isSameClass || !isWithin3h) {
                        const reasons = [];
                        if (!isWithin3h) reasons.push('Horário');
                        if (!isSameClass) reasons.push('Classe');
                        tagValidation = `Indireto (${reasons.join(' e ')})`;
                    }

                    comp._matchValidationTag = tagValidation;
                    comp._timeDiffHours = timeDiffHours;
                    closestIta._assignedComps.push(comp);
                }
            });

            // 5. Filtrar as viagens da Itapemirim para o turno selecionado
            let itaTrips = allItaTrips.filter(row => {
                if (currentTurno === 'TODOS') return true;
                const hour = parseInt(row['Hora Apenas'].split(':')[0]);
                if (currentTurno === 'MADRUGADA' && (hour >= 0 && hour < 6)) return true;
                if (currentTurno === 'MANHA' && (hour >= 6 && hour < 12)) return true;
                if (currentTurno === 'TARDE' && (hour >= 12 && hour < 18)) return true;
                if (currentTurno === 'NOITE' && (hour >= 18 && hour <= 23)) return true;
                return false;
            });

            if (itaTrips.length === 0) {
                tbody.innerHTML = `<tr><td colspan="10" style="text-align: center; color: var(--text-muted); padding: 30px; font-size: 15px;">Nenhum horário da Itapemirim encontrado para o turno e filtros selecionados.</td></tr>`;
                return;
            }

            pricingDataArray = [];
            const marketsMap = {};
            let totalWinsHorarios = 0;
            let totalLossesHorarios = 0;
            let sumHorariosWinGap = 0;
            let sumHorariosLossGap = 0;

            itaTrips.forEach((ita, index) => {
                const matchingComps = ita._assignedComps || [];
                const hasComps = matchingComps.length > 0;

                const origem = ita['Origem'];
                const destino = ita['Destino'];
                const mercadoNome = `${origem.replace('-TODOS', '')} x ${destino.replace('-TODOS', '')}`;
                const itaTimestamp = ita['Timestamp'] ? (ita['Timestamp'] * 1000) : new Date(ita['Horário Partida'].replace(' ', 'T')).getTime();
                const itaPrice = ita['Preço (R$)'];
                const itaSeats = ita['Assentos Disponíveis'];

                let distinctComps = [];
                let compPrices = [];
                let compSeatsTotal = 0;
                let compMinPrice = null;
                let compMaxPrice = null;
                let gapRS = null;
                let gapPct = null;
                let validation = 'Sem concorrência';
                let compCountText = '0 Player(s)';
                let compCountTrips = 0;

                if (hasComps) {
                    distinctComps = [...new Set(matchingComps.map(c => c['Viação']))];
                    compCountText = `${distinctComps.length} Player(s)`;

                    // Prioridade absoluta para concorrentes DIRETOS
                    const directComps = matchingComps.filter(c => c._matchValidationTag === 'Direto');
                    const hasDirect = directComps.length > 0;

                    // O benchmark de preços, assentos e GAP prioriza os concorrentes diretos quando existirem
                    const benchmarkComps = hasDirect ? directComps : matchingComps;

                    compPrices = benchmarkComps.map(c => c['Preço (R$)']);
                    compSeatsTotal = benchmarkComps.reduce((acc, c) => acc + c['Assentos Disponíveis'], 0);
                    compMinPrice = Math.min(...compPrices);
                    compMaxPrice = Math.max(...compPrices);
                    compCountTrips = benchmarkComps.length;

                    // Determinação da validação da linha (sem redundâncias e com preferência Direto)
                    if (hasDirect) {
                        validation = 'Direto';
                    } else {
                        const hasHorario = matchingComps.some(c => (c._matchValidationTag || '').includes('Horário'));
                        const hasClasse = matchingComps.some(c => (c._matchValidationTag || '').includes('Classe'));
                        if (hasHorario && hasClasse) {
                            validation = 'Indireto (Horário e Classe)';
                        } else if (hasHorario) {
                            validation = 'Indireto (Horário)';
                        } else {
                            validation = 'Indireto (Classe)';
                        }
                    }

                    // GAP de Yield: Preço Itapemirim - Maior Preço Concorrência do Benchmark
                    gapRS = itaPrice - compMaxPrice;
                    gapPct = compMaxPrice > 0 ? (gapRS / compMaxPrice) * 100 : 0;

                    if (!marketsMap[mercadoNome]) {
                        marketsMap[mercadoNome] = {
                            wins: 0,
                            losses: 0,
                            totalGapWin: 0,
                            totalGapLoss: 0
                        };
                    }

                    // Saldo de Horários
                    if (itaPrice >= compMaxPrice) {
                        totalWinsHorarios++;
                        sumHorariosWinGap += (itaPrice - compMaxPrice);
                        marketsMap[mercadoNome].wins++;
                        marketsMap[mercadoNome].totalGapWin += (itaPrice - compMaxPrice);
                    } else {
                        totalLossesHorarios++;
                        sumHorariosLossGap += (compMaxPrice - itaPrice);
                        marketsMap[mercadoNome].losses++;
                        marketsMap[mercadoNome].totalGapLoss += (compMaxPrice - itaPrice);
                    }
                }

                let ringueArray = matchingComps.map(c => {
                    const shortDate = c['Horário Formatado'].substring(0, 5) + ' ' + c['Horário Formatado'].substring(11, 16);
                    return { 
                        isIta: false, 
                        viacao: c['Viação'], 
                        classe: c['Classe Padronizada'], 
                        hora: shortDate, 
                        preco: c['Preço (R$)'], 
                        assentos: c['Assentos Disponíveis'],
                        valTag: c._matchValidationTag || 'Direto'
                    };
                });
                ringueArray.push({
                    isIta: true, 
                    viacao: ita['Viação'], 
                    classe: ita['Classe Padronizada'], 
                    hora: ita['Horário Formatado'].substring(0, 5) + ' ' + ita['Horário Formatado'].substring(11, 16), 
                    preco: itaPrice, 
                    assentos: itaSeats,
                    valTag: 'Benchmark'
                });
                ringueArray.sort((a,b) => b.preco - a.preco);

                pricingDataArray.push({
                    mercado: mercadoNome,
                    totalViagensMercado: totalViagensPorMercado[mercadoNome] || 1,
                    horarioStr: ita['Horário Formatado'].substring(11, 16),
                    itaHorarioFull: ita['Horário Formatado'].substring(0, 5) + ' ' + ita['Horário Formatado'].substring(11, 16),
                    itaViacao: ita['Viação'],
                    rawTimestamp: itaTimestamp,
                    classe: ita['Classe Padronizada'],
                    hasComps: hasComps,
                    ringueArray: ringueArray,
                    compCount: distinctComps.length,
                    compCountText: compCountText,
                    compCountTrips: compCountTrips,
                    nossoPreco: itaPrice,
                    compMin: compMinPrice,
                    compMax: compMaxPrice,
                    nossoAssentos: itaSeats,
                    elesAssentos: compSeatsTotal,
                    gapRS: gapRS,
                    gapPct: gapPct,
                    validation: validation,
                    rowId: `detail-${index}-${itaTimestamp}`
                });
            });

            // 6. Consolidação dos Mercados nos Cards por SALDO DE HORÁRIOS (Vitórias vs Derrotas)
            let kpiWinsMercadosCount = 0;
            let kpiLossesMercadosCount = 0;

            Object.keys(marketsMap).forEach(mNome => {
                const mData = marketsMap[mNome];
                if (mData.wins === 0 && mData.losses === 0) return;

                if (mData.wins >= mData.losses) {
                    kpiWinsMercadosCount++;
                } else {
                    kpiLossesMercadosCount++;
                }
            });

            if (pricingDataArray.length > 0) {
                kpiContainer.style.display = 'flex';
                document.getElementById('kpi-pricing-win').innerText = kpiWinsMercadosCount;
                const avgWin = totalWinsHorarios > 0 ? (sumHorariosWinGap / totalWinsHorarios) : 0;
                document.getElementById('kpi-pricing-win-sub').innerHTML = `<b>${totalWinsHorarios}</b> horários com preço superior/igual • Vantagem média: <b style="color: var(--alert-green);">+R$ ${formatBR(avgWin, true)}</b>`;

                document.getElementById('kpi-pricing-loss').innerText = kpiLossesMercadosCount;
                const avgLoss = totalLossesHorarios > 0 ? (sumHorariosLossGap / totalLossesHorarios) : 0;
                document.getElementById('kpi-pricing-loss-sub').innerHTML = `<b>${totalLossesHorarios}</b> horários com concorrência acima • Diferença média: <b style="color: var(--alert-red);">-R$ ${formatBR(avgLoss, true)}</b>`;
            }
            activePricingKpiFilter = null;
            updatePricingKpiCardsUI();
            sortPricingTable(''); // Render inicial
        }

        function togglePricingKpiFilter(type) {
            if (activePricingKpiFilter === type) {
                activePricingKpiFilter = null;
                currentSortCol = '';
                currentSortAsc = true;
            } else {
                activePricingKpiFilter = type;
                if (type === 'win') {
                    currentSortCol = 'gapRS';
                    currentSortAsc = false; // Maior vantagem positiva primeiro (+R$ 66, +R$ 20, R$ 0)
                } else if (type === 'loss') {
                    currentSortCol = 'gapRS_loss';
                    currentSortAsc = false; // Maior oportunidade de aumento tarifário primeiro (maior desvantagem em módulo, ex: -R$ 626, -R$ 437...)
                }
            }
            updatePricingKpiCardsUI();
            sortPricingTable('');
        }

        function updatePricingKpiCardsUI() {
            const cardWin = document.getElementById('card-pricing-win');
            const cardLoss = document.getElementById('card-pricing-loss');
            const badgeWin = document.getElementById('badge-pricing-win');
            const badgeLoss = document.getElementById('badge-pricing-loss');
            
            if (!cardWin || !cardLoss) return;
            
            if (activePricingKpiFilter === 'win') {
                cardWin.style.border = '2px solid var(--alert-green)';
                cardWin.style.background = 'rgba(46, 204, 113, 0.12)';
                cardWin.style.boxShadow = '0 6px 14px rgba(46, 204, 113, 0.25)';
                if (badgeWin) badgeWin.style.display = 'inline-block';
                
                cardLoss.style.border = '1px solid var(--border-color)';
                cardLoss.style.borderLeft = '5px solid var(--alert-red)';
                cardLoss.style.background = 'var(--bg-card)';
                cardLoss.style.boxShadow = '0 4px 6px rgba(0,0,0,0.15)';
                cardLoss.style.opacity = '0.55';
                if (badgeLoss) badgeLoss.style.display = 'none';
            } else if (activePricingKpiFilter === 'loss') {
                cardLoss.style.border = '2px solid var(--alert-red)';
                cardLoss.style.background = 'rgba(231, 76, 60, 0.12)';
                cardLoss.style.boxShadow = '0 6px 14px rgba(231, 76, 60, 0.25)';
                if (badgeLoss) badgeLoss.style.display = 'inline-block';
                
                cardWin.style.border = '1px solid var(--border-color)';
                cardWin.style.borderLeft = '5px solid var(--alert-green)';
                cardWin.style.background = 'var(--bg-card)';
                cardWin.style.boxShadow = '0 4px 6px rgba(0,0,0,0.15)';
                cardWin.style.opacity = '0.55';
                if (badgeWin) badgeWin.style.display = 'none';
            } else {
                cardWin.style.border = '1px solid var(--border-color)';
                cardWin.style.borderLeft = '5px solid var(--alert-green)';
                cardWin.style.background = 'var(--bg-card)';
                cardWin.style.boxShadow = '0 4px 6px rgba(0,0,0,0.15)';
                cardWin.style.opacity = '1';
                if (badgeWin) badgeWin.style.display = 'none';
                
                cardLoss.style.border = '1px solid var(--border-color)';
                cardLoss.style.borderLeft = '5px solid var(--alert-red)';
                cardLoss.style.background = 'var(--bg-card)';
                cardLoss.style.boxShadow = '0 4px 6px rgba(0,0,0,0.15)';
                cardLoss.style.opacity = '1';
                if (badgeLoss) badgeLoss.style.display = 'none';
            }
        }

        function sortPricingTable(col) {
            if (col) {
                if (currentSortCol === col) {
                    currentSortAsc = !currentSortAsc;
                } else {
                    currentSortCol = col;
                    if (['mercado', 'classe', 'validation'].includes(col)) {
                        currentSortAsc = true;
                    } else {
                        currentSortAsc = false;
                    }
                }
            }
            
            // 1. Filtragem da lista para exibição baseada no clique dos cards
            let displayList = [...pricingDataArray];
            if (activePricingKpiFilter === 'win') {
                displayList = displayList.filter(item => item.hasComps && item.gapRS >= 0);
            } else if (activePricingKpiFilter === 'loss') {
                displayList = displayList.filter(item => item.hasComps && item.gapRS < 0);
            }

            // 2. Ordenação da lista a ser exibida
            displayList.sort((a, b) => {
                if (currentSortCol === 'gapRS_loss') {
                    const diffA = Math.abs(a.gapRS || 0);
                    const diffB = Math.abs(b.gapRS || 0);
                    return currentSortAsc ? (diffA - diffB) : (diffB - diffA);
                }

                let valA = a[currentSortCol];
                let valB = b[currentSortCol];
                if (currentSortCol === 'horarioStr') {
                    valA = a.rawTimestamp || valA;
                    valB = b.rawTimestamp || valB;
                }
                if (valA === null || valA === undefined) valA = currentSortAsc ? Infinity : -Infinity;
                if (valB === null || valB === undefined) valB = currentSortAsc ? Infinity : -Infinity;
                if (typeof valA === 'string') {
                    return currentSortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
                }
                return currentSortAsc ? (valA - valB) : (valB - valA);
            });

            window._currentPricingDisplayList = displayList;
            
            // 3. Atualizar cabeçalho da coluna Mercado para exibir SEMPRE a quantidade de horários nos filtros
            const thMercado = document.querySelector('.pricing-table th[data-col="mercado"]');
            if (thMercado) {
                const arrow = thMercado.querySelector('.sort-arrow');
                const arrowHTML = arrow ? arrow.outerHTML : '<span class="sort-arrow">↕</span>';
                
                let countBadge = '';
                if (activePricingKpiFilter === 'win') {
                    countBadge = `<span style="font-weight: 600; font-size: 11px; color: var(--alert-green); text-transform: none; margin-left: 5px;">(${displayList.length} de ${pricingDataArray.length} horários • Preço Ita ≥ Concorrência)</span>`;
                } else if (activePricingKpiFilter === 'loss') {
                    countBadge = `<span style="font-weight: 600; font-size: 11px; color: var(--alert-red); text-transform: none; margin-left: 5px;">(${displayList.length} de ${pricingDataArray.length} horários • Concorrência > Preço Ita)</span>`;
                } else {
                    countBadge = `<span style="font-weight: 600; font-size: 11px; color: var(--primary-yellow); text-transform: none; margin-left: 5px;">(${displayList.length} horários)</span>`;
                }
                
                thMercado.innerHTML = `Mercado ${countBadge} ${arrowHTML}`;
            }

            // 4. Atualizar indicadores de ordenação nos cabeçalhos
            document.querySelectorAll('.pricing-table th.sortable-th').forEach(th => {
                const c = th.getAttribute('data-col');
                const arrow = th.querySelector('.sort-arrow');
                const isColActive = (c === currentSortCol) || (c === 'gapRS' && currentSortCol === 'gapRS_loss');
                if (isColActive) {
                    th.classList.add('active-sort');
                    if (arrow) {
                        arrow.innerHTML = currentSortAsc ? ' ▲' : ' ▼';
                        arrow.style.color = 'var(--primary-yellow)';
                        arrow.style.opacity = '1';
                    }
                } else {
                    th.classList.remove('active-sort');
                    if (arrow) {
                        arrow.innerHTML = ' ↕';
                        arrow.style.color = 'var(--text-muted)';
                        arrow.style.opacity = '0.4';
                    }
                }
            });

            const tbody = document.getElementById('pricing-tbody');
            if (displayList.length === 0) {
                const msg = activePricingKpiFilter 
                    ? `Nenhum horário encontrado para este filtro de embate tarifário.`
                    : `Nenhum horário encontrado para os filtros selecionados.`;
                tbody.innerHTML = `<tr><td colspan="10" style="text-align: center; color: var(--text-muted); padding: 20px;">${msg}</td></tr>`;
                setTimeout(adjustAllLayouts, 20);
                return;
            }

            let htmlRows = '';
            displayList.forEach((item, idx) => {
                let faixaCompStr = '-';
                let assentosStr = `<b style="color: var(--primary-yellow);">${item.nossoAssentos}</b> x <span style="color: var(--text-muted);">-</span>`;
                let gapRsStr = '<span style="color: var(--text-muted);">-</span>';
                let gapPctStr = '<span style="color: var(--text-muted);">-</span>';
                let validationStr = `<span style="color: var(--text-muted); font-style: italic;">${item.validation}</span>`;
                let compBadge = `<span style="color: var(--text-muted);">0 Player(s)</span> <span style="font-size: 10px;">▼</span>`;

                if (item.hasComps) {
                    faixaCompStr = `R$ ${formatBR(item.compMin, true)} - R$ ${formatBR(item.compMax, true)}`;
                    assentosStr = `<b style="color: var(--primary-yellow);">${item.nossoAssentos}</b> x ${item.elesAssentos} <span style="font-size: 11px; color: var(--text-muted); font-weight: normal;">(${item.compCountTrips} viag.)</span>`;
                    const gapColor = item.gapRS >= 0 ? 'var(--alert-green)' : 'var(--alert-red)';
                    const gapSignal = item.gapRS > 0 ? '+' : '';
                    gapRsStr = `<span style="font-weight: bold; color: ${gapColor};">${gapSignal}R$ ${formatBR(item.gapRS, true)}</span>`;
                    gapPctStr = `<span style="font-weight: bold; color: ${gapColor};">${gapSignal}${item.gapPct.toFixed(1)}%</span>`;
                    
                    const isDirect = item.validation === 'Direto';
                    validationStr = `<b style="color: ${isDirect ? 'var(--alert-green)' : 'var(--alert-red)'}; font-style: ${isDirect ? 'normal' : 'italic'};">${item.validation}</b>`;
                    compBadge = `${item.compCountText} <span style="font-size: 10px;">▼</span>`;
                }

                htmlRows += `
                    <tr style="cursor: pointer;" onclick="togglePricingDetail('${item.rowId}', ${idx})">
                        <td><b>${item.mercado}</b></td>
                        <td><b style="color: var(--primary-yellow);">${item.horarioStr}</b></td>
                        <td>${item.classe}</td>
                        <td>${compBadge}</td>
                        <td><b style="color: var(--primary-yellow);">R$ ${formatBR(item.nossoPreco, true)}</b></td>
                        <td>${faixaCompStr}</td>
                        <td>${assentosStr}</td>
                        <td>${gapRsStr}</td>
                        <td>${gapPctStr}</td>
                        <td>${validationStr}</td>
                    </tr>
                    <tr id="${item.rowId}" style="display: none;">
                        <td colspan="10" style="padding: 0;"></td>
                    </tr>
                `;
            });
            tbody.innerHTML = htmlRows;
            setTimeout(adjustAllLayouts, 20);
        }

        function buildPricingDetailHTML(item) {
            let detailRows = '';
            if (item.hasComps && item.ringueArray) {
                item.ringueArray.forEach(t => {
                    const bgRow = t.isIta ? 'rgba(241, 196, 15, 0.1)' : 'transparent';
                    const vName = t.isIta ? `<b style="color: var(--primary-yellow);">${t.viacao}</b>` : t.viacao;
                    let gapStr = '-';
                    let tagCol = `<span style="color: var(--text-muted);">-</span>`;
                    if (!t.isIta) {
                        const diffVal = item.nossoPreco - t.preco; 
                        const gapColorSub = diffVal >= 0 ? 'var(--alert-green)' : 'var(--alert-red)';
                        gapStr = `<span style="color:${gapColorSub}; font-weight:bold;">${diffVal >= 0 ? '+' : ''} R$ ${formatBR(diffVal, true)}</span>`;
                        const isDirect = t.valTag === 'Direto';
                        tagCol = `<span style="font-weight: bold; color: ${isDirect ? 'var(--alert-green)' : 'var(--alert-red)'};">${t.valTag}</span>`;
                    } else {
                        tagCol = `<b style="color: var(--primary-yellow);">Referência Ita</b>`;
                    }
                    detailRows += `
                        <tr style="background-color: ${bgRow};">
                            <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${t.hora}</td>
                            <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${vName}</td>
                            <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${t.classe}</td>
                            <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color); font-weight:bold;">R$ ${formatBR(t.preco, true)}</td>
                            <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${t.assentos}</td>
                            <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${gapStr}</td>
                            <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${tagCol}</td>
                        </tr>
                    `;
                });
            } else {
                detailRows = `
                    <tr style="background-color: rgba(241, 196, 15, 0.1);">
                        <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${item.itaHorarioFull || item.horarioStr}</td>
                        <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);"><b style="color: var(--primary-yellow);">${item.itaViacao || 'ITAPEMIRIM'}</b></td>
                        <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${item.classe}</td>
                        <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color); font-weight:bold;">R$ ${formatBR(item.nossoPreco, true)}</td>
                        <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);">${item.nossoAssentos}</td>
                        <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color); color: var(--text-muted);">-</td>
                        <td style="padding: 12px 15px; border-bottom: 1px solid var(--border-color);"><b style="color: var(--primary-yellow);">Referência Ita</b></td>
                    </tr>
                    <tr>
                        <td colspan="7" style="text-align: center; color: var(--text-muted); padding: 12px; font-style: italic;">Nenhum concorrente operando neste trecho.</td>
                    </tr>
                `;
            }

            return `
                <div style="padding: 15px; background-color: var(--bg-body); border-radius: 0 0 8px 8px; border: 1px solid var(--border-color); border-top: none;">
                    <h4 style="margin: 0 0 10px 0; color: var(--text-muted); text-transform: uppercase; font-size: 12px;">📊 Embate Detalhado da Partida</h4>
                    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                        <thead>
                            <tr style="color: var(--text-muted); text-transform: uppercase; border-bottom: 2px solid var(--border-color); background-color: var(--input-bg);">
                                <th style="padding: 12px 15px; text-align: left;">Partida (Data/Hora)</th>
                                <th style="padding: 12px 15px; text-align: left;">Viação</th>
                                <th style="padding: 12px 15px; text-align: left;">Classe</th>
                                <th style="padding: 12px 15px; text-align: left;">Preço Cobrado</th>
                                <th style="padding: 12px 15px; text-align: left;">Vagas Livres</th>
                                <th style="padding: 12px 15px; text-align: left;">Diferença (Ita vs Concorrência)</th>
                                <th style="padding: 12px 15px; text-align: left;">Disputa</th>
                            </tr>
                        </thead>
                        <tbody>${detailRows}</tbody>
                    </table>
                </div>
            `;
        }

        function togglePricingDetail(rowId, index) {
            const row = document.getElementById(rowId);
            if (!row) return;
            if (row.style.display === 'table-row') {
                row.style.display = 'none';
            } else {
                if (!row.dataset.loaded) {
                    const item = (window._currentPricingDisplayList && window._currentPricingDisplayList[index]) || pricingDataArray[index];
                    if (item) {
                        row.children[0].innerHTML = buildPricingDetailHTML(item);
                        row.dataset.loaded = 'true';
                    }
                }
                row.style.display = 'table-row';
            }
        }

        // ==========================================
        // ABA DE CLASSES: VISUALIZAÇÃO EM CARDS COMPACTOS
        // ==========================================
        function getClasseIcon(classe) {
            const c = (classe || '').toUpperCase();
            if (c.includes('CAMA')) return '🛏️';
            if (c.includes('LEITO')) return '🛋️';
            if (c.includes('SEMILEITO')) return '💺';
            if (c.includes('CONVENCIONAL')) return '🚌';
            return '🎫';
        }

        function renderClassesTab(baseMarketData) {
            const grid = document.getElementById('classes-grid-container');
            const summaryStrip = document.getElementById('classes-summary-strip');
            grid.innerHTML = '';
            if (summaryStrip) summaryStrip.innerHTML = '';

            let classesAtivas = [...new Set(baseMarketData.map(i => i['Classe Padronizada']))];
            if (selections.classe.length > 0) classesAtivas = classesAtivas.filter(c => selections.classe.includes(c));

            if (classesAtivas.length === 0) {
                grid.innerHTML = `<p style="text-align: center; color: var(--text-muted); padding: 30px; width: 100%; font-size: 14px;">Nenhuma viagem encontrada para os filtros selecionados.</p>`;
                return;
            }

            // Ordenação Hierárquica por Conforto (Grupo 1 -> Grupo 2)
            const classOrder = {
                'CAMA INDIVIDUAL': 1,
                'CAMA DUPLO': 2,
                'LEITO DUPLO': 3,
                'LEITO': 4,
                'SEMILEITO': 5,
                'CONVENCIONAL': 6,
                'CONEXÃO': 7,
                'CONEXAO': 7
            };
            classesAtivas.sort((a, b) => {
                const orderA = classOrder[a.toUpperCase().trim()] || 99;
                const orderB = classOrder[b.toUpperCase().trim()] || 99;
                return orderA - orderB;
            });

            let html = '';
            let totalClassesMkt = classesAtivas.length;
            let totalClassesIta = 0;
            let highestTicketClass = null;
            let highestTicketVal = 0;
            let totalItaSeats = 0;
            let totalMktSeats = 0;

            const totalCards = classesAtivas.length;
            classesAtivas.forEach((cls, idx) => {
                // 1. Dados da Itapemirim para esta classe (sempre preservados como benchmark)
                let isSemileitoRef = false;
                let itaClsData = baseMarketData.filter(row => {
                    if (row['Classe Padronizada'] !== cls) return false;
                    const v = (row['Viação'] || '').toUpperCase();
                    return v.includes('ITAPEMIRIM') || v.includes('SUZANTUR');
                });

                // Regra de comparabilidade: no Convencional, se a Ita não tiver Convencional, compara com o Semileito da Itapemirim
                if (cls === 'CONVENCIONAL' && itaClsData.length === 0) {
                    itaClsData = baseMarketData.filter(row => {
                        const v = (row['Viação'] || '').toUpperCase();
                        const isIta = v.includes('ITAPEMIRIM') || v.includes('SUZANTUR');
                        return isIta && row['Classe Padronizada'] === 'SEMILEITO';
                    });
                    isSemileitoRef = true;
                }

                // 2. Dados dos Concorrentes para esta classe (estritamente concorrentes, eliminando autocontaminação)
                const compClsData = baseMarketData.filter(row => {
                    if (row['Classe Padronizada'] !== cls) return false;
                    const v = (row['Viação'] || '').toUpperCase();
                    const isOur = v.includes('ITAPEMIRIM') || v.includes('SUZANTUR');
                    if (isOur) return false;
                    if (selections.viacao.length > 0 && !selections.viacao.includes(row['Viação'])) return false;
                    return true;
                });

                // Se não há dados nem da Itapemirim nem de concorrentes para esta classe, pula
                if (itaClsData.length === 0 && compClsData.length === 0) return;

                // Preços e assentos Itapemirim
                const pIta = itaClsData.map(i => i['Preço (R$)']).filter(p => p > 0);
                const sIta = itaClsData.map(i => i['Assentos Disponíveis']).filter(s => s > 0);
                const itaMin = pIta.length ? Math.min(...pIta) : 0;
                const itaMedia = pIta.length ? (pIta.reduce((a,b)=>a+b, 0) / pIta.length) : 0;
                const itaMax = pIta.length ? Math.max(...pIta) : 0;
                const itaAssentos = sIta.length ? sIta.reduce((a,b)=>a+b, 0) : 0;
                const itaViagens = itaClsData.length;

                // Preços e assentos Concorrência / Mercado (estritamente concorrentes externos)
                const pComp = compClsData.map(i => i['Preço (R$)']).filter(p => p > 0);
                const sComp = compClsData.map(i => i['Assentos Disponíveis']).filter(s => s > 0);
                const mktMin = pComp.length ? Math.min(...pComp) : 0;
                const mktMedia = pComp.length ? (pComp.reduce((a,b)=>a+b, 0) / pComp.length) : 0;
                const mktMax = pComp.length ? Math.max(...pComp) : 0;
                const mktAssentos = sComp.length ? sComp.reduce((a,b)=>a+b, 0) : 0;
                const mktViagens = compClsData.length;

                // Total de volume para cálculo de shares (Mercado total da rota/classe = Concorrência + Itapemirim)
                const totalAssentos = mktAssentos + itaAssentos;
                const totalViagens = mktViagens + itaViagens;
                const shareAssentosNum = (totalAssentos > 0 && itaAssentos > 0) ? ((itaAssentos / totalAssentos) * 100) : 0;
                const shareViagensNum = (totalViagens > 0 && itaViagens > 0) ? ((itaViagens / totalViagens) * 100) : 0;


                // 1. Ticket Médio e GAP
                let diffMedioPct = 0;
                let hasMedioGap = false;
                if (itaMedia > 0 && mktMedia > 0) {
                    diffMedioPct = ((itaMedia - mktMedia) / mktMedia) * 100;
                    hasMedioGap = true;
                }

                // Status de Posicionamento Estratégico (Badge Minimalista no Header com tolerância de 5% sobre Ticket Médio)
                let statusTag = '';
                let statusClass = '';
                if (itaViagens === 0) {
                    statusTag = 'Sem Operação Ita';
                    statusClass = '';
                } else if (compClsData.length === 0) {
                    statusTag = 'Exclusivo Ita';
                    statusClass = 'competitive';
                } else if (Math.abs(diffMedioPct) <= 5.0) {
                    statusTag = 'Competitivo';
                    statusClass = 'competitive';
                } else if (diffMedioPct > 5.0) {
                    statusTag = 'Premium (+Yield)';
                    statusClass = 'premium';
                } else {
                    statusTag = 'Agressivo (Piso)';
                    statusClass = 'aggressive';
                }

                // 2. Teto Máximo e GAP
                let diffTetoPct = 0;
                let hasTetoGap = false;
                if (itaMax > 0 && mktMax > 0) {
                    diffTetoPct = ((itaMax - mktMax) / mktMax) * 100;
                    hasTetoGap = true;
                }

                // 3. Piso Mínimo e GAP
                let diffMinPct = 0;
                let hasMinGap = false;
                if (itaMin > 0 && mktMin > 0) {
                    diffMinPct = ((itaMin - mktMin) / mktMin) * 100;
                    hasMinGap = true;
                }

                function formatGapBadge(diffPct, hasGap) {
                    if (!hasGap) return '';
                    if (diffPct > 0.05) {
                        return `<span class="gap-badge-inline gap-pos">▲ +${diffPct.toFixed(1)}%</span>`;
                    } else if (diffPct < -0.05) {
                        return `<span class="gap-badge-inline gap-neg">▼ ${diffPct.toFixed(1)}%</span>`;
                    } else {
                        return `<span class="gap-badge-inline gap-neutral">0.0%</span>`;
                    }
                }

                const medioBadgeHtml = formatGapBadge(diffMedioPct, hasMedioGap);
                const tetoBadgeHtml = formatGapBadge(diffTetoPct, hasTetoGap);
                const minBadgeHtml = formatGapBadge(diffMinPct, hasMinGap);

                // Valores formatados Mercado
                const mktMinDisplay = mktMin > 0 ? `R$ ${formatBR(mktMin, true)}` : '-';
                const mktMediaDisplay = mktMedia > 0 ? `R$ ${formatBR(mktMedia, true)}` : '-';
                const mktMaxDisplay = mktMax > 0 ? `R$ ${formatBR(mktMax, true)}` : '-';

                // Valores formatados Itapemirim
                let itaMinDisplay = '-';
                let itaMediaDisplay = '-';
                let itaMaxDisplay = '-';
                let itaValClass = 'ita';

                if (itaViagens > 0) {
                    itaMinDisplay = `R$ ${formatBR(itaMin, true)}`;
                    itaMediaDisplay = `R$ ${formatBR(itaMedia, true)}`;
                    itaMaxDisplay = `R$ ${formatBR(itaMax, true)}`;
                } else {
                    itaValClass = 'ita no-op';
                }

                // Share de Oferta (Assentos e Viagens)
                const shareGaugeWidth = Math.min(100, Math.max(0, shareAssentosNum)).toFixed(1);
                const shareViagensWidth = Math.min(100, Math.max(0, shareViagensNum)).toFixed(1);

                // Acúmulos para a barra de resumo executivo
                if (itaViagens > 0) {
                    totalClassesIta++;
                    if (itaMedia > highestTicketVal) {
                        highestTicketVal = itaMedia;
                        highestTicketClass = cls;
                    }
                }
                totalItaSeats += itaAssentos;
                totalMktSeats += totalAssentos;

                // Classe CSS de layout para Grid em 2 Linhas
                let cardLayoutClass = '';
                if (totalCards >= 7) {
                    cardLayoutClass = (idx < 4) ? 'card-r4' : 'card-r3';
                } else if (totalCards === 6) {
                    cardLayoutClass = 'card-r3';
                } else if (totalCards === 5) {
                    cardLayoutClass = (idx < 3) ? 'card-r3' : 'card-r2';
                } else if (totalCards === 4) {
                    cardLayoutClass = 'card-r4';
                } else if (totalCards === 3) {
                    cardLayoutClass = 'card-r3';
                } else if (totalCards === 2) {
                    cardLayoutClass = 'card-r2-wide';
                } else {
                    cardLayoutClass = 'card-r1-solo';
                }

                html += `
                    <div class="class-panoramic-card ${cardLayoutClass}">
                        <div class="class-card-header">
                            <span class="class-title" title="${cls}">${cls}</span>
                            <span class="status-tag ${statusClass}">${statusTag}</span>
                        </div>

                        <div class="price-table-box">
                            <div class="price-table-header">
                                <span>Métrica</span>
                                <span class="th-mkt">🌍 Mercado</span>
                                <span class="th-ita">🚌 Itapemirim</span>
                            </div>
                            <div class="price-row">
                                <span class="price-lbl">Ticket Médio</span>
                                <span class="price-val-mkt">${mktMediaDisplay}</span>
                                <span class="price-val-ita ${itaValClass}">
                                    ${itaMediaDisplay}
                                    ${medioBadgeHtml}
                                </span>
                            </div>
                            <div class="price-row">
                                <span class="price-lbl">Teto Máximo</span>
                                <span class="price-val-mkt">${mktMaxDisplay}</span>
                                <span class="price-val-ita ${itaValClass}">
                                    ${itaMaxDisplay}
                                    ${tetoBadgeHtml}
                                </span>
                            </div>
                            <div class="price-row">
                                <span class="price-lbl">Piso Mínimo</span>
                                <span class="price-val-mkt">${mktMinDisplay}</span>
                                <span class="price-val-ita ${itaValClass}">
                                    ${itaMinDisplay}
                                    ${minBadgeHtml}
                                </span>
                            </div>
                        </div>

                        <div class="capacity-container">
                            <div class="capacity-card">
                                <div class="capacity-header">
                                    <span class="capacity-title">💺 Assentos Disponíveis</span>
                                    <span class="capacity-share-badge" style="color: ${shareAssentosNum > 40 ? 'var(--alert-red)' : (shareAssentosNum > 25 ? 'var(--primary-yellow)' : 'var(--alert-green)')};">${(totalAssentos > 0 && itaAssentos > 0) ? shareAssentosNum.toFixed(1) + '%' : '0.0%'}</span>
                                </div>
                                <div class="capacity-numbers">
                                    <span>${formatBR(itaAssentos, false)} <span class="sub-total">/ ${formatBR(totalAssentos, false)} vagos</span></span>
                                </div>
                                <div class="capacity-track">
                                    <div class="capacity-fill" style="width: ${shareGaugeWidth}%; background-color: ${shareAssentosNum > 40 ? 'var(--alert-red)' : (shareAssentosNum > 25 ? 'var(--primary-yellow)' : 'var(--alert-green)')};"></div>
                                </div>
                            </div>

                            <div class="capacity-card">
                                <div class="capacity-header">
                                    <span class="capacity-title">🚌 Viagens Ofertadas</span>
                                    <span class="capacity-share-badge">${(totalViagens > 0 && itaViagens > 0) ? shareViagensNum.toFixed(1) + '%' : '0.0%'}</span>
                                </div>
                                <div class="capacity-numbers">
                                    <span>${formatBR(itaViagens, false)} <span class="sub-total">/ ${formatBR(totalViagens, false)}</span></span>
                                </div>
                                <div class="capacity-track">
                                    <div class="capacity-fill" style="width: ${shareViagensWidth}%;"></div>
                                </div>
                            </div>
                        </div>

                        <div class="class-card-footer">
                            <button class="ghost-action-btn" onclick="focusClassInPricing('${cls}')" title="Filtrar '${cls}' e analisar no Pricing">
                                Analisar no Pricing ➔
                            </button>
                        </div>
                    </div>
                `;
            });

            grid.innerHTML = html;

            if (summaryStrip) {
                const shareGeralAssentos = totalMktSeats > 0 ? ((totalItaSeats / totalMktSeats) * 100) : 0;

                // Semântica de Ociosidade: Quanto maior o share de assentos vagos da Itapemirim frente ao mercado, pior (ônibus rodando vazio).
                let statusColor = 'var(--alert-green)';
                let statusBadge = '<span style="font-size:10px; padding:2px 6px; border-radius:3px; background:rgba(46,204,113,0.15); color:var(--alert-green); font-weight:700;">🟢 Baixa Ociosidade</span>';
                
                if (shareGeralAssentos > 40.0) {
                    statusColor = 'var(--alert-red)';
                    statusBadge = '<span style="font-size:10px; padding:2px 6px; border-radius:3px; background:rgba(231,76,60,0.15); color:var(--alert-red); font-weight:700;">🚨 Alta Ociosidade (Assentos Vagos)</span>';
                } else if (shareGeralAssentos > 25.0) {
                    statusColor = 'var(--primary-yellow)';
                    statusBadge = '<span style="font-size:10px; padding:2px 6px; border-radius:3px; background:rgba(241,196,15,0.15); color:var(--primary-yellow); font-weight:700;">⚠️ Ociosidade Moderada</span>';
                }

                summaryStrip.innerHTML = `
                    <div class="summary-chip" style="min-width: 250px; border-left: 3px solid ${statusColor};" title="Assentos disponíveis/vagos na plataforma. Quanto maior o share, maior a proporção de assentos não vendidos da Itapemirim em relação ao mercado.">
                        <div style="display:flex; justify-content:space-between; align-items:center; gap:8px;">
                            <span class="chip-lbl">Share Geral de Assentos Disponíveis (Vagos)</span>
                            ${statusBadge}
                        </div>
                        <div style="display:flex; align-items:baseline; gap:8px; margin-top:2px;">
                            <span class="chip-val"><b style="color: ${statusColor}; font-size: 15px;">${shareGeralAssentos.toFixed(1)}%</b></span>
                            <span style="font-size: 11px; color: var(--text-muted);">${formatBR(totalItaSeats, false)} de ${formatBR(totalMktSeats, false)} assentos disponíveis no mercado</span>
                        </div>
                    </div>
                `;
            }
            setTimeout(adjustAllLayouts, 20);
        }

        function focusClassInPricing(cls) {
            if (cls === 'CONVENCIONAL') {
                selections.classe = ['CONVENCIONAL', 'SEMILEITO'];
            } else {
                selections.classe = [cls];
            }
            const optClasse = document.getElementById('opt-classe');
            if (optClasse) {
                optClasse.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                    cb.checked = selections.classe.includes(cb.value);
                });
                const allCb = optClasse.querySelector('input[value="ALL"]');
                if (allCb) allCb.checked = false;
            }
            updateButtonText('classe');
            switchTab('tab-pricing');
            applyDataFilter();
        }

        function applyDataFilter() {
            const dtInicio = dtInicioInput.value;
            const dtFim = dtFimInput.value;

            const baseMarketData = rawData.filter(row => {
                const rowDate = row['Horário Partida'].substring(0, 10);
                if (dtInicio && rowDate < dtInicio) return false;
                if (dtFim && rowDate > dtFim) return false;
                if (selections.origem.length > 0 && !selections.origem.includes(row['Origem'])) return false;
                if (selections.destino.length > 0 && !selections.destino.includes(row['Destino'])) return false;
                return true;
            });

            // Concorrência selecionada para os cards de Visão de Mercado (sempre estritamente concorrentes)
            const compRows = baseMarketData.filter(row => {
                if (selections.classe.length > 0 && !selections.classe.includes(row['Classe Padronizada'])) return false;
                const vUpper = (row['Viação'] || '').toUpperCase();
                const isOur = vUpper.includes('ITAPEMIRIM') || vUpper.includes('SUZANTUR');
                if (selections.viacao.length > 0) {
                    return selections.viacao.includes(row['Viação']) && !isOur;
                }
                return !isOur;
            });

            // filteredData inclui SEMPRE a Itapemirim (como benchmark) + Concorrentes selecionados
            window.filteredData = baseMarketData.filter(row => {
                if (selections.classe.length > 0 && !selections.classe.includes(row['Classe Padronizada'])) return false;
                if (selections.viacao.length > 0) {
                    const vUpper = (row['Viação'] || '').toUpperCase();
                    const isOur = vUpper.includes('ITAPEMIRIM') || vUpper.includes('SUZANTUR');
                    if (!isOur && !selections.viacao.includes(row['Viação'])) return false;
                }
                return true;
            });

            window.itaData = baseMarketData.filter(row => {
                if (selections.classe.length > 0 && !selections.classe.includes(row['Classe Padronizada'])) return false;
                const v = row['Viação'].toUpperCase();
                return v.includes('ITAPEMIRIM') || v.includes('SUZANTUR');
            });

            const baseAssentos = baseMarketData.map(i => i['Assentos Disponíveis']).filter(s => s > 0).reduce((a,b)=>a+b, 0);
            const baseViagens = baseMarketData.length;

            const pMkt = compRows.map(i => i['Preço (R$)']).filter(p => p > 0);
            const sMkt = compRows.map(i => i['Assentos Disponíveis']).filter(s => s > 0);
            
            const mktMin = pMkt.length ? Math.min(...pMkt) : 0;
            const mktMedia = pMkt.length ? pMkt.reduce((a,b)=>a+b)/pMkt.length : 0;
            const mktMax = pMkt.length ? Math.max(...pMkt) : 0;
            const mktAssentos = sMkt.length ? sMkt.reduce((a,b)=>a+b, 0) : 0;
            const mktViagens = compRows.length;
            
            document.getElementById('mkt-min').innerText = formatBR(mktMin, true);
            document.getElementById('mkt-media').innerText = formatBR(mktMedia, true);
            document.getElementById('mkt-max').innerText = formatBR(mktMax, true);
            document.getElementById('mkt-assentos').innerText = formatBR(mktAssentos, false);
            document.getElementById('mkt-viagens').innerText = formatBR(mktViagens, false);

            if (selections.viacao.length > 0 || selections.classe.length > 0) {
                document.getElementById('mkt-assentos-sub').innerHTML = `<span style="color: var(--text-muted);">(${(baseAssentos>0 ? (mktAssentos/baseAssentos)*100 : 0).toFixed(1)}% do mercado global)</span>`;
                document.getElementById('mkt-viagens-sub').innerHTML = `<span style="color: var(--text-muted);">(${(baseViagens>0 ? (mktViagens/baseViagens)*100 : 0).toFixed(1)}% do mercado global)</span>`;
            } else {
                document.getElementById('mkt-assentos-sub').innerHTML = '';
                document.getElementById('mkt-viagens-sub').innerHTML = '';
            }

            const pIta = window.itaData.map(i => i['Preço (R$)']).filter(p => p > 0);
            const sIta = window.itaData.map(i => i['Assentos Disponíveis']).filter(s => s > 0);
            
            const itaMin = pIta.length ? Math.min(...pIta) : 0;
            const itaMedia = pIta.length ? pIta.reduce((a,b)=>a+b)/pIta.length : 0;
            itapemirimMediaGlobal = itaMedia; 
            const itaMax = pIta.length ? Math.max(...pIta) : 0;
            const itaAssentos = sIta.length ? sIta.reduce((a,b)=>a+b, 0) : 0;
            const itaViagens = window.itaData.length;
            
            document.getElementById('ita-min').innerText = formatBR(itaMin, true);
            document.getElementById('ita-media').innerText = formatBR(itaMedia, true);
            document.getElementById('ita-max').innerText = formatBR(itaMax, true);
            document.getElementById('ita-assentos').innerText = formatBR(itaAssentos, false);
            document.getElementById('ita-viagens').innerText = formatBR(itaViagens, false);

            const formatVar = (val) => {
                if (val === 0 || isNaN(val)) return '';
                const color = val > 0 ? 'var(--alert-red)' : 'var(--alert-green)'; 
                const sign = val > 0 ? '▲' : '▼';
                return `<span style="color: ${color};">${sign} ${Math.abs(val).toFixed(1)}%</span>`;
            };
            
            let baseAssentosItaComp = baseAssentos;
            let baseViagensItaComp = baseViagens;
            
            if (selections.classe.length > 0) {
                const baseFilteredClass = baseMarketData.filter(row => selections.classe.includes(row['Classe Padronizada']));
                baseAssentosItaComp = baseFilteredClass.map(i => i['Assentos Disponíveis']).filter(s => s > 0).reduce((a,b)=>a+b, 0);
                baseViagensItaComp = baseFilteredClass.length;
            }
            
            const formatShare = (val) => { return (val===0||isNaN(val)) ? '' : `<span style="color: var(--text-muted);">(${val.toFixed(1)}% fatia)</span>`; };

            document.getElementById('ita-min-sub').innerHTML = formatVar(mktMin>0 ? ((itaMin/mktMin)-1)*100 : 0);
            document.getElementById('ita-media-sub').innerHTML = formatVar(mktMedia>0 ? ((itaMedia/mktMedia)-1)*100 : 0);
            document.getElementById('ita-max-sub').innerHTML = formatVar(mktMax>0 ? ((itaMax/mktMax)-1)*100 : 0);
            document.getElementById('ita-assentos-sub').innerHTML = formatShare(baseAssentosItaComp>0 ? (itaAssentos/baseAssentosItaComp)*100 : 0);
            document.getElementById('ita-viagens-sub').innerHTML = formatShare(baseViagensItaComp>0 ? (itaViagens/baseViagensItaComp)*100 : 0);
            
            window._lastBaseMarketData = baseMarketData;
            renderCurrentTab();
        }

        function drawCharts() {
            if (currentTab !== 'tab-mercado') return;
            
            const data = window.filteredData || rawData;
            if (data.length === 0) { Plotly.purge('chart-timeline'); Plotly.purge('chart-combo'); return; }

            const viacoes = [...new Set(data.map(i => i['Viação']))].sort();
            const heatColors = [[0, '#2ecc71'], [0.5, '#f1c40f'], [1, '#e74c3c']]; 
            const fontColor = isLightMode ? '#7f8c8d' : '#a0a0a0';
            const gridColor = isLightMode ? '#e1e8ed' : '#333333';
            const titleColor = isLightMode ? '#2c3e50' : '#ffffff';

            const viacoesLabels = viacoes.map(v => {
                if (v.toUpperCase().includes('ITAPEMIRIM') || v.toUpperCase().includes('SUZANTUR')) {
                    return `<b style="color: #f1c40f;">${v}</b>`;
                }
                return `<b>${v}</b>`;
            });

            const pricesPlot = data.map(i => i['Preço (R$)']).filter(p => p > 0);
            const localMinPrice = pricesPlot.length ? Math.min(...pricesPlot) : 0;
            const localMaxPrice = pricesPlot.length ? Math.max(...pricesPlot) : 1;

            const tracesTimeline = [];
            viacoes.forEach(v => {
                const vData = data.filter(i => i['Viação'] === v && i['Horário Partida'] !== '');
                if (vData.length === 0) return;
                
                const sizes = vData.map(i => {
                    let cap = getCapacidadeEstimada(i['Classe Padronizada']);
                    let lugaresOcupados = Math.max(0, cap - i['Assentos Disponíveis']);
                    let ocupacaoPercentual = lugaresOcupados / cap;
                    return (ocupacaoPercentual * 25) + 6; 
                });

                const customInfo = vData.map(i => {
                    let cap = getCapacidadeEstimada(i['Classe Padronizada']);
                    let lugaresOcupados = Math.max(0, cap - i['Assentos Disponíveis']);
                    let ocup = Math.min(100, Math.max(0, (lugaresOcupados / cap) * 100));
                    return [i['Classe Padronizada'], i['Assentos Disponíveis'], i['Preço (R$)'], ocup, cap];
                });

                tracesTimeline.push({
                    x: vData.map(i => i['Horário Partida']),
                    y: vData.map(i => i['Viação']),
                    type: 'scatter',
                    mode: 'markers', name: v,
                    marker: { 
                        size: sizes, 
                        color: vData.map(i => i['Preço (R$)']), 
                        coloraxis: "coloraxis", 
                        line: {width: 1, color: isLightMode ? '#ffffff' : '#121212'} 
                    },
                    customdata: customInfo,
                    hovertemplate: 
                        "<b>%{y}</b><br><br>" +
                        "💰 Preço: <b>R$ %{customdata[2]:,.2f}</b><br>" +
                        "👥 Ocupação Est.: <b>%{customdata[3]:.0f}%</b><br>" +
                        "-----------------------<br>" +
                        "Partida: %{x}<br>" +
                        "Classe: %{customdata[0]}<br>" +
                        "Assentos Livres: %{customdata[1]} de %{customdata[4]}<extra></extra>"
                });
            });
            
            const layoutTimeline = {
                title: { 
                    text: 'Dispersão de Preços e Volume de Ocupação<br><span style="font-size: 13px; color: ' + fontColor + ';">ℹ️ <b>Legenda:</b> Bolhas menores = Ônibus vazios | Bolhas maiores = Ônibus lotando</span>', 
                    font: { color: titleColor, size: 18 } 
                },
                paper_bgcolor: 'transparent', plot_bgcolor: 'transparent', dragmode: isPanMode ? 'pan' : 'zoom',
                margin: {t: 75, l: 150, r: 20, b: 50}, 
                font: { color: fontColor, size: 11 }, 
                xaxis: { title: 'Horário de Partida', gridcolor: gridColor },
                yaxis: { 
                    type: 'category', 
                    categoryorder: 'category descending', 
                    gridcolor: gridColor, 
                    tickmode: 'array',
                    tickvals: viacoes,
                    ticktext: viacoesLabels 
                },
                coloraxis: { colorscale: heatColors, cmin: localMinPrice, cmax: localMaxPrice, colorbar: { title: 'Preço R$', separatethousands: true } },
                showlegend: false
            };
            Plotly.newPlot('chart-timeline', tracesTimeline, layoutTimeline, {scrollZoom: true, displayModeBar: false});

            const volViagens = [];
            const tktMedio = [];
            const barColors = [];
            
            viacoes.forEach(v => {
                const vData = data.filter(i => i['Viação'] === v);
                volViagens.push(vData.length);
                const prices = vData.map(i => i['Preço (R$)']).filter(p => p > 0);
                tktMedio.push(prices.length ? prices.reduce((a,b)=>a+b)/prices.length : 0);
                
                if (v.toUpperCase().includes('ITAPEMIRIM') || v.toUpperCase().includes('SUZANTUR')) {
                    barColors.push('#f1c40f');
                } else {
                    barColors.push('#3498db');
                }
            });

            const traceBar = { x: viacoes, y: volViagens, type: 'bar', name: 'Volume de Viagens', marker: {color: barColors} };
            const traceLine = { x: viacoes, y: tktMedio, type: 'scatter', mode: 'lines+markers', name: 'Ticket Médio', yaxis: 'y2', marker: {color: '#f1c40f', size: 10, line: {width: 2, color: '#121212'}}, line: {width: 3, color: '#f1c40f'} };

            const layoutCombo = {
                title: { text: 'Comparativo de Competitividade: Volume x Preço Médio', font: { color: titleColor, size: 18 } },
                paper_bgcolor: 'transparent', plot_bgcolor: 'transparent', dragmode: 'pan',
                margin: {t: 50, l: 60, r: 60, b: 50}, font: { color: fontColor, size: 11 },
                xaxis: { 
                    gridcolor: gridColor,
                    tickmode: 'array',
                    tickvals: viacoes,
                    ticktext: viacoesLabels
                },
                yaxis: { title: 'Volume (Qtd. Viagens)', gridcolor: gridColor, side: 'left' },
                yaxis2: { title: 'Ticket Médio (R$)', overlaying: 'y', side: 'right', showgrid: false },
                showlegend: true, legend: {x: 1.05, y: 1}
            };
            Plotly.newPlot('chart-combo', [traceBar, traceLine], layoutCombo, {scrollZoom: true, displayModeBar: false});
        }

        let viagensSortCol = 'Horário Partida';
        let viagensSortAsc = true;
        let viagensPage = 1;
        const viagensPageSize = 50;

        function changeViagensPage(delta) {
            viagensPage += delta;
            drawTable();
        }

        function sortViagensTable(col) {
            if (viagensSortCol === col) {
                viagensSortAsc = !viagensSortAsc;
            } else {
                viagensSortCol = col;
                viagensSortAsc = ['Origem', 'Destino', 'Classe Padronizada', 'Viação'].includes(col);
            }
            viagensPage = 1;
            drawTable();
        }

        function drawTable() {
            if (currentTab !== 'tab-viagens') return;

            const data = window.filteredData || rawData;
            const tbody = document.getElementById('table-body');
            tbody.innerHTML = '';
            
            const totalItems = data.length;
            const totalPages = Math.max(1, Math.ceil(totalItems / viagensPageSize));
            if (viagensPage > totalPages) viagensPage = totalPages;
            if (viagensPage < 1) viagensPage = 1;

            const sortedData = [...data].sort((a, b) => {
                let valA = a[viagensSortCol];
                let valB = b[viagensSortCol];
                if (typeof valA === 'string') {
                    return viagensSortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
                }
                return viagensSortAsc ? (valA - valB) : (valB - valA);
            });

            // Atualiza setas no cabeçalho de viagens
            document.querySelectorAll('.data-table th.sortable-th').forEach(th => {
                const c = th.getAttribute('data-col');
                const arrow = th.querySelector('.sort-arrow');
                if (c === viagensSortCol) {
                    th.classList.add('active-sort');
                    if (arrow) {
                        arrow.innerHTML = viagensSortAsc ? ' ▲' : ' ▼';
                        arrow.style.color = 'var(--primary-yellow)';
                        arrow.style.opacity = '1';
                    }
                } else {
                    th.classList.remove('active-sort');
                    if (arrow) {
                        arrow.innerHTML = ' ↕';
                        arrow.style.color = 'var(--text-muted)';
                        arrow.style.opacity = '0.4';
                    }
                }
            });

            const startIndex = (viagensPage - 1) * viagensPageSize;
            const pageData = sortedData.slice(startIndex, startIndex + viagensPageSize);

            let rowsHtml = '';
            pageData.forEach(row => {
                const p = row['Preço (R$)'];
                const s = row['Assentos Disponíveis'];
                const v = (row['Viação'] || '').toUpperCase();
                
                let priceStyle = '';
                let seatStyle = '';
                
                let viacaoText = `<b>${row['Viação']}</b>`;
                if (v.includes('ITAPEMIRIM') || v.includes('SUZANTUR')) {
                    viacaoText = `<b style="color: var(--primary-yellow);">${row['Viação']}</b>`;
                }
                
                if (!v.includes('ITAPEMIRIM') && !v.includes('SUZANTUR') && p < itapemirimMediaGlobal && itapemirimMediaGlobal > 0) {
                    priceStyle = 'color: var(--alert-red); font-weight: bold;';
                }
                if (s > 0 && s <= 5) {
                    seatStyle = 'color: var(--alert-green); font-weight: bold;';
                }

                rowsHtml += `
                    <tr>
                        <td>${row['Horário Formatado']}</td>
                        <td>${viacaoText}</td>
                        <td>${row['Origem']}</td>
                        <td>${row['Destino']}</td>
                        <td>${row['Classe Padronizada']}</td>
                        <td style="${priceStyle}">R$ ${formatBR(p, true)}</td>
                        <td style="${seatStyle}">${s}</td>
                    </tr>
                `;
            });
            tbody.innerHTML = rowsHtml;

            // Atualiza barra de paginação
            const pageInfo = document.getElementById('viagens-page-info');
            if (pageInfo) {
                const startNum = totalItems === 0 ? 0 : startIndex + 1;
                const endNum = Math.min(startIndex + pageData.length, totalItems);
                pageInfo.innerHTML = `Mostrando <b>${startNum}-${endNum}</b> de <b>${formatBR(totalItems)}</b> viagens • Página <b>${viagensPage}</b> de <b>${totalPages}</b>`;
                const btnPrev = document.getElementById('btn-viagens-prev');
                const btnNext = document.getElementById('btn-viagens-next');
                if (btnPrev) {
                    btnPrev.disabled = (viagensPage <= 1);
                    btnPrev.style.opacity = viagensPage <= 1 ? '0.4' : '1';
                    btnPrev.style.cursor = viagensPage <= 1 ? 'not-allowed' : 'pointer';
                }
                if (btnNext) {
                    btnNext.disabled = (viagensPage >= totalPages);
                    btnNext.style.opacity = viagensPage >= totalPages ? '0.4' : '1';
                    btnNext.style.cursor = viagensPage >= totalPages ? 'not-allowed' : 'pointer';
                }
            }
            setTimeout(adjustAllLayouts, 20);
        }

        // ==========================================
        // AJUSTE DINÂMICO DE ALTURA DAS TABELAS E CARDS (FULL SCREEN VIEWPORT)
        // ==========================================
        function adjustAllLayouts() {
            const windowH = window.innerHeight;
            
            // 1. Aba Pricing: tabela ocupa até o fim da tela
            if (currentTab === 'tab-pricing') {
                const pricingContainer = document.querySelector('.pricing-container');
                if (pricingContainer) {
                    const rect = pricingContainer.getBoundingClientRect();
                    const availableH = windowH - rect.top - 16;
                    if (availableH > 220) {
                        pricingContainer.style.height = `${Math.floor(availableH)}px`;
                        pricingContainer.style.maxHeight = `${Math.floor(availableH)}px`;
                    }
                }
            } 
            // 2. Aba Visão por Classes: cards expandem verticalmente até o fim da tela
            else if (currentTab === 'tab-classes') {
                const container = document.getElementById('classes-grid-container');
                if (container) {
                    const cards = container.querySelectorAll('.class-panoramic-card');
                    if (cards.length > 0) {
                        const rect = container.getBoundingClientRect();
                        const availableH = windowH - rect.top - 16;
                        if (availableH > 300) {
                            const isTwoRows = cards.length > 4;
                            let cardH;
                            if (isTwoRows) {
                                cardH = Math.floor((availableH - 12) / 2);
                            } else {
                                cardH = Math.min(450, Math.floor(availableH));
                            }
                            cards.forEach(c => {
                                c.style.height = `${cardH}px`;
                                c.style.minHeight = `${cardH}px`;
                            });
                        }
                    }
                }
            }
            // 3. Aba Tabela de Viagens: tabela ocupa até o topo da barra de paginação
            else if (currentTab === 'tab-viagens') {
                const tableWrapper = document.querySelector('.table-wrapper');
                if (tableWrapper) {
                    const rect = tableWrapper.getBoundingClientRect();
                    const paginationBar = document.getElementById('viagens-pagination-bar');
                    const pagH = paginationBar ? (paginationBar.offsetHeight + 10) : 55;
                    const availableH = windowH - rect.top - pagH - 16;
                    if (availableH > 220) {
                        tableWrapper.style.height = `${Math.floor(availableH)}px`;
                        tableWrapper.style.maxHeight = `${Math.floor(availableH)}px`;
                    }
                }
            }
        }

        window.addEventListener('resize', adjustAllLayouts);
        initDropdowns();
        setDefaultDates();
        applyDataFilter();
        setTimeout(adjustAllLayouts, 100);
    </script>
</body>
</html>
"""

html_template = html_parte1 + html_parte2
html_final = (
    html_template
    .replace("{DATA_PLACEHOLDER}", json_data)
    .replace("{DATA_MINIMA}", str(data_min))
    .replace("{DATA_MAXIMA}", str(data_max))
    .replace("{DATA_PADRAO}", str(data_padrao_inicial))
    .replace("{LOGO_IMG_TAG}", logo_img_tag)
)

with open(caminho_html, 'w', encoding='utf-8') as f:
    f.write(html_final)

print("\n[SUCESSO] Código gerado 100% completo com DataGrid de Classes Compacto e Tabela Ordenável!")