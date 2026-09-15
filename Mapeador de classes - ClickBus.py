from playwright.sync_api import sync_playwright
import pandas as pd
import os
import time
import random
import json
from datetime import date, timedelta

pasta_atual = os.path.dirname(os.path.abspath(__file__))
caminho_historico = os.path.join(pasta_atual, 'relatorio_precos_clickbus.xlsx')
caminho_saida_frota = os.path.join(pasta_atual, 'capacidade_frota.xlsx')

# ==========================================
# 1. CÉREBRO E SANIDADE
# ==========================================
DICIONARIO_CLASSES = {
    "CONVENCIONAL COM AR": "CONVENCIONAL",
    "CONV.": "CONVENCIONAL",
    "EXECUTIVO": "SEMILEITO",
    "SEMI-LEITO": "SEMILEITO",
    "SEMI_LEITO": "SEMILEITO",
    "SEMI-LEITO**": "SEMILEITO",
    "SEMILEITO DD": "SEMILEITO",
    "LEITO DUPLO": "LEITO DUPLO",
    "LEITO INDIVIDUAL": "LEITO INDIVIDUAL",
    "LEITO INDIVIDUAL*": "LEITO INDIVIDUAL",
    "LEITO DD": "LEITO",
    "CAMA DUPLO": "CAMA DUPLO",
    "CAMA INDIVIDUAL": "CAMA INDIVIDUAL",
    "LEITO CAMA": "CAMA",
    "PRIMEIRA CLASSE": "CAMA",
    "CONEXÃO": "CONEXÃO"
}

REGRAS_DE_SANIDADE = {
    "CONVENCIONAL":     {"min": 38, "max": 54},
    "SEMILEITO":        {"min": 38, "max": 54},
    "LEITO":            {"min": 8,  "max": 30},
    "LEITO DUPLO":      {"min": 8,  "max": 30},
    "LEITO INDIVIDUAL": {"min": 4,  "max": 15},
    "CAMA":             {"min": 6,  "max": 20},
    "CAMA DUPLO":       {"min": 6,  "max": 20},
    "CAMA INDIVIDUAL":  {"min": 4,  "max": 10},
    "CONEXÃO":          {"min": 0,  "max": 100}
}

def padronizar_classe(nome_clickbus):
    nome_limpo = str(nome_clickbus).strip().upper()
    return DICIONARIO_CLASSES.get(nome_limpo, nome_limpo)

def valida_sanidade(classe, capacidade):
    if capacidade == 0: return False
    regra = REGRAS_DE_SANIDADE.get(classe)
    if not regra: return True
    return regra["min"] <= capacidade <= regra["max"]

# ==========================================
# 2. INTELIGÊNCIA DE ROTAS POR VIAÇÃO
# ==========================================
print("Lendo histórico para criar o roteiro de cada Viação...")
try:
    df_historico = pd.read_excel(caminho_historico)
    # Limpa linhas vazias na Viação
    df_historico = df_historico.dropna(subset=['Viação'])
    
    viacoes_alvo = list(df_historico['Viação'].unique())
    # Cria um dicionário onde a Chave é a Viação e o Valor é uma lista das Rotas que ela opera
    rotas_por_viacao = {}
    
    for viacao in viacoes_alvo:
        # Filtra a base para a viação atual e pega as 3 rotas mais frequentes dela (para não gastar tempo tentando rotas obscuras)
        rotas_da_viacao = df_historico[df_historico['Viação'] == viacao].groupby(['Origem', 'Destino']).size().sort_values(ascending=False).head(3).index.tolist()
        rotas_por_viacao[viacao] = [{'Origem': str(r[0]).strip().lower(), 'Destino': str(r[1]).strip().lower()} for r in rotas_da_viacao]
        
    print(f"-> {len(viacoes_alvo)} viações detectadas para mapeamento cirúrgico.")
except Exception as e:
    print(f"[ERRO] Falha ao ler histórico: {e}")
    exit()

datas_para_buscar = []
data_base = date(2026, 11, 10)
while len(datas_para_buscar) < 7:
    if data_base.weekday() in [1, 2]:
        datas_para_buscar.append(data_base.strftime('%Y-%m-%d'))
    data_base += timedelta(days=1)

memoria_frota = {}

def formatar_data_hora(dado):
    if isinstance(dado, dict) and 'schedule' in dado:
        return f"{dado['schedule'].get('date', '')} {dado['schedule'].get('time', '')}".strip()
    return str(dado)

def farejar_viagens(dado):
    if isinstance(dado, dict):
        for valor in dado.values():
            if isinstance(valor, list) and len(valor) > 0 and isinstance(valor[0], dict) and 'price' in valor[0]: return valor
            res = farejar_viagens(valor)
            if res: return res
    elif isinstance(dado, list):
        if len(dado) > 0 and isinstance(dado[0], dict) and 'price' in dado[0]: return dado
        for item in dado:
            res = farejar_viagens(item)
            if res: return res
    return None

# ==========================================
# 3. EXECUÇÃO DO ROBÔ SNIPER
# ==========================================
print("\n=== INICIANDO MAPEAMENTO POR VIAÇÃO ===")

with sync_playwright() as p:
    navegador = p.chromium.launch(headless=False, args=['--disable-blink-features=AutomationControlled', '--no-sandbox'])
    contexto = navegador.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    contexto.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    # Agora iteramos sobre as VIAÇÕES, não sobre todas as rotas
    viacoes_restantes = viacoes_alvo.copy()
    
    for idx, viacao in enumerate(viacoes_alvo, 1):
        if viacao not in viacoes_restantes:
            continue # Já foi mapeada e removida da lista
            
        rotas_da_vez = rotas_por_viacao.get(viacao, [])
        if not rotas_da_vez: continue
        
        print(f"\n[{idx}/{len(viacoes_alvo)}] Rastreiando a frota da: {viacao}")
        
        # O robô vai tentar achar a viação usando a sua rota mais popular. 
        # Se a rota 1 não der resultado em nenhuma data, ele pula para a rota 2 dela.
        viacao_mapeada = False
        
        for rota in rotas_da_vez:
            if viacao_mapeada: break
            
            origem = rota['Origem']
            destino = rota['Destino']
            print(f"   -> Viajando para o mercado: {origem} x {destino}")

            for data in datas_para_buscar:
                pagina = contexto.new_page()
                jsons_da_rede = []
                
                def interceptar(resposta):
                    if "application/json" in resposta.headers.get("content-type", ""):
                        try: jsons_da_rede.append(resposta.json())
                        except: pass
                
                pagina.on("response", interceptar)
                
                try:
                    pagina.goto("https://www.clickbus.com.br/", timeout=45000)
                    time.sleep(random.uniform(1.5, 3)) # Calibração de segurança mantida
                    url_site = f"https://www.clickbus.com.br/onibus/{origem}/{destino}?departureDate={data}"
                    pagina.goto(url_site, wait_until="networkidle", timeout=45000)
                    pagina.remove_listener("response", interceptar)
                    
                    viagens_encontradas = None
                    for j in jsons_da_rede:
                        viagens_encontradas = farejar_viagens(j)
                        if viagens_encontradas: break
                    
                    if not viagens_encontradas:
                        for texto in pagina.locator('script[type="application/json"]').all_inner_texts():
                            try:
                                viagens_encontradas = farejar_viagens(json.loads(texto))
                                if viagens_encontradas: break
                            except: pass
                    
                    if viagens_encontradas:
                        onibus_fisicos = {}
                        
                        for v in viagens_encontradas:
                            assentos = v.get('availableSeats', 0)
                            parte = v['parts'][0] if 'parts' in v and len(v['parts']) > 0 else v
                            
                            viacao_achada = parte.get('travelCompany', {}).get('name', parte.get('company', {}).get('name', 'N/A')).upper()
                            
                            # Filtro de Sniper: Só nos importamos se for a viação que viemos buscar
                            if viacao_achada != str(viacao).upper():
                                continue
                                
                            classe_suja = parte.get('serviceClass', {}).get('name', parte.get('seatClass', 'N/A'))
                            classe_limpa = padronizar_classe(classe_suja)
                            
                            saida = formatar_data_hora(parte.get('departure', parte.get('departureTime', {})))
                            chave_onibus = f"{viacao_achada}_{saida}"
                            
                            if chave_onibus not in onibus_fisicos: onibus_fisicos[chave_onibus] = []
                            onibus_fisicos[chave_onibus].append({'classe': classe_limpa, 'assentos': assentos})
                        
                        for chave, classes_vendidas in onibus_fisicos.items():
                            v_nome = chave.split('_')[0]
                            classes_vendidas = sorted(classes_vendidas, key=lambda x: x['assentos'], reverse=True)
                            
                            if len(classes_vendidas) == 1:
                                tipo, c1, cap1 = "SINGLE", classes_vendidas[0]['classe'], classes_vendidas[0]['assentos']
                                c2, cap2 = "N/A", 0
                            else:
                                tipo = "DOUBLE DECKER"
                                c1, cap1 = classes_vendidas[0]['classe'], classes_vendidas[0]['assentos']
                                c2, cap2 = classes_vendidas[1]['classe'], classes_vendidas[1]['assentos']
                            
                            chave_mestre = (v_nome, tipo, c1, c2)
                            if chave_mestre not in memoria_frota:
                                memoria_frota[chave_mestre] = {'cap1_amostras': [], 'cap2_amostras': []}
                            
                            memoria_frota[chave_mestre]['cap1_amostras'].append(cap1)
                            if cap2 > 0: memoria_frota[chave_mestre]['cap2_amostras'].append(cap2)
                            
                            # Correção do erro max() vazio
                            max_c1 = max(memoria_frota[chave_mestre]['cap1_amostras']) if memoria_frota[chave_mestre]['cap1_amostras'] else 0
                            max_c2 = max(memoria_frota[chave_mestre]['cap2_amostras']) if memoria_frota[chave_mestre]['cap2_amostras'] else 0

                            valido_c1 = valida_sanidade(c1, max_c1)
                            valido_c2 = valida_sanidade(c2, max_c2) if c2 != "N/A" else True
                            
                            # Se conseguiu mapear com segurança, avisa e sai da rota
                            if len(memoria_frota[chave_mestre]['cap1_amostras']) >= 4 and (valido_c1 and valido_c2):
                                if viacao in viacoes_restantes:
                                    viacoes_restantes.remove(viacao)
                                    viacao_mapeada = True
                                    print(f"      [OK] Frota mapeada com Sanity Check (Max Cap1: {max_c1})")
                                    break # Sai do loop das classes

                except Exception as e:
                    pass # Ignora erros de rede e continua
                finally:
                    pagina.close()
                time.sleep(random.uniform(2, 3))
                
                if viacao_mapeada: break # Sai do loop das datas se já mapeou
            if viacao_mapeada: break # Sai do loop das rotas se já mapeou

    navegador.close()

# ==========================================
# 4. EXPORTAÇÃO
# ==========================================
if memoria_frota:
    dados_exportacao = []
    for (viacao, tipo, c1, c2), dados in memoria_frota.items():
        max_cap1 = max(dados['cap1_amostras']) if dados['cap1_amostras'] else 0
        max_cap2 = max(dados['cap2_amostras']) if dados['cap2_amostras'] else 0
        total = max_cap1 + max_cap2
        
        status = "OK"
        if not valida_sanidade(c1, max_cap1): status = f"Revisar Classe 1 ({c1} c/ {max_cap1} lug.)"
        if c2 != "N/A" and not valida_sanidade(c2, max_cap2): status = f"Revisar Classe 2 ({c2} c/ {max_cap2} lug.)"

        dados_exportacao.append({
            'Viação': viacao,
            'Tipo de Veículo': tipo,
            'Classe 1': c1,
            'Capacidade 1': max_cap1,
            'Classe 2': c2,
            'Capacidade 2': max_cap2,
            'Capacidade Total': total,
            'Amostras': len(dados['cap1_amostras']),
            'Status': status
        })
    
    df_frota = pd.DataFrame(dados_exportacao)
    df_frota = df_frota.sort_values(by=['Viação', 'Tipo de Veículo'])
    df_frota.to_excel(caminho_saida_frota, index=False)
    print(f"\n[SUCESSO] Planilha salva em: {caminho_saida_frota}")
else:
    print("\n[!] Nenhuma frota foi mapeada com sucesso.")