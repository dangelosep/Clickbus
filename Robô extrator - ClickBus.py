from playwright.sync_api import sync_playwright
import pandas as pd
import os
import time
import random
import json
import datetime
import shutil

pasta_atual = os.path.dirname(os.path.abspath(__file__))
caminho_excel = os.path.join(pasta_atual, 'relatorio_precos_clickbus.xlsx')
caminho_backup = os.path.join(pasta_atual, 'relatorio_precos_clickbus_backup.xlsx')
caminho_entrada = os.path.join(pasta_atual, 'Slugs.xlsx')

# ==========================================
# 0. BACKUP AUTOMÁTICO DE SEGURANÇA
# ==========================================
if os.path.exists(caminho_excel):
    try:
        shutil.copy2(caminho_excel, caminho_backup)
        print(f"[*] Backup prévio de segurança criado com sucesso em:\n    -> {caminho_backup}")
    except Exception as e_bkp:
        print(f"[!] Aviso: Não foi possível criar backup: {e_bkp}")

# Carregamento prévio da base existente na memória (evita reabrir arquivo de 40.000 linhas 742 vezes no disco)
df_acumulado = None
if os.path.exists(caminho_excel):
    try:
        df_acumulado = pd.read_excel(caminho_excel)
        print(f"[*] Base existente carregada na memória: {len(df_acumulado)} registros.")
    except Exception as e_load:
        print(f"[!] Aviso: Não foi possível ler base existente: {e_load}")
        df_acumulado = None

# ==========================================
# 1. CARREGAMENTO DIRETO DA PLANILHA ÚNICA
# ==========================================
try:
    df_slugs = pd.read_excel(caminho_entrada)
    rotas = []
    
    for idx, row in df_slugs.iterrows():
        origem = str(row.iloc[0]).strip()
        destino = str(row.iloc[1]).strip()
        
        if origem.upper() in ['NAN', 'NONE', ''] or destino.upper() in ['NAN', 'NONE', '']:
            continue
            
        rotas.append({
            'origem_slug': origem,
            'destino_slug': destino
        })
        
    print(f"[*] Planilha carregada! {len(rotas)} rotas prontas para pesquisa.")
except Exception as e:
    print(f"[ERRO] Falha ao processar a planilha. Detalhe: {e}")
    rotas = []

# ==========================================
# 2. CALENDÁRIO DINÂMICO (Hoje + 7 dias à frente)
# ==========================================
hoje = datetime.date.today()
datas_para_buscar = [(hoje + datetime.timedelta(days=i)).strftime('%Y-%m-%d') for i in range(8)]
print(f"[*] Período que será pesquisado ({len(datas_para_buscar)} dias): de {datas_para_buscar[0]} até {datas_para_buscar[-1]}")

# ==========================================
# 3. FUNÇÕES AUXILIARES E FORMATAÇÃO DE DADOS
# ==========================================
def formatar_data_hora(dado):
    if not isinstance(dado, dict):
        return str(dado)
    
    if 'schedule' in dado and isinstance(dado['schedule'], dict):
        sched = dado['schedule']
        d = sched.get('date', '')
        t = sched.get('time', '')
        return f"{d} {t}".strip()
    
    if 'date' in dado and 'time' in dado:
        return f"{dado['date']} {dado['time']}".strip()
    if 'time' in dado:
        return str(dado['time'])
    
    return str(dado)

def extrair_trip_id(viagem):
    """
    Extrai o identificador único da viagem no JSON da ClickBus
    para viabilizar a deduplicação definitiva e diferenciar ônibus gêmeos.
    """
    if not isinstance(viagem, dict):
        return 'N/A'
    
    # 1. Tenta direto na raiz da viagem
    for k in ['id', 'tripId', 'serviceId', 'service_id', 'trip_id']:
        val = viagem.get(k)
        if val is not None and str(val).strip() != '':
            return str(val).strip()
            
    # 2. Tenta dentro do primeiro segmento (parts[0])
    if 'parts' in viagem and isinstance(viagem['parts'], list) and len(viagem['parts']) > 0:
        p0 = viagem['parts'][0]
        if isinstance(p0, dict):
            for k in ['tripId', 'serviceId', 'id', 'service_id', 'trip_id']:
                val = p0.get(k)
                if val is not None and str(val).strip() != '':
                    return str(val).strip()
                    
    return 'N/A'

def eh_lista_de_viagens(lista):
    if not isinstance(lista, list) or len(lista) == 0: return False
    primeiro = lista[0]
    if isinstance(primeiro, dict) and 'price' in primeiro:
        if any(chave in primeiro for chave in ['parts', 'company', 'travelCompany', 'departure']):
            return True
    return False

def farejar_viagens(dado):
    if isinstance(dado, dict):
        for valor in dado.values():
            if eh_lista_de_viagens(valor): return valor
            resultado = farejar_viagens(valor)
            if resultado: return resultado
    elif isinstance(dado, list):
        if eh_lista_de_viagens(dado): return dado
        for item in dado:
            resultado = farejar_viagens(item)
            if resultado: return resultado
    return None

# ==========================================
# 4. EXECUÇÃO DO ROBÔ
# ==========================================
if len(rotas) > 0:
    print("=== INICIANDO O ROBÔ DE CAPTURA ===")

    with sync_playwright() as p:
        navegador = p.chromium.launch(
            headless=False,
            args=['--disable-blink-features=AutomationControlled', '--disable-infobars', '--no-sandbox']
        )
        contexto = navegador.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 768},
            locale="pt-BR",
            timezone_id="America/Sao_Paulo"
        )
        contexto.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

        for idx, rota in enumerate(rotas, 1):
            print(f"\n[{idx}/{len(rotas)}] Rota: {rota['origem_slug']} -> {rota['destino_slug']}")
            passagens_desta_rota = []

            for data in datas_para_buscar:
                pagina = contexto.new_page()
                jsons_da_rede = []
                
                def interceptar(resposta):
                    if "application/json" in resposta.headers.get("content-type", ""):
                        try: jsons_da_rede.append(resposta.json())
                        except: pass
                
                pagina.on("response", interceptar)
                
                try:
                    # Aquecimento essencial de sessão na Home para renovar cookies anti-bot
                    pagina.goto("https://www.clickbus.com.br/", timeout=45000)
                    time.sleep(random.uniform(1.5, 3))
                    
                    url_site = f"https://www.clickbus.com.br/onibus/{rota['origem_slug']}/{rota['destino_slug']}?departureDate={data}"
                    resp_navegacao = pagina.goto(url_site, wait_until="networkidle", timeout=45000)
                    pagina.remove_listener("response", interceptar)
                    
                    # Pausa preventiva silenciosa caso a ClickBus retorne erro de rate limit (Circuit Breaker)
                    if resp_navegacao and resp_navegacao.status in [403, 429]:
                        print(f"\n   [ALERTA DE SEGURANÇA] Resposta HTTP {resp_navegacao.status} detectada na ClickBus!")
                        print("   [PAUSA PREVENTIVA] Pausando por 10 minutos para resfriar o IP e evitar bloqueio...")
                        time.sleep(600)
                        
                    if "Just a moment..." in pagina.title() or "Attention Required! | Cloudflare" in pagina.title():
                        print("\n   [ALERTA DE SEGURANÇA] Desafio Cloudflare detectado na tela!")
                        print("   [PAUSA PREVENTIVA] Pausando por 10 minutos para resfriar o IP...")
                        time.sleep(600)

                    viagens_encontradas = None
                    for j in jsons_da_rede:
                        viagens_encontradas = farejar_viagens(j)
                        if viagens_encontradas: break
                    
                    if not viagens_encontradas:
                        textos_ocultos = pagina.locator('script[type="application/json"]').all_inner_texts()
                        for texto in textos_ocultos:
                            try:
                                viagens_encontradas = farejar_viagens(json.loads(texto))
                                if viagens_encontradas: break
                            except: pass
                    
                    if not viagens_encontradas:
                        print(f"   [!] Sem passagens para {data}.")
                    else:
                        for viagem in viagens_encontradas:
                            preco = viagem.get('price')
                            viacao, classe, duracao = 'N/A', 'N/A', 'N/A'
                            assentos = viagem.get('availableSeats', 0)
                            
                            # Extração do Trip ID único para viabilizar ônibus gêmeos e deduplicação limpa
                            trip_id = extrair_trip_id(viagem)
                            
                            if 'parts' in viagem and len(viagem['parts']) > 0:
                                parte = viagem['parts'][0]
                                viacao = parte.get('travelCompany', {}).get('name', 'N/A')
                                classe = parte.get('serviceClass', {}).get('name', 'N/A')
                                saida = formatar_data_hora(parte.get('departure', {}))
                                chegada = formatar_data_hora(parte.get('arrival', {}))
                                duracao = parte.get('duration', 'N/A')
                            else:
                                viacao = viagem.get('company', {}).get('name', viagem.get('travelCompany', {}).get('name', 'N/A'))
                                classe = viagem.get('serviceClass', {}).get('name', viagem.get('seatClass', 'N/A'))
                                saida = formatar_data_hora(viagem.get('departure', viagem.get('departureTime', {})))
                                chegada = formatar_data_hora(viagem.get('arrival', viagem.get('arrivalTime', {})))
                                duracao = viagem.get('duration', 'N/A')

                            passagens_desta_rota.append({
                                'Trip ID': str(trip_id),
                                'Origem': rota['origem_slug'].upper(),
                                'Destino': rota['destino_slug'].upper(),
                                'Data': data,
                                'Viação': viacao,
                                'Classe': classe,
                                'Saída': saida,
                                'Chegada': chegada,
                                'Duração': duracao,
                                'Preço (R$)': preco,
                                'Assentos Disponíveis': assentos,
                                'Status da Busca': 'Sucesso'
                            })
                        print(f"   [OK] Capturou {len(viagens_encontradas)} viagens para {data}.")
                        
                except Exception as e:
                     print(f"   [ERRO] Falha em {data}: {e}")
                finally:
                    pagina.close()
                
                # Intervalo aleatório de segurança para comportamento humano
                time.sleep(random.uniform(2, 4))

            # ==========================================
            # 5. SALVAMENTO INCREMENTAL SEGURO E ATÔMICO
            # ==========================================
            if len(passagens_desta_rota) > 0:
                df_novo = pd.DataFrame(passagens_desta_rota)
                if df_acumulado is not None and not df_acumulado.empty:
                    df_acumulado = pd.concat([df_acumulado, df_novo], ignore_index=True)
                else:
                    df_acumulado = df_novo

                # Deduplicação inteligente no robô
                if 'Trip ID' in df_acumulado.columns and (df_acumulado['Trip ID'] != 'N/A').any():
                    m_valido = (df_acumulado['Trip ID'].notna()) & (~df_acumulado['Trip ID'].astype(str).str.upper().isin(['N/A', 'NAN', '']))
                    df_com_id = df_acumulado[m_valido].drop_duplicates(subset=['Trip ID', 'Data'], keep='last')
                    df_sem_id = df_acumulado[~m_valido].drop_duplicates(subset=['Origem', 'Destino', 'Data', 'Viação', 'Classe', 'Saída'], keep='last')
                    df_acumulado = pd.concat([df_com_id, df_sem_id], ignore_index=True)
                else:
                    df_acumulado = df_acumulado.drop_duplicates(subset=['Origem', 'Destino', 'Data', 'Viação', 'Classe', 'Saída'], keep='last')

                # Gravação atômica via arquivo temporário (impossível corromper se der Ctrl+C)
                try:
                    caminho_tmp = caminho_excel + '.tmp'
                    df_acumulado.to_excel(caminho_tmp, index=False)
                    if os.path.exists(caminho_tmp):
                        os.replace(caminho_tmp, caminho_excel)
                    print(f"   -> [Salvo] +{len(df_novo)} viagens gravadas. Total acumulado: {len(df_acumulado)} registros.")
                except Exception as e_save:
                    print(f"   [AVISO] Erro na gravação do Excel: {e_save}")

        navegador.close()
    print("\n==========================================")
    print("CAPTURA FINALIZADA COM SUCESSO!")
    print("==========================================")
else:
    print("\nNenhuma rota processada.")