import selenium
import re
import pandas as pd
import json
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.relative_locator import locate_with
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from nltk.tokenize import RegexpTokenizer
import openpyxl
from openpyxl.styles import PatternFill
from time import sleep
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from unidecode import unidecode
from bs4 import BeautifulSoup


# ======================== CONFIGURAÇÕES ==========================================

# Caminho do ChromeDriver
caminho = Service('/Users/Marcella/PycharmProjects/webscrapper101/chromedriver')

# Dataframe com os processos trabalhistas
df = pd.read_parquet('/Users/Marcella/pesquisa_insper_luciana/trt4_similariedade/data/interim/dftrt4_g1_ritoordinario.parquet')
print(f'Total de processos a scrappear: {len(df)}')
df = df.head()
print(f'Total de processos a scrappear: {len(df)}')

# Link base para pesquisa 
link_base = 'https://comunica.pje.jus.br/consulta?dataDisponibilizacaoInicio=2022-12-31&dataDisponibilizacaoFim=2026-05-09&numeroProcesso='

# Criar o Navegador
navegador = webdriver.Chrome(service=caminho)

# listas para acumular resultados
passivo_rows = []
ativo_rows = []
adv_rows = []

try:
    for processo in df['numero_processo']:
    # monta o link da busca 
        link_da_pesquisa = link_base + str(processo)

        #acessa a página de pesquisa do processo
        navegador.get(link_da_pesquisa)
        sleep(5)

        # tentar extrair usando Selenium (mais confiável em páginas dinâmicas)
        def clean_text(text):
            if not text:
                return ''
            txt = ILLEGAL_CHARACTERS_RE.sub('', text)
            txt = unidecode(txt)
            return re.sub(r"\s+", ' ', txt).strip()

        partes = []
        advogados = []

        try:
            # Pegar apenas o PRIMEIRO card/article com comunicações
            first_card = navegador.find_element(By.CSS_SELECTOR, 'article.card')
            
            # Extrair partes apenas do primeiro card
            elems = first_card.find_elements(By.CSS_SELECTOR, 'div.info-sumary.d-flex.align-items-center')
            for el in elems:
                polo_text = ''
                nome = ''
                
                # Obter TODO o texto do elemento (incluindo conteúdo de pseudo-elementos do Angular)
                texto_completo = el.get_attribute('textContent') or el.text
                texto_completo = clean_text(texto_completo)
                
                # Procurar por "Polo Ativo" ou "Polo Passivo" no texto completo
                if 'polo passivo' in texto_completo.lower():
                    polo_text = 'Polo Passivo'
                elif 'polo ativo' in texto_completo.lower():
                    polo_text = 'Polo Ativo'
                
                # Procurar nome: procurar todos os spans diretos (fora de tooltip-polo)
                try:
                    spans = el.find_elements(By.TAG_NAME, 'span')
                    for sp in spans:
                        txt = clean_text(sp.text)
                        # Pular spans vazios ou que contenham "polo"
                        if not txt or 'polo' in txt.lower():
                            continue
                        nome = txt
                        break
                except Exception:
                    pass
                
                if nome:
                    partes.append({'polo': polo_text, 'nome': nome})

            # Extrair advogados via Selenium: procurar todos os col-md-10 dentro do primeiro card que contêm OAB
            try:
                adv_nodes = first_card.find_elements(By.CSS_SELECTOR, 'div.col-md-10')
                for an in adv_nodes:
                    text = clean_text(an.text)
                    if not text:
                        continue
                    # Filtrar apenas se contém padrão de OAB
                    if ' - OAB ' not in text.upper():
                        continue
                    m = re.match(r"^(.*?)\s*-\s*OAB\s*(.+)$", text, flags=re.IGNORECASE)
                    if m:
                        nome_adv = m.group(1).strip()
                        oab = m.group(2).strip()
                    else:
                        nome_adv = text
                        oab = None
                    advogados.append({'nome': nome_adv, 'oab': oab})
            except Exception as e:
                advogados = []

        except Exception as e:
            # Fallback: procurar aside.card-sumary como alternativa
            first_card = None
            try:
                first_card = navegador.find_element(By.CSS_SELECTOR, 'aside.card-sumary')
            except Exception:
                pass
            
            if first_card:
                # Tentar extrair com aside.card-sumary
                try:
                    elems = first_card.find_elements(By.CSS_SELECTOR, 'div.info-sumary.d-flex.align-items-center')
                    for el in elems:
                        polo_text = ''
                        nome = ''
                        
                        # Procurar especificamente div.tooltip-polo > span.tooltip-text
                        try:
                            polo_span = el.find_element(By.XPATH, ".//div[@class='tooltip-polo']//span[@class='tooltip-text']")
                            polo_text = clean_text(polo_span.text)
                        except Exception:
                            pass
                        
                        # Procurar nome: span que não seja tooltip-text
                        try:
                            spans = el.find_elements(By.TAG_NAME, 'span')
                            for sp in spans:
                                cls = sp.get_attribute('class') or ''
                                if 'tooltip-text' in cls:
                                    continue
                                txt = clean_text(sp.text)
                                if txt:
                                    nome = txt
                                    break
                        except Exception:
                            pass
                        
                        if nome:
                            partes.append({'polo': polo_text, 'nome': nome})
                except Exception:
                    pass
                except Exception:
                    pass
                
                # Extrair advogados do aside.card-sumary
                try:
                    adv_nodes = first_card.find_elements(By.CSS_SELECTOR, 'div.col-md-10')
                    for an in adv_nodes:
                        text = clean_text(an.text)
                        if not text or ' - OAB ' not in text.upper():
                            continue
                        m = re.match(r"^(.*?)\s*-\s*OAB\s*(.+)$", text, flags=re.IGNORECASE)
                        if m:
                            nome_adv = m.group(1).strip()
                            oab = m.group(2).strip()
                        else:
                            nome_adv = text
                            oab = None
                        advogados.append({'nome': nome_adv, 'oab': oab})
                except Exception:
                    pass
            else:
                # Fallback final: BeautifulSoup
                soup = BeautifulSoup(navegador.page_source, 'html.parser')
                first_card_soup = soup.select_one('article.card') or soup.select_one('aside.card-sumary')
                if first_card_soup:
                    for info in first_card_soup.select('div.info-sumary.d-flex.align-items-center'):
                        # Procurar especificamente span.tooltip-text dentro de div.tooltip-polo
                        polo_text = ''
                        tooltip_div = info.find('div', class_='tooltip-polo')
                        if tooltip_div:
                            tooltip_span = tooltip_div.find('span', class_='tooltip-text')
                            if tooltip_span:
                                polo_text = clean_text(tooltip_span.get_text(strip=True))
                        
                        # Procurar nome: span que não seja tooltip-text
                        nome = None
                        spans = info.find_all('span')
                        for sp in spans:
                            cls = sp.get('class')
                            cls_str = ' '.join(cls) if cls else ''
                            if 'tooltip-text' in cls_str:
                                continue
                            txt = clean_text(sp.get_text(strip=True))
                            if txt:
                                nome = txt
                                break
                    
                        if nome:
                            partes.append({'polo': polo_text, 'nome': nome})
                    for col in first_card_soup.select('div.col-md-10'):
                        text = clean_text(col.get_text(separator=' ', strip=True))
                        if not text or ' - OAB ' not in text.upper():
                            continue
                        m = re.match(r"^(.*?)\s*-\s*OAB\s*(.+)$", text, flags=re.IGNORECASE)
                        if m:
                            nome_adv = m.group(1).strip()
                            oab = m.group(2).strip()
                        else:
                            nome_adv = text
                            oab = None
                        advogados.append({'nome': nome_adv, 'oab': oab})

        # armazenar resultados no formato solicitado
        for p in partes:
            polo = (p.get('polo') or '').lower().strip()
            nome = (p.get('nome') or '').strip()
            if 'passivo' in polo or 'passiv' in polo:
                passivo_rows.append({
                    'link': link_da_pesquisa,
                    'numero_processo': processo,
                    'nome_polo_passivo': nome
                })
            elif 'ativo' in polo or 'ativ' in polo:
                ativo_rows.append({
                    'link': link_da_pesquisa,
                    'numero_processo': processo,
                    'nome_polo_ativo': nome
                })

        for a in advogados:
            adv_rows.append({
                'link': link_da_pesquisa,
                'numero_processo': processo,
                'nome_advogado': a.get('nome'),
                'oab': a.get('oab')
            })

        # prints mínimos para acompanhamento
        print(f'Processo: {processo} — partes: {len(partes)} — advs: {len(advogados)}')
finally:
    try:
        navegador.quit()
    except Exception:
        pass

# Criar DataFrames finais
df_passivo = pd.DataFrame(passivo_rows, columns=['link', 'numero_processo', 'nome_polo_passivo'])
df_ativo = pd.DataFrame(ativo_rows, columns=['link', 'numero_processo', 'nome_polo_ativo'])
df_advogados = pd.DataFrame(adv_rows, columns=['link', 'numero_processo', 'nome_advogado', 'oab'])

# garantir diretório de saída
out_dir = os.path.join('data', 'webscraping', 'trt4_rito_ord')
os.makedirs(out_dir, exist_ok=True)

# salvar em parquet
df_passivo.to_csv(os.path.join(out_dir, 'df_passivo_trt4_rito_ord.csv'), index=False)
df_ativo.to_csv(os.path.join(out_dir, 'df_ativo_trt4_rito_ord.csv'), index=False)
df_advogados.to_csv(os.path.join(out_dir, 'df_advogados_trt4_rito_ord.csv'), index=False)

print('Salvos:', os.path.join(out_dir, 'df_passivo_trt4_rito_ord.csv'), os.path.join(out_dir, 'df_ativo_trt4_rito_ord.csv'), os.path.join(out_dir, 'df_advogados_trt4_rito_ord.csv'))

    


