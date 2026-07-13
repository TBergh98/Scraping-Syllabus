#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Syllabus Offerta Formativa Scraper
Questo script effettua lo scraping dei corsi presenti su Syllabus (www.syllabus.gov.it),
rileva le novità rispetto alle esecuzioni precedenti e genera un report JSON.
"""

import os
import sys
import json
import time
import random
import argparse
import datetime
import urllib.parse
import re
import requests
# pyrefly: ignore [missing-import]
from bs4 import BeautifulSoup

# Configurazione default
DEFAULT_START_URL = "https://www.syllabus.gov.it/portale/web/syllabus/offerta-formativa"
DEFAULT_DB_FILE = os.path.join("data", "corsi_storico.json")
DEFAULT_TIMEOUT = 60.0  # Timeout incrementato a 60 secondi come da richiesta utente
DEFAULT_DELAY_MIN = 1.5
DEFAULT_DELAY_MAX = 2.5
DEFAULT_RETRIES = 3

# Forza la codifica dello standard output su UTF-8 in ambienti Windows
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def get_headers():
    """Ritorna gli header HTTP simulando un browser reale."""
    return {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7",
        "Connection": "keep-alive"
    }


def fetch_with_retry(url, retries=DEFAULT_RETRIES, timeout=DEFAULT_TIMEOUT):
    """Scarica il contenuto di una pagina web gestendo i tentativi in caso di errore."""
    headers = get_headers()
    
    for attempt in range(retries):
        try:
            # print(f"  Richiesta a: {url} (Tentativo {attempt + 1}/{retries})...")
            response = requests.get(url, headers=headers, timeout=timeout)
            
            # Se la risposta è 200, la restituiamo
            if response.status_code == 200:
                return response
            
            print(f"  [AVVISO] Codice di stato HTTP {response.status_code} per URL: {url}")
            if response.status_code in [403, 404]:
                # Errori fatali non temporanei, inutile riprovare
                break
                
        except requests.exceptions.RequestException as e:
            print(f"  [AVVISO] Errore di rete durante il caricamento di {url}: {e}")
            
        # Backoff esponenziale con jitter
        if attempt < retries - 1:
            sleep_time = (2 ** attempt) + random.uniform(1.0, 3.0)
            print(f"  Attesa di {sleep_time:.2f} secondi prima del prossimo tentativo...")
            time.sleep(sleep_time)
            
    return None


def parse_index_page(html_content):
    """Analizza la pagina principale per estrarre le card dei corsi."""
    soup = BeautifulSoup(html_content, 'html.parser')
    cards = soup.find_all(class_="card-rainbow")
    
    courses = []
    for card in cards:
        # 1. Estrazione Titolo
        title_el = card.find('h3', class_='font-bold')
        title = title_el.get_text(strip=True) if title_el else "N/A"
        
        # 2. Estrazione Descrizione Breve
        desc_div = card.find('div', {'data-lfr-editable-id': lambda x: x and 'pragraph-id' in x})
        if not desc_div:
            desc_div = card.find('div', {'data-lfr-editable-type': 'rich-text'})
        desc_short = desc_div.get_text(strip=True) if desc_div else "N/A"
        
        # 3. Estrazione Link Dettagli
        link_el = card.find('a', class_='button')
        relative_link = link_el.get('href') if link_el else None
        
        if relative_link:
            full_link = urllib.parse.urljoin(DEFAULT_START_URL, relative_link)
        else:
            full_link = None
            
        if title != "N/A" and full_link:
            courses.append({
                "title": title,
                "desc_short": desc_short,
                "detail_url": full_link
            })
            
    return courses


def parse_details_page(html_content, url):
    """Analizza la pagina dei dettagli di un singolo corso."""
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 1. Titolo Esteso
    detail_title_el = soup.find('h2', {'data-lfr-editable-id': lambda x: x and x.endswith('-title-text')})
    if not detail_title_el:
        detail_title_el = soup.find('h2', class_='portlet-title-text')
    if not detail_title_el:
        # Fallback al primo h2 visibile
        for h2 in soup.find_all('h2'):
            classes = h2.get('class', [])
            if not any(c in classes for c in ['sr-only', 'visually-hidden', 'hide-accessible']):
                detail_title_el = h2
                break
                
    title = detail_title_el.get_text(strip=True) if detail_title_el else "N/A"
    
    # 2. Descrizione Estesa (paragrafi non nell'accordion)
    desc_paragraphs = []
    for p in soup.find_all('p'):
        # Escludiamo paragrafi all'interno delle sezioni accordion o footer
        if p.find_parent(class_='accordion-ctn') or p.find_parent('section', id=lambda x: x and x.startswith('accTxt-panel')):
            continue
        # Escludiamo footer/cookie banner
        text = p.get_text(strip=True)
        if text and len(text) > 12:
            if "cookie" not in text.lower() and "syllabus" not in text.lower() or "corso" in text.lower():
                desc_paragraphs.append(text)
                
    desc_extended = "\n\n".join(desc_paragraphs)
    
    # 3. Sezioni Accordion (Programma, Durate, Livelli)
    accordions = soup.find_all(class_='accordion-ctn')
    
    durations = []
    mastery_levels = set()
    syllabus_sections = []
    
    for idx, acc in enumerate(accordions):
        # Titolo sezione accordion
        header_el = acc.find(class_='header')
        header_txt = header_el.get_text(strip=True) if header_el else f"Modulo {idx + 1}"
        
        panel = acc.find('section')
        if not panel:
            continue
            
        # Cerca durata
        duration_txt = ""
        duration_el = panel.find(class_='intro')
        if duration_el:
            duration_txt = duration_el.get_text(strip=True)
            # Rimuoviamo intestazioni ripetitive se presenti
            duration_txt = duration_txt.replace("Durata media percorso da introduttivo a intermedio: ", "").replace("Durata: ", "")
            durations.append(duration_txt)
        else:
            # Fallback testuale
            dur_search = panel.find(text=re.compile(r'Durata', re.I))
            if dur_search:
                duration_txt = dur_search.strip()
                durations.append(duration_txt)
                
        # Cerca Livelli di Padronanza e Syllabus
        ctn_tables = panel.find_all(class_='ctn-table')
        acc_details = []
        
        for table in ctn_tables:
            # Livello
            level_txt = "N/A"
            level_el = table.find(class_='intro-table')
            if level_el:
                div_p = level_el.find(class_='paragraph')
                if div_p:
                    level_txt = div_p.get_text(strip=True)
                    mastery_levels.add(level_txt)
            
            # Obiettivi / Unità
            goals = []
            table_div = table.find(class_='table')
            h4_txt = "Dettagli"
            if table_div:
                h4_el = table_div.find('h4')
                if h4_el:
                    h4_txt = h4_el.get_text(strip=True)
                
                ul = table_div.find('ul', class_='rows') or table_div.find('ul')
                if ul:
                    lis = ul.find_all('li')
                    for li in lis:
                        goals.append(li.get_text(strip=True))
                        
            if level_txt != "N/A" or goals:
                acc_details.append({
                    "level": level_txt,
                    "title": h4_txt,
                    "items": goals
                })
                
        syllabus_sections.append({
            "header": header_txt,
            "duration": duration_txt,
            "syllabus": acc_details
        })
        
    # Formattazione sintetica
    durations_str = ", ".join(durations) if durations else "N/A"
    levels_str = ", ".join(sorted(list(mastery_levels))) if mastery_levels else "N/A"
    
    # Formattazione del programma completo in formato testuale strutturato
    syllabus_formatted = format_syllabus_text(syllabus_sections)
    
    return {
        "title": title,
        "desc_extended": desc_extended,
        "duration": durations_str,
        "levels": levels_str,
        "syllabus_raw": syllabus_sections,
        "syllabus_text": syllabus_formatted
    }


def format_syllabus_text(syllabus_sections):
    """Formatta i moduli del syllabus in una stringa di testo strutturata."""
    if not syllabus_sections:
        return "N/A"
        
    lines = []
    for section in syllabus_sections:
        lines.append(f"◆◆◆ MODULO: {section['header']} ◆◆◆")
        if section['duration']:
            lines.append(f"Durata modulo: {section['duration']}")
            
        for s in section['syllabus']:
            lines.append(f"Livello: {s['level']}")
            lines.append(f"{s['title']}:")
            for item in s['items']:
                lines.append(f"  • {item}")
        lines.append("-" * 35)
        
    return "\n".join(lines).strip()


def load_history(db_path):
    """Carica lo storico dei corsi già visitati."""
    if os.path.exists(db_path):
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[ERRORE] Impossibile caricare lo storico {db_path}: {e}")
            print("Verrà inizializzato un nuovo storico.")
    return {}


def save_history(db_path, data):
    """Salva lo storico dei corsi aggiornato."""
    try:
        db_dir = os.path.dirname(db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)
        with open(db_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[ERRORE] Impossibile salvare lo storico in {db_path}: {e}")


def main():
    parser = argparse.ArgumentParser(description="Scraper per i corsi del portale Syllabus (Syllabus Gov).")
    parser.add_argument("--url", default=DEFAULT_START_URL, help="URL di partenza per lo scraping.")
    parser.add_argument("--db", default=DEFAULT_DB_FILE, help="Nome del file JSON dello storico dei corsi.")
    parser.add_argument("--force", "--all", action="store_true", help="Forza lo scraping di tutti i dettagli dei corsi, anche se già visitati.")
    parser.add_argument("--clear", action="store_true", help="Azzera lo storico dei corsi prima di iniziare.")
    parser.add_argument("--delay", type=float, default=2.0, help="Ritardo medio di cortesia in secondi tra le richieste.")
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("      SYLLABUS GOV IT - SCRAPER OFFERTA FORMATIVA")
    print("=" * 60)
    print(f"Data/Ora: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"URL di partenza: {args.url}")
    print(f"Storico: {args.db}")
    
    # 1. Gestione comando Clear
    if args.clear:
        print("[STORICO] Azzittimento storico richiesto. Ripristino database vuoto...")
        if os.path.exists(args.db):
            try:
                os.remove(args.db)
                print("  Storico precedente eliminato con successo.")
            except Exception as e:
                print(f"  [ERRORE] Impossibile eliminare {args.db}: {e}")
                
    # 2. Caricamento storico
    history = load_history(args.db)
    print(f"[STORICO] Caricati {len(history)} corsi memorizzati in precedenza.")
    
    # 3. Scaricamento pagina principale
    print("\n[DOWNLOAD] Caricamento pagina offerta formativa principale...")
    r = fetch_with_retry(args.url)
    if not r:
        print("[ERRORE] Impossibile scaricare la pagina principale. Interruzione.")
        sys.exit(1)
        
    # 4. Parsing card dei corsi
    available_courses = parse_index_page(r.content)
    print(f"[PARSING] Trovati {len(available_courses)} corsi attivi nell'offerta formativa.")
    
    if not available_courses:
        print("[AVVISO] Nessuna card di corso identificata sulla pagina. Verifica i selettori HTML.")
        sys.exit(0)
        
    # 5. Individuazione dei nuovi corsi
    new_courses_to_scrape = []
    for c in available_courses:
        url = c['detail_url']
        is_new = url not in history
        
        if is_new or args.force:
            new_courses_to_scrape.append(c)
        else:
            # Per i corsi non nuovi, ci assicuriamo che rimangano nello storico
            pass
            
    total_new = len(new_courses_to_scrape)
    if args.force:
        print(f"\n[MODALITÀ FORZATA] Verranno estratti i dettagli di TUTTI i {total_new} corsi.")
    else:
        print(f"\n[ANALISI] Individuati {total_new} nuovi corsi da analizzare (non visti alla run precedente).")
        
    # 6. Scaricamento e parsing dettagli per ciascun corso da analizzare
    scraped_new_courses = []
    
    if total_new > 0:
        print("\n[DETTAGLI] Avvio dello scraping dei dettagli dei corsi...")
        for i, course in enumerate(new_courses_to_scrape):
            url = course['detail_url']
            title_idx = course['title']
            
            print(f"[{i + 1}/{total_new}] Scraping corso: '{title_idx}'")
            # print(f"  URL: {url}")
            
            # Cortesia / Rate Limit
            if i > 0:
                actual_delay = random.uniform(args.delay - 0.5, args.delay + 0.5)
                # print(f"  Attesa di cortesia di {actual_delay:.2f} secondi...")
                time.sleep(actual_delay)
                
            response = fetch_with_retry(url)
            if not response:
                print(f"  [ERRORE] Impossibile scaricare i dettagli per '{title_idx}'. Corso saltato.")
                continue
                
            # Analisi dettagli
            try:
                details = parse_details_page(response.content, url)
                
                # Se il titolo esteso fallisce, usiamo quello dell'indice
                final_title = details['title'] if details['title'] != "N/A" else title_idx
                
                # Aggreghiamo le informazioni
                full_course_data = {
                    "title": final_title,
                    "desc_short": course['desc_short'],
                    "desc_extended": details['desc_extended'],
                    "detail_url": url,
                    "duration": details['duration'],
                    "levels": details['levels'],
                    "syllabus_text": details['syllabus_text'],
                    "scraped_at": datetime.datetime.now().strftime('%Y-%m-%d')
                }
                
                scraped_new_courses.append(full_course_data)
                print(f"  [OK] Estratto con successo. Livelli: [{details['levels']}] - Durata: [{details['duration']}]")
                
            except Exception as e:
                print(f"  [ERRORE] Errore imprevisto durante il parsing di '{title_idx}': {e}")
                
    # 7. Aggiornamento storico e salvataggio dei risultati
    if scraped_new_courses:
        # Aggiorna il dizionario dello storico
        for course in scraped_new_courses:
            url = course['detail_url']
            history[url] = {
                "title": course['title'],
                "desc_short": course['desc_short'],
                "desc_extended": course['desc_extended'],
                "detail_url": course['detail_url'],
                "duration": course['duration'],
                "levels": course['levels'],
                "syllabus_text": course['syllabus_text'],
                "scraped_at": course['scraped_at']
            }
            
        save_history(args.db, history)
        print(f"\n[STORICO] Storico aggiornato. Ora contiene {len(history)} corsi in totale.")
        
        # Generazione file JSON
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        json_dir = os.path.join("data", "json")
        os.makedirs(json_dir, exist_ok=True)
        json_filepath = os.path.join(json_dir, f"nuovi_corsi_{timestamp}.json")
        
        print(f"\n[OUTPUT] Scrittura dei nuovi corsi nel file JSON: {json_filepath}...")
        try:
            formatted_courses = []
            for course in scraped_new_courses:
                formatted_courses.append({
                    "Titolo": course['title'],
                    "Descrizione Breve": course['desc_short'],
                    "Descrizione Estesa": course['desc_extended'],
                    "Link Dettagli": course['detail_url'],
                    "Durata": course['duration'],
                    "Livelli di Padronanza": course['levels'],
                    "Dettaglio Percorso (Syllabus)": course['syllabus_text']
                })
            
            with open(json_filepath, mode="w", encoding="utf-8") as f:
                json.dump(formatted_courses, f, ensure_ascii=False, indent=2)
            
            print(f"  [OK] File JSON creato correttamente. Percorso assoluto: {os.path.abspath(json_filepath)}")
            
            # Generazione automatica Excel
            excel_dir = os.path.join("data", "excel")
            os.makedirs(excel_dir, exist_ok=True)
            excel_filepath = os.path.join(excel_dir, f"nuovi_corsi_{timestamp}.xlsx")
            print(f"\n[OUTPUT] Conversione automatica in Excel: {excel_filepath}...")
            try:
                from json_to_excel import convert_json_to_excel
                convert_json_to_excel(json_filepath, excel_filepath)
            except Exception as e:
                print(f"  [AVVISO] Impossibile generare automaticamente l'Excel: {e}")
            
        except Exception as e:
            print(f"  [ERRORE] Impossibile scrivere il file JSON: {e}")
            
    else:
        print("\n[RISULTATO] Nessun nuovo corso rilevato rispetto allo storico precedente.")
        print("  Non è stato generato alcun file JSON.")
        
    print("\n" + "=" * 60)
    print("            PROCESSO DI SCRAPING COMPLETATO")
    print("=" * 60)


if __name__ == "__main__":
    main()
