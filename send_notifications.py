#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Gestore Notifiche Email con Onboarding Nuovi Destinatari - Syllabus
Identifica i nuovi iscritti tramite hash SHA-256 e differenzia le email:
- Nuovi iscritti: ricevono il catalogo storico completo (+ nuovi corsi se presenti)
- Iscritti già registrati: ricevono solo i nuovi corsi (se presenti)
- Nessuna novità e nessun nuovo iscritto: nessuna email inviata
"""

import os
import sys
import json
import glob
import hashlib
import smtplib
import argparse
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from json_to_excel import generate_full_catalog_excel

# Forza UTF-8 per console Windows
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def normalize_email(email_str):
    """Normalizza l'indirizzo email (rimuove spazi e converte in minuscolo)."""
    return email_str.strip().lower()


def hash_email(email_str):
    """Calcola l'hash SHA-256 di un indirizzo email normalizzato."""
    normalized = normalize_email(email_str)
    return hashlib.sha256(normalized.encode('utf-8')).hexdigest()


def parse_recipients(recipients_raw):
    """Estrae una lista di indirizzi email univoci da stringhe separate da virgola, punto e virgola o a capo."""
    if not recipients_raw:
        return []
    
    # Sostituisce punto e virgola e ritorni a capo con virgole
    cleaned = recipients_raw.replace(';', ',').replace('\n', ',').replace('\r', '')
    emails = []
    seen = set()
    for token in cleaned.split(','):
        email = normalize_email(token)
        if email and '@' in email and '.' in email:
            if email not in seen:
                seen.add(email)
                emails.append(email)
    return emails


def load_recipients_history(history_file):
    """Carica il registro degli hash delle email già notificate in passato."""
    if os.path.exists(history_file):
        try:
            with open(history_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and "known_hashes" in data:
                    return set(data["known_hashes"])
                elif isinstance(data, list):
                    return set(data)
        except Exception as e:
            print(f"[AVVISO] Impossibile caricare il registro destinatari '{history_file}': {e}")
    return set()


def save_recipients_history(history_file, known_hashes):
    """Salva il registro aggiornato degli hash delle email notificate."""
    try:
        os.makedirs(os.path.dirname(history_file), exist_ok=True)
        data = {
            "known_hashes": sorted(list(known_hashes)),
            "last_updated": datetime.datetime.now().isoformat()
        }
        with open(history_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"[STORICO] Registro destinatari aggiornato con successo: {history_file}")
        return True
    except Exception as e:
        print(f"[ERRORE] Impossibile salvare il registro destinatari in '{history_file}': {e}")
        return False


def get_latest_new_files(excel_dir, json_dir):
    """Individua gli ultimi file dei nuovi corsi generati dallo scraper."""
    excel_files = sorted(glob.glob(os.path.join(excel_dir, "nuovi_corsi_*.xlsx")), reverse=True)
    json_files = sorted(glob.glob(os.path.join(json_dir, "nuovi_corsi_*.json")), reverse=True)
    
    latest_files = []
    if excel_files:
        latest_files.append(excel_files[0])
    if json_files:
        latest_files.append(json_files[0])
    return latest_files


def send_email_smtp(smtp_server, smtp_port, username, password, sender_name, to_email, subject, body_text, attachment_paths, dry_run=False):
    """Invia una singola email tramite server SMTP con allegati."""
    if dry_run:
        print(f"  [DRY-RUN] Invio simulato a: {to_email}")
        print(f"            Oggetto: {subject}")
        print(f"            Allegati ({len(attachment_paths)}): {[os.path.basename(p) for p in attachment_paths]}")
        return True

    msg = MIMEMultipart()
    sender_header = f"{sender_name} <{username}>" if sender_name else username
    msg['From'] = sender_header
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body_text, 'plain', 'utf-8'))

    # Aggiungi allegati
    for file_path in attachment_paths:
        if not os.path.exists(file_path):
            print(f"  [AVVISO] File allegato non trovato, saltato: {file_path}")
            continue
        try:
            with open(file_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(file_path)}"'
            )
            msg.attach(part)
        except Exception as e:
            print(f"  [ERRORE] Impossibile allegare il file {file_path}: {e}")

    # Connessione al server SMTP
    try:
        if int(smtp_port) == 465:
            with smtplib.SMTP_SSL(smtp_server, int(smtp_port), timeout=60) as server:
                server.login(username, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(smtp_server, int(smtp_port), timeout=60) as server:
                server.starttls()
                server.login(username, password)
                server.send_message(msg)
        print(f"  [OK] Email inviata con successo a: {to_email}")
        return True
    except Exception as e:
        print(f"  [ERRORE] Invio fallito a {to_email}: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Invio notifiche email differenziate per nuovi corsi e nuovi iscritti.")
    parser.add_argument("--dry-run", action="store_true", help="Simula l'invio senza effettuare connessioni SMTP.")
    parser.add_argument("--recipients-file", default=os.path.join("data", "recipients_history.json"), help="Percorso del registro hash dei destinatari.")
    parser.add_argument("--history-courses", default=os.path.join("data", "corsi_storico.json"), help="Percorso del file corsi_storico.json.")
    parser.add_argument("--excel-dir", default=os.path.join("data", "excel"), help="Cartella dei file Excel.")
    parser.add_argument("--json-dir", default=os.path.join("data", "json"), help="Cartella dei file JSON.")
    args = parser.parse_args()

    print("=" * 60)
    print("      SYLLABUS - GESTIONE NOTIFICHE E ONBOARDING ISCRITTI")
    print("=" * 60)

    # 1. Recupera configurazione SMTP e destinatari da variabili d'ambiente
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = os.environ.get("SMTP_PORT", "465")
    smtp_username = os.environ.get("SMTP_USERNAME", "")
    smtp_password = os.environ.get("SMTP_PASSWORD", "")
    sender_name = os.environ.get("SENDER_NAME", "formazione_40ore_Syllabus_IZSVe")
    raw_recipients = os.environ.get("EMAIL_RECIPIENTS", "")
    custom_body = os.environ.get("EMAIL_BODY", "")

    if not args.dry_run and (not smtp_username or not smtp_password):
        print("[ERRORE] Credenziali SMTP non configurate nelle variabili d'ambiente (SMTP_USERNAME, SMTP_PASSWORD).")
        print("         Usa --dry-run per eseguire una prova locale simulata.")
        sys.exit(1)

    recipients_list = parse_recipients(raw_recipients)
    if not recipients_list:
        print("[AVVISO] Nessun indirizzo email valido specificato in EMAIL_RECIPIENTS.")
        sys.exit(0)

    print(f"[INFO] Trovati {len(recipients_list)} indirizzi email configurati.")

    # 2. Carica lo storico degli hash
    known_hashes = load_recipients_history(args.recipients_file)
    print(f"[INFO] Registro destinatari esistenti: {len(known_hashes)} indirizzi già censiti.")

    # 3. Classifica i destinatari in Nuovi vs Vecchi
    new_recipients = []
    old_recipients = []

    for email in recipients_list:
        h = hash_email(email)
        if h in known_hashes:
            old_recipients.append(email)
        else:
            new_recipients.append(email)

    print(f"[ANALISI] Nuovi destinatari da registrare (onboarding): {len(new_recipients)}")
    print(f"[ANALISI] Destinatari già noti (report periodico): {len(old_recipients)}")

    # 4. Verifica disponibilità di nuovi corsi
    latest_new_files = get_latest_new_files(args.excel_dir, args.json_dir)
    has_new_courses = len(latest_new_files) > 0
    if has_new_courses:
        print(f"[CORSI] Rilevati nuovi corsi generati: {[os.path.basename(f) for f in latest_new_files]}")
    else:
        print("[CORSI] Nessun file di nuovi corsi rilevato in questa esecuzione.")

    # 5. Generazione o aggiornamento del Catalogo Completo in Excel (se esiste lo storico corsi)
    full_catalog_excel_path = os.path.join(args.excel_dir, "catalogo_completo_syllabus.xlsx")
    if os.path.exists(args.history_courses):
        print("\n[CATALOGO] Generazione/Aggiornamento file catalogo completo...")
        generated_catalog = generate_full_catalog_excel(args.history_courses, full_catalog_excel_path)
        if not generated_catalog:
            print("[AVVISO] Impossibile generare l'Excel del catalogo completo.")
    else:
        print(f"[AVVISO] File storico '{args.history_courses}' non trovato.")

    # 6. Esecuzione logica invio differenziato
    # Scenario A: Nuovi corsi presenti
    #   - Nuovi destinatari: ricevono Catalogo Completo + Nuovi Corsi
    #   - Vecchi destinatari: ricevono solo Nuovi Corsi
    # Scenario B: Nessun nuovo corso, ma ci sono Nuovi destinatari
    #   - Nuovi destinatari: ricevono Catalogo Completo
    #   - Vecchi destinatari: nessuna email
    # Scenario C: Nessun nuovo corso e nessun nuovo destinatario
    #   - Nessuna email a nessuno

    emails_to_send = []

    # Preparazione email per NUOVI destinatari
    if new_recipients:
        new_user_attachments = []
        if os.path.exists(full_catalog_excel_path):
            new_user_attachments.append(full_catalog_excel_path)
        if has_new_courses:
            new_user_attachments.extend(latest_new_files)

        if has_new_courses:
            subj = "Benvenuto - Catalogo Completo e Nuovi Corsi Syllabus"
            body = (
                "Gentile collega,\n\n"
                "Benvenuto/a al servizio di aggiornamento automatico per i corsi della piattaforma Syllabus (www.syllabus.gov.it).\n\n"
                "In allegato a questa email troverai:\n"
                "1. Il Catalogo Completo di tutti i percorsi formativi disponibili finora ('catalogo_completo_syllabus.xlsx');\n"
                "2. Il report con i Nuovi Corsi pubblicati nell'ultimo periodo.\n\n"
                "A partire dalla prossima schedulazione periodica, riceverai unicamente gli aggiornamenti con le nuove uscite.\n\n"
                "Cordiali saluti,\n"
                "Staff Formazione Continua"
            )
        else:
            subj = "Benvenuto - Catalogo Completo Corsi Syllabus"
            body = (
                "Gentile collega,\n\n"
                "Benvenuto/a al servizio di aggiornamento automatico per i corsi della piattaforma Syllabus (www.syllabus.gov.it).\n\n"
                "In allegato a questa email troverai il Catalogo Completo con tutti i corsi attualmente disponibili ('catalogo_completo_syllabus.xlsx').\n\n"
                "Nel periodo corrente non sono stati pubblicati ulteriori corsi; riceverai una nuova notifica alla prossima occasione in cui verranno rilevate novità formative.\n\n"
                "Cordiali saluti,\n"
                "Staff Formazione Continua"
            )

        for email in new_recipients:
            emails_to_send.append({
                "type": "new_recipient",
                "email": email,
                "subject": subj,
                "body": body,
                "attachments": new_user_attachments
            })

    # Preparazione email per VECCHI destinatari (solo se ci sono nuovi corsi)
    if has_new_courses and old_recipients:
        subj = "Report Trimestrale Automatico - Nuovi Corsi Syllabus"
        default_old_body = (
            "Gentile collega,\n\n"
            "Ti informiamo che durante il controllo periodico della piattaforma Syllabus sono stati rilevati nuovi corsi.\n\n"
            "In allegato trovi il report dettagliato con titoli, durata, livelli di padronanza e programma dei corsi appena pubblicati.\n\n"
            "Cordiali saluti,\n"
            "Staff Formazione Continua"
        )
        body = custom_body if custom_body else default_old_body

        for email in old_recipients:
            emails_to_send.append({
                "type": "old_recipient",
                "email": email,
                "subject": subj,
                "body": body,
                "attachments": latest_new_files
            })

    # 7. Esecuzione effettiva dell'invio
    if not emails_to_send:
        print("\n[ESITO] Nessun nuovo corso e nessun nuovo destinatario rilevato. Nessuna email da inviare.")
    else:
        print(f"\n[INVIO] Avvio invio di {len(emails_to_send)} email in corso...")
        successfully_notified_new_hashes = set()

        for item in emails_to_send:
            success = send_email_smtp(
                smtp_server=smtp_server,
                smtp_port=smtp_port,
                username=smtp_username,
                password=smtp_password,
                sender_name=sender_name,
                to_email=item["email"],
                subject=item["subject"],
                body_text=item["body"],
                attachment_paths=item["attachments"],
                dry_run=args.dry_run
            )

            if success and item["type"] == "new_recipient":
                successfully_notified_new_hashes.add(hash_email(item["email"]))

        # 8. Aggiorna il registro destinatari se ci sono stati nuovi iscritti notificati
        recipients_updated = False
        if successfully_notified_new_hashes:
            known_hashes.update(successfully_notified_new_hashes)
            save_recipients_history(args.recipients_file, known_hashes)
            recipients_updated = True

        # Imposta output per GitHub Actions se eseguito in CI
        github_output = os.environ.get("GITHUB_OUTPUT")
        if github_output:
            with open(github_output, "a", encoding="utf-8") as f:
                f.write(f"recipients_updated={'true' if recipients_updated else 'false'}\n")
                f.write(f"emails_sent={'true' if emails_to_send else 'false'}\n")

    print("\n" + "=" * 60)
    print("      PROCESSO GESTIONE NOTIFICHE COMPLETATO")
    print("=" * 60)


if __name__ == "__main__":
    main()

