import imaplib
import email
import os
import re
from email.header import decode_header
import asana
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from dotenv import load_dotenv

load_dotenv()

# ---------- CONFIGURATION ----------
EMAIL_ACCOUNT = os.getenv("EMAIL_ACCOUNT")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")
IMAP_FOLDER = f'"{os.getenv("IMAP_FOLDER")}"'

CARINGTON_DRIVE = os.path.expanduser(os.getenv("CARINGTON_DRIVE"))

ASANA_TOKEN = os.getenv("ASANA_TOKEN")
PROJECT_ID_DEMANDES = os.getenv("PROJECT_ID_DEMANDES")
SECTION_ID_DEMANDES = os.getenv("SECTION_ID_DEMANDES")
SECTION_ID_QUITUS = os.getenv("SECTION_ID_QUITUS")

FIELD_DOCS_MANQUANTS = os.getenv("FIELD_DOCS_MANQUANTS")
FIELD_DOCS_MANQUANTS_QUITUS = os.getenv("FIELD_DOCS_MANQUANTS_QUITUS")

OPTION_DEMANDE_IMMAT = os.getenv("OPTION_DEMANDE_IMMAT")
OPTION_MANDAT = os.getenv("OPTION_MANDAT")
OPTION_MANDAT_QUITUS = os.getenv("OPTION_MANDAT_QUITUS")

# ---------- FONCTIONS ----------
def decode_mime_words(s):
    decoded = decode_header(s)
    return ''.join([
        frag.decode(charset or 'utf-8') if isinstance(frag, bytes) else frag
        for frag, charset in decoded
    ])

def extract_dossier_num(subject):
    match = re.search(r'n[\u00ba\u00b0]?\s*(\d{5})', subject, re.IGNORECASE)
    if match:
        return match.group(1)
    match = re.search(r'\b\d{5}\b', subject)
    if match:
        return match.group(0)
    return None

def find_dossier_path(base_path, dossier_num):
    for root, dirs, files in os.walk(base_path):
        for dir_name in dirs:
            if dossier_num in dir_name:
                return os.path.join(root, dir_name)
    return None

def find_task_by_dossier_number(client, dossier_num, project_id):
    tasks = client.tasks.get_tasks_for_project(project_id, opt_fields="name,gid")
    for task in tasks:
        if dossier_num in task['name']:
            log(f"✅ Tâche trouvée : {task['name']} (GID : {task['gid']})")
            return task['gid']
    return None

def task_is_in_project(task_details, project_gid):
    return any(project['gid'] == project_gid for project in task_details.get("projects", []))

def log(message):
    console.insert(tk.END, message + "\n")
    console.see(tk.END)

def process_mails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    mail.select(IMAP_FOLDER)
    status, messages = mail.search(None, 'ALL')
    email_ids = messages[0].split()

    client = asana.Client.access_token(ASANA_TOKEN)

    for e_id in email_ids:
        status, msg_data = mail.fetch(e_id, '(RFC822)')
        raw_email = msg_data[0][1]
        message = email.message_from_bytes(raw_email)

        raw_subject = message.get("Subject", "")
        subject = decode_mime_words(raw_subject)
        log(f"\n📨 Objet : {subject}")

        dossier_num = extract_dossier_num(subject)
        if not dossier_num:
            log("❌ Aucun numéro de dossier trouvé.")
            continue

        log(f"📁 Numéro de dossier : {dossier_num}")

        dossier_path = find_dossier_path(CARINGTON_DRIVE, dossier_num)
        if not dossier_path:
            log("❌ Aucun dossier trouvé dans le Drive.")
            continue

        has_attachment = False
        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            if filename:
                has_attachment = True
                filename = decode_mime_words(filename)
                filepath = os.path.join(dossier_path, filename)
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                log(f"📎 Pièce jointe enregistrée dans : {filepath}")

        if not has_attachment:
            log("⚠️ Aucun fichier joint détecté, mail supprimé.")
            mail.store(e_id, '+FLAGS', '\\Deleted')
            continue

        task_gid = find_task_by_dossier_number(client, dossier_num, PROJECT_ID_DEMANDES)
        if not task_gid:
            log("❌ Aucune tâche trouvée pour ce dossier.")
            continue

        task_details = client.tasks.find_by_id(task_gid, opt_fields="custom_fields,projects")
        current_docs_manquants = []
        current_docs_quitus = []

        for field in task_details["custom_fields"]:
            if field["gid"] == FIELD_DOCS_MANQUANTS:
                current_docs_manquants = field.get("multi_enum_values", [])
            if field["gid"] == FIELD_DOCS_MANQUANTS_QUITUS:
                current_docs_quitus = field.get("multi_enum_values", [])

        updated_docs_manquants = []
        for opt in current_docs_manquants:
            if opt["gid"] in {OPTION_DEMANDE_IMMAT, OPTION_MANDAT}:
                log(f"🧹 Retrait : ❌ {opt['name']} (Docs manquants)")
            else:
                updated_docs_manquants.append(opt["gid"])

        custom_fields_update = {
            FIELD_DOCS_MANQUANTS: updated_docs_manquants
        }

        success = True
        if task_is_in_project(task_details, SECTION_ID_QUITUS):
            log("📌 Projet QUITUS détecté ✅")
            updated_docs_quitus = []
            for opt in current_docs_quitus:
                if opt["gid"] == OPTION_MANDAT_QUITUS:
                    log(f"🧹 Retrait : ❌ {opt['name']} (Docs manquants QUITUS)")
                else:
                    updated_docs_quitus.append(opt["gid"])
            custom_fields_update[FIELD_DOCS_MANQUANTS_QUITUS] = updated_docs_quitus

        client.tasks.update_task(task_gid, {
            "custom_fields": custom_fields_update
        })
        log("✅ Champs personnalisés mis à jour.")

        client.sections.add_task_for_section(SECTION_ID_DEMANDES, {"task": task_gid})
        log("📦 Tâche déplacée dans ENREGISTREMENT (DEMANDES)")

        if task_is_in_project(task_details, SECTION_ID_QUITUS):
            client.sections.add_task_for_section(SECTION_ID_QUITUS, {"task": task_gid})
            log("📦 Tâche déplacée dans CERFA QUITUS A FAIRE (QUITUS)")

        if success:
            mail.store(e_id, '+FLAGS', '\\Deleted')
            log("🗑️ Mail supprimé.")

    mail.expunge()
    mail.logout()
    log("\n🎉 Script terminé avec succès.")

# ---------- INTERFACE GRAPHIQUE ----------
def launch_gui():
    global console
    root = tk.Tk()
    root.title("QuadraFlow")
    root.geometry("600x600")
    root.configure(bg="black")

    try:
        logo = Image.open("QuadraFlow_logo.png")
        logo = logo.resize((120, 120), Image.LANCZOS)
        logo_img = ImageTk.PhotoImage(logo)
        logo_label = tk.Label(root, image=logo_img, bg="black")
        logo_label.image = logo_img
        logo_label.pack(pady=(10, 0))
    except:
        pass

    label = tk.Label(root, text="Bienvenue dans QuadraFlow", font=("Arial", 16, "bold"), bg="black", fg="white")
    label.pack(pady=10)

    btn = ttk.Button(root, text="Lancer le traitement", command=process_mails)
    btn.pack(pady=10)

    console_frame = tk.Frame(root)
    console_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    console = tk.Text(console_frame, bg="black", fg="white", insertbackground="white", font=("Courier New", 10))
    console.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(console_frame, command=console.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    console.config(yscrollcommand=scrollbar.set)

    root.mainloop()

if __name__ == '__main__':
    launch_gui()
