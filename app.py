
from flask import Flask, render_template, request, send_from_directory
from openpyxl import Workbook, load_workbook
import os
import dropbox

app = Flask(__name__)

# Percorso del file Excel
EXCEL_FILE = "7tmp7adesioni.xlsx"

# Access Token di Dropbox (copia l'access token ottenuto dal tuo account Dropbox)
DROPBOX_ACCESS_TOKEN = 'your_dropbox_access_token'

# Creazione iniziale del file Excel se non esiste
if not os.path.exists(EXCEL_FILE):
    print("Creazione del file Excel...")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Adesioni"
    # Intestazioni del file Excel
    sheet.append(["Cognome", "Nome", "Data di Nascita", "Luogo di Nascita", "Codice Fiscale", "Email", "Cellulare"])
    workbook.save(EXCEL_FILE)
    print("File Excel creato!")
else:
    print("File Excel già esistente.")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Raccogliere i dati inviati dal modulo
    cognome = request.form.get('surname')
    nome = request.form.get('name')
    data_nascita = request.form.get('birthdate')
    luogo_nascita = request.form.get('birthplace')
    codice_fiscale = request.form.get('fiscalcode')
    email = request.form.get('email')
    cellulare = request.form.get('phone')

    print("Dati ricevuti dal modulo:", cognome, nome, data_nascita, luogo_nascita, codice_fiscale, email, cellulare)

    try:
        # Aprire o creare il file Excel
        if not os.path.exists(EXCEL_FILE):
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Adesioni"
            sheet.append(["Cognome", "Nome", "Data di Nascita", "Luogo di Nascita", "Codice Fiscale", "Email", "Cellulare"])
        else:
            workbook = load_workbook(EXCEL_FILE)
            sheet = workbook.active

        # Aggiungere i dati
        sheet.append([cognome, nome, data_nascita, luogo_nascita, codice_fiscale, email, cellulare])
        workbook.save(EXCEL_FILE)
        print("Dati salvati con successo!")

        # Caricare il file su Dropbox
        upload_to_dropbox(EXCEL_FILE)
    except Exception as e:
        print(f"Errore durante il salvataggio: {e}")
        return f"Errore durante il salvataggio: {e}"

    return "Iscrizione completata con successo!"

# Funzione per caricare il file su Dropbox
def upload_to_dropbox(file_path):
    # Creazione di una connessione con Dropbox
    dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)

    try:
        # Aprire il file e caricarlo su Dropbox
        with open(file_path, "rb") as file:
            # Carica il file su Dropbox (ad esempio, nella cartella 'uploads')
            dropbox_path = f"/{file_path}"  # Il file sarà caricato con lo stesso nome
            dbx.files_upload(file.read(), dropbox_path, mute=True)
            print(f"File {file_path} caricato su Dropbox!")
    except Exception as e:
        print(f"Errore durante il caricamento su Dropbox: {e}")

# Endpoint per il download del file Excel
@app.route('/download', methods=['GET'])
def download_file():
    # Invia il file Excel dal server al client
    return send_from_directory(directory=os.getcwd(), filename=EXCEL_FILE)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)), debug=True)
