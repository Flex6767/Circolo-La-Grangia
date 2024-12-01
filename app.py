from flask import Flask, render_template, request
from openpyxl import Workbook, load_workbook
import os

app = Flask(__name__)

# Percorso del file Excel
EXCEL_FILE = "adesioni.xlsx"

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

    # Debug: Stampa i dati ricevuti
    print("Dati ricevuti dal modulo:")
    print(f"Cognome: {cognome}, Nome: {nome}, Data di Nascita: {data_nascita}, Luogo di Nascita: {luogo_nascita}")
    print(f"Codice Fiscale: {codice_fiscale}, Email: {email}, Cellulare: {cellulare}")

    # Verifica che i dati non siano vuoti
    if not all([cognome, nome, data_nascita, luogo_nascita, codice_fiscale, email, cellulare]):
        print("Errore: Dati mancanti!")
        return "Errore: Tutti i campi sono obbligatori."

    try:
        # Aprire il file Excel e aggiungere i dati
        print("Apertura del file Excel...")
        workbook = load_workbook(EXCEL_FILE)
        sheet = workbook.active
        sheet.append([cognome, nome, data_nascita, luogo_nascita, codice_fiscale, email, cellulare])
        workbook.save(EXCEL_FILE)
        print("Dati salvati con successo!")

    except Exception as e:
        print(f"Errore durante il salvataggio: {e}")
        return f"Errore durante il salvataggio: {e}"

    return "Iscrizione completata con successo!"

if __name__ == '__main__':
    app.run(debug=True)
