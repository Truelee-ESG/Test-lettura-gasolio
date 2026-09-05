import os
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

# Librerie per la lettura dei PDF
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


def estrai_dati_testo(testo):
    """Estrae periodo, quantità e unità di misura usando le regular expression."""
    dati = {
        "periodo": "Non trovato",
        "quantita": "Non trovata",
        "unita_misura": "Non trovata"
    }
    
    # Pulizia del testo
    testo_pulito = " ".join(testo.split())

    # 1. Ricerca Unità di Misura e Quantità (es. 500 litri, 1.250 L, 450 mc)
    # Cerca numeri seguiti da unità comuni per il gasolio
    match_qta = re.search(r'(\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s*(litri|l\b|mc|metri cubi|kg)', testo_pulito, re.IGNORECASE)
    if match_qta:
        dati["quantita"] = match_qta.group(1)
        dati["unita_misura"] = match_qta.group(2).lower()

    # 2. Ricerca Periodo di riferimento (es. Gennaio 2026, dal 01/01/2026 al 31/01/2026)
    match_periodo = re.search(r'(?:periodo|competenza|dal)[:\s]*([0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2,4}(?:\s+al\s+[0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2,4})?|[A-Za-z]+\s+[0-9]{4})', testo_pulito, re.IGNORECASE)
    if match_periodo:
        dati["periodo"] = match_periodo.group(1)

    return dati

Dati_estratti = []

def elabora_fatture():
    root = tk.Tk()
    root.withdraw()
    
    # Selezione cartella di input
    cartella_input = filedialog.askdirectory(title="Seleziona la cartella contenente le fatture del gasolio")
    if not cartella_input:
        return

    file_pdf = list(Path(cartella_input).glob("*.pdf")) + list(Path(cartella_input).glob("*.PDF"))
    
    if not file_pdf:
        messagebox.showwarning("Attenzione", "Nessun file PDF trovato nella cartella selezionata.")
        return

    risultati = []
    
    for file_path in file_pdf:
        testo_totale = ""
        try:
            # Tentativo 1: Lettura testo nativo con pypdf
            if PdfReader:
                reader = PdfReader(str(file_path))
                for pagina in reader.pages:
                    testo_pagina = pagina.extract_text()
                    if testo_pagina:
                        testo_totale += testo_pagina + "\n"
            
            # Tentativo 2: Se il testo è vuoto e OCR è disponibile, usa Tesseract
            if not testo_totale.strip() and OCR_AVAILABLE:
                immagini = convert_from_path(str(file_path))
                for img in immagini:
                    testo_totale += pytesseract.image_to_string(img, language='ita') + "\n"
            
            # Estrazione informazioni
            info = estrai_dati_testo(testo_totale)
            risultati.append(f"File: {file_path.name}\n- Periodo: {info['periodo']}\n- Quantità: {info['quantita']} {info['unita_misura']}\n" + "-"*40)
            
        except Exception as e:
            risultati.append(f"File: {file_path.name} -> Errore di lettura: {str(e)}\n" + "-"*40)

    # Scrittura del file di testo sul Desktop
    desktop_path = Path.home() / "Desktop" / "riepilogo_fatture_gasolio.txt"
    with open(desktop_path, "w", encoding="utf-8") as f:
        f.write("\n".join(risultati))

    messagebox.Success("Completato", f"Elaborazione terminata!\nFile salvato sul desktop:\n{desktop_path.name}")

if __name__ == "__main__":
    elabora_fatture()
