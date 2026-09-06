import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pdfplumber
import requests


def extract_text_from_pdf(pdf_path):
  text = ""
  with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
      extracted = page.extract_text()
      if extracted:
        text += extracted + "\n"
  return text


def analyze_bill(pdf_path, result_box, status_label, btn_analyze):
  try:
    status_label.config(text="Estrazione testo dal PDF in corso...")
    raw_text = extract_text_from_pdf(pdf_path)

    if not raw_text.strip():
      messagebox.showerror(
          "Errore",
          "Impossibile estrarre testo dal PDF. Potrebbe essere un'immagine"
          " scansionata.",
      )
      status_label.config(text="Pronto")
      btn_analyze.config(state=tk.normal)
      return

    status_label.config(
        text="Invio a Ollama in corso (assicurati che sia avviato)..."
    )

    prompt = f"""
Sei un assistente energetico esperto. Analizza il testo della seguente bolletta e restituisci un JSON puro (senza blocchi di codice markdown o testo aggiuntivo) con i seguenti campi esatti:
- "fornitore": nome dell'azienda energetica
- "tipo": "Luce" o "Gas" o "Altro"
- "importo_totale": importo in euro (es. "45.50 €")
- "scadenza": data di scadenza (es. "DD/MM/AAAA")
- "consumo": consumo totale (es. "150 kWh" o "80 Smc")

Testo della bolletta:
{raw_text[:4000]}
"""

    response = requests.post(
        "http://localhost:11434/api/generate",
        json={"model": "llama3", "prompt": prompt, "stream": False},
        timeout=60,
    )

    if response.status_code == 200:
      res_json = response.json()
      ai_response = res_json.get("response", "")

      result_box.delete(1.0, tk.END)
      result_box.insert(tk.END, ai_response)
      status_label.config(text="Analisi completata con successo!")
    else:
      messagebox.showerror(
          "Errore Ollama",
          "Impossibile comunicare con Ollama. È attivo in background?",
      )
      status_label.config(text="Errore di connessione")

  except Exception as e:
    messagebox.showerror("Errore", str(e))
    status_label.config(text="Errore imprevisto")

  finally:
    btn_analyze.config(state=tk.normal)


def select_file(result_box, status_label, btn_analyze):
  file_path = filedialog.askopenfilename(
      title="Seleziona la bolletta", filetypes=[("File PDF", "*.pdf")]
  )
  if file_path:
    btn_analyze.config(state=tk.disabled)
    threading.Thread(
        target=analyze_bill,
        args=(file_path, result_box, status_label, btn_analyze),
        daemon=True,
    ).start()


def main():
  root = tk.Tk()
  root.title("Lettore Bollette IA (Locale)")
  root.geometry("600x500")

  label = tk.Label(
      root, text="Analizzatore Bollette con IA Locale", font=("Arial", 14, "bold")
  )
  label.pack(pady=10)

  status_label = tk.Label(root, text="Pronto", font=("Arial", 10), fg="gray")
  status_label.pack(pady=5)

  btn_analyze = tk.Button(
      root,
      text="Seleziona Bolletta PDF",
      font=("Arial", 12),
      bg="#4CAF50",
      fg="white",
      padx=10,
      pady=5,
      command=lambda: select_file(result_box, status_label, btn_analyze),
  )
  btn_analyze.pack(pady=10)

  result_box = scrolledtext.Text(root, wrap=tk.WORD, width=70, height=20)
  result_box.pack(pady=10, padx=10)

  root.mainloop()


if __name__ == "__main__":
  main()
