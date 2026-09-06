from datetime import datetime
import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
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


def save_to_excel(data_dict):
  try:
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    excel_path = os.path.join(desktop_path, "bollette_analisi.xlsx")

    df_new = pd.DataFrame([{
        "Data Analisi": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Fornitore": data_dict.get("fornitore", "N/D"),
        "Tipo": data_dict.get("tipo", "N/D"),
        "Importo Totale": data_dict.get("importo_totale", "N/D"),
        "Scadenza": data_dict.get("scadenza", "N/D"),
        "Quantità / Consumo": data_dict.get("consumo", "N/D"),
    }])

    if os.path.exists(excel_path):
      df_existing = pd.read_excel(excel_path)
      df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else:
      df_combined = df_new

    df_combined.to_excel(excel_path, index=False)
    return excel_path
  except Exception as e:
    print("Errore salvataggio Excel:", e)
    return None


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
      btn_analyze.config(state=tk.NORMAL)
      return

    status_label.config(text="Invio a Ollama in corso (attendere)...")

    prompt = f"""
Sei un assistente energetico esperto. Analizza il testo della seguente bolletta e restituisci UNICAMENTE un oggetto JSON valido (senza blocchi di codice markdown ```json, senza commenti o testo aggiuntivo prima o dopo) con esattamente questi campi:
- "fornitore": nome dell'azienda energetica (stringa)
- "tipo": "Luce" o "Gas" o "Altro" (stringa)
- "importo_totale": importo in euro (es. "45.50 €") (stringa)
- "scadenza": data di scadenza (es. "DD/MM/AAAA") (stringa)
- "consumo": consumo totale comprensivo di unità di misura (es. "150 kWh" o "80 Smc") (stringa)

Testo della bolletta:
{raw_text[:4000]}
"""

    response = requests.post(
        "http://localhost:11434/api/generate",
        json={"model": "llama3", "prompt": prompt, "stream": False},
        timeout=120,
    )

    if response.status_code == 200:
      res_json = response.json()
      ai_response = res_json.get("response", "").strip()

      if ai_response.startswith("```"):
        ai_response = ai_response.split("```")[1]
        if ai_response.startswith("json"):
          ai_response = ai_response[4:]
        ai_response = ai_response.strip()

      try:
        data_dict = json.loads(ai_response)
      except:
        data_dict = {"raw_output": ai_response}

      result_box.delete(1.0, tk.END)
      result_box.insert(
          tk.END, json.dumps(data_dict, indent=4, ensure_ascii=False)
      )

      # Salvataggio automatico Excel sul Desktop
      excel_file = save_to_excel(
          data_dict if isinstance(data_dict, dict) else {}
      )
      if excel_file:
        status_label.config(
            text=f"Completato! Salvato in: {os.path.basename(excel_file)}"
        )
        messagebox.showinfo(
            "Successo",
            "Analisi completata!\nFile Excel aggiornato sul Desktop:\n"
            f"{excel_file}",
        )
      else:
        status_label.config(text="Analisi completata (Errore salvataggio Excel)")
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
    btn_analyze.config(state=tk.NORMAL)


def select_file(result_box, status_label, btn_analyze):
  file_path = filedialog.askopenfilename(
      title="Seleziona la bolletta", filetypes=[("File PDF", "*.pdf")]
  )
  if file_path:
    btn_analyze.config(state=tk.DISABLED)
    status_label.config(text="Elaborazione in corso...")
    threading.Thread(
        target=analyze_bill,
        args=(file_path, result_box, status_label, btn_analyze),
        daemon=True,
    ).start()


def main():
  root = tk.Tk()
  root.title("Lettore Bollette IA (Locale)")
  root.geometry("650x550")

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

  result_box = scrolledtext.Text(root, wrap=tk.WORD, width=75, height=20)
  result_box.pack(pady=10, padx=10)

  root.mainloop()


if __name__ == "__main__":
  main()
