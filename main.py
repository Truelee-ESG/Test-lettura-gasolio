import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import glob
import re

class DieselInvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Estrai Dati Fatture Gasolio - MD Sustainify")
        self.root.geometry("550px 380px")
        self.root.minsize(500, 350)
        self.root.configure(f"#f4f6f8")

        # Variables
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # Title Frame
        title_frame = tk.Frame(self.root, bg="#1e293b", padx=15, pady=15)
        title_frame.pack(fill=tk.X)

        title_label = tk.Label(
            title_frame, 
            text="Analizzatore Fatture Gasolio", 
            font=("Arial", 14, "bold"), 
            bg="#1e293b", 
            fg="white"
        )
        title_label.pack(anchor="w")

        subtitle_label = tk.Label(
            title_frame, 
            text="Estrai quantità e unità di misura in un file Excel (.xlsx)", 
            font=("Arial", 9), 
            bg="#1e293b", 
            fg="#94a3b8"
        )
        subtitle_label.pack(anchor="w", pady=(2, 0))

        # Main Content Frame
        content_frame = tk.Frame(self.root, bg="#f4f6f8", padx=20, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Input Directory Selection
        tk.Label(content_frame, text="Cartella Fatture (Input):", font=("Arial", 10, "bold"), bg="#f4f6f8", fg="#334155").pack(anchor="w", pady=(0, 5))
        
        in_frame = tk.Frame(content_frame, bg="#f4f6f8")
        in_frame.pack(fill=tk.X, pady=(0, 15))

        self.in_entry = tk.Entry(in_frame, textvariable=self.input_dir, font=("Arial", 10), bg="white", relief="solid", bd=1)
        self.in_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4, padx=(0, 8))

        in_btn = tk.Button(in_frame, text="Sfoglia...", font=("Arial", 9, "bold"), bg="#3b82f6", fg="white", relief="flat", padx=12, pady=4, command=self.browse_input)
        in_btn.pack(side=tk.RIGHT)

        # Output Directory Selection
        tk.Label(content_frame, text="Cartella di Salvataggio Report (Output):", font=("Arial", 10, "bold"), bg="#f4f6f8", fg="#334155").pack(anchor="w", pady=(0, 5))
        
        out_frame = tk.Frame(content_frame, bg="#f4f6f8")
        out_frame.pack(fill=tk.X, pady=(0, 20))

        self.out_entry = tk.Entry(out_frame, textvariable=self.output_dir, font=("Arial", 10), bg="white", relief="solid", bd=1)
        self.out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4, padx=(0, 8))

        out_btn = tk.Button(out_frame, text="Sfoglia...", font=("Arial", 9, "bold"), bg="#3b82f6", fg="white", relief="flat", padx=12, pady=4, command=self.browse_output)
        out_btn.pack(side=tk.RIGHT)

        # Action Button (Crea Report)
        action_btn = tk.Button(
            content_frame, 
            text="Crea Report Excel", 
            font=("Arial", 11, "bold"), 
            bg="#10b981", 
            fg="white", 
            relief="flat", 
            pady=8,
            command=self.generate_report
        )
        action_btn.pack(fill=tk.X)

    def browse_input(self):
        dir_path = filedialog.askdirectory(title="Seleziona la cartella contenente le fatture")
        if dir_path:
            self.input_dir.set(dir_path)

    def browse_output(self.root_self=None):
        dir_path = filedialog.askdirectory(title="Seleziona la cartella di salvataggio del report")
        if dir_path:
            app.output_dir.set(dir_path)

    def generate_report(self):
        in_path = self.input_dir.get()
        out_path = self.output_dir.get()

        if not in_path or not os.path.exists(in_path):
            messagebox.showerror("Errore", "Seleziona una cartella di input valida.")
            return

        if not out_path or not os.path.exists(out_path):
            messagebox.showerror("Errore", "Seleziona una cartella di output valida.")
            return

        try:
            # Search for files (PDF, XML, TXT)
            files = []
            for ext in ("*.pdf", "*.xml", "*.txt"):
                files.extend(glob.glob(os.path.join(in_path, ext)))
                files.extend(glob.glob(os.path.join(in_path, ext.upper())))

            if not files:
                messagebox.showwarning("Attenzione", "Nessun file supportato (PDF, XML, TXT) trovato nella cartella selezionata.")
                return

            extracted_data = []

            for file_path in files:
                filename = os.path.basename(file_path)
                # Simple heuristic extraction logic (can be extended based on specific formats like FatturaPA XML or PDF text)
                qty, unit = self.parse_invoice(file_path)
                
                extracted_data.append({
                    "Nome File": filename,
                    "Quantità": qty,
                    "Unità di Misura": unit,
                    "Percorso File": file_path
                })

            # Create DataFrame and save to Excel
            df = pd.DataFrame(extracted_data)
            output_file = os.path.join(out_path, "Report_Fatture_Gasolio.xlsx")
            
            df.to_excel(output_file, index=False, sheet_name="Gasolio")
            
            messagebox. successo = messagebox.showinfo("Successo", f"Report generato con successo!\nSalvato in:\n{output_file}")

        except Exception as e:
            messagebox.showerror("Errore di Elaborazione", f"Si è verificato un errore durante la lettura dei file:\n{str(e)}")

    def parse_invoice(self, file_path):
        """
        Funzione di parsing di esempio. Estrae dati mock o cerca pattern di litri/kg in file XML o PDF.
        """
        qty = 0.0
        unit = "Litri"

        try:
            if file_path.lower().endswith(".xml"):
                # Esempio parsing XML FatturaPA
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                    # Cerca tag tipici quantità nelle fatture elettroniche
                    matches = re.findall(r'<Quantita>([\d\.]+)</Quantita>', content)
                    if matches:
                        qty = float(matches[0])
            else:
                # Per PDF o altri file, mock o ricerca testuale di base se si usa librerie dedicate
                # Qui restituisce un valore di esempio o analizza il nome file per test
                qty = 1000.0 
        except Exception:
            qty = 0.0

        return qty, unit

if __name__ == "__main__":
    root = tk.Tk()
    app = DieselInvoiceApp(root)
    root.mainloop()
