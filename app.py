import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

def estrai_testo_sicuro(pdf_path):
    """
    Estrae il testo da un PDF in modo sicuro:
    1. Prova prima l'estrazione nativa (veloce, per PDF digitali).
    2. Se il PDF è una scansione o un'immagine, converte la pagina in immagine 
       e applica l'OCR per leggere il testo in modo pulito.
    """
    testo_totale = ""
    
    try:
        doc = fitz.open(pdf_path)
        
        for numero_pagina, pagina in enumerate(doc):
            # Tentativo 1: Estrazione del testo nativo
            testo_pagina = pagina.get_text()
            
            # Se il testo è quasi vuoto, significa che è una scansione/foto
            if len(testo_pagina.strip()) < 50:
                print(f"Pagina {numero_pagina + 1}: Rilevata scansione/immagine, avvio OCR...")
                
                # Converte la pagina PDF in un'immagine ad alta risoluzione (300 DPI)
                pix = pagina.get_pixmap(dpi=300)
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))
                
                # Esegue l'OCR in italiano per riconoscere il testo dell'immagine
                testo_pagina = pytesseract.image_to_string(img, lang='ita')
            
            testo_totale += f"\n--- Pagina {numero_pagina + 1} ---\n" + testo_pagina
            
        doc.close()
        return testo_totale
        
    except Exception as e:
        print(f"Errore durante l'elaborazione del PDF: {e}")
        return None

# Esempio di utilizzo:
# testo_estratto = estrai_testo_sicuro("percorso/bolletta.pdf")
# print(testo_estratto)
