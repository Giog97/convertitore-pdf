# Convertitore PDF

Applicazione desktop con interfaccia grafica (due schede) per:

1. **Convertire in PDF** immagini e documenti Word
2. **Dividere un PDF** in più file

Interfaccia disponibile in **Italiano** e **Inglese** (selettore 🌐 in alto a destra).

## Avvio rapido

- **Se vuoi solo usarlo:** apri la cartella `dist` e fai doppio clic su
  **`Convertitore PDF.exe`** (file unico, non serve installare nulla).
- **Se vuoi avviarlo dal codice:** doppio clic su **`avvia.bat`**
  (al primo avvio prepara l'ambiente e installa le librerie).

---

## Scheda “Converti in PDF”

- **Immagini** (JPG, PNG, BMP, GIF, TIFF, WEBP, ICO, JFIF…) → convertite con Pillow
- **Documenti Word** (DOCX, DOCM) → convertiti **senza Microsoft Word** (mammoth + xhtml2pdf), quindi funziona su qualsiasi PC. *(I formati più vecchi .doc/.rtf/.odt non sono supportati: salvali prima come .docx.)*
- **Anteprima del PDF risultante** (non del file caricato): per le immagini mostra come apparirà sulla pagina scelta e si aggiorna quando cambi le opzioni; per i documenti genera in background la vera anteprima del PDF
- **Rotazione** dall'anteprima con i pulsanti ⟲ ⟳: il PDF finale viene ruotato di conseguenza e la pagina si orienta automaticamente
- **Auto-orientamento EXIF**: molte foto da telefono vengono raddrizzate automaticamente
- Spazio **diviso a metà** tra elenco file e anteprima, ridimensionabile con la finestra
- **Trascinamento (drag & drop)** di file e cartelle nella finestra
- **Unisci tutti i file in un unico PDF** (immagini + documenti insieme, nell'ordine dell'elenco)
- **Riordina i file** con i pulsanti ▲ ▼ **o trascinando il nome** nell'elenco
- **Formato pagina immagini**: *Adatta all'immagine*, *A4* o *Lettera* (con orientamento automatico)
- Scelta della **cartella di destinazione** (di default la stessa dei file di origine)

## Scheda “Dividi PDF”

- Scegli un PDF (anche col drag & drop) e ne vedi il **numero di pagine**
- **Anteprima delle pagine** con navigazione ◀ ▶
- Tre modalità di divisione:
  - **Una pagina per file** (ogni pagina diventa un PDF separato)
  - **Ogni N pagine** per file
  - **Estrai pagine** specifiche, es. `1-3, 5, 8`
- Scelta della cartella di destinazione

> I PDF non sovrascrivono mai file esistenti: in caso di nome già presente
> viene aggiunto un numero, es. `documento (2).pdf`.

---

## Requisiti

- **Windows**
- Per l'uso dal codice: **Python 3** nel PATH.
- **Nessun altro programma**: la conversione Word→PDF non richiede Microsoft Word.

## Portare il programma su un altro computer

**Modo più semplice:** copia il solo file `dist\Convertitore PDF.exe`
(es. su una chiavetta USB). Sull'altro PC basta un doppio clic, senza installare
Python né Microsoft Word: la conversione dei documenti funziona comunque.

## Ricreare l'eseguibile

Se modifichi il codice e vuoi rigenerare l'`.exe`, fai doppio clic su
**`crea_exe.bat`**: ricrea `dist\Convertitore PDF.exe`.

## Caricare il progetto su GitHub

Il file `.gitignore` è già configurato per **non** caricare l'ambiente (`.venv`),
la cache e la cartella `dist/` (l'eseguibile). Comandi da eseguire nella cartella
del progetto:

```bash
git init
git add .
git commit -m "Primo commit: Convertitore PDF"
git branch -M main
git remote add origin https://github.com/TUO-UTENTE/convertitore-pdf.git
git push -u origin main
```

Prima devi creare un repository vuoto su GitHub (senza README) e incollarne l'URL
al posto di quello sopra.

**L'eseguibile** (`dist/Convertitore PDF.exe`) è escluso dal repository perché è
grande e va ricreato dal codice. Il modo consigliato per distribuirlo è tramite
**GitHub → Releases**: crea una release e allega il file `.exe` come asset, così
chi scarica trova sia il codice sia il programma pronto.

## Struttura del progetto

| File / cartella      | Descrizione                                              |
|----------------------|----------------------------------------------------------|
| `convertitore.py`    | Applicazione principale (interfaccia + logica)           |
| `avvia.bat`          | Avvio dal codice con installazione automatica            |
| `crea_exe.bat`       | Crea il file `.exe` unico (doppio clic)                  |
| `build_exe.py`       | Script di build usato da `crea_exe.bat`                  |
| `requirements.txt`   | Librerie richieste                                       |
| `.gitignore`         | File esclusi dal repository GitHub                       |
| `assets/`            | Icona dell'applicazione                                  |
| `dist/`              | Contiene l'eseguibile `Convertitore PDF.exe`             |

## Librerie usate

Pillow (immagini) · mammoth + xhtml2pdf (Word→PDF senza Word) · tkinterdnd2 (drag & drop) ·
pypdf (unione/divisione/rotazione PDF) · pypdfium2 (anteprima pagine PDF).
