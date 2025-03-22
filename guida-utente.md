# Guida Utente – Milestone-LPR-Export.ps1

Questa guida descrive come utilizzare lo script `Milestone-LPR-Export.ps1` per esportare eventi di riconoscimento targhe (LPR) dal database Milestone, acquisire snapshot dalle telecamere e creare un report Excel con immagini integrate.

---

## 🔧 Requisiti

- Windows 10/11 o Windows Server 2012 o superiore
- PowerShell 5.1+
- Database SQL Server con accesso al database `Surveillance`
- Milestone XProtect Corporate configurato e operativo
- Modulo PowerShell:
  - [`MilestonePSTools`](https://www.milestonepstools.com) per snapshot
  - [`ImportExcel`](https://github.com/dfinke/ImportExcel) per generare file `.xlsx`

---

## 📦 Dipendenze

### 📌 1. MilestonePSTools

Installa con:
```powershell
Install-Module MilestonePSTools -Scope CurrentUser -Force
```

Collega lo script al management server con:
```powershell
Connect-ManagementServer -ShowDialog -AcceptEula
```

Per guida completa:  
👉 https://www.milestonepstools.com/getting-started

---

### 📌 2. ImportExcel

Installa con:
```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

---

## 🚀 Come funziona lo script

### 1. Selezioni iniziali (runtime)
Lo script chiede:
- Qualità snapshot (`piccola`, `media`, `grande`)
- Una o più telecamere da includere (oppure `ALL`)
- L’intervallo temporale (24h, 7 giorni, 30 giorni, tutto, o personalizzato)
- Il nome del server viene prelevato automaticamente

### 2. Query SQL eseguita

Lo script interroga la tabella:
```sql
[Surveillance].[Central].[Event_Active]
```
Filtrando i campi:
- `[Type]` = `'LPR Event'`
- `[UtcCreated]` compreso nell’intervallo selezionato
- `[ObjectValue]` → Targa
- `[SourceName]` → ID della telecamera (non il nome leggibile)
- `[ObjectId]` → CameraId usato per il comando `Get-Snapshot`

⚠️ Solo i record in cui **[UtcCreated]**, **[ObjectValue]**, e **[ObjectId]** non sono nulli vengono elaborati.

---

## 📸 Acquisizione snapshot

Per ogni evento LPR, viene chiamato:
```powershell
Get-Snapshot -CameraId ... -Timestamp ... -Width ... -Height ...
```
L’immagine viene salvata in JPEG e collegata a una riga del report.

---

## 📊 Output generato

- Report Excel `.xlsx` in una cartella `LPR_Export_<timestamp>`
- Colonne:
  - Data/Ora (convertita da UTC a ora locale)
  - Targa
  - Nome della telecamera (ricavato da mapping iniziale)
  - Immagine incorporata
  - Eventuali note (campo vuoto modificabile)

---

## 📂 Struttura cartelle

Al termine lo script crea:
```
/LPR_Export_<data>/
├── Snapshots/         # Immagini JPEG
└── <SERVER>_Eventi_LPR_<timestamp>.xlsx
```

---

## 🛠️ Risoluzione problemi

- **Nessun evento trovato**:
  - Intervallo temporale errato o nessun evento LPR presente
- **Immagini non acquisite**:
  - Verificare che la telecamera sia abilitata alla registrazione
  - Se l’archivio è stato cancellato, lo snapshot **non è recuperabile**
- **Errore nel salvataggio Excel**:
  - Verificare che ImportExcel sia correttamente installato
- **Errore SQL**:
  - Controllare accesso al database `Surveillance`
  - Controllare che `Invoke-Sqlcmd` sia disponibile (modulo `SqlServer`)

---

## 📬 Contatti

Autore: [https://github.com/maestrir](https://github.com/maestrir)

---

## 🧾 Licenza

Distribuito sotto licenza MIT. Consulta il file `LICENSE` per i dettagli.