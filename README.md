# Milestone-LPR-Export

Script PowerShell per esportare eventi LPR (License Plate Recognition) da sistemi Milestone XProtect, con generazione automatica di snapshot e report Excel completo di immagini.

## ✨ Caratteristiche

- Connessione a Management Server di Milestone XProtect
- Estrazione degli eventi "LPR Event" dal database SQL
- Supporto per selezione periodo e telecamere (anche multiple o tutte)
- Conversione automatica UTC → orario locale per snapshot precisi
- Salvataggio snapshot in 3 formati (piccolo, medio, grande)
- Report Excel con immagini embeddate nella tabella
- Nome file Excel contenente il nome del server
- Nessuna dipendenza da SDK, utilizza solo `PSTools` e `ImportExcel`

## 🧰 Requisiti

- PowerShell 5.1 o superiore
- [ImportExcel](https://github.com/dfinke/ImportExcel)
- [Milestone PSTools](https://www.powershellgallery.com/packages/MilestonePSTools)
- Accesso al database SQL Milestone
- Utente con permessi di lettura sul DB e accesso al Management Server

## 📦 Installazione moduli richiesti

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
Install-Module MilestonePSTools -Scope CurrentUser -Force
```

## 🚀 Utilizzo

1. Posiziona lo script `LPR_Export.ps1` in una cartella locale
2. Esegui da PowerShell con:

```powershell
.\LPR_Export.ps1
```

3. Lo script guida l’operatore:
   - Connessione al server Milestone
   - Selezione telecamere
   - Periodo di esportazione
   - Qualità immagini (piccolo/medio/grande)

### 📁 Output generato

```
LPR_Export_20250321_102233\
├── Snapshots\
│   ├── img_0001.jpg
│   └── img_0002.jpg
└── MILCORPDELL26_Eventi_LPR_20250321_102233.xlsx
```

## 🖼️ Report Excel con immagini

| DataOra                   | Targa   | Telecamera                             | Note | Immagine |
|---------------------------|---------|----------------------------------------|------|----------|
| 2025-03-21 08:44:12.123   | AA123BB | GV-LPR2800-DL (192.168.24.110)         |      | 📷        |

## ⚙️ Parametri interattivi

- 📸 Qualità immagini:
  - 1 = 320x180 (piccolo)
  - 2 = 640x360 (medio)
  - 3 = 1280x720 (grande)

- 🕒 Periodo esportazione:
  - Ultime 24 ore / 7 giorni / 30 giorni / Tutto
  - Oppure personalizzato con formato: `yyyy-MM-dd HH:mm:ss.fff`

## 📋 Licenza

Questo script è distribuito sotto licenza MIT. Vedi file `LICENSE`.

## 🙌 Autore

Sviluppato da **Roberto Mestri** per l’ottimizzazione e semplificazione delle attività LPR in ambienti Milestone XProtect.
