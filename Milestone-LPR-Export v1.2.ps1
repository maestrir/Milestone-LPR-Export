Connect-ManagementServer -ShowDialog -AcceptEula

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Funzione Add-ExcelImage (già integrata)
function Add-ExcelImage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)] [OfficeOpenXml.ExcelWorksheet] $WorkSheet,
        [Parameter(Mandatory)] [System.Drawing.Image] $Image,
        [Parameter(Mandatory)] [int] $Row,
        [Parameter(Mandatory)] [int] $Column,
        [Parameter()] [string] $Name,
        [Parameter()] [int] $RowOffset = 1,
        [Parameter()] [int] $ColumnOffset = 1,
        [Parameter()] [switch] $ResizeCell
    )
    begin {
        $widthFactor = 1 / 7
        $heightFactor = 3 / 4
    }
    process {
        if ([string]::IsNullOrWhiteSpace($Name)) { $Name = (New-Guid).ToString() }
        $picture = $WorkSheet.Drawings.AddPicture($Name, $Image)
        $picture.SetPosition($Row - 1, $RowOffset, $Column - 1, $ColumnOffset)
        if ($ResizeCell) {
            $WorkSheet.Column($Column).Width = [Math]::Max($widthFactor * ($Image.Width + 1), $WorkSheet.Column($Column).Width)
            $WorkSheet.Row($Row).Height = [Math]::Max($heightFactor * ($Image.Height + 1), $WorkSheet.Row($Row).Height)
        }
    }
}

# ➕ Scelta qualità immagini
Write-Host "`nSeleziona la qualità delle immagini:"
Write-Host "1) Piccola (320x180)"
Write-Host "2) Media   (640x360)"
Write-Host "3) Grande  (1280x720)"
$sceltaQualita = Read-Host "Inserisci un valore (1-3)"
switch ($sceltaQualita) {
    "1" { $imgWidth = 320; $imgHeight = 180 }
    "2" { $imgWidth = 640; $imgHeight = 360 }
    "3" { $imgWidth = 1280; $imgHeight = 720 }
    default { Write-Host "Scelta non valida. Uso default 640x360."; $imgWidth = 640; $imgHeight = 360 }
}

# 📦 Nome server per il file Excel
$serverName = ($env:COMPUTERNAME).ToUpper()

# 📷 Selezione telecamere
$telecamere = Get-VmsCameraReport | Select Name, Id
$telecamere | ForEach-Object { Write-Host "$($_.Name)" }
$scelteCam = Read-Host "Inserisci una o più telecamere (separate da virgola) oppure 'ALL'"

# 📅 Periodo
Write-Host "`nSeleziona il periodo:"
Write-Host "1) Ultime 24 ore  2) Ultimi 7 giorni  3) Ultimi 30 giorni  4) Tutto  5) Personalizzato"
$scelta = Read-Host "Opzione (1-5)"
switch ($scelta) {
    "1" { $start = (Get-Date).AddDays(-1); $end = Get-Date }
    "2" { $start = (Get-Date).AddDays(-7); $end = Get-Date }
    "3" { $start = (Get-Date).AddDays(-30); $end = Get-Date }
    "4" { $start = [datetime]"1900-01-01"; $end = Get-Date }
    "5" {
        do {
            $dataOK = $true
            $startInput = Read-Host "Data inizio (yyyy-MM-dd HH:mm:ss.fff)"
            $endInput = Read-Host "Data fine (yyyy-MM-dd HH:mm:ss.fff)"
            try {
                $start = [datetime]::ParseExact($startInput, "yyyy-MM-dd HH:mm:ss.fff", $null)
                $end = [datetime]::ParseExact($endInput, "yyyy-MM-dd HH:mm:ss.fff", $null)
                if ($start -ge $end) {
                    Write-Host "⚠️ Data inizio >= fine." -ForegroundColor Yellow
                    $dataOK = $false
                }
            } catch {
                Write-Host "❌ Formato data non valido (usa yyyy-MM-dd HH:mm:ss.fff)" -ForegroundColor Red
                $dataOK = $false
            }
        } while (-not $dataOK)
    }
    default { Write-Host "Scelta non valida. Esco."; exit }
}

# 🔎 Query SQL
$server = "127.0.0.1"
$db = "Surveillance"
$filtroCam = ""
if ($scelteCam -ne "ALL") {
    $nomi = ($scelteCam.Split(",") | ForEach-Object { "'$($_.Trim())'" }) -join ","
    $filtroCam = "AND [SourceName] IN ($nomi)"
}
$query = @"
SELECT [UtcCreated], [ObjectValue] AS Targa, [SourceName] AS Telecamera, [ObjectId] AS CameraId
FROM [Surveillance].[Central].[Event_Active]
WHERE [Type]='LPR Event'
AND [UtcCreated] BETWEEN '$($start.ToString("yyyy-MM-dd HH:mm:ss.fff"))' AND '$($end.ToString("yyyy-MM-dd HH:mm:ss.fff"))'
$filtroCam
ORDER BY [UtcCreated]
"@

try {
    $dati = Invoke-Sqlcmd -ServerInstance $server -Database $db -Query $query -ErrorAction Stop
} catch {
    Write-Host "❌ Errore SQL: $_" -ForegroundColor Red
    exit
}
if (-not $dati -or $dati.Count -eq 0) {
    Write-Host "⚠️ Nessun evento trovato." -ForegroundColor Yellow
    exit
}

# 📂 Cartelle output
$timestampFolder = Get-Date -Format 'yyyyMMdd_HHmmss'
$base = ".\LPR_Export_$timestampFolder"
$imgFolder = "$base\Snapshots"
New-Item -ItemType Directory -Force -Path $imgFolder | Out-Null

# ⚙️ Elaborazione eventi
$report = @()
$count = 0
foreach ($row in $dati) {
    if (-not $row.UtcCreated -or -not $row.Targa -or -not $row.CameraId) {
        Write-Host "⏭️ Record ignorato per dati nulli." -ForegroundColor DarkGray
        continue
    }

    $count++
    $utc = [datetime]$row.UtcCreated
    $local = $utc.ToLocalTime()
    $imgName = "img_$($count.ToString("D4")).jpg"
    $imgPath = Join-Path $imgFolder $imgName

    Write-Host "📸 [$count] UTC: $utc → Locale: $local - Targa: $($row.Targa)"

    try {
        $snapshot = Get-Snapshot -CameraId $row.CameraId `
            -Timestamp $local `
            -Quality 100 -Width $imgWidth -Height $imgHeight -LiftPrivacyMask -LocalTimestamp

        if ($snapshot.Bytes.Count -gt 0) {
            [System.IO.File]::WriteAllBytes($imgPath, $snapshot.Bytes)
            Write-Host "✅ Snapshot salvato: $imgName" -ForegroundColor Green
        } else {
            $imgPath = $null
            Write-Host "⚠️ Nessun dato immagine" -ForegroundColor Yellow
        }
    } catch {
        $imgPath = $null
        Write-Host "❌ Errore snapshot: $_" -ForegroundColor Red
    }

    $report += [PSCustomObject]@{
        DataOra    = $local.ToString("yyyy-MM-dd HH:mm:ss.fff")
        Targa      = $row.Targa
        Telecamera = $row.Telecamera
        Note       = ""
        Immagine   = $imgPath
    }
}

# 📊 1️⃣ Esportazione file principale (solo dati, senza immagini)
$excelFileMain = "$base\${serverName}_Eventi_LPR_Tutti.xlsx"
$report | Export-Excel -Path $excelFileMain -WorksheetName "Eventi" -AutoSize
Write-Host "✅ File principale salvato (senza immagini): $excelFileMain" -ForegroundColor Cyan

# 🧮 2️⃣ Suddivisione in blocchi da 750 eventi
$bloccoMax = 750
$blocchi = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $report.Count; $i += $bloccoMax) {
    $blocchi.Add($report[$i..([Math]::Min($i + $bloccoMax - 1, $report.Count - 1))])
}

# 🔀 3️⃣ Esportazione dei blocchi con immagini
$bloccoIndex = 1
foreach ($blocco in $blocchi) {
    # 📅 Periodo di riferimento blocco
    $startDate = ($blocco[0].DataOra -as [datetime]).ToString("ddMMyy")
    $endDate   = ($blocco[-1].DataOra -as [datetime]).ToString("ddMMyy")
    $rangeStr  = "${startDate}_${endDate}"

    # 📁 Percorso immagini e file
    $bloccoCode = $bloccoIndex.ToString("D3")
    $imgSubfolder = Join-Path $imgFolder $bloccoCode
    New-Item -ItemType Directory -Force -Path $imgSubfolder | Out-Null

    $excelFile = "$base\${serverName}_LPR_${bloccoCode}_$rangeStr.xlsx"
    $excel = $blocco | Export-Excel -Path $excelFile -WorksheetName "LPR" -AutoSize -PassThru
    $ws = $excel.Workbook.Worksheets["LPR"]

    # 📷 Aggiunta immagini nel file Excel
    $row = 2
    foreach ($r in $blocco) {
        if ($r.Immagine -and (Test-Path $r.Immagine)) {
            $imgNameOnly = [System.IO.Path]::GetFileName($r.Immagine)
            $destImgPath = Join-Path $imgSubfolder $imgNameOnly

            # Copia immagine nella sottocartella del blocco
            Copy-Item $r.Immagine -Destination $destImgPath -Force

            # Inserisci immagine nel file Excel
            try {
                $img = [System.Drawing.Image]::FromFile($destImgPath)
                $ws | Add-ExcelImage -Image $img -Row $row -Column 5 -ResizeCell
            } catch {
                Write-Host "⚠️ Errore nell'inserimento immagine: $_" -ForegroundColor Yellow
            }
        }
        $row++
    }

    Close-ExcelPackage $excel
    Write-Host "📁 [$bloccoCode] Esportato con immagini: $excelFile" -ForegroundColor Green
    $bloccoIndex++
}

Write-Host "`n✅ Esportazione completata! Tutti i file sono pronti nella cartella: $base" -ForegroundColor Green
