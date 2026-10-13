# Apre con Excel il file di prova (dati INVENTATI) e la sua versione anonima, e controlla che l'anonima sia sana
# (nessuna riparazione all'apertura) e che formattazione, numeri e oggetti siano gli stessi dell'originale.
# Esce con codice 1 se un controllo fallisce.
# Uso: powershell -File apri_in_excel.ps1 <prova_formattata.xlsx> <prova_formattata_anonima.xlsx>
param([Parameter(Mandatory = $true)][string]$Originale, [Parameter(Mandatory = $true)][string]$File)
$ErrorActionPreference = 'Stop'

function Leggi($x, $percorso) {
    $wb = $x.Workbooks.Open($percorso, 0, $true)
    $x.CalculateFull()
    $c = $wb.Worksheets.Item(1)
    $n = $wb.Worksheets.Item(2)
    $r = $wb.Worksheets.Item(3)
    $v = [ordered]@{
        finestra            = $x.Windows.Item(1).Caption
        fogli               = $wb.Worksheets.Count
        nascosto            = $n.Visible
        titolo_unito        = $c.Range('A1').MergeCells
        titolo_grassetto    = $c.Range('A1').Font.Bold
        titolo_colore       = $c.Range('A1').Interior.Color
        larghezza_colonna_A = $c.Columns.Item(1).ColumnWidth
        formato_fatturato   = $c.Range('F4').NumberFormat
        fatturato           = ($c.Range('F4:F9').Value2 | ForEach-Object { $_ }) -join ';'
        # il grassetto di «Ferrero» dentro B5 non si confronta: il nome intero diventa un codice
        formula_G4          = $c.Range('G4').HasFormula
        collegamenti        = $c.Hyperlinks.Count
        commenti            = $c.Comments.Count
        caselle             = $c.Shapes.Count
        valore_A2           = $r.Range('A2').Value2
        grafici             = $r.ChartObjects().Count
        pivot               = $r.PivotTables().Count
        pivot_totale        = $r.PivotTables(1).GetPivotData('Somma di Fatturato 2025').Value2
        tabelle             = $c.ListObjects.Count
        nomi_definiti       = $wb.Names.Count
        importo_B5          = $n.Range('B5').Value2
        data_B10            = $n.Range('B10').Value2
        formato_data_B10    = $n.Range('B10').NumberFormat
    }
    $wb.Close($false)
    return $v
}

$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$falliti = 0
try {
    $o = Leggi $x (Resolve-Path $Originale).Path
    $a = Leggi $x (Resolve-Path $File).Path
    # Excel aggiunge al titolo «[Ripristinato]» o «(Repaired)»: si guarda il titolo senza il nome del file
    $resto = $a.finestra.Replace((Split-Path $File -Leaf), '')
    if ($resto -match 'ripar|repair|ripristin') { "FAIL Excel ha riparato il file: " + $a.finestra; $falliti++ }
    else { "PASS aperto senza riparazioni" }
    foreach ($k in $o.Keys) {
        if ($k -eq 'finestra') { continue }
        if ([string]$o[$k] -ceq [string]$a[$k]) { "PASS $k = " + $a[$k] }
        else { "FAIL $k : originale " + $o[$k] + " / anonimo " + $a[$k]; $falliti++ }
    }
} finally {
    $x.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
if ($falliti) { "`n$falliti controlli FALLITI"; exit 1 } else { "`ntutti i controlli superati" }
