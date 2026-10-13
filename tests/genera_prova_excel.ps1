# Crea con Excel un file di prova con dati INVENTATI e tutti i posti dove un nome puo' nascondersi:
# formattazione, testo ricco, commento, collegamento email, formule con testo e riferimenti ad altri fogli,
# elenco a discesa, intestazione di stampa, foglio nascosto, tabella, grafico, tabella pivot, nomi definiti.
# Anche: autore del commento e del file, intestazioni con i codici di Excel, formato numerico con un nome, casella di
# testo, messaggio di convalida, P.IVA salvata come numero, importo che sembra una P.IVA, accento scritto separato.
# Uso: powershell -File genera_prova_excel.ps1 <cartella>   (nella cartella ci deve essere sensibili.json)
# I caratteri non ASCII si scrivono con [char]: Windows PowerShell 5.1 legge i .ps1 senza BOM come ANSI.
param([Parameter(Mandatory = $true)][string]$Cartella)
$ErrorActionPreference = 'Stop'
trap { "ERRORE alla riga " + $_.InvocationInfo.ScriptLineNumber + ": " + $_; break }
$s = Get-Content -Raw -Encoding UTF8 (Join-Path $Cartella 'sensibili.json') | ConvertFrom-Json
$pive = $s.pive
$euro = [string][char]0x20AC
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
# nome dell'utente di Excel: diventa l'autore del commento e del file; si rimette com'era alla fine
$utenteExcel = $x.UserName
$x.UserName = 'Sara Moretti'
try {
    $wb = $x.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $c = $wb.Worksheets.Item(1); $c.Name = 'Clienti'
    $n = $wb.Worksheets.Item(2); $n.Name = 'Note Edilnord'
    $r = $wb.Worksheets.Item(3); $r.Name = 'Riepilogo'

    $c.Range('A1').Value2 = 'Bilancio clienti 2025 - preparato da Laura Verdi'
    $c.Range('A1:F1').Merge()
    $c.Range('A1').Font.Bold = $true
    $c.Range('A1').Font.Size = 14
    $c.Range('A1').Interior.Color = 15773696
    $intest = @('Ragione sociale', 'Referente', 'Email', 'Telefono', 'P.IVA', 'Fatturato 2025')
    for ($j = 0; $j -lt 6; $j++) { $c.Range([string][char](65 + $j) + '3').Value2 = [string]$intest[$j] }
    $c.Range('A3:F3').Font.Bold = $true
    $righe = @(
        @('Edilnord Srl', 'Marco Galli', 'm.galli@edilnord.it', '011 5551234', $pive[0], 152340.5),
        @('Tecnofarm S.p.A.', 'Giovanna Ferrero', 'amministrazione@tecnofarm.com', '+39 347 1234567', $pive[1], 98760),
        @('Bianchi & Figli Snc', 'Paolo Bianchi', 'paolo@bianchiefigli.it', '02 48001234', $pive[2], 45000),
        @('Studio Associato Rinaldi', 'Elena Rinaldi', 'elena.rinaldi@studiorinaldi.it', '333 9876543', $pive[3], 12500.75),
        @('Alfa Logistica Srl', 'Davide Conti', 'd.conti@alfalogistica.it', '0125 612345', $pive[4], 230000),
        @('Mondo Verde Coop', 'Sara Moretti', 'info@mondoverde.coop', '011 7778899', $pive[5], 8700)
    )
    $c.Range('D4:E9').NumberFormat = '@'
    for ($i = 0; $i -lt 6; $i++) {
        for ($j = 0; $j -lt 6; $j++) {
            $v = $righe[$i][$j]
            $cella = $c.Range([string][char](65 + $j) + [string](4 + $i))
            if ($v -is [double] -or $v -is [int]) { $cella.Value2 = [double]$v } else { $cella.Value2 = [string]$v }
        }
    }
    $c.Range('F4:F8').NumberFormat = '#,##0.00 "' + $euro + '"'
    $c.Range('F9').NumberFormat = '#,##0 "rif. Davide Conti"'
    $c.Columns.Item(1).ColumnWidth = 28
    $c.Columns.Item(3).ColumnWidth = 32
    $c.Cells.Item(5, 2).Characters(10, 7).Font.Bold = $true                # 'Ferrero' in grassetto dentro la cella
    [void]$c.Hyperlinks.Add($c.Range('C4'), 'mailto:m.galli@edilnord.it', '', '', 'm.galli@edilnord.it')
    [void]$c.Range('F4').AddComment("Sara Moretti:`nDato confermato da Marco Galli")
    $c.Range('G3').Value2 = 'Etichetta'
    $c.Range('G4').Formula = '="Cliente: "&A4'
    $c.Range('G5').Formula = '="Referente Marco Galli"'
    [void]$c.Range('H4').Validation.Add(3, 1, 1, 'Edilnord Srl,Tecnofarm S.p.A.')
    $c.Range('H4').Validation.InputTitle = 'Nota'
    $c.Range('H4').Validation.InputMessage = 'Chiedere a Elena Rinaldi'
    $c.PageSetup.LeftHeader = 'Marco Galli'
    $c.PageSetup.CenterHeader = 'Riservato Laura Verdi'
    $c.PageSetup.RightFooter = 'Bianchi && Figli Snc'
    $c.Shapes.AddTextbox(1, 420, 10, 220, 40).TextFrame2.TextRange.Text = 'Approvato da Davide Conti'

    $n.Range('A1').Value2 = 'Riunione con Marco Galli e Sara Moretti'
    $n.Range('A3').Value2 = 'P.IVA salvata come numero'
    $n.Range('B3').Value2 = [double]$pive[1]
    $n.Range('A4').Value2 = 'Partita IVA'
    $n.Range('B4').Value2 = [double]$s.piva_numero                 # c'e' solo come numero: lo zero iniziale si perde
    $n.Range('A5').Value2 = 'Incassi in centesimi'
    $n.Range('B5').Value2 = [double]$s.importo_come_piva              # passa il controllo della P.IVA per caso
    $n.Range('A6').Value2 = 'Telefonata con Nicolo' + [char]0x0300 + ' Ferre' + [char]0x0300   # accento separato
    $n.Range('A7').Value2 = 'Nicol' + [char]0x00F2 + ' Ferr' + [char]0x00E8
    $n.Range('A8').Value2 = 'Bonifico IBAN ' + $s.iban
    $n.Range('A9').Value2 = 'Codice fiscale ' + $s.cf[0]
    $n.Range('A10').Value2 = 'Data riunione'
    $n.Range('B10').Value = [datetime]'2025-01-15'                    # Excel sceglie da solo il formato data

    $r.Range('A1').Formula = "='Note Edilnord'!A1"
    $r.Range('A2').Formula = '=Clienti!F4*2'

    $lo = $c.ListObjects.Add(1, $c.Range('A3:F9'), $null, 1)
    $lo.Name = 'TabClienti'
    $ch = $r.ChartObjects().Add(300, 20, 420, 260)
    # due aree in un solo intervallo: Union non va con Windows PowerShell 5.1, e il separatore dipende dalla lingua
    $sep = $x.International(5)
    $ch.Chart.SetSourceData($c.Range('A3:A9' + $sep + 'F3:F9'))
    $ch.Chart.ChartType = 57
    $ch.Chart.HasTitle = $true
    $ch.Chart.ChartTitle.Text = 'Fatturato - Laura Verdi'
    $pc = $wb.PivotCaches().Create(1, 'TabClienti')
    $pt = $pc.CreatePivotTable($r.Range('A20'), 'PivotClienti')
    $pt.PivotFields('Ragione sociale').Orientation = 1
    [void]$pt.AddDataField($pt.PivotFields('Fatturato 2025'))

    [void]$wb.Names.Add('Contatto', '="Marco Galli"')
    [void]$wb.Names.Add('NotaRif', "='Note Edilnord'!`$A`$1")
    $n.Visible = 0

    $percorso = Join-Path $Cartella 'prova_formattata.xlsx'
    if (Test-Path $percorso) { Remove-Item $percorso }
    $wb.SaveAs($percorso, 51)
    $wb.Close($false)
    "ok $percorso"
} finally {
    $x.UserName = $utenteExcel
    $x.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
