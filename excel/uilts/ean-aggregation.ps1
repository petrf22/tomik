# Proměnná $colLastNonEanLetter obsahuje poslední sloupec který není EAN kódem (po něm následují EAN kódy)

$excel = New-Object -Com Excel.Application
$excel.Visible = $true
$colLastNonEanLetter = 'Y'
$xlDown = -4121
$xlToRight = -4161

# Definice třídy
class Context {
    [Object]$sheet
    [int]$row
    [int]$col

    Context([Object]$sheet) {
        $this.sheet = $sheet
        $this.row = 1
        $this.col = 1
    }
}

function Update-Sheet-1 {
  param (
    [Parameter(Mandatory=$true)] [Context]$ctx,
    [Parameter(Mandatory=$true)] [Context]$ctxOrig,
    [Parameter(Mandatory=$true)] [boolean]$isNumber
  )

  if ($isNumber -eq $True) {
    if ($ctx.sheet.Cells.Item($ctx.row, 1).Text -eq '') {
      $ctx.sheet.Cells.Item($ctx.row, 1).NumberFormat = "@"
      $ctx.sheet.Cells.Item($ctx.row, 1).Value = $ctxOrig.sheet.Cells.Item($ctxOrig.row, 1).Text # EAN
      $ctx.col++
    }

    $ctx.sheet.Cells.Item($ctx.row, $ctx.col) = $ctxOrig.sheet.Cells.Item(1, $ctxOrig.col)

    if ($ctx.sheet.Cells.Item(1, $ctx.col).Text -eq '') {
      $ctx.sheet.Cells.Item(1, $ctx.col).Value = 'EAN ' + ($ctx.col - 1)
    }

    $ctx.col++
  }
}

function Update-Sheet-2 {
  param (
    [Parameter(Mandatory=$true)] [Context]$ctx,
    [Parameter(Mandatory=$true)] [Context]$ctxOrig,
    [Parameter(Mandatory=$true)] [boolean]$isNumber
  )

  if ($item -eq $True) {
    $ctx.sheet.Cells.Item($ctx.row, 1).NumberFormat = "@"
    $ctx.sheet.Cells.Item($ctx.row, 1).Value = $ctxOrig.sheet.Cells.Item($ctxOrig.row, 1).Text # EAN
    $ctx.sheet.Cells.Item($ctx.row, 2).NumberFormat = "@"
    $ctx.sheet.Cells.Item($ctx.row, 2).Value = $ctxOrig.sheet.Cells.Item(1, $ctxOrig.col).Text # Nadpis (EAN) z prvního řádku
    $ctx.sheet.Cells.Item($ctx.row, 3) = $ctxOrig.sheet.Cells.Item($ctxOrig.row, $ctxOrig.col)
    $ctx.row++
  }

  $ctxOrig.col++
}


try {
  $importFile = 'c:\Users\Petr\github\petrf\tomik\excel\uilts\component-export-21012025.xlsx'
  # $importFile = $excel.GetOpenFilename("Excel files (*.xlsx*), *.xlsx*")

  if ($importFile -eq $false) {
    return
  }

  $excel.Interactive = $false

  $wbOrig = $excel.Workbooks.Open($importFile)
  $ctxOrig = [Context]::new($wbOrig.sheets.item(1))

  # Odstrannění NON EAN sloupců
  $ctxOrig.sheet.Columns("B:$($colLastNonEanLetter)").Delete() | out-null

  # Odstrannění duplicitních řádků
  $ctxOrig.sheet.UsedRange.RemoveDuplicates(1) | out-null

  $origRowsCount = $ctxOrig.sheet.Columns("A:A").End($xlDown).Row
  $origColsCount = $ctxOrig.sheet.Rows("1:1").End($xlToRight).Column

  Write-Host $ctxOrig.sheet.UsedRange.columns.count
  Write-Host $ctxOrig.sheet.UsedRange.rows.count

  $wb = $excel.Workbooks.Add()
  $ctx1 = [Context]::new($wb.sheets.item(1))

  # Získá poslední list v sešitu
  $lastSheet = $wb.Worksheets.Item($wb.Worksheets.Count)

  # Přidá nový list před poslední list a tím ho posune na konec
  # $newSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
  $ctx2 = [Context]::new($wb.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet))

  #$ctx1.sheet.Select | out-null
  $wb.Sheets($ctx1.sheet.Name).Select | out-null

  # List 1 - První řádek
  $ctx1.sheet.Cells.Item(1, 1) = $ctxOrig.sheet.Cells.Item(1, 1)

  # List 2 - První řádek
  $ctx2.sheet.Cells.Item(1, 1).Value = 'Dil'
  $ctx2.sheet.Cells.Item(1, 2).Value = 'Material'
  $ctx2.sheet.Cells.Item(1, 3).Value = 'Mnozstvi'

  $ctx1.row = 2
  $ctx2.row = 2
  $startDate = Get-Date
  $estimateText = ''
  #$origRowsCount = $ctxOrig.sheet.UsedRange.rows.count
  $colFirstEan = 2 # $ctxOrig.sheet.Columns($colFirstEanLetter).Column

  for ($ctxOrig.row = 2; $ctxOrig.row -le $origRowsCount; $ctxOrig.row++)
  {
    $ctxOrig.col = 1
    $ean = $ctxOrig.sheet.Cells($ctxOrig.row, $ctxOrig.col).Text
    # $ctxOrig.sheet.Cells.Item(1, 1).text
    # Write-Host "EAN: $($ean)"
    if ($ctxOrig.row -gt 100 -and $ctxOrig.row % 10 -eq 0) {
      $endDate = Get-Date
      $totalSeconds = $(New-TimeSpan $startDate $endDate).TotalSeconds
      $rowPerTime = $totalSeconds / $ctxOrig.row
      $estimateSec = ($origRowsCount - $ctxOrig.row) * $rowPerTime
      $estimateTime =  [timespan]::fromseconds($estimateSec)
      $estimateText = "(odhad: $("{0:hh\:mm\:ss\,fff}" -f $estimateTime))";
    }

    $proc = 100 / $origRowsCount * $ctxOrig.row
    Write-Progress -Activity "Vydrzte, stroje pracuji za vas ..." -Status "$("{0:N3}" -f [Math]::Round($proc, 3))% $($estimateText)" `
                   -PercentComplete $proc -CurrentOperation "Radek cislo $($ctxOrig.row) z $($origRowsCount), EAN: $($ean)"

    $ctx1.col = 1
    $ctx2.col = 1

    # $ctxOrig.col++
    $ctxOrig.col = $colFirstEan

    $range = $ctxOrig.sheet.Range($ctxOrig.sheet.Cells($ctxOrig.row, $ctxOrig.col), $ctxOrig.sheet.Cells($ctxOrig.row, $origColsCount))
    $arrayIsNumber = $excel.WorksheetFunction.IsNumber($range)

    foreach ($item in $arrayIsNumber) {
      Update-Sheet-1 -ctx $ctx1 -ctxOrig $ctxOrig -isNumber $item
      Update-Sheet-2 -ctx $ctx2 -ctxOrig $ctxOrig -isNumber $item

      $ctxOrig.col++
    }

    $ctx1.row++
  }
} finally {
  $excel.Interactive = $true
  if ($null -ne $wbOrig) {
    $wbOrig.Close($false)
  }

  # if ($null -ne $wb) {
  #   $wb.Close($true)
  # }

  $excel.Quit()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  if ($null -ne $wbOrig) {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbOrig) | out-null
  }

  if ($null -ne $wb) {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
  }

  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null

  Remove-Variable -Name excel
}
