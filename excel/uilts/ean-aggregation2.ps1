# Soubor musí být uložel ve znakové sadì Windows 1250 (ANSI)
# Promìnná $colLastNonEanLetter obsahuje poslední sloupec který není EAN kódem (po nìm následují EAN kódy)

$excel = New-Object -Com Excel.Application
$excel.Visible = $true
$wbOrig = $null
$wb = $null
$colLastNonEanLetter = 'Y'
$xlDown = -4121
$xlToRight = -4161

try {
  # $importFile = 'c:\Users\Petr\github\petrf\tomik\excel\uilts\component-export-21012025.xlsx'
  $importFile = $excel.GetOpenFilename("Excel files (*.xlsx*), *.xlsx*")

  if ($importFile -eq $false) {
    return
  }

  $excel.Interactive = $false

  $wbOrig = $excel.Workbooks.Open($importFile)
  $wsOrig = $wbOrig.sheets.item(1)

  # Odstrannìní NON EAN sloupcù
  $wsOrig.Columns("B:$($colLastNonEanLetter)").Delete()

  # Odstrannìní duplicitních øádkù
  $wsOrig.UsedRange.RemoveDuplicates(1)

  $origRowsCount = $wsOrig.Columns("A:A").End($xlDown).Row
  $origColsCount = $wsOrig.Rows("1:1").End($xlToRight).Column

  # Write-Host $wsOrig.UsedRange.columns.count
  # Write-Host $wsOrig.UsedRange.rows.count

  $wb = $excel.Workbooks.Add()
  $ws = $wb.sheets.item(1)

  $ws.Cells.Item(1, 1).Value = 'Díl'
  $ws.Cells.Item(1, 2).Value = 'Materiál'
  $ws.Cells.Item(1, 3).Value = 'Množství'

  $row = 2
  $ean = ''
  $lastEan = ''
  $startDate = Get-Date
  $estimateText = ''
  #$origRowsCount = $wsOrig.UsedRange.rows.count
  $colFirstEan = $wsOrig.Columns($colFirstEanLetter).Column

  for ($rowOrig = 2; $rowOrig -le $origRowsCount; $rowOrig++)
  {
    $colOrig = 1
    $ean = $wsOrig.Cells($rowOrig, $colOrig).Text
    # $wsOrig.Cells.Item(1, 1).text
    # Write-Host "EAN: $($ean)"
    if ($rowOrig -gt 100 -and $rowOrig % 10 -eq 0) {
      $endDate = Get-Date
      $totalSeconds = $(New-TimeSpan $startDate $endDate).TotalSeconds
      $rowPerTime = $totalSeconds / $rowOrig
      $estimateSec = ($origRowsCount - $rowOrig) * $rowPerTime
      $estimateTime =  [timespan]::fromseconds($estimateSec)
      $estimateText = "(odhad: $("{0:hh\:mm\:ss\,fff}" -f $estimateTime))";
    }

    $proc = 100 / $origRowsCount * $rowOrig
    Write-Progress -Activity "Vydrzte, stroje pracuji za vas ..." -Status "$("{0:N3}" -f [Math]::Round($proc, 3))% $($estimateText)" `
                   -PercentComplete $proc -CurrentOperation "Radek cislo $($rowOrig) z $($origRowsCount), EAN: $($ean)"

    if ($ean -eq $lastEan) {
      # Duplicitní EAN
      # Write-Host "Duplicitní EAN: $($ean)"
      continue
    }

    # Write-Host "rowOrig: $($rowOrig)"

    $range = $wsOrig.Range($wsOrig.Cells($rowOrig, $colFirstEan), $wsOrig.Cells($rowOrig, $origColsCount))
    # Write-Host "rowOrig: $($range.Formula2R1C1)"

    $arrayIsNumber = $excel.WorksheetFunction.IsNumber($range)
    $colOrig = $colFirstEan

    foreach ($item in $arrayIsNumber) {
      if ($item -eq $True) {
        $ws.Cells.Item($row, 1).NumberFormat = "@"
        $ws.Cells.Item($row, 1).Value = $ean
        $ws.Cells.Item($row, 2).NumberFormat = "@"
        $ws.Cells.Item($row, 2).Value = $wsOrig.Cells.Item(1, $colOrig).Text # Nadpis (EAN) z prvního øádku
        $ws.Cells.Item($row, 3) = $wsOrig.Cells.Item($rowOrig, $colOrig)
        $row++
      }
      $colOrig++
    }

    $lastEan = $ean
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
