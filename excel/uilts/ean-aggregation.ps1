# Proměnná $colLastNonEanLetter obsahuje poslední sloupec který není EAN kódem (po něm následují EAN kódy)

$excel = New-Object -Com Excel.Application
$excel.Visible = $true
$wbOrig = $null
$wb = $null
$colLastNonEanLetter = 'Y'
$xlDown = -4121
$xlToRight = -4161

try {
  #$importFile = 'c:\Users\Petr\Downloads\Sesit1comp.xlsx'
  $importFile = $excel.GetOpenFilename("Excel files (*.xlsx*), *.xlsx*")

  if ($importFile -eq $false) {
    return
  }

  $wbOrig = $excel.Workbooks.Open($importFile)
  $wsOrig = $wbOrig.sheets.item(1)

  # Odstrannění NON EAN sloupců
  $wsOrig.Columns("B:$($colLastNonEanLetter)").Delete()

  # Odstrannění duplicitních řádků
  $wsOrig.UsedRange.RemoveDuplicates(1)

  $origRowsCount = $wsOrig.Columns("A:A").End($xlDown).Row
  $origColsCount = $wsOrig.Rows("1:1").End($xlToRight).Column

  # Write-Host $wsOrig.UsedRange.columns.count
  # Write-Host $wsOrig.UsedRange.rows.count

  $wb = $excel.Workbooks.Add()
  $ws = $wb.sheets.item(1)

  $ws.Cells.Item(1, 1).NumberFormat = "@"
  $ws.Cells.Item(1, 1).Value = $wsOrig.Cells.Item(1, 1).Text

  $col = 1
  $row = 2
  $ean = ''
  $lastEan = ''
  $startDate = Get-Date
  $estimateText = ''
  #$origRowsCount = $wsOrig.UsedRange.rows.count
  $colFirstEan = 2 # $wsOrig.Columns($colFirstEanLetter).Column

  for ($rowOrig = 2; $rowOrig -le $origRowsCount; $rowOrig++)
  {
    $colOrig = 1
    $col = 1
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

    $ws.Cells.Item($row, $col).NumberFormat = "@"
    $ws.Cells.Item($row, $col).Value = $ean

    $col++
    # $colOrig++
    $colOrig = $colFirstEan

    $range = $wsOrig.Range($wsOrig.Cells($rowOrig, $colOrig), $wsOrig.Cells($rowOrig, $origColsCount))
    $arrayIsNumber = $excel.WorksheetFunction.IsNumber($range)

    foreach ($item in $arrayIsNumber) {
      if ($item -eq $True) {
        $ws.Cells.Item($row, $col).NumberFormat = "@"
        $ws.Cells.Item($row, $col).Value = $wsOrig.Cells.Item(1, $colOrig).Text
        $col++
      }
      $colOrig++
    }

    $lastEan = $ean
    $row++
  }
} finally {
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
