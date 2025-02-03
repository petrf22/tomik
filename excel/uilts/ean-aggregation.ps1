# Proměnná $colLastNonEanLetter obsahuje poslední sloupec který není EAN kódem (po něm následují EAN kódy)

$excel = New-Object -Com Excel.Application
$excel.Visible = $true
$wbOrig = $null
$wb = $null
$colLastNonEanLetter = 'Y'
$xlDown = -4121
$xlToRight = -4161

# Definice třídy
class Context {
    [__ComObject]$sheet
    [int]$row
    [int]$col

    Context([__ComObject]$sheet) {
        $this.sheet = $sheet
        $this.row = 1
        $this.col = 1
    }
}

function ConvertTo-Windows1250 {
  param (
      [Parameter(Mandatory=$true)]
      [string]$Utf8Text
  )

  # Převod na bajty v UTF-8
  $utf8Bytes = [System.Text.Encoding]::UTF8.GetBytes($Utf8Text)

  # Převod bajtů na text v Windows-1250
  $windows1250Text = [System.Text.Encoding]::GetEncoding("windows-1250").GetString($utf8Bytes)

  return $windows1250Text
}

try {
  $importFile = 'c:\Users\Petr\github\petrf\tomik\excel\uilts\component-export-21012025.xlsx'
  # $importFile = $excel.GetOpenFilename("Excel files (*.xlsx*), *.xlsx*")

  if ($importFile -eq $false) {
    return
  }

  $excel.Interactive = $false

  $wbOrig = $excel.Workbooks.Open($importFile)

  # # Odstrannění NON EAN sloupců
  # $wsOrig.Columns("B:$($colLastNonEanLetter)").Delete()

  # # Odstrannění duplicitních řádků
  # $wsOrig.UsedRange.RemoveDuplicates(1)

  # $origRowsCount = $wsOrig.Columns("A:A").End($xlDown).Row
  # $origColsCount = $wsOrig.Rows("1:1").End($xlToRight).Column

  # Write-Host $wsOrig.UsedRange.columns.count
  # Write-Host $wsOrig.UsedRange.rows.count

  $wb = $excel.Workbooks.Add()
  $sheet = $wb.sheets.item(1)
  $ctx1 = [Context]::new($sheet)

  $sheet = $wb.sheets.item(1)
  $ctx1.sheet = $sheet

  # Získá poslední list v sešitu
  $lastSheet = $wb.Worksheets.Item($wb.Worksheets.Count)

  # Přidá nový list před poslední list a tím ho posune na konec
  # $newSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
  $ctx2 = [Context]::new($wb.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet))

  # List 1 - První řádek
  $ctx1.sheet.Cells.Item(1, 1) = $wsOrig.Cells.Item(1, 1)

  # List 2 - První řádek
  $ctx2.Cells.Item(1, 1).Value = ConvertTo-Windows1250 -Utf8Text 'Díl'
  $ctx2.Cells.Item(1, 2).Value = ConvertTo-Windows1250 -Utf8Text 'Materiál'
  $ctx2.Cells.Item(1, 3).Value = ConvertTo-Windows1250 -Utf8Text 'Množství'

  $ctx1.sheet.Activate

  return

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

    $col = 1
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
