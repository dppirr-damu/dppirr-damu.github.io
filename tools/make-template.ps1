param([string]$Out, [string]$Html)
# Шаблон Google Таблицы строится из блока DEFAULT_DATA в index.html — данные совпадают с сайтом.
$ErrorActionPreference = "Stop"

$src = [IO.File]::ReadAllText($Html, [Text.Encoding]::UTF8)
$a = $src.IndexOf("const DEFAULT_DATA = {"); $b = $src.IndexOf("const STATUS = {")
$js = $src.Substring($a, $b - $a)
$js = $js.Substring($js.IndexOf("{")).Trim().TrimEnd(";")
$js = [regex]::Replace($js, "/\*[\s\S]*?\*/", "")
$js = [regex]::Replace($js, "(?m)^(\s*)([a-zA-Z]+):", '$1"$2":')
$js = [regex]::Replace($js, "([{,]\s*)([a-zA-Z]+):", '$1"$2":')
$js = [regex]::Replace($js, ",(\s*[\]}])", '$1')
$data = $js | ConvertFrom-Json

$STATUS = @{ active = "Реализуется"; process = "Набирает обороты"; plan = "Планируется" }
$STAGE  = @{ done = "Выполнен"; now = "В работе"; next = "Впереди" }

$YELLOW = 10092543   # RGB(255,255,153) в формате BGR
$HEAD   = 14277081
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $sep = $xl.International(5)
  $n = $data.programs.Count
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt $n + 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }

  function Head($ws, $vals) {
    for ($i = 0; $i -lt $vals.Count; $i++) { $ws.Cells.Item(1, $i + 1).Value2 = $vals[$i] }
    $r = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$vals.Count))
    $r.Font.Bold = $true; $r.Interior.Color = $HEAD
  }
  function SetText($cell, $v) { $cell.NumberFormat = "@"; $cell.Value2 = [string]$v }

  # --- Инструкция ---
  $ws = $wb.Worksheets.Item(1); $ws.Name = "Инструкция"
  $lines = @(
    "Шаблон данных дашборда нефинансовой поддержки Фонда «Даму»",
    "",
    "1. Каждая программа — на своём листе «01 …», «02 …» и далее. Порядок на сайте — по номеру в названии листа.",
    "2. Менять значения только в жёлтых ячейках. Колонку A («Раздел») не переименовывать.",
    "3. Новая строка показателя, этапа или пункта: вставить строку и выбрать «Раздел» из выпадающего списка.",
    "4. Статус программы: Реализуется / Набирает обороты / Планируется.",
    "5. Статус этапа: Выполнен / В работе / Впереди; дата — в формате ДД.ММ.ГГГГ.",
    "6. Охват: в колонке C — факт (число), в D — план (число). Диаграмма строится, только если оба — числа.",
    "7. Нет подтверждённого значения — ставить «—». Не выдумывать и не экстраполировать цифры.",
    "8. Описание — не более 2 предложений, вывод — 1 предложение.",
    "9. Лист «Сводка» заполняет ответственный: дата актуальности и сводные KPI.",
    "10. Строки с другим текстом в колонке A (пояснения, заметки) сайт игнорирует.",
    "11. «Ход реализации» — пункты для справки к встрече с ПП, по одному пункту в строке; строк можно добавлять сколько нужно.",
    "12. «Для прессы» — готовый текст для СМИ и интервью (если пусто, сайт соберёт его из описания и цифр). «Итоги года» — абзац для годового отчёта.",
    "",
    "ВНИМАНИЕ: таблица опубликована в интернете и доступна любому, у кого есть ссылка.",
    "Не вносить сюда внутренние сведения: ФИО ответственных, проблемные обращения, переписку с госорганами, вопросы к Председателю Правления, персональные данные.",
    "Такие сведения добавляются в справку вручную в Word после копирования текста с сайта.",
    "",
    "Ответственный — Департамент по продвижению программ и регионального развития."
  )
  for ($i = 0; $i -lt $lines.Count; $i++) { $ws.Cells.Item($i + 1, 1).Value2 = $lines[$i] }
  $ws.Cells.Item(1,1).Font.Bold = $true; $ws.Cells.Item(1,1).Font.Size = 13
  $ws.Range("A16:A18").Font.Bold = $true
  $ws.Columns.Item(1).ColumnWidth = 120

  # --- Сводка ---
  $ws = $wb.Worksheets.Item(2); $ws.Name = "Сводка"
  Head $ws @("Показатель", "Значение")
  $sv = @(@("Дата актуальности данных", $data.date), @("Участников всего (2026)", $data.kpi.participants), @("Регионов охвата", $data.kpi.regions))
  for ($i = 0; $i -lt 3; $i++) {
    $ws.Cells.Item($i + 2, 1).Value2 = $sv[$i][0]
    SetText $ws.Cells.Item($i + 2, 2) $sv[$i][1]
    $ws.Cells.Item($i + 2, 2).Interior.Color = $YELLOW
  }
  $ws.Columns.Item(1).ColumnWidth = 32; $ws.Columns.Item(2).ColumnWidth = 20

  # --- Программы ---
  $sections = @("Название","Статус","Описание","Охват","География","Этап","Вывод","Ход реализации","Для прессы","Итоги года") -join $sep
  for ($k = 0; $k -lt $n; $k++) {
    $p = $data.programs[$k]
    $ws = $wb.Worksheets.Item($k + 3)
    # Имя листа в Excel — не длиннее 31 символа: первые два слова + последнее («Центр поддержки бизнеса»)
    $short = $p.name -replace '[\[\]:*?/\\]', ''
    if ($short.Length -gt 27) {
      $w = $short -split '\s+'
      $short = (@($w[0], $w[1], $w[-1]) -join ' ')
      if ($short.Length -gt 27) { $short = $short.Substring(0, 27).Trim() }
    }
    $ws.Name = ("{0:D2} {1}" -f ($k + 1), $short)
    Head $ws @("Раздел", "Показатель / этап / текст", "Факт / значение / статус этапа", "План / дата этапа")
    $rows = New-Object System.Collections.ArrayList
    [void]$rows.Add(@("Название", $p.name, "", ""))
    $st = $STATUS[$p.status]; if (-not $st) { $st = "" }
    [void]$rows.Add(@("Статус", $st, "", ""))
    [void]$rows.Add(@("Описание", $p.desc, "", ""))
    foreach ($r in $p.reach)  { [void]$rows.Add(@("Охват", $r[0], $r[1], $r[2])) }
    foreach ($g in $p.geo)    { [void]$rows.Add(@("География", $g[0], $g[1], "")) }
    foreach ($s in $p.stages) { [void]$rows.Add(@("Этап", $s.name, $STAGE[$s.state], $s.date)) }
    [void]$rows.Add(@("Вывод", $p.conclusion, "", ""))
    $prog = @($p.progress | Where-Object { $_ })
    foreach ($t in $prog) { [void]$rows.Add(@("Ход реализации", $t, "", "")) }
    if ($prog.Count -eq 0) { [void]$rows.Add(@("Ход реализации", "", "", "")) }
    [void]$rows.Add(@("Для прессы", [string]$p.press, "", ""))
    [void]$rows.Add(@("Итоги года", [string]$p.annual, "", ""))
    for ($i = 0; $i -lt $rows.Count; $i++) {
      $row = $rows[$i]; $ri = $i + 2; $sec = $row[0]
      $ws.Cells.Item($ri, 1).Value2 = $sec
      $cols = switch ($sec) { "Охват" {3} "География" {2} "Этап" {3} default {1} }
      for ($c = 1; $c -le $cols; $c++) {
        $cell = $ws.Cells.Item($ri, $c + 1)
        $v = [string]$row[$c]
        if ($v -ne "") {
          if ($sec -eq "Охват" -and $c -ge 2 -and $v -match '^\d[\d\s ]*(,\d+)?$') {
            $cell.Value2 = [double]($v -replace '[\s ]', '' -replace ',', '.')   # число — для диаграммы
          } else { SetText $cell $v }
        }
        $cell.Interior.Color = $YELLOW
      }
      if ($sec -eq "Статус") { [void]$ws.Cells.Item($ri, 2).Validation.Add(3, 1, 1, (@("Реализуется","Набирает обороты","Планируется") -join $sep)) }
      if ($sec -eq "Этап")   { [void]$ws.Cells.Item($ri, 3).Validation.Add(3, 1, 1, (@("Выполнен","В работе","Впереди") -join $sep)) }
    }
    [void]$ws.Range("A2:A80").Validation.Add(3, 2, 1, $sections)
    $ws.Columns.Item(1).ColumnWidth = 15
    $ws.Columns.Item(2).ColumnWidth = 70
    $ws.Columns.Item(3).ColumnWidth = 30
    $ws.Columns.Item(4).ColumnWidth = 18
    $ws.Range("B:C").WrapText = $true
    $ws.Range("A:D").VerticalAlignment = -4160
  }
  $wb.Worksheets.Item(1).Activate()
  if (Test-Path $Out) { Remove-Item $Out -Force }
  $wb.SaveAs($Out, 51)
  $wb.Close($false)
  "OK: $Out ($n программ, данные на $($data.date))"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
