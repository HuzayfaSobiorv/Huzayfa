Add-Type -AssemblyName System.IO.Compression.FileSystem

$srcList = 'C:\Users\PC_HP\Downloads\Xiaoshou_Qingdan_2026-08-08.xlsx'
$srcTP   = 'C:\Users\PC_HP\Downloads\Metalmart_Jadvallar_Jamlanma.xlsx'
$out     = Join-Path $PSScriptRoot 'BOT_Yuklatish_Xiaoshou_2026-08-08.xlsx'

function Get-ZipText($zip, [string]$name) {
    $entry = $zip.Entries | Where-Object FullName -eq $name | Select-Object -First 1
    if (-not $entry) { return $null }
    $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-ColIndex([string]$cellRef) {
    $letters = ([regex]::Match($cellRef, '^[A-Z]+')).Value
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1) }
    return $n - 1
}

function Read-XlsxRows([string]$path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $shared = @()
        $sharedXml = Get-ZipText $zip 'xl/sharedStrings.xml'
        if ($sharedXml) {
            [xml]$ss = $sharedXml
            foreach ($si in @($ss.sst.si)) { $shared += ((@($si.t) + @($si.r.t)) -join '') }
        }
        [xml]$sheet = Get-ZipText $zip 'xl/worksheets/sheet1.xml'
        $result = [Collections.Generic.List[object[]]]::new()
        foreach ($row in @($sheet.worksheet.sheetData.row)) {
            $last = -1
            foreach ($cell in @($row.c)) { $last = [Math]::Max($last, (Get-ColIndex $cell.r)) }
            if ($last -lt 0) { $result.Add(@()); continue }
            $values = New-Object object[] ($last + 1)
            foreach ($cell in @($row.c)) {
                $col = Get-ColIndex $cell.r
                if ($cell.t -eq 'inlineStr') { $values[$col] = ((@($cell.is.t) + @($cell.is.r.t)) -join '') }
                elseif ($cell.t -eq 's') { $values[$col] = $shared[[int]$cell.v] }
                elseif ($cell.v -ne $null) {
                    $raw = [string]$cell.v
                    $num = 0.0
                    $values[$col] = if ([double]::TryParse($raw, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$num)) { $num } else { $raw }
                }
            }
            $result.Add($values)
        }
        return $result
    } finally { $zip.Dispose() }
}

function Cell([object[]]$row, [int]$index) { if ($index -lt $row.Count -and $null -ne $row[$index]) { return $row[$index] }; return '' }
function Number([object]$value) {
    $out = 0.0
    [void][double]::TryParse(([string]$value).Replace(',', '.'), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$out)
    return $out
}
function DecComma([object]$value) { return ([double](Number $value)).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture).Replace('.', ',') }
function Round10Up([object]$value) { return [int]([Math]::Ceiling((Number $value) / 10.0) * 10) }
function Xml([string]$value) { return [Security.SecurityElement]::Escape($value) }
function ColName([int]$zeroIndex) {
    $n = $zeroIndex + 1; $s = ''
    while ($n -gt 0) { $r = ($n - 1) % 26; $s = [char]([int][char]'A' + $r) + $s; $n = [Math]::Floor(($n - 1) / 26) }
    return $s
}
function Sheet-Xml([Collections.Generic.List[object[]]]$rows) {
    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    $rn = 1
    foreach ($row in $rows) {
        [void]$sb.Append(('<row r="{0}">' -f $rn))
        for ($i = 0; $i -lt $row.Count; $i++) {
            $value = $row[$i]
            if ($null -eq $value -or [string]$value -eq '') { continue }
            $ref = (ColName $i) + $rn
            if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) {
                $v = ([double]$value).ToString('0.###############', [Globalization.CultureInfo]::InvariantCulture)
                [void]$sb.Append(('<c r="{0}"><v>{1}</v></c>' -f $ref, $v))
            } else {
                [void]$sb.Append(('<c r="{0}" t="inlineStr"><is><t>{1}</t></is></c>' -f $ref, (Xml ([string]$value))))
            }
        }
        [void]$sb.Append('</row>'); $rn++
    }
    [void]$sb.Append('</sheetData></worksheet>')
    return $sb.ToString()
}
function Add-ZipText($zip, [string]$name, [string]$content) {
    $entry = $zip.CreateEntry($name)
    $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
    try { $writer.Write($content) } finally { $writer.Dispose() }
}

$rawList = Read-XlsxRows $srcList
$rawTP   = Read-XlsxRows $srcTP
$botRows = [Collections.Generic.List[object[]]]::new()
$botRows.Add(@('Tovar nomi', 'Buyurtma', 'Tayyor', '1 dona vazni (kg)', 'Izoh'))

# Xiaoshou: buyurtma miqdori berilmagan, faqat mavjud miqdor "Tayyor"ga yoziladi.
for ($i = 3; $i -lt $rawList.Count; $i++) {
    $r = $rawList[$i]; $mat = [string](Cell $r 1); $surfaceRaw = [string](Cell $r 2); $spec = [string](Cell $r 3); $qty = Number (Cell $r 6)
    $m = [regex]::Match($spec, '^\s*([0-9.]+)\s*\*\s*([0-9.]+)\s*\*\s*([0-9.]+)\s*$')
    if ($qty -le 0 -or -not $m.Success) { continue }
    $markaM = [regex]::Match($mat, '(201|304|316|321|430)'); $marka = if ($markaM.Success) { $markaM.Value } else { '201' }
    $surface = if ($surfaceRaw -like '*钛金*') { 'Голд' } elseif ($surfaceRaw -like '*精磨8K*' -or $surfaceRaw -like '*8K*') { 'Глянцевый' } elseif ($surfaceRaw -like '*砂板*') { 'Матовый' } else { 'Матовый' }
    $name = "Лист-$(DecComma $m.Groups[1].Value) ($(Round10Up $m.Groups[2].Value)х$(Round10Up $m.Groups[3].Value)) ($surface) ($marka марка)"
    $botRows.Add(@($name, $null, $qty, $null, "Xiaoshou row $($i + 1) | xom: $mat | $surfaceRaw | $spec"))
}

# Jamlanma: Ombor ustunidagi haqiqiy mavjud son "Tayyor"ga yoziladi.
for ($i = 1; $i -lt $rawTP.Count; $i++) {
    $r = $rawTP[$i]; $spec = [string](Cell $r 1); $realThick = [string](Cell $r 2); $lenRaw = [string](Cell $r 3); $unitWeight = Number (Cell $r 4); $qty = Number (Cell $r 7)
    if ($qty -le 0) { continue }
    $lm = [regex]::Match($lenRaw, '([0-9]+(?:[.,][0-9]+)?)'); $len = if ($lm.Success) { DecComma $lm.Groups[1].Value } else { '' }
    $name = $null
    $tm = [regex]::Match($spec, 'Φ\s*([0-9]+)\s*cT\s*([0-9.,]+)')
    $pm = [regex]::Match($spec, 'KB\s*([0-9]+)\s*[xх]\s*([0-9]+)\s*CT\s*([0-9.,]+)')
    if ($tm.Success) { $name = "Ф-$($tm.Groups[1].Value) ст $(DecComma $tm.Groups[2].Value) ($len м) (201 марка)" }
    elseif ($pm.Success) { $name = "Пр. $($pm.Groups[1].Value)х$($pm.Groups[2].Value) ст $(DecComma $pm.Groups[3].Value) ($len м) (201 марка)" }
    if ($name) { $botRows.Add(@($name, $null, $qty, $unitWeight, "Jamlanma row $($i + 1) | xom spec: $spec | real qalinlik: $realThick | uzunlik: $lenRaw")) }
}

if (Test-Path $out) { Remove-Item -LiteralPath $out -Force }
$zip = [IO.Compression.ZipFile]::Open($out, [IO.Compression.ZipArchiveMode]::Create)
try {
    Add-ZipText $zip '[Content_Types].xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    Add-ZipText $zip '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    Add-ZipText $zip 'xl/workbook.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="BOT_Yuklash" sheetId="1" r:id="rId1"/><sheet name="Xom_Xiaoshou" sheetId="2" r:id="rId2"/><sheet name="Xom_Jamlanma" sheetId="3" r:id="rId3"/></sheets></workbook>'
    Add-ZipText $zip 'xl/_rels/workbook.xml.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    Add-ZipText $zip 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>'
    Add-ZipText $zip 'xl/worksheets/sheet1.xml' (Sheet-Xml $botRows)
    Add-ZipText $zip 'xl/worksheets/sheet2.xml' (Sheet-Xml $rawList)
    Add-ZipText $zip 'xl/worksheets/sheet3.xml' (Sheet-Xml $rawTP)
} finally { $zip.Dispose() }
Write-Output "CREATED=$out"
Write-Output "BOT_ROWS=$($botRows.Count - 1)"
