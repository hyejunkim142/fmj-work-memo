# 업무지시 입력창
# - 새 업무: 입력하면 같은 폴더의 '업무지시.txt'에 한 줄씩 저장 (메모장으로 열어 볼 수 있음)
# - 기존 업무: 같은 폴더의 가장 최근 '결재_미마감_업무지시_*.xlsx'를 읽어 목록에 함께 보여줌
# - 완료: 메모 업무는 업무지시.txt에서 [ ] → [v],  결재 업무는 '완료기록.txt'에 기록하고 목록에서 숨김
#   (엑셀은 매일 자동으로 새로 만들어지므로 엑셀 안에는 표시하지 않음)
# - 이 파일은 반드시 'UTF-8 (BOM 포함)'으로 저장해야 합니다. BOM이 없으면 Windows PowerShell 5.1이 한글을 깨뜨립니다.

$ErrorActionPreference = 'Stop'

if ($PSScriptRoot) { $baseDir = $PSScriptRoot }
else { $baseDir = Join-Path ([Environment]::GetFolderPath('Desktop')) '업무지시' }

$dataFile  = Join-Path $baseDir '업무지시.txt'
$doneFile  = Join-Path $baseDir '완료기록.txt'
$logFile   = Join-Path $baseDir '입력창_오류기록.txt'
$excelMask = '결재_미마감_업무지시_*.xlsx'
$utf8Bom   = New-Object System.Text.UTF8Encoding($true)
$header    = @(
    '# 업무지시 목록 - 입력창이 자동으로 관리합니다.',
    '# [ ] = 진행중, [v] = 완료.  한 줄이 한 건입니다. 메모장에서 직접 고쳐도 됩니다.',
    ''
)

function Write-ErrorLog([string]$msg) {
    try {
        $line = '{0}  {1}{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg, "`r`n"
        [IO.File]::AppendAllText($logFile, $line, $utf8Bom)
    } catch { }
}

function Clean-Text([string]$s) {
    if ($null -eq $s) { return '' }
    return (($s -replace '\s*\r?\n\s*', ' / ') -replace '\t', ' ').Trim()
}

# ---------------- 메모 업무 (업무지시.txt) ----------------

function Read-Lines {
    if (-not (Test-Path -LiteralPath $dataFile)) { return ,([string[]]$header) }
    return ,([IO.File]::ReadAllLines($dataFile, $utf8Bom))
}

function Save-Lines([string[]]$lines) {
    # 임시 파일에 먼저 쓰고 덮어써서, 저장 도중 오류가 나도 기존 파일이 깨지지 않게 함
    $tmp = $dataFile + '.tmp'
    [IO.File]::WriteAllLines($tmp, $lines, $utf8Bom)
    $lastErr = $null
    for ($i = 0; $i -lt 5; $i++) {
        try {
            [IO.File]::Copy($tmp, $dataFile, $true)
            [IO.File]::Delete($tmp)
            return
        } catch {
            $lastErr = $_
            Start-Sleep -Milliseconds 100
        }
    }
    throw $lastErr
}

function Append-Text([string]$path, [string]$text) {
    # 파일 끝에 덧붙이기. 마지막 줄바꿈이 빠져 있으면 앞 줄에 붙지 않게 줄바꿈부터 넣음
    if (Test-Path -LiteralPath $path) {
        $fs = [IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
        try {
            if ($fs.Length -gt 0) {
                [void]$fs.Seek(-1, 'End')
                if ($fs.ReadByte() -ne 10) { $text = "`r`n" + $text }
            }
        } finally { $fs.Dispose() }
    }
    $lastErr = $null
    for ($i = 0; $i -lt 5; $i++) {
        try { [IO.File]::AppendAllText($path, $text, $utf8Bom); return }
        catch { $lastErr = $_; Start-Sleep -Milliseconds 100 }
    }
    throw $lastErr
}

function Get-OpenTasks {
    $result = [Collections.Generic.List[string]]::new()
    foreach ($l in (Read-Lines)) {
        if ($l.StartsWith('[ ] ')) { $result.Add($l) }
    }
    return ,$result.ToArray()
}

function New-TaskLine([string]$content, [string]$owner, $due, [bool]$urgent) {
    $text = Clean-Text $content
    if ($urgent) { $text = '[긴급] ' + $text }
    $o = $owner.Trim(); if (-not $o) { $o = '-' }
    $d = '-'; if ($due) { $d = $due.ToString('yyyy-MM-dd') }
    return '[ ] 등록 {0} | 마감 {1} | 담당 {2} | {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $d, $o, $text
}

function Add-Task([string]$line) {
    $text = $line + "`r`n"
    if (-not (Test-Path -LiteralPath $dataFile)) { $text = ($header -join "`r`n") + "`r`n" + $text }
    Append-Text $dataFile $text
}

function Complete-Task([string]$line) {
    $lines = [string[]](Read-Lines)
    $idx = [Array]::IndexOf($lines, $line)
    if ($idx -lt 0) { return $false }   # 그 사이 메모장 등에서 바뀐 경우
    $lines[$idx] = '[v]' + $line.Substring(3) + ' | 완료 ' + (Get-Date -Format 'yyyy-MM-dd HH:mm')
    Save-Lines $lines
    return $true
}

# ---------------- 결재 업무 (엑셀, 읽기 전용) ----------------

function Get-LatestExcel {
    return Get-ChildItem -LiteralPath $baseDir -Filter $excelMask -File -ErrorAction SilentlyContinue |
        Where-Object { -not $_.Name.StartsWith('~$') } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

function Read-XlsxRows([string]$path) {
    # 엑셀 없이 xlsx 안의 XML을 직접 읽음 (엑셀로 파일을 열어 둔 상태에서도 읽힘). 첫 번째 시트만 읽음
    Add-Type -AssemblyName System.IO.Compression
    $fs = [IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
    try {
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Read)
        $getXml = {
            param([string]$name)
            $e = $zip.GetEntry($name)
            if (-not $e) { return $null }
            $r = New-Object IO.StreamReader($e.Open())
            try { return [xml]$r.ReadToEnd() } finally { $r.Dispose() }
        }

        $strings = [Collections.Generic.List[string]]::new()
        $ss = & $getXml 'xl/sharedStrings.xml'
        if ($ss) {
            foreach ($si in $ss.DocumentElement.ChildNodes) {
                $parts = $si.SelectNodes('*[local-name()="t"] | *[local-name()="r"]/*[local-name()="t"]')
                $strings.Add((($parts | ForEach-Object { $_.InnerText }) -join ''))
            }
        }

        # 날짜 서식이 걸린 칸 찾기 (엑셀은 날짜를 숫자로 저장함)
        $isDate = @{}
        $st = & $getXml 'xl/styles.xml'
        if ($st) {
            $custom = @{}
            foreach ($nf in @($st.styleSheet.numFmts.numFmt)) {
                if ($nf) { $custom[[int]$nf.numFmtId] = [string]$nf.formatCode }
            }
            $i = 0
            foreach ($xf in @($st.styleSheet.cellXfs.xf)) {
                $id = [int]$xf.numFmtId
                $d = ($id -ge 14 -and $id -le 22) -or ($id -ge 45 -and $id -le 47)
                if (-not $d -and $custom.ContainsKey($id)) {
                    $code = $custom[$id] -replace '"[^"]*"', '' -replace '\[[^\]]*\]', ''
                    $d = $code -match '[ymdhs]'
                }
                $isDate[$i] = $d
                $i++
            }
        }

        $wb = & $getXml 'xl/workbook.xml'
        $first = @($wb.workbook.sheets.sheet)[0]
        $rid = $first.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $rels = & $getXml 'xl/_rels/workbook.xml.rels'
        $target = (@($rels.Relationships.Relationship) | Where-Object { $_.Id -eq $rid } | Select-Object -First 1).Target
        if ($target.StartsWith('/')) { $sheetPath = $target.Substring(1) } else { $sheetPath = 'xl/' + $target }
        $sh = & $getXml $sheetPath

        $inv = [Globalization.CultureInfo]::InvariantCulture
        $rows = [Collections.Generic.List[object]]::new()
        foreach ($row in @($sh.worksheet.sheetData.row)) {
            if (-not $row) { continue }
            $cells = @{}
            $max = -1
            $auto = 0
            foreach ($c in @($row.c)) {
                if (-not $c) { continue }
                $col = $auto
                if ($c.r -match '^([A-Z]+)') {
                    $col = 0
                    foreach ($ch in $Matches[1].ToCharArray()) { $col = $col * 26 + ([int]$ch - 64) }
                    $col--
                }
                $auto = $col + 1
                $vNode = $c.SelectSingleNode('*[local-name()="v"]')
                $v = $null; if ($vNode) { $v = $vNode.InnerText }
                $t = $c.GetAttribute('t')
                $val = ''
                if ($t -eq 's' -and $v) { $val = $strings[[int]$v] }
                elseif ($t -eq 'inlineStr') { $n = $c.SelectSingleNode('*[local-name()="is"]'); if ($n) { $val = $n.InnerText } }
                elseif ($t -eq 'b') { if ($v -eq '1') { $val = 'TRUE' } else { $val = 'FALSE' } }
                elseif ($v) {
                    $val = $v
                    $s = $c.GetAttribute('s')
                    $num = 0.0
                    if ($s -and $isDate[[int]$s] -and [double]::TryParse($v, [Globalization.NumberStyles]::Float, $inv, [ref]$num)) {
                        try {
                            $dt = [DateTime]::FromOADate($num)
                            if ($dt.TimeOfDay.TotalSeconds -ge 1) { $val = $dt.ToString('yyyy-MM-dd HH:mm') }
                            else { $val = $dt.ToString('yyyy-MM-dd') }
                        } catch { }
                    }
                }
                $cells[$col] = Clean-Text $val
                if ($col -gt $max) { $max = $col }
            }
            if ($max -lt 0) { continue }
            $arr = New-Object string[] ($max + 1)
            for ($k = 0; $k -le $max; $k++) { if ($cells.ContainsKey($k)) { $arr[$k] = $cells[$k] } else { $arr[$k] = '' } }
            if ((($arr -join '').Trim()).Length -gt 0) { $rows.Add($arr) }
        }
        return ,$rows.ToArray()
    } finally {
        $fs.Dispose()
    }
}

function Get-DoneKeys {
    $set = [Collections.Generic.HashSet[string]]::new()
    if (Test-Path -LiteralPath $doneFile) {
        foreach ($l in [IO.File]::ReadAllLines($doneFile, $utf8Bom)) {
            $p = $l.Split("`t")
            if ($p.Length -ge 2 -and $p[1]) { [void]$set.Add($p[1]) }
        }
    }
    return ,$set
}

function Get-ExcelItems {
    # 반환: @{ File = 파일정보; Items = 미완료 항목들; DoneCount = 숨긴 건수 }
    $file = Get-LatestExcel
    $items = [Collections.Generic.List[object]]::new()
    $hidden = 0
    if ($file) {
        $rows = Read-XlsxRows $file.FullName
        if ($rows.Length -gt 1) {
            $head = $rows[0]
            # '번호/ID/문서/코드'가 들어간 칸이 있으면 그 칸 값으로 같은 건인지 구분, 없으면 줄 전체로 구분
            $keyCol = -1
            for ($k = 0; $k -lt $head.Length; $k++) {
                if ($head[$k] -match '번호|ID|문서|코드') { $keyCol = $k; break }
            }
            $done = Get-DoneKeys
            for ($r = 1; $r -lt $rows.Length; $r++) {
                $row = $rows[$r]
                $text = (($row | Where-Object { $_ }) -join ' | ')
                $key = $text
                if ($keyCol -ge 0 -and $keyCol -lt $row.Length -and $row[$keyCol]) { $key = $head[$keyCol] + '=' + $row[$keyCol] }
                if ($done.Contains($key)) { $hidden++; continue }
                $items.Add([pscustomobject]@{ Source = '결재'; Key = $key; Text = $text; Display = '[결재] ' + $text })
            }
        }
    }
    return @{ File = $file; Items = $items.ToArray(); DoneCount = $hidden }
}

function Complete-ExcelItem($item) {
    $line = '{0}{1}{2}{1}{3}{4}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), "`t", $item.Key, $item.Text, "`r`n"
    Append-Text $doneFile $line
}

# 테스트에서 함수만 불러 쓸 때는 창을 띄우지 않음
if ($env:TASK_INPUT_NO_UI -eq '1') { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# 입력창이 두 개 뜨지 않게 막음
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, 'Local\FMJ_TaskInput', [ref]$createdNew)
if (-not $createdNew) {
    [System.Windows.Forms.MessageBox]::Show('입력창이 이미 열려 있습니다.', '업무지시') | Out-Null
    return
}

try {
    if (-not (Test-Path -LiteralPath $baseDir)) { New-Item -ItemType Directory -Path $baseDir | Out-Null }

    $W = 700
    $font = New-Object System.Drawing.Font('맑은 고딕', 10)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = '업무지시 입력'
    $form.Font = $font
    $form.ClientSize = New-Object System.Drawing.Size($W, 640)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox = $false
    $form.KeyPreview = $true

    function Add-Label([string]$text, [int]$x, [int]$y) {
        $lb = New-Object System.Windows.Forms.Label
        $lb.Text = $text; $lb.AutoSize = $true
        $lb.Location = New-Object System.Drawing.Point($x, $y)
        $form.Controls.Add($lb)
        return $lb
    }

    [void](Add-Label '업무 내용' 15 12)
    $txtContent = New-Object System.Windows.Forms.TextBox
    $txtContent.Multiline = $true
    $txtContent.ScrollBars = 'Vertical'
    $txtContent.AcceptsReturn = $true
    $txtContent.Location = New-Object System.Drawing.Point(15, 35)
    $txtContent.Size = New-Object System.Drawing.Size(($W - 30), 90)
    $form.Controls.Add($txtContent)

    [void](Add-Label '담당자' 15 138)
    $txtOwner = New-Object System.Windows.Forms.TextBox
    $txtOwner.Location = New-Object System.Drawing.Point(75, 135)
    $txtOwner.Size = New-Object System.Drawing.Size(120, 25)
    $form.Controls.Add($txtOwner)

    [void](Add-Label '마감' 210 138)
    $dtDue = New-Object System.Windows.Forms.DateTimePicker
    $dtDue.Format = 'Custom'
    $dtDue.CustomFormat = 'yyyy-MM-dd (ddd)'
    $dtDue.ShowCheckBox = $true
    $dtDue.Checked = $false
    $dtDue.Location = New-Object System.Drawing.Point(250, 135)
    $dtDue.Size = New-Object System.Drawing.Size(175, 25)
    $form.Controls.Add($dtDue)

    $chkUrgent = New-Object System.Windows.Forms.CheckBox
    $chkUrgent.Text = '긴급'
    $chkUrgent.AutoSize = $true
    $chkUrgent.Location = New-Object System.Drawing.Point(440, 137)
    $form.Controls.Add($chkUrgent)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = '저장  (Ctrl+Enter)'
    $btnSave.Location = New-Object System.Drawing.Point(15, 170)
    $btnSave.Size = New-Object System.Drawing.Size(($W - 30), 36)
    $form.Controls.Add($btnSave)

    $lblList = Add-Label '진행중 업무  (더블클릭 = 완료)' 15 223
    [void](Add-Label '찾기' 420 223)
    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(460, 220)
    $txtSearch.Size = New-Object System.Drawing.Size(($W - 475), 25)
    $form.Controls.Add($txtSearch)

    $lstOpen = New-Object System.Windows.Forms.ListBox
    $lstOpen.HorizontalScrollbar = $true
    $lstOpen.Location = New-Object System.Drawing.Point(15, 250)
    $lstOpen.Size = New-Object System.Drawing.Size(($W - 30), 300)
    $form.Controls.Add($lstOpen)

    $btnDone = New-Object System.Windows.Forms.Button
    $btnDone.Text = '선택 항목 완료'
    $btnDone.Location = New-Object System.Drawing.Point(15, 558)
    $btnDone.Size = New-Object System.Drawing.Size(160, 32)
    $form.Controls.Add($btnDone)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = '새로고침'
    $btnRefresh.Location = New-Object System.Drawing.Point(180, 558)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 32)
    $form.Controls.Add($btnRefresh)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Text = '메모장으로 열기'
    $btnOpen.Location = New-Object System.Drawing.Point(($W - 175), 558)
    $btnOpen.Size = New-Object System.Drawing.Size(160, 32)
    $form.Controls.Add($btnOpen)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.AutoSize = $false
    $lblStatus.Location = New-Object System.Drawing.Point(15, 600)
    $lblStatus.Size = New-Object System.Drawing.Size(($W - 30), 30)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblStatus)

    $script:allItems = [Collections.Generic.List[object]]::new()
    $script:shown = @()
    $script:excelNote = ''

    function Set-Status([string]$msg, [bool]$isError) {
        $lblStatus.Text = $msg
        if ($isError) { $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick }
        else { $lblStatus.ForeColor = [System.Drawing.Color]::SeaGreen }
    }

    function Show-Failure([string]$what, $err) {
        $detail = $err.Exception.Message
        Write-ErrorLog ('{0}: {1}' -f $what, $detail)
        Set-Status ('{0} 실패 - 오류기록 파일을 확인하세요.' -f $what) $true
        [System.Windows.Forms.MessageBox]::Show(
            ("{0}에 실패했습니다.`n`n{1}" -f $what, $detail), '업무지시', 'OK', 'Error') | Out-Null
    }

    function New-MemoItem([string]$line) {
        return [pscustomobject]@{ Source = '메모'; Key = $line; Text = $line; Display = '[메모] ' + $line.Substring(4) }
    }

    function Apply-Filter {
        $words = @($txtSearch.Text.Trim() -split '\s+' | Where-Object { $_ })
        $list = [Collections.Generic.List[object]]::new()
        foreach ($it in $script:allItems) {
            $ok = $true
            foreach ($w in $words) {
                if ($it.Display.IndexOf($w, [StringComparison]::OrdinalIgnoreCase) -lt 0) { $ok = $false; break }
            }
            if ($ok) { $list.Add($it) }
        }
        $script:shown = $list.ToArray()
        $lstOpen.BeginUpdate()
        $lstOpen.Items.Clear()
        foreach ($it in $script:shown) { [void]$lstOpen.Items.Add($it.Display) }
        $lstOpen.EndUpdate()
        if ($words.Count -gt 0) { $lblList.Text = '진행중 업무  (찾은 결과 {0} / 전체 {1}건)' -f $script:shown.Length, $script:allItems.Count }
        else { $lblList.Text = '진행중 업무  (전체 {0}건, 더블클릭 = 완료)' -f $script:allItems.Count }
    }

    function Update-List {
        $script:allItems = [Collections.Generic.List[object]]::new()
        $msgs = @()
        $isErr = $false
        try {
            $ex = Get-ExcelItems
            foreach ($it in $ex.Items) { $script:allItems.Add($it) }
            if ($ex.File) { $msgs += ('결재: {0} ({1}건, 완료해서 숨김 {2}건)' -f $ex.File.Name, $ex.Items.Length, $ex.DoneCount) }
            else { $msgs += '결재 엑셀 파일 없음'; $isErr = $true }
        } catch {
            Write-ErrorLog ('엑셀 읽기: ' + $_.Exception.Message)
            $msgs += '결재 엑셀 읽기 실패 (오류기록 확인)'; $isErr = $true
        }
        try {
            $memo = Get-OpenTasks
            foreach ($t in $memo) { $script:allItems.Add((New-MemoItem $t)) }
            $msgs += ('메모: {0}건' -f $memo.Length)
        } catch {
            Write-ErrorLog ('메모 읽기: ' + $_.Exception.Message)
            $msgs += '메모 읽기 실패'; $isErr = $true
        }
        $script:excelNote = $msgs -join '  /  '
        Apply-Filter
        Set-Status $script:excelNote $isErr
    }

    $doSave = {
        $content = $txtContent.Text.Trim()
        if (-not $content) {
            Set-Status '업무 내용을 입력하세요.' $true
            $txtContent.Focus() | Out-Null
            return
        }
        $due = $null
        if ($dtDue.Checked) { $due = $dtDue.Value }
        try {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $line = New-TaskLine $content $txtOwner.Text $due $chkUrgent.Checked
            Add-Task $line
            $sw.Stop()
            $script:allItems.Add((New-MemoItem $line))
            Apply-Filter
            $txtContent.Clear()
            $chkUrgent.Checked = $false
            $dtDue.Checked = $false
            Set-Status ('저장했습니다. ({0}, {1:N2}초)' -f (Get-Date -Format 'HH:mm:ss'), $sw.Elapsed.TotalSeconds) $false
            $txtContent.Focus() | Out-Null
        } catch {
            Show-Failure '저장' $_
        }
    }

    $doComplete = {
        $idx = $lstOpen.SelectedIndex
        if ($idx -lt 0 -or $idx -ge $script:shown.Length) { Set-Status '완료할 항목을 먼저 고르세요. (찾기 칸에 글자를 넣으면 목록이 줄어듭니다)' $true; return }
        $item = $script:shown[$idx]
        try {
            if ($item.Source -eq '결재') {
                Complete-ExcelItem $item
                [void]$script:allItems.Remove($item)
                Apply-Filter
                Set-Status ('완료 처리했습니다: ' + $item.Text) $false
            } elseif (Complete-Task $item.Key) {
                [void]$script:allItems.Remove($item)
                Apply-Filter
                Set-Status ('완료 처리했습니다: ' + $item.Display) $false
            } else {
                Update-List
                Set-Status '파일 내용이 바뀌어 항목을 찾지 못했습니다. 목록을 새로 불러왔습니다.' $true
            }
            if ($lstOpen.Items.Count -gt 0) { $lstOpen.SelectedIndex = [Math]::Min($idx, $lstOpen.Items.Count - 1) }
        } catch {
            Show-Failure '완료 처리' $_
        }
    }

    $btnSave.Add_Click($doSave)
    $btnDone.Add_Click($doComplete)
    $lstOpen.Add_DoubleClick($doComplete)
    $txtSearch.Add_TextChanged({ Apply-Filter })
    $btnRefresh.Add_Click({ Update-List })
    $btnOpen.Add_Click({
        try {
            if (-not (Test-Path -LiteralPath $dataFile)) { Save-Lines $header }
            Start-Process notepad.exe -ArgumentList ('"{0}"' -f $dataFile)
        } catch { Show-Failure '메모장 열기' $_ }
    })
    $form.Add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; & $doSave }
        elseif ($e.Control -and $e.KeyCode -eq 'F') { $e.SuppressKeyPress = $true; $txtSearch.Focus() | Out-Null; $txtSearch.SelectAll() }
        elseif ($e.KeyCode -eq 'Enter' -and $lstOpen.Focused) { $e.SuppressKeyPress = $true; & $doComplete }
        elseif ($e.KeyCode -eq 'Escape') {
            if ($txtSearch.Text) { $txtSearch.Clear() } else { $form.Close() }
        }
    })
    $form.Add_Shown({ Update-List; $txtContent.Focus() | Out-Null; $form.Activate() })

    [void]$form.ShowDialog()
} catch {
    # 창이 숨김으로 실행되므로, 예상 못 한 오류도 반드시 화면과 기록 파일에 남김
    Write-ErrorLog ('실행 오류: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
    [System.Windows.Forms.MessageBox]::Show(
        ("입력창 실행 중 오류가 났습니다.`n`n{0}`n`n기록: {1}" -f $_.Exception.Message, $logFile),
        '업무지시', 'OK', 'Error') | Out-Null
} finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
