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

# 모든 값이 정수이고 위에서부터 계속 커지는 칸 = 줄 순번 칸. 없으면 -1
function Get-SequenceColumn($rows) {
    if ($rows.Length -lt 2) { return -1 }
    $cols = ($rows | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    for ($k = 0; $k -lt $cols; $k++) {
        $prev = [int]::MinValue; $ok = $true
        for ($r = 1; $r -lt $rows.Length; $r++) {
            $v = 0
            if ($k -ge $rows[$r].Length -or -not [int]::TryParse($rows[$r][$k], [ref]$v) -or $v -le $prev) { $ok = $false; break }
            $prev = $v
        }
        if ($ok) { return $k }
    }
    return -1
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
            # 단, 1,2,3…처럼 줄 순번인 칸은 매일 바뀔 수 있으므로 구분 기준에서 뺌
            $seqCol = Get-SequenceColumn $rows
            $keyCol = -1
            for ($k = 0; $k -lt $head.Length; $k++) {
                if ($k -ne $seqCol -and $head[$k] -match '번호|ID|문서|코드') { $keyCol = $k; break }
            }
            $done = Get-DoneKeys
            for ($r = 1; $r -lt $rows.Length; $r++) {
                $row = $rows[$r]
                $text = (($row | Where-Object { $_ }) -join ' | ')
                $keyParts = for ($k = 0; $k -lt $row.Length; $k++) { if ($k -ne $seqCol -and $row[$k]) { $row[$k] } }
                $key = (@($keyParts) -join ' | ')
                if ($keyCol -ge 0 -and $keyCol -lt $row.Length -and $row[$keyCol]) { $key = $head[$keyCol] + '=' + $row[$keyCol] }
                if ($done.Contains($key)) { $hidden++; continue }
                $detail = for ($k = 0; $k -lt $row.Length; $k++) {
                    if (-not $row[$k]) { continue }
                    $name = if ($k -lt $head.Length -and $head[$k]) { $head[$k] } else { '{0}열' -f ($k + 1) }
                    $name + ' : ' + $row[$k].Replace(' / ', "`r`n    ")
                }
                $seq = if ($seqCol -ge 0 -and $seqCol -lt $row.Length) { [string]$row[$seqCol] } else { '' }
                $items.Add([pscustomobject]@{ Source = '결재'; Key = $key; Text = $text; Display = '[결재] ' + $text
                    Seq = $seq; Detail = (@($detail) -join "`r`n") })
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

    # 창 (가로 420 x 세로 630). 테두리를 끌어서 크기를 늘릴 수 있음
    # $Dy = 업무 내용 칸을 기본(44)보다 키운 만큼 아래 칸들이 내려감
    $Dy = 300   # 창 크기 = 스티커 메모 정도
    $W = 420; $H = 330 + $Dy
    $small = New-Object System.Drawing.Font('맑은 고딕', 8.5)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = '업무지시 입력'
    $form.Font = New-Object System.Drawing.Font('맑은 고딕', 9)
    $form.ClientSize = New-Object System.Drawing.Size($W, $H)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.KeyPreview = $true
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.AutoPopDelay = 20000; $tip.InitialDelay = 300

    function Add-Ctl($ctl, [int]$x, [int]$y, [int]$w, [int]$h, [string]$anchor) {
        $ctl.Location = New-Object System.Drawing.Point($x, $y)
        if ($w -gt 0) { $ctl.Size = New-Object System.Drawing.Size($w, $h) }
        if ($anchor) { $ctl.Anchor = $anchor }
        $form.Controls.Add($ctl)
        return $ctl
    }
    function New-Label([string]$text, [bool]$isSmall) {
        $lb = New-Object System.Windows.Forms.Label
        $lb.Text = $text; $lb.AutoSize = $true
        if ($isSmall) { $lb.Font = $small; $lb.ForeColor = [System.Drawing.Color]::DimGray }
        return $lb
    }
    function New-Button([string]$text) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text
        return $b
    }

    [void](Add-Ctl (New-Label '업무 내용  (Ctrl+Enter = 저장)' $true) 6 6 0 0 '')
    $txtContent = New-Object System.Windows.Forms.TextBox
    $txtContent.Multiline = $true
    $txtContent.ScrollBars = 'Vertical'
    $txtContent.AcceptsReturn = $true
    [void](Add-Ctl $txtContent 8 24 ($W - 16) (44 + $Dy) 'Top, Left, Right')

    [void](Add-Ctl (New-Label '담당' $false) 6 (78 + $Dy) 0 0 '')
    $txtOwner = New-Object System.Windows.Forms.TextBox
    [void](Add-Ctl $txtOwner 38 (75 + $Dy) 62 23 'Top, Left')

    $dtDue = New-Object System.Windows.Forms.DateTimePicker
    $dtDue.Format = 'Custom'
    $dtDue.CustomFormat = 'MM-dd (ddd)'
    $dtDue.ShowCheckBox = $true
    $dtDue.Checked = $false
    [void](Add-Ctl $dtDue 104 (75 + $Dy) 112 23 'Top, Left')
    $tip.SetToolTip($dtDue, '마감일 (체크해야 저장됨)')

    $chkUrgent = New-Object System.Windows.Forms.CheckBox
    $chkUrgent.Text = '긴급'
    $chkUrgent.AutoSize = $true
    [void](Add-Ctl $chkUrgent 222 (77 + $Dy) 0 0 'Top, Left')

    $btnSave = Add-Ctl (New-Button '저장') ($W - 68) (74 + $Dy) 60 25 'Top, Right'

    [void](Add-Ctl (New-Label '찾기' $false) 6 (110 + $Dy) 0 0 '')
    $txtSearch = New-Object System.Windows.Forms.TextBox
    [void](Add-Ctl $txtSearch 38 (107 + $Dy) ($W - 46 - 80) 23 'Top, Left, Right')
    $tip.SetToolTip($txtSearch, '글자를 넣으면 맞는 업무만 목록에 나옵니다. 순번(숫자)을 넣으면 그 업무 내용이 나옵니다 (Ctrl+F, Esc = 지우기)')
    $btnView = Add-Ctl (New-Button '내용보기') ($W - 82) (106 + $Dy) 74 25 'Top, Right'

    # 목록 영역 (접었다 폈다 함). 평소에는 입력칸 + 찾기 칸만 보이는 작은 창
    $TopH = 134 + $Dy; $StatusH = 26; $CollapsedH = $TopH + $StatusH
    $script:expanded = $true
    $script:expandedH = $H
    $pnlList = New-Object System.Windows.Forms.Panel
    $pnlList.Visible = $false
    [void](Add-Ctl $pnlList 0 $TopH $W ($H - $TopH - $StatusH) 'Top, Left, Right')
    $ph = $pnlList.Height
    function Add-PCtl($ctl, [int]$x, [int]$y, [int]$w, [int]$h, [string]$anchor) {
        $ctl.Location = New-Object System.Drawing.Point($x, $y)
        if ($w -gt 0) { $ctl.Size = New-Object System.Drawing.Size($w, $h) }
        if ($anchor) { $ctl.Anchor = $anchor }
        $pnlList.Controls.Add($ctl)
        return $ctl
    }

    $lblList = Add-PCtl (New-Label '' $true) 6 0 0 0 ''

    $lstOpen = New-Object System.Windows.Forms.ListBox
    $lstOpen.HorizontalScrollbar = $true
    $lstOpen.IntegralHeight = $false
    [void](Add-PCtl $lstOpen 8 18 ($W - 16) ($ph - 18 - 34) 'Top, Bottom, Left, Right')

    # 찾기 칸에 순번을 넣으면 목록 자리에 그 업무의 전체 내용을 보여줌
    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Multiline = $true; $txtDetail.ReadOnly = $true; $txtDetail.WordWrap = $true
    $txtDetail.ScrollBars = 'Vertical'; $txtDetail.BackColor = [System.Drawing.Color]::White; $txtDetail.Visible = $false
    [void](Add-PCtl $txtDetail 8 18 ($W - 16) ($ph - 18 - 34) 'Top, Bottom, Left, Right')
    $script:detailItem = $null

    $btnDone    = Add-PCtl (New-Button '완료') 8 ($ph - 30) 60 26 'Bottom, Left'
    $btnRefresh = Add-PCtl (New-Button '새로고침') 72 ($ph - 30) 70 26 'Bottom, Left'
    $btnOpen    = Add-PCtl (New-Button '메모장') ($W - 72) ($ph - 30) 64 26 'Bottom, Right'

    # 목록 펼치기/접기. 접으면 세로 크기를 고정하고, 펴면 마지막으로 쓰던 높이로 돌아감
    function Set-Expanded([bool]$open) {
        if ($open -eq $script:expanded -and $pnlList.Visible -eq $open) { return }
        $form.SuspendLayout()
        if ($open) {
            $form.MaximumSize = [System.Drawing.Size]::Empty
            $form.MinimumSize = [System.Drawing.Size]::Empty
            $form.ClientSize = New-Object System.Drawing.Size($form.ClientSize.Width, $script:expandedH)
            $pnlList.Size = New-Object System.Drawing.Size($form.ClientSize.Width, ($script:expandedH - $TopH - $StatusH))
            $pnlList.Anchor = 'Top, Bottom, Left, Right'
            $pnlList.Visible = $true
            $form.MinimumSize = New-Object System.Drawing.Size($form.Width, ($form.Height - $form.ClientSize.Height + $CollapsedH + 140))
            $btnView.Text = '내용닫기'
        } else {
            if ($script:expanded) { $script:expandedH = $form.ClientSize.Height }
            $pnlList.Visible = $false
            $pnlList.Anchor = 'Top, Left, Right'
            $form.MaximumSize = [System.Drawing.Size]::Empty
            $form.MinimumSize = [System.Drawing.Size]::Empty
            $form.ClientSize = New-Object System.Drawing.Size($form.ClientSize.Width, $CollapsedH)
            $form.MinimumSize = $form.Size
            # 가로는 제한 없음. 0을 넣으면 Windows가 창 폭을 최소 폭으로 줄여 버리므로 화면 최대 폭을 넣음
            $form.MaximumSize = New-Object System.Drawing.Size([System.Windows.Forms.SystemInformation]::MaxWindowTrackSize.Width, $form.Height)
            $btnView.Text = '내용보기'
        }
        $script:expanded = $open
        $form.ResumeLayout($true)
    }

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.AutoSize = $false
    $lblStatus.AutoEllipsis = $true
    $lblStatus.Font = $small
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    [void](Add-Ctl $lblStatus 6 ($H - 24) ($W - 12) 20 'Bottom, Left, Right')
    $lblStatus.Add_TextChanged({ $tip.SetToolTip($lblStatus, $lblStatus.Text) })

    # 창이 좁아 잘리는 내용은 마우스를 올리면 전체가 보이게 함
    $script:tipIndex = -1
    $lstOpen.Add_MouseMove({
        param($s, $e)
        $i = $lstOpen.IndexFromPoint($e.Location)
        if ($i -eq $script:tipIndex) { return }
        $script:tipIndex = $i
        if ($i -ge 0 -and $i -lt $script:shown.Length) { $tip.SetToolTip($lstOpen, $script:shown[$i].Display) }
        else { $tip.SetToolTip($lstOpen, '') }
    })

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
        return [pscustomobject]@{ Source = '메모'; Key = $line; Text = $line; Display = '[메모] ' + $line.Substring(4)
            Seq = ''; Detail = $line.Substring(4).Replace(' | ', "`r`n").Replace(' / ', "`r`n    ") }
    }

    function Apply-Filter {
        # 찾기 칸에 숫자만 있고 그 순번의 결재 항목이 있으면 목록 대신 내용을 보여줌
        $q = $txtSearch.Text.Trim()
        $script:detailItem = $null
        if ($q -match '^[0-9]+$') {
            foreach ($it in $script:allItems) { if ($it.Seq -eq $q) { $script:detailItem = $it; break } }
        }
        $lstOpen.Visible = ($null -eq $script:detailItem)
        $txtDetail.Visible = ($null -ne $script:detailItem)
        if ($script:detailItem) {
            $txtDetail.Text = '[' + $script:detailItem.Source + "]`r`n" + $script:detailItem.Detail
            $lblList.Text = '순번 {0} 업무 내용  ·  [완료] = 완료 처리' -f $q
            return
        }
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
        if ($words.Count -gt 0) { $lblList.Text = '찾은 결과 {0} / 전체 {1}건  ·  더블클릭 = 완료' -f $script:shown.Length, $script:allItems.Count }
        else { $lblList.Text = '진행중 {0}건  ·  더블클릭 = 완료' -f $script:allItems.Count }
    }

    function Update-List {
        $script:allItems = [Collections.Generic.List[object]]::new()
        $msgs = @()
        $isErr = $false
        try {
            $ex = Get-ExcelItems
            foreach ($it in $ex.Items) { $script:allItems.Add($it) }
            if ($ex.File) { $msgs += ('결재 {1}건 (완료 숨김 {2}) - {0}' -f $ex.File.Name, $ex.Items.Length, $ex.DoneCount) }
            else { $msgs += '결재 엑셀 파일 없음'; $isErr = $true }
        } catch {
            Write-ErrorLog ('엑셀 읽기: ' + $_.Exception.Message)
            $msgs += '결재 엑셀 읽기 실패 (오류기록 확인)'; $isErr = $true
        }
        try {
            $memo = Get-OpenTasks
            foreach ($t in $memo) { $script:allItems.Add((New-MemoItem $t)) }
            $msgs += ('메모 {0}건' -f $memo.Length)
        } catch {
            Write-ErrorLog ('메모 읽기: ' + $_.Exception.Message)
            $msgs += '메모 읽기 실패'; $isErr = $true
        }
        $script:excelNote = $msgs -join '  ·  '
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
        $item = $script:detailItem
        $fromDetail = ($null -ne $item)
        if (-not $fromDetail) {
            if ($idx -lt 0 -or $idx -ge $script:shown.Length) { Set-Status '완료할 항목을 먼저 고르세요.' $true; return }
            $item = $script:shown[$idx]
        }
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
            if ($fromDetail) { $txtSearch.Clear(); $txtSearch.Focus() | Out-Null }   # 번호로 찾아 완료했으면 찾기 칸을 비워 목록을 접음
            elseif ($lstOpen.Items.Count -gt 0) { $lstOpen.SelectedIndex = [Math]::Min([Math]::Max($idx, 0), $lstOpen.Items.Count - 1) }
        } catch {
            Show-Failure '완료 처리' $_
        }
    }

    $btnSave.Add_Click($doSave)
    $btnDone.Add_Click($doComplete)
    $lstOpen.Add_DoubleClick($doComplete)
    $txtSearch.Add_TextChanged({
        Apply-Filter
        # 찾기 칸이 비어 있으면 목록을 접고, 글자를 넣으면 펼침
        $has = [bool]$txtSearch.Text.Trim()
        if ($has -ne $script:expanded) { Set-Expanded $has }
    })
    $btnView.Add_Click({ Set-Expanded (-not $script:expanded) })
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
            if ($txtSearch.Text) { $txtSearch.Clear() }
            elseif ($script:expanded) { Set-Expanded $false }
            else { $form.Close() }
        }
    })
    $form.Add_Shown({ Update-List; $txtContent.Focus() | Out-Null; $form.Activate() })

    Set-Expanded $false
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
