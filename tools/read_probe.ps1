# Read back the geometry PowerPoint reports for the probe deck built by
# tools/probe_metrics.py (ASCII only: Windows PowerShell 5.1 reads .ps1 as ANSI,
# so non-ASCII comments can break parsing).
#
# usage: powershell -NoProfile -File tools/read_probe.ps1 [probe.pptx]
#
# NOTE: the per-glyph divisor below (16) must stay in sync with GLYPHS in
# tools/probe_metrics.py.
param(
    [string]$Path = "D:\OpenClaw\2026-10-13-21-23-42\tools\_probe\probe_metrics.pptx"
)
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($Path, $true, $false, $false)
$sl = $pres.Slides.Item(1)
$ptcm = 28.3464567

$A = @{}
$B = @{}
foreach ($sh in $sl.Shapes) {
    $n = $sh.Name
    if ($n -like "A_*") {
        $key = ($n -replace "^A_", "" -replace "_n\d+$", "")
        $nn = [int]($n -replace "^.*_n", "")
        $tr = $sh.TextFrame.TextRange
        if (-not $A.ContainsKey($key)) { $A[$key] = @{} }
        $A[$key][$nn] = $tr.BoundHeight
    } elseif ($n -like "B_*") {
        $sz = [int]($n -replace "^B_", "")
        $B[$sz] = $sh.TextFrame.TextRange.BoundWidth
    }
}

"=== line pitch (font 3.45pt; pitch = (H(n2)-H(n1))/(n2-n1)) ==="
"{0,-10} {1,10} {2,10} {3,10} {4,10}" -f "variant", "H(2)pt", "H(6)pt", "pitch pt", "pitch cm"
foreach ($k in ($A.Keys | Sort-Object)) {
    $h1 = $A[$k][2]
    $h2 = $A[$k][6]
    $pitch = ($h2 - $h1) / 4.0
    "{0,-10} {1,10:N3} {2,10:N3} {3,10:N4} {4,10:N4}" -f $k, $h1, $h2, $pitch, ($pitch / $ptcm)
}

""
"=== advance width (one line of 16 full-width CJK glyphs) ==="
"{0,8} {1,12} {2,12} {3,12}" -f "size pt", "line pt", "per glyph", "per glyph cm"
foreach ($sz in ($B.Keys | Sort-Object)) {
    $w = $B[$sz] / 16.0
    "{0,8:N2} {1,12:N3} {2,12:N4} {3,12:N5}" -f ($sz / 100.0), $B[$sz], $w, ($w / $ptcm)
}

$pres.Close(); $ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
