# Dump per-shape geometry of one pptx to a numeric TSV (pure ASCII on purpose:
# Windows PowerShell 5.1 reads .ps1 as ANSI, so non-ASCII comments can break parsing).
# Bound* values are the RENDERER's own text ink box in points.
#
# columns: idx kind hasTextFrame left top width height boundLeft boundTop boundWidth boundHeight textLen nPara name
#
# usage: powershell -NoProfile -File tools/ppt_dump.ps1 <pptx> <out.tsv>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = "Stop"
$ppt = New-Object -ComObject PowerPoint.Application
$lines = New-Object System.Collections.Generic.List[string]
try {
    $pres = $ppt.Presentations.Open($Path, $true, $false, $false)
    $sl = $pres.Slides.Item(1)
    $i = 1
    foreach ($sh in $sl.Shapes) {
        $kind = "OTHER"
        switch ($sh.Type) {
            13 { $kind = "PICTURE" }
            17 { $kind = "TEXTBOX" }
            1  { $kind = "AUTO" }
            9  { $kind = "LINE" }
            default { $kind = "OTHER" }
        }
        $hasTF = "0"
        $tlen = 0.0
        $bt = 0.0; $bl = 0.0; $bw = 0.0; $bh = 0.0; $npara = 0.0
        try {
            if ($sh.HasTextFrame -eq -1) {
                $hasTF = "1"
                $tr = $sh.TextFrame.TextRange
                $tlen = $tr.Text.Length
                $bt = $tr.BoundTop
                $bl = $tr.BoundLeft
                $bw = $tr.BoundWidth
                $bh = $tr.BoundHeight
                $npara = $sh.TextFrame2.TextRange.Paragraphs.Count
            }
        } catch { }
        $lines.Add(("{0}`t{1}`t{2}`t{3:N4}`t{4:N4}`t{5:N4}`t{6:N4}`t{7:N4}`t{8:N4}`t{9:N4}`t{10:N4}`t{11:N4}`t{12:N4}`t{13}" -f `
                    $i, $kind, $hasTF, $sh.Left, $sh.Top, $sh.Width, $sh.Height, `
                    $bl, $bt, $bw, $bh, $tlen, $npara, $sh.Name))
        $i++
    }
    $pres.Close()
    [System.IO.File]::WriteAllLines($Out, $lines)
    Write-Output ("dumped {0} shapes -> {1}" -f $lines.Count, $Out)
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    # ReleaseComObject alone is NOT enough: PowerShell keeps its own RCW alive, so
    # POWERPNT.EXE lingers (it also keeps the "~$<name>.pptx" lock in dist/).
    # Force the finalizers to run before this process exits.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
