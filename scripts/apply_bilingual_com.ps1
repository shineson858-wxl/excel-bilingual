param(
  [Parameter(Mandatory=$true)][string]$Source,
  [Parameter(Mandatory=$true)][string]$Out,
  [Parameter(Mandatory=$true)][string]$Glossary,
  [string]$Rules = "",
  [switch]$NoHeight,
  [switch]$AllowInteractive
)
# Same-cell bilingual writer that goes through Excel itself (COM).
#
# Use it when audit_features.py returns RISKY: openpyxl would drop text shapes,
# slicers, form controls, pivot tables, VBA or x14 validations, while Excel keeps
# them because the file is never re-serialised by a third-party library.
#
# It writes "<source>" + newline + "<translation>" into the same cell, turns
# formula cells whose cached result is text into static bilingual text, and also
# translates text inside shapes / text boxes (which openpyxl cannot even see).
# Messages are ASCII on purpose: Windows PowerShell 5.1 mis-decodes non-ASCII
# bytes in script files saved without a BOM.

$ErrorActionPreference = 'Stop'

# Excel refuses scripted calls with RPC_E_CALL_REJECTED / VBA_E_IGNORE while it is
# busy (recalculating, refreshing links, showing a hidden dialog). The only cure
# is an OLE message filter that asks COM to retry instead of failing: install it
# once, before the first Excel object is created.
$filterSrc = @'
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
  [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
  [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
  [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

public class ExcelMessageFilter : IOleMessageFilter {
  [DllImport("ole32.dll")]
  private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

  public static bool Register() {
    IOleMessageFilter old = null;
    int hr = CoRegisterMessageFilter(new ExcelMessageFilter(), out old);
    return hr == 0;
  }

  public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }

  public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) {
    if (dwRejectType == 2) { return dwTickCount > 60000 ? -1 : 200; }
    return -1;
  }

  public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }
}
'@
try {
  Add-Type -TypeDefinition $filterSrc -ErrorAction Stop
  if ([ExcelMessageFilter]::Register()) { Write-Output "message filter installed" }
  else { Write-Output "WARN: message filter registration refused (thread not STA)" }
} catch {
  Write-Output ("WARN: message filter unavailable: " + $_.Exception.Message.Split("`n")[0])
}

function NormKey([string]$s) {
  if ($null -eq $s) { return '' }
  $t = $s.Normalize([Text.NormalizationForm]::FormC)
  $t = $t -replace "`r`n", "`n"
  $t = $t -replace "`r", "`n"
  $t = [regex]::Replace($t, '\s+', ' ')
  return $t.Trim().ToLowerInvariant()
}

function LoadGlossary([string]$path) {
  $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
  $obj = $raw | ConvertFrom-Json
  $map = @{}
  foreach ($p in $obj.PSObject.Properties) {
    if ($p.Value -is [string] -and $p.Value.Trim().Length -gt 0) {
      $k = NormKey($p.Name)
      if (-not $map.ContainsKey($k)) { $map[$k] = $p.Value }
    }
  }
  return $map
}

$gloss = LoadGlossary $Glossary
Write-Output ("glossary keys: " + $gloss.Count)

# An interactive Excel window makes automation fail with RPC_E_CALL_REJECTED
# (0x80010001) the moment a dialog or a mouse drag is in progress. Orphaned
# hidden EXCEL.EXE from an earlier failed automation poisons the ROT as well.
$busy = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' })
if ($busy.Count -gt 0) {
  if (-not $AllowInteractive) {
    Write-Output ("ERR: " + $busy.Count + " visible Excel window(s) are open; close Excel (or pass -AllowInteractive) before running this script")
    exit 2
  }
} else {
  $orphans = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
  if ($orphans.Count -gt 0) {
    Write-Output ("cleaning " + $orphans.Count + " orphaned hidden Excel process(es)")
    foreach ($p in $orphans) { try { Stop-Process -Id $p.Id -Force } catch {} }
    Start-Sleep -Seconds 3
  }
}

function Invoke-Retry([scriptblock]$sb, [int]$tries = 20, [switch]$Quiet) {
  for ($n = 1; $n -le $tries; $n++) {
    try {
      if ($Quiet) { & $sb *> $null; return $null }
      return & $sb
    } catch {
      $msg = $_.Exception.Message
      if ($msg -like '*RPC_E_CALL_REJECTED*' -or $msg -like '*rejected by callee*' -or $msg -like '*0x800AC472*' -or $msg -like '*0x80010001*') {
        Start-Sleep -Milliseconds (700 * $n)
        continue
      }
      throw
    }
  }
  throw "Excel kept rejecting calls; close Excel and remove lingering EXCEL processes, then retry"
}

$ruleList = @()
if ($Rules -ne '') {
  $rj = (Get-Content -LiteralPath $Rules -Raw -Encoding UTF8) | ConvertFrom-Json
  foreach ($r in $rj) {
    # the rules file is Python-flavoured: back-references are "\1". .NET regex
    # expects "$1", so without this translation the replacement lands literally.
    $repl = [regex]::Replace([string]$r.replacement, '\\(\d+)', '$$$1')
    $ruleList += ,@([regex]::new($r.pattern, 'IgnoreCase'), $repl)
  }
}

if (-not (Test-Path -LiteralPath $Source)) { Write-Output "ERR: source not found"; exit 1 }
Copy-Item -LiteralPath $Source -Destination $Out -Force
Write-Output ("working copy: " + $Out)

$SEP = "`n"
$xl = $null
$wb = $null
$changed = 0
$formula2text = 0
$shapeChanged = 0
$errCells = 0
$missing = @{}
$touchedRows = @{}
$ok = $false

try {
  # Excel 2013 sometimes starts up so slowly that even Workbooks.Open is refused.
  # Wait until a trivial call answers, and if the instance stays dead, replace it.
  $opened = $false
  for ($inst = 1; $inst -le 3 -and -not $opened; $inst++) {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    $xl.ScreenUpdating = $false
    $xl.EnableEvents = $false
    $ready = $false
    for ($n = 1; $n -le 20 -and -not $ready; $n++) {
      try { $null = $xl.Version; $ready = $true } catch { Start-Sleep -Seconds 1 }
    }
    if (-not $ready) {
      Write-Output ("instance " + $inst + ": never became responsive, replacing")
    } else {
      for ($n = 1; $n -le 6 -and -not $opened; $n++) {
        try { $wb = $xl.Workbooks.Open($Out, 0, $false); $opened = $true }
        catch {
          $msg = $_.Exception.Message
          Write-Output ("open attempt " + $n + " failed: " + $msg.Split("`n")[0])
          Start-Sleep -Seconds (2 * $n)
        }
      }
    }
    if (-not $opened) {
      try { $xl.Quit() } catch {}
      foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' })) {
        try { Stop-Process -Id $p.Id -Force } catch {}
      }
      Start-Sleep -Seconds 3
    }
  }
  if (-not $opened) { throw "cannot open the workbook through Excel COM (3 instances failed)" }
  # manual calculation: in automatic mode Excel returns VBA_E_IGNORE (0x800AC472)
  # while it recalculates, which happens after almost every single write
  try { $wb.Application.Calculation = -4135 } catch { Write-Output ("WARN: could not switch to manual calculation: " + $_.Exception.Message.Split("`n")[0]) }
  Write-Output ("sheets: " + $wb.Worksheets.Count)

  foreach ($ws in $wb.Worksheets) {
    $ur = $ws.UsedRange
    # UsedRange often spans tens of thousands of columns because some cell far
    # right carries formatting only. Clamp to the area that really holds values
    # (constants + formulas), otherwise the per-cell loop runs for millions of
    # empty cells and the run never finishes.
    $r0 = $ur.Row
    $c0 = $ur.Column
    $nr = $ur.Rows.Count
    $nc = $ur.Columns.Count
    $minR = [int]::MaxValue; $maxR = 0; $minC = [int]::MaxValue; $maxC = 0
    foreach ($t in @(2, -4123)) {
      try {
        $areas = $ur.SpecialCells($t)
        foreach ($a in $areas.Areas) {
          $ar = $a.Row; $ac = $a.Column
          $aer = $ar + $a.Rows.Count - 1; $aec = $ac + $a.Columns.Count - 1
          if ($ar -lt $minR) { $minR = $ar }
          if ($aer -gt $maxR) { $maxR = $aer }
          if ($ac -lt $minC) { $minC = $ac }
          if ($aec -gt $maxC) { $maxC = $aec }
        }
      } catch {}
    }
    if ($maxC -gt 0 -and $maxR -gt 0) {
      $r0 = $minR; $c0 = $minC
      $nr = $maxR - $minR + 1
      $nc = $maxC - $minC + 1
    }
    Write-Output ("sheet " + $ws.Name + " range " + $nr + "x" + $nc)

    # one COM read for the whole data area; Value2 comes back as a SAFEARRAY
    # whose bounds are 1-based, so always index relative to GetLowerBound.
    $band = $ws.Range($ws.Cells($r0, $c0), $ws.Cells($r0 + $nr - 1, $c0 + $nc - 1))
    $vals = $band.Value2
    if ($null -eq $vals) { continue }
    $rank = 0
    if ($vals -is [Array]) { $rank = $vals.Rank }
    $lb0 = 0
    $lb1 = 0
    if ($rank -eq 2) { $lb0 = $vals.GetLowerBound(0); $lb1 = $vals.GetLowerBound(1) }

    for ($i = 0; $i -lt $nr; $i++) {
      $rr = $r0 + $i
      for ($j = 0; $j -lt $nc; $j++) {
        if ($rank -eq 2) { $v = $vals.GetValue($lb0 + $i, $lb1 + $j) }
        elseif ($rank -eq 1) { if ($nr -eq 1) { $v = $vals.GetValue($vals.GetLowerBound(0) + $j) } else { $v = $vals.GetValue($vals.GetLowerBound(0) + $i) } }
        else { $v = $vals }
        if ($null -eq $v) { continue }
        if ($v -isnot [string]) { continue }
        if ($v.Trim().Length -eq 0) { continue }
        if ($v.StartsWith('=')) { continue }

        # $ws.Cells.Item(r, c) can come back null through the PS binding; the
        # default member ($ws.Cells(r, c)) is the reliable way to get a Range.
        $cell = $ws.Cells($rr, $c0 + $j)
        if ($null -eq $cell) { continue }
        try {
        # single-cell HasFormula is safe; the range-level one returns Null for
        # mixed ranges and .Formula would make Excel resolve the external links
        $isFormula = $false
        try { if ($cell.HasFormula -eq $true) { $isFormula = $true } } catch {}

        $key = NormKey($v)
        $zh = $null
        if ($gloss.ContainsKey($key)) { $zh = $gloss[$key] }
        elseif ($ruleList.Count -gt 0) {
          $flat = [regex]::Replace($v.Normalize([Text.NormalizationForm]::FormC), '\s+', ' ').Trim()
          foreach ($rule in $ruleList) {
            $cand = $rule[0].Replace($flat, $rule[1])
            if ($cand -ne $flat) { $zh = $cand; break }
          }
        }
        if ($null -eq $zh) {
          if ($missing.ContainsKey($key)) { $missing[$key] = $missing[$key] + 1 } else { $missing[$key] = 1 }
          continue
        }
        if ($zh.Trim().Length -eq 0) { continue }
        if ((NormKey($zh)) -eq $key) { continue }

        $text = $v.TrimEnd() + $SEP + $zh
        $null = Invoke-Retry { $cell.Value2 = $text; $cell.WrapText = $true } -tries 6 -Quiet
        $changed++
        if ($isFormula) { $formula2text++ }
        $rk = $ws.Name + "|" + $rr
        if ($touchedRows.ContainsKey($rk)) { $touchedRows[$rk] = $touchedRows[$rk] + 1 } else { $touchedRows[$rk] = 1 }
        } catch {
          $errCells++
          if ($errCells -le 10) {
            $addr = "?"
            try { $addr = $cell.Address(0, 0) } catch {}
            Write-Output ("CELL ERR " + $ws.Name + "!" + $addr + ": " + $_.Exception.Message.Split("`n")[0])
          }
        }
      }
    }

    # text inside shapes / text boxes: invisible to openpyxl
    $sc = $ws.Shapes.Count
    if ($sc -gt 0) {
      for ($s = 1; $s -le $sc; $s++) {
        try {
          $sh = $ws.Shapes.Item($s)
          if ($sh.TextFrame2.HasText -ne $true) { continue }
          $txt = [string]$sh.TextFrame2.TextRange.Text
          if ($txt.Trim().Length -eq 0) { continue }
          $key = NormKey($txt)
          if (-not $gloss.ContainsKey($key)) { continue }
          $sh.TextFrame2.TextRange.Text = $txt.TrimEnd() + $SEP + $gloss[$key]
          $shapeChanged++
        } catch { Write-Output ("shape " + $s + " skipped: " + $_.Exception.Message) }
      }
    }
  }

  if (-not $NoHeight) {
    foreach ($ws in $wb.Worksheets) {
      $prefix = $ws.Name + "|"
      foreach ($k in @($touchedRows.Keys)) {
        if (-not $k.StartsWith($prefix)) { continue }
        $r = [int]($k.Substring($prefix.Length))
        try {
          $rowRange = $ws.Rows.Item($r)
          if ($rowRange.MergeCells -eq $true) {
            # merged rows cannot AutoFit; approximate by doubling the stored height
            $old = $rowRange.RowHeight
            $rowRange.RowHeight = [Math]::Min(409, [Math]::Max($old, $old * 2))
            $null = $rowRange.RowHeight
          } else {
            $null = $rowRange.AutoFit()
          }
        } catch {}
      }
    }
  }

  $null = Invoke-Retry { $wb.Save() } -tries 10 -Quiet
  $ok = $true
  Write-Output ("cells updated: " + $changed)
  Write-Output ("formula cells turned into static text: " + $formula2text)
  Write-Output ("shape texts updated: " + $shapeChanged)
  Write-Output ("cell errors skipped: " + $errCells)
  Write-Output ("untranslated unique: " + $missing.Count + " cells: " + (($missing.Values | Measure-Object -Sum).Sum))
  foreach ($k in $missing.Keys) { Write-Output ("MISSING " + $missing[$k] + "`t" + $k) }
}
catch {
  Write-Output ("FATAL line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message.Split("`n")[0])
}
finally {
  if ($wb -ne $null) {
    if ($ok) { try { $wb.Close($true) } catch {} } else { try { $wb.Close($false) } catch {} }
  }
  if ($xl -ne $null) {
    $quitOk = $false
    try { $xl.Quit(); $quitOk = $true } catch {}
    # a stuck instance keeps the ROT poisoned for every later run: kill it
    if (-not $quitOk) {
      foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' })) {
        try { Stop-Process -Id $p.Id -Force } catch {}
      }
    }
  }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
if (-not $ok) { Write-Output "ERR: aborted, nothing saved"; exit 1 }


