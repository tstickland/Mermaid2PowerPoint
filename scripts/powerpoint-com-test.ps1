# Inserts every SVG from test-output/ into a real PowerPoint presentation through COM, one slide
# per diagram at the position Insert would use, and tags it the way the add-in does. Saves
# all-diagrams.pptx / .pdf and a PNG of each slide (test-output/slides/), then reopens the file
# and checks that the source survived in each picture's tag and alt text.
# Run after `npm run test:render`.
#
# PowerPoint has a single COM instance: if PowerPoint is already open this script uses it, closes
# only the presentations it opened and leaves PowerPoint running. If PowerPoint stops responding,
# it is stopped after the timeout only if this script started it.
param(
  [string]$OutDir = (Join-Path $PSScriptRoot "..\test-output"),
  [int]$TimeoutSeconds = 300
)
$OutDir = (Resolve-Path $OutDir).Path
$before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)

$job = Start-Job -ArgumentList $OutDir, ($before.Count -gt 0) -ScriptBlock {
  param($OutDir, $wasRunning)
  # Any error (a COM call that throws, a missing file) ends the job, which the caller reports.
  $ErrorActionPreference = "Stop"
  $failures = 0
  $tag = "MERMAID_DIAGRAM"
  $placements = Get-Content -Raw "$OutDir\placements.json" | ConvertFrom-Json
  $slidesDir = Join-Path $OutDir "slides"
  New-Item -ItemType Directory -Force $slidesDir | Out-Null

  function Get-Src([string]$meta) {
    $header = ($meta -split "\r\n|\r|\n|\v")[0]
    $json = $header.Substring($header.IndexOf("{")) | ConvertFrom-Json
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($json.src))
  }

  $ppt = New-Object -ComObject PowerPoint.Application
  try {
    $pres = $ppt.Presentations.Add(0) # no window
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540
    $longMeta = "%% mermaid-word {`"v`":1,`"src`":`"`"}`n" + ("x" * 60000)

    $i = 0
    foreach ($p in $placements) {
      $i++
      $slide = $pres.Slides.Add($i, 12) # ppLayoutBlank
      $label = $slide.Shapes.AddTextbox(1, 4, 2, 600, 20)
      $label.TextFrame.TextRange.Text = "$($p.name) ($($p.id))"
      $label.TextFrame.TextRange.Font.Size = 10
      $file = Join-Path $OutDir "$($p.id).svg"
      if (-not (Test-Path $file)) { $file = Join-Path $OutDir "$($p.id).png" }
      $pic = $slide.Shapes.AddPicture($file, 0, -1, $p.left, $p.top, $p.width, $p.height)
      $meta = [IO.File]::ReadAllText((Join-Path $OutDir "$($p.id).meta.txt"))
      $pic.Name = "Mermaid diagram"
      $pic.Tags.Add($tag, $meta)
      $pic.AlternativeText = $meta
    }
    # A very long source, to check tag capacity.
    $pres.Slides.Item(1).Shapes.Item(2).Tags.Add("LONG_TEST", $longMeta)

    $pptx = Join-Path $OutDir "all-diagrams.pptx"
    $pres.SaveAs($pptx, 24) # ppSaveAsOpenXMLPresentation
    $pres.SaveCopyAs((Join-Path $OutDir "all-diagrams.pdf"), 32) # ppSaveAsPDF
    $i = 0
    foreach ($p in $placements) { $i++; $pres.Slides.Item($i).Export((Join-Path $slidesDir "$($p.id).png"), "PNG", 1920, 1080) }
    $pres.Close()
    "Saved all-diagrams.pptx, all-diagrams.pdf and $($placements.Count) slide images"

    # Reopen and check what survived.
    $pres = $ppt.Presentations.Open($pptx, -1, 0, 0) # read-only, no window
    $i = 0
    foreach ($p in $placements) {
      $i++
      $original = [IO.File]::ReadAllText((Join-Path $OutDir "$($p.id).meta.txt"))
      $pic = $pres.Slides.Item($i).Shapes.Item(2)
      $fromTag = $pic.Tags.Item($tag)
      $alt = $pic.AlternativeText
      $tagExact = $fromTag -ceq $original
      $altSrcOk = $false
      try { $altSrcOk = (Get-Src $alt) -ceq (Get-Src $original) } catch {}
      $altExact = $alt -ceq $original
      $kind = switch ($pic.Type) { 28 { "svg" } 13 { "picture" } default { "type$($pic.Type)" } }
      $ok = $tagExact -and $altSrcOk
      if (-not $ok) { $failures++ }
      "{0}  {1,-19} {2,-8} tag exact={3}  alt src ok={4} exact={5}" -f ($(if ($ok) { "ok  " } else { "FAIL" })), $p.id, $kind, $tagExact, $altSrcOk, $altExact
    }
    $long = $pres.Slides.Item(1).Shapes.Item(2).Tags.Item("LONG_TEST")
    "Long tag ($($longMeta.Length) chars) survived: $($long -ceq $longMeta)"
    if ($long -cne $longMeta) { $failures++ }
    $pres.Close()
  } finally {
    if (-not $wasRunning) { $ppt.Quit() }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
  "FAILURES=$failures"
}

if (-not (Wait-Job $job -Timeout $TimeoutSeconds)) {
  Write-Host "PowerPoint did not finish within $TimeoutSeconds s."
  $started = Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id }
  if ($started) { Write-Host "Stopping the PowerPoint instance this script started."; $started | Stop-Process -Force }
  Stop-Job $job
  Receive-Job $job | Write-Host
  exit 1
}
$jobState = $job.State
$jobErrors = @()
$output = Receive-Job $job -ErrorVariable jobErrors -ErrorAction SilentlyContinue
Remove-Job $job
# Quit() can leave a windowless POWERPNT behind; stop it only if this script started it.
if (-not $before.Count) {
  for ($i = 0; $i -lt 30 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id -and $_.MainWindowHandle -eq 0 } | Stop-Process -Force
}
$output | Where-Object { $_ -notlike "FAILURES=*" } | Write-Host
$summary = $output | Where-Object { $_ -like "FAILURES=*" } | Select-Object -Last 1
# A job that threw (missing placements.json, a failed AddPicture, no PowerPoint) never prints
# FAILURES=; treat that as a failure rather than as zero problems.
if ($jobState -ne "Completed" -or -not $summary) {
  $jobErrors | ForEach-Object { Write-Host "ERROR: $_" }
  Write-Host "The test did not finish (job state: $jobState)."
  exit 1
}
$failures = [int]($summary -replace "FAILURES=", "")
Write-Host "$failures problem(s)"
exit $failures
