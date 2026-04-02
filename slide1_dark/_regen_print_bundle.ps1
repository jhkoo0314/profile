$ErrorActionPreference = 'Stop'
$root = 'C:\profile\slide1_dark'
$outFile = Join-Path $root 'print-all-slides.html'
$entries = @(
  @{ Path='slide-01.html'; Mode='static' },
  @{ Path='slide-02.html'; Mode='static' },
  @{ Path='slide-03.html'; Mode='static' },
  @{ Path='slide-04.html'; Mode='static' },
  @{ Path='slide-05.html'; Mode='static' },
  @{ Path='slide-06.html'; Mode='static' },
  @{ Path='slide-07.html'; Mode='static' },
  @{ Path='slide-08.html'; Mode='static' },
  @{ Path='slide-09.html'; Mode='static' },
  @{ Path='slide-10.html'; Mode='static' },
  @{ Path='slide-11.html'; Mode='static' },
  @{ Path='slide-12.html'; Mode='static' },
  @{ Path='slide-13.html'; Mode='static' },
  @{ Path='slide-14.html'; Mode='static' },
  @{ Path='slide-15.html'; Mode='static' },
  @{ Path='slide-16.html'; Mode='static' },
  @{ Path='slide-17.html'; Mode='static' },
  @{ Path='slide-18.html'; Mode='static' },
  @{ Path='slide-19.html'; Mode='static' },
  @{ Path='slide-20.html'; Mode='static' },
  @{ Path='slide-21.html'; Mode='static' },
  @{ Path='slide-22.html'; Mode='static' },
  @{ Path='slide-23.html'; Mode='static' },
  @{ Path='slide-24.html'; Mode='dynamic' },
  @{ Path='slide-25.html'; Mode='dynamic' },
  @{ Path='slide-26.html'; Mode='dynamic' },
  @{ Path='slide-27.html'; Mode='dynamic' },
  @{ Path='slide-28.html'; Mode='dynamic' }
)

function Get-BodyInnerHtml([string]$html) {
  $m = [regex]::Match($html, '<body[^>]*>([\s\S]*?)</body>', 'IgnoreCase')
  if ($m.Success) { return $m.Groups[1].Value.Trim() }
  return ''
}

function Get-StyleBlocks([string]$html) {
  [regex]::Matches($html, '<style[^>]*>([\s\S]*?)</style>', 'IgnoreCase') | ForEach-Object { $_.Groups[1].Value.Trim() }
}

function Get-InlineScripts([string]$html) {
  [regex]::Matches($html, '<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)</script>', 'IgnoreCase') | ForEach-Object { $_.Groups[1].Value.Trim() }
}

function Get-ExternalAssets([string]$html) {
  $assets = New-Object System.Collections.Generic.List[string]
  [regex]::Matches($html, '<script[^>]+src=["''][^"'']+["''][^>]*></script>', 'IgnoreCase') | ForEach-Object { $assets.Add($_.Value.Trim()) }
  [regex]::Matches($html, '<link[^>]+(?:rel=["'']stylesheet["'']|rel=["'']preconnect["'']|rel=["'']preload["''])[^>]*>', 'IgnoreCase') | ForEach-Object { $assets.Add($_.Value.Trim()) }
  return $assets
}

function Invoke-ScopeCss([string]$css, [string]$prefix) {
  $css = [regex]::Replace($css, '/\*[\s\S]*?\*/', '')
  $keyframes = [regex]::Matches($css, '(?ms)@keyframes\s+[^{]+\{(?:[^{}]|\{[^{}]*\})*\}') | ForEach-Object { $_.Value.Trim() }
  $css = [regex]::Replace($css, '(?ms)@keyframes\s+[^{]+\{(?:[^{}]|\{[^{}]*\})*\}', '')
  $result = [regex]::Replace($css, '(?ms)([^{}@]+)\{([^{}]*)\}', {
    param($m)
    $selectorText = $m.Groups[1].Value.Trim()
    $body = $m.Groups[2].Value.Trim()
    if ([string]::IsNullOrWhiteSpace($selectorText) -or [string]::IsNullOrWhiteSpace($body)) { return '' }
    $selectors = $selectorText.Split(',') | ForEach-Object {
      $s = $_.Trim()
      if (-not $s) { return }
      $s = $s -replace ':root', $prefix
      $s = $s -replace '(^|\s)html\b', ('$1' + $prefix)
      $s = $s -replace '(^|\s)body\b', ('$1' + $prefix)
      if ($s -notmatch [regex]::Escape($prefix)) { "$prefix $s" } else { $s }
    }
    $selectors = $selectors | Where-Object { $_ -and $_.Trim() -ne '' }
    if (-not $selectors) { return '' }
    return (($selectors -join ', ') + " {`n" + $body + "`n}")
  })
  return (($result.Trim()) + "`n" + ($keyframes -join "`n")).Trim()
}

function Invoke-PrefixScriptIds([string]$script, [string]$id) {
  $script = $script -replace "document\.getElementById\('([^']+)'\)", "document.querySelector('#$id #`$1')"
  $script = $script -replace 'document\.getElementById\("([^"]+)"\)', 'document.querySelector("#$id #$1")'
  return $script
}

$globalStyles = @"
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: #111; }
body { font-family: 'Outfit', sans-serif; }
.deck { display: flex; flex-direction: column; align-items: center; gap: 0; }
.slide-page {
  width: 16in;
  height: 9in;
  margin: 0;
  padding: 0;
  overflow: hidden;
  page-break-after: always;
  break-after: page;
  background: #000;
  position: relative;
}
.slide-page:last-child { page-break-after: auto; break-after: auto; }
.slide-root {
  width: 100%;
  height: 100%;
  position: relative;
  opacity: 1 !important;
  visibility: visible !important;
  overflow: hidden;
}
.slide-root, .slide-root * {
  animation: none !important;
  transition: none !important;
}
.slide-root .fade-up,
.slide-root .stagger-up,
.slide-root .fade-in,
.slide-root .pop-in,
.slide-root .slide-in,
.slide-root .slide-in-right,
.slide-root [class*='delay-'] {
  opacity: 1 !important;
  visibility: visible !important;
  transform: none !important;
  filter: none !important;
}
@page { size: 16in 9in; margin: 0; }
@media screen {
  body { padding: 20px 0; }
  .slide-page { box-shadow: 0 6px 28px rgba(0, 0, 0, 0.45); }
}
@media print {
  html, body { background: #fff; }
  .deck { display: block; }
  .slide-page { box-shadow: none; }

  /* Print-to-PDF stability (Chrome):
     Avoid backdrop-filter in print and force exact color adjustments. */
  .slide-page, .slide-root {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }

  /* Reported missing "KPI cards": slide 02 / 10 / 13 */
  #slide-02 .principle-card,
  #slide-10 .flow-node,
  #slide-13 .step-card {
    -webkit-backdrop-filter: none !important;
    backdrop-filter: none !important;
    background: rgba(255, 255, 255, 0.08) !important;
    border-color: rgba(255, 255, 255, 0.18) !important;
  }

  /* Decorative moving particles are irrelevant in print. */
  #slide-10 .particle { display: none !important; }
}
"@

$styleParts = New-Object System.Collections.Generic.List[string]
$sections = New-Object System.Collections.Generic.List[string]
$scripts = New-Object System.Collections.Generic.List[string]
$assets = New-Object 'System.Collections.Generic.HashSet[string]'
$index = 0
foreach ($entry in $entries) {
  $index++
  $path = Join-Path $root $entry.Path
  $html = Get-Content $path -Raw -Encoding UTF8
  $id = ('slide-{0:d2}' -f $index)
  $prefix = "#$id"
  foreach ($asset in (Get-ExternalAssets $html)) { [void]$assets.Add($asset) }
  $styles = (Get-StyleBlocks $html | ForEach-Object { Invoke-ScopeCss $_ $prefix }) -join "`n"
  $bodyInner = Get-BodyInnerHtml $html
  $styleParts.Add($styles)
  $sections.Add(@"
<section class="slide-page">
  <div class="slide-root" id="$id" data-source="$($entry.Path)">
$bodyInner
  </div>
</section>
"@)
  if ($entry.Mode -eq 'dynamic') {
    foreach ($inlineScript in (Get-InlineScripts $html)) {
      $scripts.Add(@"
(function() {
  const root = document.getElementById('$id');
  if (!root) return;
  ${inlineScript}
})();
"@)
    }
  }
}

$assetText = ($assets | Sort-Object) -join "`n  "
$scriptText = ($scripts -join "`n")

$output = @"
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>slide1_dark - PDF Print Bundle</title>
  $assetText
  <style>
$globalStyles
$($styleParts -join "`n")
  </style>
</head>
<body>
  <main class="deck" aria-label="slide1_dark print bundle">
$($sections -join "`n")
  </main>
  <script>
$scriptText
  </script>
</body>
</html>
"@
Set-Content -Path $outFile -Value $output -Encoding UTF8
Write-Output "Regenerated with reports: $outFile"
