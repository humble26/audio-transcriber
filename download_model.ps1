param(
  [string]$Repo = "Systran/faster-whisper-small",
  [string]$Dest = "",
  [string]$Files = "config.json,model.bin,tokenizer.json,vocabulary.txt"
)
$ErrorActionPreference = "Continue"
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$base = "https://hf-mirror.com/$Repo/resolve/main"
foreach ($f in ($Files -split ',')) {
  $f = $f.Trim()
  if (-not $f) { continue }
  $out = Join-Path $Dest $f
  Write-Host "[$(Get-Date -Format 'HH:mm:ss')] downloading $f ..."
  curl.exe -L --retry 5 --retry-delay 2 -C - --connect-timeout 20 --no-progress-meter -o "$out" "$base/$f"
  if (Test-Path $out) {
    Write-Host ("    -> {0:N2} MB" -f ((Get-Item $out).Length/1MB))
  } else {
    Write-Host "    -> FAILED"
  }
}
Write-Host "[$(Get-Date -Format 'HH:mm:ss')] ALL DONE: $Dest"