# g0.4-versiyon-bul-powershell.ps1
# ASCII-safe. Bu script buildVars.py icinden addon_version degerini okur
# ve ayni klasore 'versiyon.txt' olusturur. Icerik tek satir: 1.8.2 veya NO_MATCH veya PYERR:...

$ErrorActionPreference = "Stop"

# Dosya yolu: buildVars.py ayni klasorde olmalidir
$buildVarsPath = Join-Path -Path $PSScriptRoot -ChildPath "buildVars.py"
$outPath = Join-Path -Path $PSScriptRoot -ChildPath "versiyon.txt"

try {
    if (-not (Test-Path -Path $buildVarsPath)) {
        "NO_FILE" | Out-File -FilePath $outPath -Encoding ASCII
        exit 1
    }

    # raw oku, UTF8 ve BOM icin guvenli
    $s = Get-Content -Raw -Encoding UTF8 -Path $buildVarsPath

    # regex ile eslestir
    if ($s -match 'addon_version\s*=\s*["'']([^"'']+)["'']') {
        $matches[1] | Out-File -FilePath $outPath -Encoding ASCII
        exit 0
    } else {
        "NO_MATCH" | Out-File -FilePath $outPath -Encoding ASCII
        exit 2
    }
}
catch {
    # hata mesaji versiyon dosyasina yazilsin
    ("PYERR:" + $_.Exception.Message) | Out-File -FilePath $outPath -Encoding ASCII
    exit 3
}
