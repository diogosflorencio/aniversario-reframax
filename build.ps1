$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

Write-Host "Compilando Aniversariantes..."
py -m PyInstaller Aniversariantes.spec --noconfirm --clean
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$Dist = Join-Path $Root "dist\Aniversariantes"
$Pastas = @("modelos", "fontes", "layouts", "planilhas", "outputs")

foreach ($pasta in $Pastas) {
    $origem = Join-Path $Root $pasta
    $destino = Join-Path $Dist $pasta
    if (Test-Path $destino) { Remove-Item $destino -Recurse -Force }
    Copy-Item $origem $destino -Recurse -Force
    Write-Host "Copiado: $pasta/"
}

Copy-Item (Join-Path $Root "server.json") (Join-Path $Dist "server.json") -Force
Copy-Item (Join-Path $Root "INSTALACAO.txt") (Join-Path $Dist "INSTALACAO.txt") -Force

Write-Host ""
Write-Host "Build concluido: $Dist"
Write-Host "Execute Aniversariantes.exe - pastas editaveis ficam na raiz do app."
