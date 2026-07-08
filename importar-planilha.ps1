param(
    [Parameter(Mandatory = $true)]
    [string]$Planilha
)

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not (Test-Path "node_modules")) {
    Write-Host "Instalando dependencias..."
    npm install
}

npm run import:firebase -- $Planilha
