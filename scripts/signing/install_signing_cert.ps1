# Встановлює відкриту частину сертифіката (.cer) на робочому ПК, щоб Windows
# довіряла підпису EXE повідомлень. Приватного ключа тут немає й не потрібно.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File install_signing_cert.ps1 -CerPath NAT_code_signing.cer
#   ... -AllUsers   - для всіх користувачів ПК (потрібні права адміністратора)

param(
    [string]$CerPath = (Join-Path $PSScriptRoot "NAT_code_signing.cer"),
    [switch]$AllUsers
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $CerPath)) {
    Write-Host "Файл сертифіката не знайдено: $CerPath"
    exit 1
}
$scope = if ($AllUsers) { "LocalMachine" } else { "CurrentUser" }
Import-Certificate -FilePath $CerPath -CertStoreLocation "Cert:\$scope\Root" | Out-Null
Import-Certificate -FilePath $CerPath -CertStoreLocation "Cert:\$scope\TrustedPublisher" | Out-Null
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CerPath)
Write-Host "Встановлено ($scope): $($cert.Subject), відбиток $($cert.Thumbprint)"
