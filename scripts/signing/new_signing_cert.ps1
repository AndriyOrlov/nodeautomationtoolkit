# Створює власний (самопідписаний) сертифікат підпису коду для EXE повідомлень.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\signing\new_signing_cert.ps1
#
# Що робить:
#   - кладе сертифікат із приватним ключем у Cert:\CurrentUser\My
#     (звідти його бере sign_exe.ps1 і build_messages_qt.bat);
#   - експортує ВІДКРИТУ частину в signing\NAT_code_signing.cer поруч із проєктом -
#     її встановлюють на робочих ПК, щоб Windows довіряла підпису;
#   - на цьому ПК одразу додає .cer до «Довірених кореневих центрів» і
#     «Довірених видавців» поточного користувача (Windows спитає підтвердження).
# Приватний ключ із ПК не виходить; .pfx не створюється, якщо не вказати -ExportPfx.

param(
    [string]$Subject = "CN=NAT Povidomlennia Code Signing",
    [int]$Years = 5,
    [switch]$ExportPfx,
    [switch]$NoTrust
)

$ErrorActionPreference = "Stop"
$root = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$outDir = Join-Path $root "signing"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject `
    -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 3072 `
    -HashAlgorithm SHA256 -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears($Years)

$cerPath = Join-Path $outDir "NAT_code_signing.cer"
Export-Certificate -Cert $cert -FilePath $cerPath -Type CERT | Out-Null
Write-Host "Сертифікат створено: $($cert.Subject)"
Write-Host "Відбиток: $($cert.Thumbprint)   дійсний до $($cert.NotAfter.ToString('dd.MM.yyyy'))"
Write-Host "Відкрита частина для робочих ПК: $cerPath"

if ($ExportPfx) {
    $password = Read-Host "Пароль для .pfx (резервна копія ключа)" -AsSecureString
    $pfxPath = Join-Path $outDir "NAT_code_signing.pfx"
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password | Out-Null
    Write-Host "Резервна копія з приватним ключем: $pfxPath - НЕ передавайте її й не кладіть у git."
}

if (-not $NoTrust) {
    Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
    Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
    Write-Host "На цьому ПК сертифікату довіряє поточний користувач."
}
