# Підписує EXE цифровим підписом Authenticode (SHA-256).
#
# Викликає build_messages_qt.bat після збірки; можна й вручну:
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\signing\sign_exe.ps1 -Path dist\PovidomlenniaKadrovi.exe
#
# Звідки береться сертифікат (перше, що знайдеться):
#   1) -PfxPath або змінна NAT_SIGN_PFX — файл .pfx (пароль: -PfxPassword або NAT_SIGN_PFX_PASSWORD);
#   2) -Thumbprint або змінна NAT_SIGN_THUMBPRINT — сертифікат зі сховища «Особисті»;
#   3) перший сертифікат підпису коду з приватним ключем у Cert:\CurrentUser\My
#      (його створює new_signing_cert.ps1).
# Мітка часу (-TimestampServer) додається, якщо сервер доступний: тоді підпис
# лишається чинним і після закінчення строку сертифіката. Без мережі підпис
# ставиться без мітки, з попередженням.
# Код повернення: 0 — підписано; 1 — помилка; 2 — сертифіката немає (EXE не підписано).

param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$PfxPath = $env:NAT_SIGN_PFX,
    [string]$PfxPassword = $env:NAT_SIGN_PFX_PASSWORD,
    [string]$Thumbprint = $env:NAT_SIGN_THUMBPRINT,
    [string]$TimestampServer = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Host "Файл для підпису не знайдено: $Path"
    exit 1
}

$cert = $null
if ($PfxPath) {
    if (-not (Test-Path -LiteralPath $PfxPath)) {
        Write-Host "Файл сертифіката не знайдено: $PfxPath"
        exit 1
    }
    $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $PfxPassword, $flags)
} elseif ($Thumbprint) {
    $clean = ($Thumbprint -replace "\s", "").ToUpper()
    $cert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My |
        Where-Object { $_.Thumbprint -eq $clean } | Select-Object -First 1
    if (-not $cert) {
        Write-Host "Сертифікат із відбитком $clean не знайдено в сховищі «Особисті»."
        exit 1
    }
} else {
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object { $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
}

if (-not $cert) {
    Write-Host "Сертифіката для підпису коду немає - EXE НЕ підписано."
    Write-Host "Створіть власний: scripts\signing\new_signing_cert.ps1 (див. scripts\signing\README.md)."
    exit 2
}
if (-not $cert.HasPrivateKey) {
    Write-Host "У сертифіката $($cert.Subject) немає приватного ключа - підписати ним не можна."
    exit 1
}

Write-Host "Сертифікат: $($cert.Subject)  (дійсний до $($cert.NotAfter.ToString('dd.MM.yyyy')), відбиток $($cert.Thumbprint))"

$result = $null
if ($TimestampServer) {
    try {
        $result = Set-AuthenticodeSignature -LiteralPath $Path -Certificate $cert -HashAlgorithm SHA256 -TimestampServer $TimestampServer
        if ($result.Status -eq "UnknownError" -and -not $result.SignerCertificate) { $result = $null }
    } catch {
        $result = $null
    }
    if (-not $result -or -not $result.SignerCertificate) {
        Write-Host "Сервер мітки часу недоступний ($TimestampServer) - підписую без мітки часу."
        $result = $null
    }
}
if (-not $result) {
    $result = Set-AuthenticodeSignature -LiteralPath $Path -Certificate $cert -HashAlgorithm SHA256
}

$check = Get-AuthenticodeSignature -LiteralPath $Path
if (-not $check.SignerCertificate) {
    Write-Host "Підписати не вдалося: $($result.StatusMessage)"
    exit 1
}
Write-Host "Підписано: $Path"
if ($check.TimeStamperCertificate) {
    Write-Host "Мітка часу: $($check.TimeStamperCertificate.Subject)"
}
if ($check.Status -eq "Valid") {
    Write-Host "Перевірка підпису: дійсний, цьому ПК видавець довіряє."
} else {
    # Самопідписаний сертифікат: підпис цілий, але Windows ще не довіряє видавцю.
    Write-Host "Перевірка підпису: $($check.Status) - $($check.StatusMessage)"
    Write-Host "Для власного сертифіката це очікувано: на кожному ПК треба один раз"
    Write-Host "встановити його відкриту частину (.cer) - scripts\signing\README.md."
}
exit 0
