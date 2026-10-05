<#
.SYNOPSIS
    Arctera (Veritas) Enterprise Vault: kullanıcının adı soyadından arşivini bulur ve export'u başlatır.

.DESCRIPTION
    1. Enterprise Vault PowerShell snap-in / modülünü yükler.
    2. Get-EVArchive ile tüm arşivleri çeker, verilen ad-soyad ile eşleşenleri bulur
       (büyük/küçük harf ve Türkçe karakter duyarsız: "Şerif Alıkavak" == "serif alikavak",
       "Ad Soyad" ve "Soyad, Ad" formatlarının ikisi de bulunur).
    3. Tek eşleşme varsa direkt, birden fazla varsa listeden seçtirerek
       Export-EVArchive ile export'u başlatır.

    Enterprise Vault sunucusunda, EV yönetici hesabıyla (Vault Service Account
    ya da export yetkisi olan bir rol) çalıştırılmalıdır.

.EXAMPLE
    .\Start-EVUserExport.ps1 -AdSoyad "Ahmet Yılmaz"

.EXAMPLE
    .\Start-EVUserExport.ps1 "Ahmet Yılmaz" -OutputDirectory "E:\Export" -MaxPSTSizeMB 10240

.EXAMPLE
    .\Start-EVUserExport.ps1 "Ahmet Yılmaz" -StartDate 2020-01-01 -EndDate 2023-12-31

.EXAMPLE
    # Sadece arşivi bul, export başlatma
    .\Start-EVUserExport.ps1 "Ahmet Yılmaz" -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Kullanıcının adı soyadı, örn: Ahmet Yılmaz')]
    [string]$AdSoyad,

    [string]$OutputDirectory = 'D:\EVExport',

    [ValidateSet('PST', 'MSG')]
    [string]$Format = 'PST',

    # PST dosya başına maksimum boyut (MB). Aşılırsa yeni PST dosyası açılır.
    [int]$MaxPSTSizeMB = 20480,

    [datetime]$StartDate,

    [datetime]$EndDate,

    # İsteğe bağlı arşiv tipi filtresi (örn: 'Exchange'). Boşsa tüm tipler aranır.
    [string]$ArchiveType = ''
)

$ErrorActionPreference = 'Stop'

function Get-NormalizedText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    # Türkçe ı/İ, NFD ile ayrışmadığı için önce elle çevrilir
    $t = $Text.Replace('ı', 'i').Replace('İ', 'I')
    $t = $t.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $t.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    # Noktalama -> boşluk, çoklu boşlukları tekle
    $t = $sb.ToString().ToLowerInvariant() -replace '[^a-z0-9]+', ' '
    return $t.Trim()
}

function Import-EVPowerShell {
    if (Get-Command Get-EVArchive -ErrorAction SilentlyContinue) { return }

    # EV 12+ : modül olarak gelebilir
    $module = Get-Module -ListAvailable | Where-Object { $_.Name -like '*EnterpriseVault*' } | Select-Object -First 1
    if ($module) {
        Import-Module $module.Name -DisableNameChecking
    }

    # Klasik snap-in (Symantec/Veritas/Arctera sürümleri)
    if (-not (Get-Command Get-EVArchive -ErrorAction SilentlyContinue)) {
        $snapin = Get-PSSnapin -Registered -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like '*EnterpriseVault*' } | Select-Object -First 1
        if ($snapin) {
            Add-PSSnapin $snapin.Name -ErrorAction SilentlyContinue
        }
    }

    if (-not (Get-Command Get-EVArchive -ErrorAction SilentlyContinue)) {
        throw 'Enterprise Vault PowerShell cmdlet''leri bulunamadı. Scripti EV sunucusunda ya da "Enterprise Vault Management Shell" içinden çalıştırın.'
    }
}

# --- 1. EV cmdlet'lerini yükle -------------------------------------------------
Import-EVPowerShell

# --- 2. Arşivi bul -------------------------------------------------------------
$search = Get-NormalizedText $AdSoyad
$tokens = @($search -split ' ' | Where-Object { $_ })
if ($tokens.Count -eq 0) { throw 'Geçerli bir ad soyad girin.' }

Write-Host "Arşivler taranıyor: '$AdSoyad' ..." -ForegroundColor Cyan

# Get-EVArchive'in parametreleri EV sürümüne göre değiştiği için filtre istemci tarafında uygulanır
$allArchives = @(Get-EVArchive)
if ($ArchiveType) {
    $allArchives = @($allArchives | Where-Object { "$($_.ArchiveType)" -like "*$ArchiveType*" })
}

# Önce birebir eşleşme ("Ad Soyad" veya "Soyad, Ad"), yoksa tüm kelimeleri içerenler
$reversed = $search
if ($tokens.Count -gt 1) {
    $reversed = (@($tokens[-1]) + $tokens[0..($tokens.Count - 2)]) -join ' '
}
$exact = @($allArchives | Where-Object {
        $n = Get-NormalizedText $_.ArchiveName
        $n -eq $search -or $n -eq $reversed
    })

if ($exact.Count -gt 0) {
    $matches_ = $exact
}
else {
    $matches_ = @($allArchives | Where-Object {
            $n = ' ' + (Get-NormalizedText $_.ArchiveName) + ' '
            $all = $true
            foreach ($tk in $tokens) { if ($n -notlike "* $tk*") { $all = $false; break } }
            $all
        })
}

if ($matches_.Count -eq 0) {
    throw "'$AdSoyad' için arşiv bulunamadı ($($allArchives.Count) arşiv tarandı)."
}

if ($matches_.Count -eq 1) {
    $archive = $matches_[0]
}
else {
    Write-Host "`nBirden fazla arşiv bulundu:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $matches_.Count; $i++) {
        $a = $matches_[$i]
        Write-Host ("  [{0}] {1}  |  {2}  |  {3}" -f ($i + 1), $a.ArchiveName, $a.ArchiveType, $a.ArchiveId)
    }
    do {
        $choice = Read-Host "Export edilecek arşivin numarası (1-$($matches_.Count))"
        $idx = 0
        $valid = [int]::TryParse($choice, [ref]$idx) -and $idx -ge 1 -and $idx -le $matches_.Count
    } until ($valid)
    $archive = $matches_[$idx - 1]
}

Write-Host "`nArşiv    : $($archive.ArchiveName)" -ForegroundColor Green
Write-Host "ArchiveId: $($archive.ArchiveId)" -ForegroundColor Green

# --- 3. Export'u başlat --------------------------------------------------------
$safeName = ($archive.ArchiveName -replace '[\\/:*?"<>|,]', '_').Trim()
$targetDir = Join-Path $OutputDirectory ("{0}_{1:yyyyMMdd_HHmmss}" -f $safeName, (Get-Date))

$exportArgs = @{
    ArchiveId       = $archive.ArchiveId
    OutputDirectory = $targetDir
    Format          = $Format
}
if ($Format -eq 'PST') { $exportArgs['MaxPSTSizeMB'] = $MaxPSTSizeMB }
if ($PSBoundParameters.ContainsKey('StartDate')) { $exportArgs['StartDate'] = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $exportArgs['EndDate'] = $EndDate }

# Bu EV sürümünde Export-EVArchive'in desteklemediği parametreleri ayıkla
$supported = (Get-Command Export-EVArchive).Parameters.Keys
foreach ($key in @($exportArgs.Keys)) {
    if ($supported -notcontains $key) {
        if ($key -in 'ArchiveId', 'OutputDirectory') {
            throw "Export-EVArchive '-$key' parametresini desteklemiyor. Desteklenenler: $($supported -join ', ')"
        }
        Write-Warning "Export-EVArchive '-$key' parametresini desteklemiyor, atlanıyor."
        $exportArgs.Remove($key)
    }
}

if ($PSCmdlet.ShouldProcess("$($archive.ArchiveName) [$($archive.ArchiveId)]", "Export-EVArchive -> $targetDir")) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    $log = Join-Path $targetDir 'export.log'
    Start-Transcript -Path $log | Out-Null
    try {
        Write-Host "Export başlıyor -> $targetDir ($Format)" -ForegroundColor Cyan
        $result = Export-EVArchive @exportArgs
        $result | Format-List | Out-String | Write-Host
        Write-Host "Export tamamlandı: $targetDir" -ForegroundColor Green
    }
    finally {
        Stop-Transcript | Out-Null
    }
}
