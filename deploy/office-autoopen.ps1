<#
.SYNOPSIS
    Office-Anhänge aus der MultiData Werkbank öffnen sich in Chrome und Edge automatisch im lokalen Office.

.DESCRIPTION
    Setzt die Browser-Richtlinien AutoOpenFileTypes und AutoOpenAllowedForURLs (HKLM) für Google Chrome
    und Microsoft Edge. Downloads dieser Dateitypen werden dann ohne Klick auf die Download-Leiste
    geöffnet – aber nur, wenn sie von MultiData kommen (AutoOpenAllowedForURLs).

    - Vorhandene Einträge in beiden Listen bleiben erhalten; es wird nur ergänzt (bzw. mit -Remove entfernt).
    - Mehrfach ausführbar (idempotent).
    - Benötigt Administratorrechte. Für Intune als Plattformskript im Systemkontext (64-Bit) verteilen.
    - Chrome und Edge werten diese Richtlinien unter Windows nur auf verwalteten Geräten aus
      (AD-Domäne, Entra ID oder Geräteverwaltung/Intune).
    - Makrofähige Formate (docm, xlsm, pptm, dotm, xlsb …) sind bewusst nicht in der Standardliste:
      Ticket-Anhänge stammen oft von externen Absendern.

.PARAMETER Url
    Herkunft, für die das automatische Öffnen gilt. Standard: csp.synaforce.com

.PARAMETER Extensions
    Dateiendungen ohne Punkt.

.PARAMETER Browser
    Chrome, Edge oder beide (Standard).

.PARAMETER Remove
    Die angegebenen Endungen und die URL wieder aus den Richtlinien entfernen.

.EXAMPLE
    .\office-autoopen.ps1
.EXAMPLE
    .\office-autoopen.ps1 -Browser Edge -WhatIf
.EXAMPLE
    .\office-autoopen.ps1 -Remove
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]   $Url        = 'csp.synaforce.com',
    [string[]] $Extensions = @('doc','docx','dot','dotx','rtf','odt','xls','xlsx','xlt','xltx','csv','ods','ppt','pptx','pps','ppsx','pot','potx','odp','vsd','vsdx'),
    [ValidateSet('Chrome','Edge')]
    [string[]] $Browser    = @('Chrome','Edge'),
    [switch]   $Remove
)

$ErrorActionPreference = 'Stop'

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -and -not $WhatIfPreference) {
    throw 'Bitte als Administrator ausführen (bzw. per Intune im Systemkontext).'
}

$roots = @{
    Chrome = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
    Edge   = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
}

# Listen-Richtlinie (Werte "1", "2", … vom Typ REG_SZ) lesen
function Get-PolicyList([string] $Key) {
    if (-not (Test-Path $Key)) { return @() }
    $item = Get-Item $Key
    $item.GetValueNames() |
        Where-Object { $_ -match '^\d+$' } |
        Sort-Object { [int]$_ } |
        ForEach-Object { [string]$item.GetValue($_) } |
        Where-Object { $_ }
}

# Listen-Richtlinie vollständig neu schreiben (lückenlos nummeriert)
function Set-PolicyList([string] $Key, [string[]] $Values) {
    if (-not (Test-Path $Key)) { New-Item -Path $Key -Force | Out-Null }
    $item = Get-Item $Key
    foreach ($n in $item.GetValueNames() | Where-Object { $_ -match '^\d+$' }) {
        Remove-ItemProperty -Path $Key -Name $n
    }
    $i = 1
    foreach ($v in $Values) {
        New-ItemProperty -Path $Key -Name "$i" -PropertyType String -Value $v -Force | Out-Null
        $i++
    }
}

function Update-PolicyList([string] $Key, [string[]] $Items, [bool] $RemoveItems) {
    $current = @(Get-PolicyList $Key)
    $norm    = $Items | ForEach-Object { $_.Trim().TrimStart('.').ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique
    if ($RemoveItems) {
        $new = @($current | Where-Object { $norm -notcontains $_.ToLowerInvariant() })
    } else {
        $new = @($current) + @($norm | Where-Object { ($current | ForEach-Object { $_.ToLowerInvariant() }) -notcontains $_ })
    }
    if (($current -join "`n") -eq ($new -join "`n")) { Write-Host "  unverändert: $Key"; return }
    if ($PSCmdlet.ShouldProcess($Key, ($(if ($RemoveItems) { 'Einträge entfernen' } else { 'Einträge ergänzen' })))) {
        if ($new.Count) { Set-PolicyList $Key $new } elseif (Test-Path $Key) { Remove-Item $Key -Recurse -Force }
        Write-Host "  geschrieben: $Key -> $($new -join ', ')"
    }
}

foreach ($b in $Browser) {
    $root = $roots[$b]
    Write-Host "$b"
    Update-PolicyList "$root\AutoOpenFileTypes"      $Extensions $Remove.IsPresent
    Update-PolicyList "$root\AutoOpenAllowedForURLs" @($Url)     $Remove.IsPresent
}

Write-Host ''
Write-Host 'Fertig. Wirksam nach Neustart des Browsers (oder in chrome://policy bzw. edge://policy auf "Richtlinien neu laden").'
