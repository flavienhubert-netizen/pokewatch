# PokéWatch - surveille l'UPC et l'ETB Pokémon 30 ans sur les boutiques en ligne
# et envoie une notification (PC + téléphone via l'appli ntfy) dès qu'il y a du nouveau.
#
# Lancer :              double-clic sur lancer.bat
# Tester les alertes :  powershell -ExecutionPolicy Bypass -File pokewatch.ps1 -Test
# Voir l'état actuel :  powershell -ExecutionPolicy Bypass -File pokewatch.ps1 -Check   (un passage, sans alerte)
# Mode GitHub :         pwsh pokewatch.ps1 -Cloud   (un seul passage, notifications téléphone uniquement)

param([switch]$Test, [switch]$Check, [switch]$Cloud)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not $Cloud) { Add-Type -AssemblyName System.Windows.Forms, System.Drawing }

$ConfigPath = Join-Path $PSScriptRoot 'config.json'
$StatePath  = Join-Path $PSScriptRoot $(if ($Cloud) { 'etat-cloud.json' } else { 'etat.json' })
$cfg = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

# Le nom du canal ntfy reste privé : variable d'environnement (secret GitHub) ou fichier local non publié
$TopicFile = Join-Path $PSScriptRoot 'ntfy_topic.txt'
$Topic = if ($env:NTFY_TOPIC) { $env:NTFY_TOPIC.Trim() }
         elseif (Test-Path $TopicFile) { (Get-Content $TopicFile -Raw).Trim() }
         else { $cfg.ntfy_topic }

$UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36'
$Inv = [Globalization.CultureInfo]::InvariantCulture
$Tmp = [IO.Path]::GetTempPath()
$Edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
          "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
$EdgeProfile = Join-Path $Tmp 'pokewatch-edge'

# ---------- Reconnaissance des produits ----------

$Thirty = '\b30\s?(e|ème|eme|th)\b|\b30 ans\b|mentali|noctali'
$Types = @{
    upc = 'ultra[- ]?premium|\bupc\b'
    etb = "dresseur d.?\s?élite|\betb\b|elite trainer box"
}
# Versions étrangères et accessoires (protections vides, etc.) : ignorés
$Exclude = '\b(japonais|japanese|jap|jp|cor[ée]en|korean|kr|chinois|chinese|cn)\b|^\s*(protection|étui|etui|support|présentoir)|bo[iî]te vide|non inclus|empty box'

function Get-WatchFor($title) {
    $t = $title.ToLowerInvariant()
    if ($t -notmatch $Thirty -or $t -match $Exclude) { return $null }
    foreach ($w in $cfg.watches) {
        if ($w.enabled -and $t -match $Types[$w.id]) { return $w }
    }
    return $null
}

function Log($msg) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $msg) }

# ---------- Lecture des boutiques ----------

function Get-Text($url, $headers = @{}) {
    $r = Invoke-WebRequest -Uri $url -UseBasicParsing -UserAgent $UserAgent -TimeoutSec 25 -Headers $headers
    return [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
}

function New-Item2($shop, $title, $price, $available, $url) {
    [pscustomobject]@{ shop = $shop; title = $title; price = $price; available = $available; url = $url }
}

function Get-ShopifyItems($s) {
    foreach ($w in $cfg.watches | Where-Object enabled) {
        foreach ($q in $w.queries) {
            $u = "https://$($s.host)/search/suggest.json?q=$([uri]::EscapeDataString($q))&resources[type]=product&resources[limit]=10&resources[options][unavailable_products]=last"
            $j = Get-Text $u | ConvertFrom-Json
            foreach ($p in $j.resources.results.products) {
                $link = "https://$($s.host)" + ($p.url -replace '\?.*$', '')
                $price = if ($p.price) { [double]::Parse($p.price, $Inv) } else { $null }
                New-Item2 $s.name $p.title $price ([bool]$p.available) $link
            }
        }
    }
}

function Get-WooItems($s) {
    foreach ($w in $cfg.watches | Where-Object enabled) {
        foreach ($q in $w.queries) {
            $u = "https://$($s.host)/wp-json/wc/store/v1/products?search=$([uri]::EscapeDataString($q))&per_page=30"
            foreach ($p in (Get-Text $u | ConvertFrom-Json)) {
                $price = [double]$p.prices.price / [Math]::Pow(10, $p.prices.currency_minor_unit)
                New-Item2 $s.name ([Net.WebUtility]::HtmlDecode($p.name)) $price ([bool]($p.is_in_stock -and $p.is_purchasable)) $p.permalink
            }
        }
    }
}

function Get-PrestaItems($s) {
    foreach ($w in $cfg.watches | Where-Object enabled) {
        foreach ($q in $w.queries) {
            $u = "https://$($s.host)/recherche?controller=search&s=$([uri]::EscapeDataString($q))"
            $j = Get-Text $u @{ Accept = 'application/json'; 'X-Requested-With' = 'XMLHttpRequest' } | ConvertFrom-Json
            foreach ($p in $j.products) {
                New-Item2 $s.name $p.name ([double]$p.price_amount) ($p.availability -ne 'unavailable') $p.url
            }
        }
    }
}

# Grandes enseignes : pas d'info de stock fiable, on alerte quand le produit APPARAIT dans la recherche
function Get-PageItems($s) {
    if ($s.browser) { $raw = Get-PageEdge $s.url } else { $raw = Get-Text $s.url }
    if (-not $raw -or $raw.Length -lt 3000) { throw 'page vide ou bloquée par le site' }
    $page = [Net.WebUtility]::HtmlDecode($raw).ToLowerInvariant()
    foreach ($w in $cfg.watches | Where-Object { $_.enabled -and $_.id -in $s.watches }) {
        $type = $Types[$w.id]
        $m = [regex]::Match($page, "($type)[^<]{0,150}($Thirty)|($Thirty)[^<]{0,150}($type)")
        if ($m.Success -and $m.Value -notmatch $Exclude) {
            New-Item2 $s.name "$($w.name) (annonce : « $($m.Value.Trim()) »)" $null $true "$($s.url)#$($w.id)"
        }
    }
}

function Stop-LeftoverEdge {
    for ($i = 0; $i -lt 20; $i++) {
        $left = @(Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" |
                  Where-Object { $_.CommandLine -like "*pokewatch-edge*" })
        if ($left.Count -eq 0) { return }
        if ($i -ge 10) { $left | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } }
        Start-Sleep -Milliseconds 500
    }
}

function Invoke-EdgeDump($url) {
    Stop-LeftoverEdge
    $id = [guid]::NewGuid().ToString('N')
    $out = Join-Path $Tmp "pokewatch-$id.html"
    $err = Join-Path $Tmp "pokewatch-$id.log"
    $edgeArgs = @('--headless=new', '--disable-gpu', '--no-first-run', '--mute-audio',
                  "--user-data-dir=`"$EdgeProfile`"", '--virtual-time-budget=15000', '--dump-dom', "`"$url`"")
    try {
        $p = Start-Process $Edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden `
            -RedirectStandardOutput $out -RedirectStandardError $err
        if (-not $p.WaitForExit(60000)) { $p.Kill(); throw 'Edge : délai dépassé' }
        if (Test-Path $out) { return Get-Content $out -Raw -Encoding UTF8 }
        return $null
    } finally {
        Remove-Item $out, $err -ErrorAction SilentlyContinue
    }
}

function Get-PageEdge($url) {
    if (-not $Edge) { throw 'Microsoft Edge introuvable' }
    $raw = Invoke-EdgeDump $url
    if (-not $raw -or $raw.Length -lt 3000) { Start-Sleep 3; $raw = Invoke-EdgeDump $url }
    return $raw
}

# Renvoie les annonces UPC/ETB 30 ans d'une boutique (dédoublonnées)
function Get-ShopListings($s) {
    $items = switch ($s.type) {
        'shopify' { Get-ShopifyItems $s }
        'woo'     { Get-WooItems $s }
        'presta'  { Get-PrestaItems $s }
        'page'    { Get-PageItems $s }
    }
    $seen = @{}
    foreach ($it in $items) {
        if ($seen.ContainsKey($it.url)) { continue }
        $seen[$it.url] = $true
        $w = if ($s.type -eq 'page') { $cfg.watches | Where-Object { $it.url.EndsWith("#$($_.id)") } } else { Get-WatchFor $it.title }
        if ($w) {
            $it | Add-Member -NotePropertyName watch -NotePropertyValue $w
            $it | Add-Member -NotePropertyName ok -NotePropertyValue ($it.available -and ($null -eq $it.price -or $it.price -le $w.max_price))
            $it
        }
    }
}

# ---------- Notifications ----------

function Show-Toast($title, $msg) {
    if ($Cloud) { return }
    if ($script:LastIcon) { $script:LastIcon.Dispose() }
    $n = New-Object System.Windows.Forms.NotifyIcon
    $n.Icon = [System.Drawing.SystemIcons]::Information
    $n.Visible = $true
    $n.ShowBalloonTip(15000, $title, $msg, [System.Windows.Forms.ToolTipIcon]::Info)
    $script:LastIcon = $n
}

function Send-Ntfy($title, $msg, $url, $urgent) {
    if (-not $Topic) { return }
    # En-tête encodé (RFC 2047) pour garder les accents dans le titre
    $t64 = '=?UTF-8?B?' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($title)) + '?='
    $headers = @{ Title = $t64; Click = $url
                  Priority = $(if ($urgent) { 'urgent' } else { 'default' })
                  Tags = $(if ($urgent) { 'rotating_light' } else { 'eyes' }) }
    try {
        Invoke-RestMethod -Method Post -Uri "https://ntfy.sh/$Topic" `
            -Body ([Text.Encoding]::UTF8.GetBytes($msg)) -Headers $headers | Out-Null
    } catch { Log "Push téléphone impossible : $($_.Exception.Message)" }
}

function Format-Price($p) { if ($null -eq $p) { 'prix ?' } else { '{0:N2} €' -f $p } }

function Send-Alert($it, [switch]$NewListing) {
    $page = $it.url -replace '#[a-z]+$', ''
    if ($NewListing) {
        $title = "Nouvelle annonce $($it.watch.name) - $($it.shop)"
        $msg = "$($it.title) - $(Format-Price $it.price) - pas encore dispo"
        Log "+ $title : $msg"
        Show-Toast $title $msg
        Send-Ntfy $title $msg $page $false
        return
    }
    $title = "$($it.watch.name) DISPO chez $($it.shop) !"
    $msg = "$($it.title) - $(Format-Price $it.price)"
    Log "*** $title $msg -> $page"
    Send-Ntfy $title $msg $page $true
    if ($Cloud) { return }
    if ($cfg.open_browser_on_stock) { Start-Process $page }
    Show-Toast $title $msg
    1..5 | ForEach-Object { [console]::Beep(1200, 250); [console]::Beep(800, 250) }
}

# ---------- État (pour ne prévenir que de ce qui est nouveau) ----------

function Read-State {
    $s = @{}
    if (Test-Path $StatePath) {
        (Get-Content $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $s[$_.Name] = [bool]$_.Value }
    }
    return $s
}

function Save-State($s) {
    # Clés triées : le fichier ne change que s'il y a vraiment du nouveau
    $o = [ordered]@{}
    $s.Keys | Sort-Object | ForEach-Object { $o[$_] = $s[$_] }
    $o | ConvertTo-Json | Out-File $StatePath -Encoding utf8
}

# ---------- Programme principal ----------

if ($Test) {
    $w = [pscustomobject]@{ name = 'UPC 30 ans'; id = 'upc' }
    Send-Alert ([pscustomobject]@{ shop = 'TEST'; title = 'Ceci est un test PokeWatch'; price = 229.99; url = 'https://www.pokemon.com/fr'; watch = $w })
    Start-Sleep 5
    if ($script:LastIcon) { $script:LastIcon.Dispose() }
    return
}

$shops = @($cfg.shops | Where-Object enabled)
# Sur GitHub : pas de navigateur, donc pas de grandes enseignes. Sur le PC, si GitHub surveille
# déjà les boutiques spécialisées, le PC ne garde que les grandes enseignes (pas de double alerte).
# Les boutiques "pc_only" refusent les serveurs de GitHub : c'est le PC qui s'en charge.
if ($Cloud) { $shops = @($shops | Where-Object { $_.type -ne 'page' -and -not $_.pc_only }) }
elseif ($cfg.cloud_handles_shops -and -not $Check) { $shops = @($shops | Where-Object { $_.type -eq 'page' -or $_.pc_only }) }

if ($Check) {
    foreach ($s in $shops) {
        try {
            $list = @(Get-ShopListings $s)
            Log "$($s.name) : $($list.Count) annonce(s)"
            foreach ($it in $list) {
                $flag = if ($it.ok) { 'DISPO' } elseif ($it.available) { 'trop cher' } else { 'épuisé' }
                Log ("    [{0}] {1} | {2} | {3}" -f $flag, $it.title, (Format-Price $it.price), $it.watch.name)
            }
        } catch { Log "$($s.name) : ERREUR $($_.Exception.Message)" }
    }
    return
}

$state = Read-State
$firstRun = $state.Count -eq 0
$interval = [Math]::Max(1, [double]$cfg.interval_minutes)
Log "Surveillance de $($shops.Count) boutiques toutes les $interval min. Ctrl+C pour arrêter."

while ($true) {
    $found = 0
    foreach ($s in $shops) {
        try {
            foreach ($it in Get-ShopListings $s) {
                $found++
                $known = $state.ContainsKey($it.url)
                if ($it.ok -and -not ($known -and $state[$it.url])) { Send-Alert $it }
                elseif (-not $known -and -not $firstRun) { Send-Alert $it -NewListing }
                $state[$it.url] = [bool]$it.ok
            }
        } catch { Log "$($s.name) : $($_.Exception.Message)" }
        Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1500)
    }
    Save-State $state
    if ($Cloud) { Log "Tour terminé : $found annonce(s) 30 ans suivies."; break }
    $firstRun = $false
    Log "Tour terminé : $found annonce(s) 30 ans suivies. Prochain passage dans $interval min."
    Start-Sleep -Seconds ([int]($interval * 60) + (Get-Random -Maximum 30))
}
