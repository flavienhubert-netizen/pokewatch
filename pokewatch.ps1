# PokéWatch - surveille l'UPC et l'ETB Pokémon 30 ans sur les boutiques en ligne
# et envoie une notification (PC + téléphone via l'appli ntfy) dès qu'il y a du nouveau.
#
# Lancer :              double-clic sur lancer.bat
# Tester les alertes :  powershell -ExecutionPolicy Bypass -File pokewatch.ps1 -Test
# Voir l'état actuel :  powershell -ExecutionPolicy Bypass -File pokewatch.ps1 -Check   (un passage, sans alerte)
# Mode GitHub :         pwsh pokewatch.ps1 -Cloud   (un seul passage, notifications téléphone uniquement)
# Récap du jour :       pwsh pokewatch.ps1 -Cloud -Recap   (9h30 et 16h30 ; -Force pour l'envoyer tout de suite)

param([switch]$Test, [switch]$Check, [switch]$Cloud, [switch]$Recap, [switch]$Force)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not $Cloud) { Add-Type -AssemblyName System.Windows.Forms, System.Drawing }

$ConfigPath = Join-Path $PSScriptRoot 'config.json'
$StatePath  = Join-Path $PSScriptRoot $(if ($Cloud) { 'etat-cloud.json' } else { 'etat.json' })
# Journal des alertes envoyées (lu par le récap) : un fichier pour GitHub, un pour le PC
$JournalPath = Join-Path $PSScriptRoot $(if ($Cloud) { 'journal-cloud.json' } else { 'journal-pc.json' })
$RecapPath   = Join-Path $PSScriptRoot 'recap-cloud.json'
$cfg = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

# Le nom du canal ntfy reste privé : variable d'environnement (secret GitHub) ou fichier local non publié
$TopicFile = Join-Path $PSScriptRoot 'ntfy_topic.txt'
$Topic = if ($env:NTFY_TOPIC) { $env:NTFY_TOPIC.Trim() }
         elseif (Test-Path $TopicFile) { (Get-Content $TopicFile -Raw).Trim() }
         else { $cfg.ntfy_topic }
# Un nom de canal ntfy ne contient que lettres, chiffres, - et _ : on retire tout caractère parasite (BOM, retour ligne...)
if ($Topic) { $Topic = $Topic -replace '[^A-Za-z0-9_-]', '' }

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

function Send-Ntfy($title, $msg, $url, $urgent, $tags) {
    if (-not $Topic) { return }
    # En-tête encodé (RFC 2047) pour garder les accents dans le titre
    $t64 = '=?UTF-8?B?' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($title)) + '?='
    if (-not $tags) { $tags = if ($urgent) { 'rotating_light' } else { 'eyes' } }
    $headers = @{ Title = $t64; Tags = $tags
                  Priority = $(if ($urgent) { 'urgent' } else { 'default' }) }
    if ($url) { $headers.Click = $url }
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

# ---------- Journal des alertes (pour le récap) ----------

$script:NewEvents = @()

# PowerShell 7 transforme les dates JSON en DateTime, PowerShell 5 les laisse en texte : on gère les deux
function Get-EventTime($e) {
    if ($e.t -is [DateTime]) { return $e.t.ToUniversalTime() }
    return [DateTime]::Parse($e.t, $Inv, [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
}

function Read-Journal($path) {
    if (Test-Path $path) { foreach ($e in (Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json)) { $e } }
}

function Add-JournalEvent($it, $kind) {
    $script:NewEvents += [pscustomobject]@{
        t = [DateTime]::UtcNow.ToString('o'); kind = $kind; shop = $it.shop
        title = $it.title; price = $it.price; url = ($it.url -replace '#[a-z]+$', '') }
}

# Ajoute les nouvelles alertes au journal (on garde 3 jours). Renvoie $true s'il a changé.
function Save-Journal {
    if ($script:NewEvents.Count -eq 0) { return $false }
    $limit = [DateTime]::UtcNow.AddDays(-3)
    $all = @(@(Read-Journal $JournalPath) + $script:NewEvents | Where-Object { (Get-EventTime $_) -gt $limit } |
        ForEach-Object { [pscustomobject]@{ t = (Get-EventTime $_).ToString('o'); kind = $_.kind; shop = $_.shop
                                            title = $_.title; price = $_.price; url = $_.url } })
    ConvertTo-Json -InputObject $all -Depth 3 | Out-File $JournalPath -Encoding utf8
    $script:NewEvents = @()
    return $true
}

# Le PC envoie son journal sur GitHub pour que le récap inclue aussi ses alertes
function Sync-PcJournal {
    if (-not (Test-Path (Join-Path $PSScriptRoot '.git'))) { return }
    $git = (Get-Command git -ErrorAction SilentlyContinue).Source
    if (-not $git) { $git = 'C:\Program Files\Git\cmd\git.exe' }
    $ErrorActionPreference = 'Continue'
    Push-Location $PSScriptRoot
    try {
        & $git add journal-pc.json *> $null
        & $git commit -q -m "Journal du PC" *> $null
        & $git pull -q --rebase --autostash *> $null
        & $git push -q *> $null
        if ($LASTEXITCODE -ne 0) { Log "Journal non envoyé sur GitHub : le prochain récap ne verra pas ces alertes du PC." }
    } finally { Pop-Location }
}

function ConvertTo-Paris($utc) {
    foreach ($id in 'Europe/Paris', 'Romance Standard Time') {
        try { return [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($utc, $id) } catch { }
    }
    return $utc.AddHours(1)
}

function Format-Hour($d) { '{0:HH}h{0:mm}' -f $d }

# ---------- Récap de 9h30 et 16h30 ----------

function Send-Recap {
    $nowUtc = [DateTime]::UtcNow
    $now = ConvertTo-Paris $nowUtc
    $mins = $now.Hour * 60 + $now.Minute
    # GitHub lance parfois ses tâches en retard : on accepte une marge après l'heure prévue
    if ($Force) { $slot = 'test'; $label = 'Récap (test)' }
    elseif ($mins -ge 9 * 60 + 15 -and $mins -lt 12 * 60) { $slot = 'matin'; $label = 'Récap 9h30' }
    elseif ($mins -ge 16 * 60 + 15 -and $mins -lt 19 * 60) { $slot = 'aprem'; $label = 'Récap 16h30' }
    else { Log "Pas l'heure d'un récap ($(Format-Hour $now) à Paris)."; return }

    $key = '{0:yyyy-MM-dd}-{1}' -f $now, $slot
    $rs = if (Test-Path $RecapPath) { Get-Content $RecapPath -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }
    if (-not $Force -and $rs -and $rs.last_slot -eq $key) { Log "Récap $key déjà envoyé."; return }

    $since = if ($rs -and $rs.last_time) { Get-EventTime ([pscustomobject]@{ t = $rs.last_time }) }
             else { $nowUtc.AddHours($(if ($slot -eq 'matin') { -17 } else { -7 })) }
    $events = @(@(foreach ($f in 'journal-cloud.json', 'journal-pc.json') { Read-Journal (Join-Path $PSScriptRoot $f) }) |
                Where-Object { (Get-EventTime $_) -gt $since } | Sort-Object { Get-EventTime $_ })

    $sinceP = ConvertTo-Paris $since
    $sinceTxt = if ($sinceP.Date -eq $now.Date) { "Depuis $(Format-Hour $sinceP)" }
                elseif ($sinceP.Date -eq $now.Date.AddDays(-1)) { "Depuis hier $(Format-Hour $sinceP)" }
                else { "Depuis le {0:dd/MM} à $(Format-Hour $sinceP)" -f $sinceP }
    $nShops = @($cfg.shops | Where-Object enabled).Count

    if ($events.Count -eq 0) {
        $title = "$label : aucun pack aujourd'hui"
        $msg = "$sinceTxt, rien de nouveau sur les $nShops boutiques (UPC et ETB 30 ans). La surveillance continue."
        $click = $null; $tags = 'zzz'
    } else {
        $dispo = @($events | Where-Object { $_.kind -eq 'dispo' })
        $news  = @($events | Where-Object { $_.kind -eq 'annonce' })
        $title = "$label : $($dispo.Count) dispo, $($news.Count) nouvelle(s) annonce(s)"
        $lines = foreach ($e in $events | Select-Object -First 15) {
            $icon = if ($e.kind -eq 'dispo') { '🟢' } else { '🆕' }
            "$icon $(Format-Hour (ConvertTo-Paris (Get-EventTime $e))) $($e.shop) - $($e.title) - $(Format-Price $e.price)"
        }
        if ($events.Count -gt 15) { $lines += "... et $($events.Count - 15) autre(s)" }
        $msg = "$sinceTxt :`n" + ($lines -join "`n")
        $click = if ($dispo) { $dispo[-1].url } else { $events[-1].url }
        $tags = 'package'
    }

    Log "$title`n$msg"
    Send-Ntfy $title $msg $click $false $tags
    if (-not $Force) {
        [pscustomobject]@{ last_slot = $key; last_time = $nowUtc.ToString('o') } | ConvertTo-Json | Out-File $RecapPath -Encoding utf8
    }
}

# ---------- Programme principal ----------

if ($Test) {
    $w = [pscustomobject]@{ name = 'UPC 30 ans'; id = 'upc' }
    Send-Alert ([pscustomobject]@{ shop = 'TEST'; title = 'Ceci est un test PokeWatch'; price = 229.99; url = 'https://www.pokemon.com/fr'; watch = $w })
    Start-Sleep 5
    if ($script:LastIcon) { $script:LastIcon.Dispose() }
    return
}

if ($Recap) { Send-Recap; return }

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
                if ($it.ok -and -not ($known -and $state[$it.url])) { Send-Alert $it; Add-JournalEvent $it 'dispo' }
                elseif (-not $known -and -not $firstRun) { Send-Alert $it -NewListing; Add-JournalEvent $it 'annonce' }
                $state[$it.url] = [bool]$it.ok
            }
        } catch { Log "$($s.name) : $($_.Exception.Message)" }
        Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1500)
    }
    Save-State $state
    $journalChanged = Save-Journal
    if ($journalChanged -and -not $Cloud) { Sync-PcJournal }
    if ($Cloud) { Log "Tour terminé : $found annonce(s) 30 ans suivies."; break }
    $firstRun = $false
    Log "Tour terminé : $found annonce(s) 30 ans suivies. Prochain passage dans $interval min."
    Start-Sleep -Seconds ([int]($interval * 60) + (Get-Random -Maximum 30))
}
