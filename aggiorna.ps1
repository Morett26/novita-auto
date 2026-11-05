# Rassegna Auto: scarica i feed RSS, aggiorna l'archivio e genera data.js per index.html.
# Il file resta solo ASCII perche' Windows PowerShell 5.1 legge gli script senza BOM come ANSI.

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Web.Extensions

$RetentionDays = 30
$SummaryLength = 220
$ArchivePath = Join-Path $PSScriptRoot 'archivio.json'
$DataPath = Join-Path $PSScriptRoot 'data.js'

# Per aggiungere una fonte basta una riga in questo elenco.
# lang: it/en. cat: forza la categoria. exclude: scarta i titoli non legati alle auto.
$Feeds = @(
    @{ name = 'alVolante';        lang = 'it'; url = 'https://www.alvolante.it/rss.xml' }
    @{ name = 'Motor1 Italia';    lang = 'it'; url = 'https://it.motor1.com/rss/news/all/' }
    @{ name = 'Autoblog';         lang = 'it'; url = 'https://www.autoblog.it/feed/' }
    @{ name = 'HDmotori';         lang = 'it'; url = 'https://www.hdmotori.it/feed/' }
    @{ name = 'Autotoday';        lang = 'it'; url = 'https://www.autotoday.it/feed/' }
    @{ name = 'Il Sole 24 Ore';   lang = 'it'; url = 'https://www.ilsole24ore.com/rss/motori.xml' }
    @{ name = 'la Repubblica';    lang = 'it'; url = 'https://www.repubblica.it/rss/motori/rss2.0.xml' }
    @{ name = 'Motorsport.com';   lang = 'it'; url = 'https://it.motorsport.com/rss/f1/news/'; cat = 'sport' }
    @{ name = 'Motor1';           lang = 'en'; url = 'https://www.motor1.com/rss/news/all/' }
    @{ name = 'Carscoops';        lang = 'en'; url = 'https://www.carscoops.com/feed/' }
    @{ name = 'Autocar';          lang = 'en'; url = 'https://www.autocar.co.uk/rss' }
    @{ name = 'Auto Express';     lang = 'en'; url = 'https://www.autoexpress.co.uk/rss' }
    @{ name = 'Car and Driver';   lang = 'en'; url = 'https://www.caranddriver.com/rss/all.xml/' }
    @{ name = 'InsideEVs';        lang = 'en'; url = 'https://insideevs.com/rss/news/all/' }
    @{ name = 'Electrek';         lang = 'en'; url = 'https://electrek.co/feed/'; exclude = 'e-?bike|electric bike|scooter|motorcycle|solar|Green Deals|wind|heat pump|Powerwall|storage|boat|podcast|lawn|mower' }
)

$ItalianMonths = @{ Gen = 'Jan'; Mag = 'May'; Giu = 'Jun'; Lug = 'Jul'; Ago = 'Aug'; Set = 'Sep'; Ott = 'Oct'; Dic = 'Dec' }

# La prima regola che corrisponde decide la categoria: l'ordine conta.
$CategoryRules = @(
    @{ cat = 'sport';     re = '\b(F1|Formula 1|MotoGP|WRC|WEC|Le Mans|Gran Premio|Grand Prix|Rally)\b' }
    @{ cat = 'prove';     re = '\bprov[ae]\b|abbiamo guidato|al volante|recensione|test drive|first drive|road test|\breview\b' }
    @{ cat = 'elettriche'; re = 'elettric|ibrid|electric|hybrid|\bEVs?\b|\bBEV\b|\bPHEV\b|ricarica|charging|\bbatter(ia|ie|y|ies)\b|Tesla' }
    @{ cat = 'mercato';   re = 'mercato|vendite|immatricolazion|incentiv|dazi|listino|stabiliment|bilancio|richiam|\bsales\b|tariff|recall|factory|market\b|Stellantis' }
    @{ cat = 'modelli';   re = '\bnuov[aeio]\b|debutt|anteprima|svela|restyling|concept|\bnew\b|reveal|unveil|facelift|debut|spy' }
)

function Get-FeedXml([string]$Url) {
    $request = [Net.HttpWebRequest]::Create($Url)
    $request.Timeout = 20000
    $request.ReadWriteTimeout = 20000
    $request.UserAgent = 'Mozilla/5.0 (compatible; RassegnaAuto/1.0)'
    $request.Accept = 'application/rss+xml, application/atom+xml, application/xml, text/xml, */*'
    $request.AutomaticDecompression = [Net.DecompressionMethods]::GZip -bor [Net.DecompressionMethods]::Deflate

    $response = $request.GetResponse()
    try {
        $buffer = New-Object IO.MemoryStream
        $response.GetResponseStream().CopyTo($buffer)
    } finally {
        $response.Close()
    }
    $buffer.Position = 0

    # XmlReader rispetta la codifica dichiarata dal feed; DTD e risorse esterne restano disattivate.
    $settings = New-Object Xml.XmlReaderSettings
    $settings.DtdProcessing = [Xml.DtdProcessing]::Ignore
    $settings.XmlResolver = $null
    $reader = [Xml.XmlReader]::Create($buffer, $settings)
    try {
        $doc = New-Object Xml.XmlDocument
        $doc.XmlResolver = $null
        $doc.Load($reader)
        return $doc
    } finally {
        $reader.Close()
    }
}

function Get-ChildText($Node, [string[]]$Names) {
    foreach ($name in $Names) {
        $child = $Node.SelectSingleNode("*[local-name()='$name']")
        if ($child -and $child.InnerText.Trim()) { return $child.InnerText.Trim() }
    }
    return ''
}

function Get-CleanText([string]$Html) {
    $text = [Net.WebUtility]::HtmlDecode($Html)
    $text = [regex]::Replace($text, '<[^>]*>', ' ')
    $text = [Net.WebUtility]::HtmlDecode($text)
    $text = [regex]::Replace($text, '(?i)\s*(The post .* appeared first on .*|Il post .* comparso .*|L.articolo .* proviene da .*)$', '')
    $text = [regex]::Replace($text, '(?i)\s*\[?(\.\.\.|…)\]?\s*$', '')
    return ([regex]::Replace($text, '\s+', ' ')).Trim()
}

function Get-Summary([string]$Text) {
    if ($Text.Length -le $SummaryLength) { return $Text }
    $cut = $Text.Substring(0, $SummaryLength)
    $lastSpace = $cut.LastIndexOf(' ')
    if ($lastSpace -gt 0) { $cut = $cut.Substring(0, $lastSpace) }
    return $cut.TrimEnd(' ', ',', ';', ':', '.') + [char]0x2026
}

function Get-IsoDate([string]$Value) {
    if (-not $Value) { return $null }
    $normalized = $Value.Trim() -replace '\bCEST$', '+02:00' -replace '\bCET$', '+01:00' -replace '\b(UT|UTC|GMT|Z)$', '+00:00' -replace '\bEDT$', '-04:00' -replace '\bEST$', '-05:00' -replace '\bPDT$', '-07:00' -replace '\bPST$', '-08:00'
    $normalized = $normalized -replace '([+-]\d{2})(\d{2})$', '$1:$2'
    # alVolante pubblica date in italiano ("Gio, 05 Nov 2026"): il giorno si scarta, il mese si traduce.
    $normalized = $normalized -replace '^[^\d,]+,\s*', ''
    foreach ($month in $ItalianMonths.Keys) {
        $normalized = $normalized -replace "\b$month\b", $ItalianMonths[$month]
    }
    $parsed = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse($normalized, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) {
        return $null
    }
    # Alcuni feed programmano articoli nel futuro: non devono finire in cima.
    $now = [DateTimeOffset]::UtcNow
    if ($parsed -gt $now) { $parsed = $now }
    return $parsed.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-WebUrl([string]$Value) {
    $url = $Value.Trim() -replace '&amp;', '&'
    if ($url.StartsWith('//')) { $url = "https:$url" }
    if ($url -match '^https?://') { return $url }
    return ''
}

function Get-ItemLink($Node) {
    $atomLink = $Node.SelectSingleNode("*[local-name()='link'][@href and (@rel='alternate' or not(@rel))]")
    if ($atomLink) { return Get-WebUrl $atomLink.GetAttribute('href') }
    return Get-WebUrl (Get-ChildText $Node @('link', 'guid'))
}

function Get-ItemImage($Node, [string]$Html) {
    foreach ($media in $Node.SelectNodes(".//*[(local-name()='content' or local-name()='thumbnail' or local-name()='enclosure') and @url]")) {
        $type = $media.GetAttribute('type') + $media.GetAttribute('medium')
        if ($type -and $type -notmatch 'image') { continue }
        $url = Get-WebUrl $media.GetAttribute('url')
        if ($url) { return $url }
    }
    $decoded = [Net.WebUtility]::HtmlDecode($Html)
    $match = [regex]::Match($decoded, '(?i)<img[^>]+src=["'']([^"'']+)["'']')
    if ($match.Success) { return Get-WebUrl $match.Groups[1].Value }
    return ''
}

function Get-Category([string]$Title, [string]$Summary, $Feed) {
    if ($Feed.cat) { return $Feed.cat }
    foreach ($text in @($Title, $Summary)) {
        foreach ($rule in $CategoryRules) {
            if ($text -match $rule.re) { return $rule.cat }
        }
    }
    return 'altro'
}

function Read-Feed($Feed) {
    $doc = Get-FeedXml $Feed.url
    $result = New-Object Collections.ArrayList
    foreach ($node in $doc.SelectNodes("//*[local-name()='item' or local-name()='entry']")) {
        $title = Get-CleanText (Get-ChildText $node @('title'))
        $link = Get-ItemLink $node
        if (-not $title -or -not $link) { continue }
        if ($Feed.exclude -and $title -match $Feed.exclude) { continue }

        $description = Get-ChildText $node @('description', 'summary')
        $body = Get-ChildText $node @('encoded', 'content')
        $summarySource = if ($description) { $description } else { $body }
        $summary = Get-Summary (Get-CleanText $summarySource)

        [void]$result.Add(@{
            title   = $title
            link    = ($link -replace '[?&]utm_[^#]*', '')
            source  = $Feed.name
            lang    = $Feed.lang
            cat     = Get-Category $title $summary $Feed
            date    = Get-IsoDate (Get-ChildText $node @('pubDate', 'published', 'updated', 'date'))
            summary = $summary
            image   = Get-ItemImage $node ($description + $body)
        })
    }
    return , $result
}

function Write-Utf8File([string]$Path, [string]$Content) {
    $temp = "$Path.tmp"
    [IO.File]::WriteAllText($temp, $Content, (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temp -Destination $Path -Force
}

$serializer = New-Object Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue

# 1. L'archivio conserva i giorni passati: i feed contengono solo gli ultimi articoli.
$byLink = @{}
if (Test-Path -LiteralPath $ArchivePath) {
    try {
        foreach ($item in $serializer.DeserializeObject([IO.File]::ReadAllText($ArchivePath))) {
            $byLink[$item['link']] = $item
        }
    } catch {
        Write-Warning "Archivio illeggibile, riparto da zero: $($_.Exception.Message)"
    }
}

# 2. Una fonte che fallisce non blocca le altre.
$nowIso = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
$sources = New-Object Collections.ArrayList
$failed = 0
foreach ($feed in $Feeds) {
    try {
        $feedItems = Read-Feed $feed
        foreach ($item in $feedItems) {
            $known = $byLink[$item['link']]
            if (-not $item['date']) {
                $item['date'] = if ($known) { $known['date'] } else { $nowIso }
            }
            $byLink[$item['link']] = $item
        }
        [void]$sources.Add(@{ name = $feed.name; lang = $feed.lang; ok = $true; count = $feedItems.Count })
        Write-Host ("OK      {0,-16} {1,4} articoli" -f $feed.name, $feedItems.Count)
    } catch {
        $failed++
        [void]$sources.Add(@{ name = $feed.name; lang = $feed.lang; ok = $false; count = 0 })
        Write-Host ("ERRORE  {0,-16} {1}" -f $feed.name, $_.Exception.Message)
    }
}

# 3. Elimina gli articoli scaduti e i doppioni con lo stesso titolo.
$cutoff = [DateTime]::UtcNow.AddDays(-$RetentionDays).ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
$seenTitles = New-Object 'Collections.Generic.HashSet[string]'
$items = New-Object Collections.ArrayList
foreach ($item in ($byLink.Values | Sort-Object { $_['date'] } -Descending)) {
    if ([string]::CompareOrdinal($item['date'], $cutoff) -lt 0) { continue }
    if ($seenTitles.Add($item['title'].ToLowerInvariant())) { [void]$items.Add($item) }
}

# 4. Se nessuna fonte risponde restano i dati precedenti.
if ($failed -eq $Feeds.Count) {
    Write-Host 'Nessuna fonte raggiungibile: dati precedenti non modificati. Controlla la connessione.'
    exit 1
}

$itemsJson = $serializer.Serialize($items)
Write-Utf8File $ArchivePath $itemsJson
Write-Utf8File $DataPath ('window.RASSEGNA = {"updated":"' + $nowIso + '","sources":' + $serializer.Serialize($sources) + ',"items":' + $itemsJson + '};')
Write-Host ''
Write-Host ("Fatto: {0} articoli in archivio, {1} fonti su {2} raggiunte." -f $items.Count, ($Feeds.Count - $failed), $Feeds.Count)
