<#
    =====================================================================
    TECH NEWS DAILY — HTML EMAIL VERSION (OUTLOOK CLIENT SEND)
    =====================================================================

    This version:
      - Generates a full HTML newsletter
      - Makes all article links clickable
      - Sends HTML email via Outlook Desktop
      - Produces a .html report file instead of markdown

    FRONT-END REQUIREMENTS:
      - Outlook Desktop installed
      - Outlook logged into your account
      - Forwarding rule in Outlook.com:
          From: yourself
          Subject contains: TechNewsDaily
          Forward to: your real email
    =====================================================================
#>

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$BasePath = "C:\TechNewsDaily"
$Date = (Get-Date).ToString("yyyy-MM-dd")
$OutputFile = "$BasePath\TechNews_$Date.html"
$LogFile = "$BasePath\TechNews.log"
$EmailFile = "$BasePath\email.txt"

$EmailFrom = Get-Content $EmailFile
$EmailTo = $EmailFrom

$EmailSubject = "TechNewsDaily"
$EmailBody = "<h2>Your Daily Tech News Report for $Date</h2><p>See attached full report.</p>"

# RSS/Atom/JSON feeds
$Feeds = @(
    "https://feeds.arstechnica.com/arstechnica/technology-lab",
    "https://www.wired.com/feed/rss",
    "https://techcrunch.com/feed/",
    "https://www.theverge.com/rss/index.xml",
    "https://www.engadget.com/rss.xml",
    "https://www.zdnet.com/news/rss.xml",
    "https://www.cnet.com/rss/news/",
    "https://krebsonsecurity.com/feed/",
    "https://www.bleepingcomputer.com/feed/",
    "https://www.darkreading.com/rss.xml",
    "https://www.microsoft.com/en-us/security/blog/feed/",
    "https://azure.microsoft.com/en-us/blog/feed/",
    "https://aws.amazon.com/about-aws/whats-new/recent/feed/",
    "https://cloud.google.com/blog/topics/rss.xml",
    "https://openai.com/blog/rss/",
    "https://ai.googleblog.com/feeds/posts/default",
    "https://developer.nvidia.com/blog/feed/"
)

# Clean summary helper
function Get-Summary {
    param($Text)
    if (-not $Text) { return "" }
    $clean = ($Text -replace '<[^>]+>', '') -replace '\s+', ' '
    $clean = $clean.Trim()
    $sentences = $clean -split '(?<=[\.!\?])\s+'
    if ($sentences.Count -gt 2) {
        $summary = ($sentences[0..1] -join ' ')
    } else {
        $summary = $clean
    }
    if ($summary.Length -gt 400) {
        $summary = $summary.Substring(0, 400) + "..."
    }
    return $summary
}

# Start HTML document
$Content = @"
<html>
<head>
<style>
body { font-family: Arial, sans-serif; line-height: 1.5; }
h1 { color: #333; }
h2 { color: #444; margin-top: 30px; }
.article { margin-bottom: 20px; padding-bottom: 10px; border-bottom: 1px solid #ccc; }
.title { font-size: 18px; font-weight: bold; }
.summary { font-style: italic; color: #555; }
</style>
</head>
<body>
<h1>Daily Technology News — $Date</h1>
"@

$Seen = @{}

foreach ($Feed in $Feeds) {
    try {
        $Headers = @{
            "User-Agent" = "Mozilla/5.0"
            "Accept"      = "*/*"
        }

        $response = Invoke-WebRequest -Uri $Feed -Headers $Headers -UseBasicParsing
        $raw = $response.Content.Trim()

        $Content += "<h2>Source: $Feed</h2>"

        # Detect feed type
        if ($raw.StartsWith("<?xml") -or $raw.Contains("<rss")) {
            $rss = [xml]$raw
            $items = $rss.rss.channel.item | Select-Object -First 10
        }
        elseif ($raw.Contains("<feed")) {
            $atom = [xml]$raw
            $items = $atom.feed.entry | Select-Object -First 10
        }
        elseif ($raw.StartsWith("{") -or $raw.StartsWith("[")) {
            $json = $raw | ConvertFrom-Json
            if ($json.items) {
                $items = $json.items | Select-Object -First 10
            } else {
                throw "JSON feed format not recognized"
            }
        }
        else {
            throw "Unsupported or blocked feed format"
        }

        foreach ($item in $items) {
            $title = $item.title
            if (-not $title) { continue }

            # Extract link
            if ($item.link.href) { $link = $item.link.href }
            elseif ($item.link) { $link = $item.link }
            elseif ($item.url) { $link = $item.url }
            else { $link = $Feed }

            # Extract description
            if ($item.summary) { $desc = $item.summary }
            elseif ($item.description) { $desc = $item.description }
            elseif ($item.content) { $desc = $item.content }
            else { $desc = "" }

            # Deduplicate
            $key = "$title|$link".ToLower()
            if ($Seen.ContainsKey($key)) { continue }
            $Seen[$key] = $true

            $summary = Get-Summary $desc

            # Add HTML article block
            $Content += @"
<div class='article'>
  <div class='title'><a href='$link'>$title</a></div>
  <div><a href='$link'>$link</a></div>
  <div class='summary'>$summary</div>
</div>
"@
        }
    }
    catch {
        Add-Content -Path $LogFile -Value "$(Get-Date) - Failed feed $Feed - $($_.Exception.Message)"
    }
}

# Close HTML
$Content += "</body></html>"

# Save HTML report
$Content | Out-File -FilePath $OutputFile -Encoding UTF8

Add-Content -Path $LogFile -Value "$(Get-Date) - Created $OutputFile"

# SEND EMAIL USING OUTLOOK CLIENT (HTML)
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    $Mail.To = $EmailTo
    $Mail.Subject = $EmailSubject
    $Mail.HTMLBody = $EmailBody
    $Mail.Attachments.Add($OutputFile)
    $Mail.Send()

    Add-Content -Path $LogFile -Value "$(Get-Date) - Email sent via Outlook client to $EmailTo"
}
catch {
    Add-Content -Path $LogFile -Value "$(Get-Date) - Email FAILED via Outlook client - $($_.Exception.Message)"
}

