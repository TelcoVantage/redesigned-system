#Requires -Version 5.1
<#
.SYNOPSIS
    Genesys Cloud CX - Agent Copilot knowledge-surfacing audit for voice conversations.

.DESCRIPTION
    For every voice conversation in a date range this script answers:
      1. Did Agent Copilot surface knowledge (suggestions), which articles, and did agents open / use / rate them?
      2. Which customer (and optionally agent) utterances SHOULD have matched an article but produced no
         suggestion, or only a low-confidence one?  (Transcript replay against the Knowledge Search API.)
      3. Which articles are dormant / low-performing versus the articles that surface most often?
      4. Which real call phrases could be added as alternative phrases to under-performing articles?
      5. Which recurring topics (wrap-up code x queue, utterance clusters, Knowledge Optimizer
         unanswered groups) have no matching article?

    Data sources (all Genesys Cloud Platform API, validated against the published swagger on 2026-11-06):
      - POST /api/v2/analytics/conversations/details/jobs (+ status + results)  -> voice conversation inventory
      - GET  /api/v2/conversations/{conversationId}/suggestions                   -> persisted Copilot suggestions
      - POST /api/v2/analytics/knowledge/aggregates/query                        -> knowledge usage counters
      - GET  /api/v2/knowledge/knowledgebases/{kbId}/documents                    -> article inventory + alternatives
      - GET  /api/v2/knowledge/knowledgebases/{kbId}/documents/feedback           -> article thumbs up / down
      - GET  /api/v2/knowledge/knowledgebases/{kbId}/unanswered/groups            -> Knowledge Optimizer unanswered
      - GET  /api/v2/speechandtextanalytics/conversations/{id}/communications/{commId}/transcripturl -> transcript
      - POST /api/v2/knowledge/knowledgebases/{kbId}/documents/search             -> replay utterances
      - GET  /api/v2/routing/queues, /api/v2/routing/wrapupcodes, /api/v2/users    -> name resolution

    The script is written for PowerShell Constrained Language Mode (CLM):
      - no Add-Type, no custom .NET classes, no COM, no Invoke-Expression, no SDK/modules
      - only Invoke-WebRequest / Invoke-RestMethod, built-in cmdlets, hashtables, PSObjects, Export-Csv
      - Base64 is implemented with core types because [Convert] method calls are blocked in CLM; objects are
        built with New-Object PSObject because the [pscustomobject] cast is also blocked on some builds.

    PII: utterance text is masked (digits redacted, truncated) unless -IncludeUtteranceText is supplied.

.PARAMETER StartDate
    Start of the reporting window (inclusive). Local time unless you append 'Z'. Converted to UTC for the API.
.PARAMETER EndDate
    End of the reporting window (exclusive).
.PARAMETER QueueId
    Optional queue ID(s). When supplied only conversations with a voice segment on those queues are included.
.PARAMETER KnowledgeBaseId
    The knowledge base used by Agent Copilot for the audited queues.
.PARAMETER OutputPath
    Folder for the CSV files, the log file and the PowerShell transcript. Created if missing.
.PARAMETER Region
    Genesys Cloud region domain. Sydney = mypurecloud.com.au (default).
.PARAMETER ClientId
    OAuth client-credentials client ID. Defaults to $env:GENESYS_CLIENT_ID.
.PARAMETER ClientSecret
    OAuth client secret as a SecureString. Defaults to $env:GENESYS_CLIENT_SECRET, otherwise prompted.
.PARAMETER IncludeUtteranceText
    OPT-IN: write full utterance text (customer speech = PII) to the CSVs. Default: masked and truncated.
.PARAMETER MaskLength
    Number of characters kept when masking utterances (default 60).
.PARAMETER SkipTranscriptReplay
    Do not download transcripts or replay utterances against knowledge search (fastest run; outputs 1, 2 and 5 only
    partially populated).
.PARAMETER SkipSuggestionsApi
    Do not call GET /conversations/{id}/suggestions (use knowledge aggregates only).
.PARAMETER SuggestionLookupMode
    'All' (default) calls GET /conversations/{id}/suggestions for every in-scope conversation (one call each).
    'AggregatesFirst' only calls it for conversations that the knowledge aggregates show had knowledge events,
    which is far cheaper on large volumes (tens of thousands of calls) at the cost of relying on aggregates.
.PARAMETER IncludeAgentUtterances
    Also replay agent (internal) utterances, not just customer (external) ones.
.PARAMETER MaxConversationsForReplay
    Cap on conversations whose transcripts are replayed (default 500). Conversations with NO live knowledge
    suggestion are replayed first because they are the most likely missed opportunities.
.PARAMETER MaxUtterancesPerConversation
    Longest N qualifying utterances per conversation sent to knowledge search (default 6).
.PARAMETER MinUtteranceWords
    Ignore utterances shorter than this many words (default 5).
.PARAMETER ReplayResultCount
    Number of ranked results requested per replayed utterance (default 3, max 30).
.PARAMETER MissedOpportunityMinConfidence
    Reporting cut-off: a replayed utterance whose best match is at/above this confidence but was never surfaced
    live is classified 'Missed' (default 0.5). This is a reporting label only, not a Copilot threshold.
.PARAMETER GapMaxConfidence
    Reporting cut-off: utterances whose best match is below this are treated as knowledge gaps (default 0.35).
.PARAMETER LowSurfaceThreshold
    Articles surfaced this many times or fewer in the window are 'Low' (default 3). 0 surfaces = 'Dormant'.
.PARAMETER AggregateAppType
    appType filter for knowledge aggregates (default 'Assistant' = Agent Copilot / Agent Assist). '' = no filter.
.PARAMETER TagReplaySearchesWithContext
    Send application/conversationContext on replay searches so they look like Copilot searches in Knowledge
    Optimizer. OFF by default so replays can be distinguished from (and do not pollute) live Copilot analytics.
.PARAMETER ApiBaseUri / LoginBaseUri
    Overrides for testing against a mock or going through a proxy. Normally derived from -Region.

.EXAMPLE
    .\Invoke-GenesysKnowledgeAudit.ps1 -StartDate '2026-10-01' -EndDate '2026-10-08' `
        -KnowledgeBaseId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' -OutputPath 'C:\Reports\KnowledgeAudit'

.NOTES
    See README.md for permissions, limitations and how to action the output.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [datetime] $StartDate,
    [Parameter(Mandatory = $true)] [datetime] $EndDate,
    [string[]] $QueueId,
    [Parameter(Mandatory = $true)] [string] $KnowledgeBaseId,
    [string] $OutputPath = (Join-Path -Path (Get-Location) -ChildPath 'KnowledgeAudit'),
    [string] $Region = 'mypurecloud.com.au',
    [string] $ClientId = $env:GENESYS_CLIENT_ID,
    [securestring] $ClientSecret,
    [switch] $IncludeUtteranceText,
    [int] $MaskLength = 60,
    [switch] $SkipTranscriptReplay,
    [switch] $SkipSuggestionsApi,
    [ValidateSet('All', 'AggregatesFirst')] [string] $SuggestionLookupMode = 'All',
    [switch] $IncludeAgentUtterances,
    [int] $MaxConversationsForReplay = 500,
    [int] $MaxUtterancesPerConversation = 6,
    [int] $MinUtteranceWords = 5,
    [ValidateRange(1, 30)] [int] $ReplayResultCount = 3,
    [double] $MissedOpportunityMinConfidence = 0.5,
    [double] $GapMaxConfidence = 0.35,
    [int] $LowSurfaceThreshold = 3,
    [string] $AggregateAppType = 'Assistant',
    [switch] $TagReplaySearchesWithContext,
    [string] $ApiBaseUri,
    [string] $LoginBaseUri
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------------------------
# region Embedded credentials (optional)
#   Fill these in to run the script without -ClientId/-ClientSecret or the GENESYS_CLIENT_ID / GENESYS_CLIENT_SECRET
#   environment variables. Precedence: parameter > environment variable > embedded value > interactive prompt.
#   WARNING: anything placed here is stored in clear text. Keep this file out of source control (or strip the
#   values before committing) and restrict NTFS permissions on it; the OAuth client should carry only the
#   read-only permissions listed in README.md.
# ---------------------------------------------------------------------------------------------------------------
$EmbeddedClientId     = ''      # e.g. '1a2b3c4d-....'
$EmbeddedClientSecret = ''      # e.g. 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region Script-level state
# ---------------------------------------------------------------------------------------------------------------
$script:ApiBase    = if ($ApiBaseUri)   { $ApiBaseUri.TrimEnd('/') }   else { "https://api.$Region" }
$script:LoginBase  = if ($LoginBaseUri) { $LoginBaseUri.TrimEnd('/') } else { "https://login.$Region" }
$script:Token      = $null
$script:TokenExpiry = [datetime]::MinValue
$script:LogFile    = $null
$script:RunStats      = @{ ApiCalls = 0; Retries = 0; RateLimited = 0; Warnings = 0 }
$script:Warnings   = @()
$script:IsPs7      = ($PSVersionTable.PSVersion.Major -ge 7)

# endregion

# ---------------------------------------------------------------------------------------------------------------
# region Logging
# ---------------------------------------------------------------------------------------------------------------
function Write-Log {
    <#  Console + file logger. Levels: INFO, WARN, ERROR, DEBUG. Never swallows errors: WARN/ERROR are counted
        and repeated in RunSummary.txt so nothing fails silently. #>
    param([string] $Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')] [string] $Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'WARN'  { Write-Warning $Message; $script:RunStats.Warnings++; $script:Warnings += $Message }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'DEBUG' { Write-Verbose $Message }
        default { Write-Host $line }
    }
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 }
}

function Add-Warning { param([string] $Message) Write-Log -Message $Message -Level WARN }
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region Encoding helpers (CLM-safe: [Convert] / [Uri] / [Text.Encoding] method calls are not allowed in CLM)
# ---------------------------------------------------------------------------------------------------------------
function ConvertTo-Base64Ascii {
    <# Base64 of an ASCII string, implemented with core types only. Used for the OAuth Basic header. #>
    param([Parameter(Mandatory = $true)] [string] $Text)
    $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
    $bytes = @()
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int][char]$ch
        if ($code -gt 127) { throw 'ConvertTo-Base64Ascii: only ASCII input is supported (client id/secret).' }
        $bytes += $code
    }
    $out = ''
    for ($i = 0; $i -lt $bytes.Count; $i += 3) {
        $b0 = $bytes[$i]
        $b1 = if (($i + 1) -lt $bytes.Count) { $bytes[$i + 1] } else { 0 }
        $b2 = if (($i + 2) -lt $bytes.Count) { $bytes[$i + 2] } else { 0 }
        $n = ($b0 -shl 16) -bor ($b1 -shl 8) -bor $b2
        $out += $alphabet[(($n -shr 18) -band 63)]
        $out += $alphabet[(($n -shr 12) -band 63)]
        $out += if (($i + 1) -lt $bytes.Count) { $alphabet[(($n -shr 6) -band 63)] } else { '=' }
        $out += if (($i + 2) -lt $bytes.Count) { $alphabet[($n -band 63)] } else { '=' }
    }
    return $out
}

function ConvertTo-UrlEncoded {
    <# RFC 3986 percent-encoding with manual UTF-8 for non-ASCII (no [Uri]::EscapeDataString in CLM). #>
    param([string] $Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    $sb = ''
    foreach ($ch in $Value.ToCharArray()) {
        $code = [int][char]$ch
        if ($ch -match '[A-Za-z0-9\-_.~]') { $sb += $ch }
        elseif ($code -lt 128)   { $sb += '%' + ('{0:X2}' -f $code) }
        elseif ($code -lt 2048)  { $sb += '%' + ('{0:X2}' -f (192 -bor ($code -shr 6))) + '%' + ('{0:X2}' -f (128 -bor ($code -band 63))) }
        else                     { $sb += '%' + ('{0:X2}' -f (224 -bor ($code -shr 12))) + '%' + ('{0:X2}' -f (128 -bor (($code -shr 6) -band 63))) + '%' + ('{0:X2}' -f (128 -bor ($code -band 63))) }
    }
    return $sb
}

function ConvertTo-QueryString {
    param([hashtable] $Query)
    if ($null -eq $Query -or $Query.Count -eq 0) { return '' }
    $parts = @()
    foreach ($k in $Query.Keys) {
        $v = $Query[$k]
        if ($null -eq $v -or "$v" -eq '') { continue }
        $parts += (ConvertTo-UrlEncoded $k) + '=' + (ConvertTo-UrlEncoded "$v")
    }
    if ($parts.Count -eq 0) { return '' }
    return '?' + ($parts -join '&')
}

function Get-Prop {
    <# Safe property read on ConvertFrom-Json objects (StrictMode-safe). Returns $null when absent. #>
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [hashtable]) { if ($Object.ContainsKey($Name)) { return $Object[$Name] } else { return $null } }
    if ($Object.PSObject.Properties.Name -contains $Name) { return $Object.$Name }
    return $null
}

function Get-UtcIso { param([datetime] $Date) return $Date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ') }
function ConvertTo-IsoText {
    <# PowerShell 7 ConvertFrom-Json turns ISO dates into [datetime]; normalise everything to ISO-8601 UTC text
       so the CSVs are culture-independent for Power BI. #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    return "$Value"
}

function Get-MinInt { param([int] $A, [int] $B) if ($A -lt $B) { return $A } else { return $B } }
function Get-MaxInt { param([int] $A, [int] $B) if ($A -gt $B) { return $A } else { return $B } }
function ConvertTo-Rounded {
    <# 3-decimal rounding without [System.Math] (blocked in CLM). #>
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    return [double]('{0:F3}' -f [double]$Value)
}
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region Authentication
# ---------------------------------------------------------------------------------------------------------------
function Get-GcToken {
    <#  OAuth 2.0 client-credentials grant: POST https://login.{region}/oauth/token
        The secret is only ever held in a SecureString / PSCredential and sent as a Basic header. #>
    param([string] $Id, [securestring] $Secret)
    $cred   = [pscredential]::new($Id, $Secret)
    $plain  = $cred.GetNetworkCredential().Password
    $basic  = ConvertTo-Base64Ascii ("{0}:{1}" -f $Id, $plain)
    $plain  = $null
    $uri    = "$script:LoginBase/oauth/token"
    Write-Log "Requesting OAuth token from $uri"
    $resp = Invoke-RestMethod -Method Post -Uri $uri -Headers @{ Authorization = "Basic $basic" } `
        -ContentType 'application/x-www-form-urlencoded' -Body 'grant_type=client_credentials'
    if (-not (Get-Prop $resp 'access_token')) { throw 'OAuth token response did not contain access_token.' }
    $script:Token = $resp.access_token
    $ttl = Get-Prop $resp 'expires_in'; if (-not $ttl) { $ttl = 3600 }
    $script:TokenExpiry = (Get-Date).AddSeconds([int]$ttl - 300)
    Write-Log ("Token acquired, expires in {0}s" -f $ttl)
}

function Resolve-ClientSecret {
    <# Secret precedence: -ClientSecret parameter > GENESYS_CLIENT_SECRET env var > $EmbeddedClientSecret > prompt. #>
    param([securestring] $Provided)
    if ($Provided) { return $Provided }
    if ($env:GENESYS_CLIENT_SECRET) { return (ConvertTo-SecureString -String $env:GENESYS_CLIENT_SECRET -AsPlainText -Force) }
    if (-not [string]::IsNullOrWhiteSpace($EmbeddedClientSecret)) {
        Write-Log 'Using the client secret embedded in the script file.' -Level WARN
        return (ConvertTo-SecureString -String $EmbeddedClientSecret -AsPlainText -Force)
    }
    return (Read-Host -Prompt 'Enter Genesys Cloud OAuth client secret' -AsSecureString)
}
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region API wrapper with retry / backoff / 429 handling
# ---------------------------------------------------------------------------------------------------------------
function Invoke-GcApi {
    <#  Single choke point for every Platform API call.
        - Adds the bearer token, JSON headers, query string encoding.
        - Honours Retry-After on HTTP 429, exponential backoff (2,4,8,16,32s) on 429/5xx/network errors.
        - Refreshes the token once on 401.
        - 404 returns $null only when -AllowNotFound is set, otherwise throws (no silent failures).
        - Works on Windows PowerShell 5.1 (WebException path) and PowerShell 7 (-SkipHttpErrorCheck path). #>
    param(
        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')] [string] $Method = 'GET',
        [string] $Path,
        [hashtable] $Query,
        [object] $Body,
        [string] $AbsoluteUri,
        [switch] $AllowNotFound,
        [switch] $NoAuth,
        [switch] $RawContent
    )
    $uri = if ($AbsoluteUri) { $AbsoluteUri } else { $script:ApiBase + $Path + (ConvertTo-QueryString $Query) }
    $jsonBody = $null
    if ($null -ne $Body) { $jsonBody = $Body | ConvertTo-Json -Depth 20 -Compress }

    $maxAttempts = 6
    $attempt = 0
    $authRetried = $false
    while ($true) {
        $attempt++
        if (-not $NoAuth -and (Get-Date) -gt $script:TokenExpiry) { Get-GcToken -Id $script:ClientIdValue -Secret $script:ClientSecretValue }
        $headers = @{ Accept = 'application/json' }
        if (-not $NoAuth) { $headers['Authorization'] = "Bearer $script:Token" }

        $status = 0; $content = ''; $retryAfter = $null; $netError = $null
        $script:RunStats.ApiCalls++
        try {
            $iwr = @{ Uri = $uri; Method = $Method; Headers = $headers; ErrorAction = 'Stop'; TimeoutSec = 120 }
            if ($jsonBody) { $iwr['Body'] = $jsonBody; $iwr['ContentType'] = 'application/json' }
            if ($script:IsPs7) {
                $resp = Invoke-WebRequest @iwr -SkipHttpErrorCheck
                $status = [int]$resp.StatusCode
                $content = $resp.Content
                try { if ($resp.Headers -and ($resp.Headers.Keys -contains 'Retry-After')) { $retryAfter = @($resp.Headers['Retry-After'])[0] } } catch { $retryAfter = $null }
            }
            else {
                try {
                    $resp = Invoke-WebRequest @iwr -UseBasicParsing
                    $status = [int]$resp.StatusCode
                    $content = $resp.Content
                }
                catch [System.Net.WebException] {
                    $r = $_.Exception.Response
                    if ($null -eq $r) { throw }
                    $status = [int]$r.StatusCode
                    try { $retryAfter = $r.Headers['Retry-After'] } catch { $retryAfter = $null }
                    $content = ''
                }
            }
        }
        catch {
            $netError = $_.Exception.Message
        }

        if ($null -eq $netError -and $status -ge 200 -and $status -lt 300) {
            if ($RawContent) { return $content }
            if ([string]::IsNullOrWhiteSpace($content)) { return $null }
            return ($content | ConvertFrom-Json)
        }
        if ($null -eq $netError -and $status -eq 404 -and $AllowNotFound) { return $null }
        if ($null -eq $netError -and $status -eq 401 -and -not $authRetried -and -not $NoAuth) {
            Write-Log 'HTTP 401 - refreshing token once' -Level DEBUG
            $authRetried = $true
            Get-GcToken -Id $script:ClientIdValue -Secret $script:ClientSecretValue
            continue
        }
        $retryable = ($null -ne $netError) -or ($status -eq 429) -or ($status -ge 500 -and $status -le 504)
        if ($retryable -and $attempt -lt $maxAttempts) {
            $script:RunStats.Retries++
            $delay = @(2, 4, 8, 16, 32, 64)[$attempt - 1]   # exponential backoff ([System.Math] is blocked in CLM)
            if ($status -eq 429) {
                $script:RunStats.RateLimited++
                $ra = 0; if ($retryAfter) { $ra = [int]("$retryAfter".Trim()) }
                if ($ra -gt $delay) { $delay = $ra }
                Write-Log ("HTTP 429 on {0} {1} - waiting {2}s (Retry-After={3})" -f $Method, $Path, $delay, $retryAfter) -Level DEBUG
            }
            else {
                Write-Log ("Transient failure on {0} {1}: status={2} error={3} - retry in {4}s" -f $Method, $Path, $status, $netError, $delay) -Level DEBUG
            }
            Start-Sleep -Seconds $delay
            continue
        }
        $snippet = if ($content) { "$content".Substring(0, (Get-MinInt 500 "$content".Length)) } else { $netError }
        throw ("API call failed: {0} {1} -> HTTP {2}. {3}" -f $Method, ($(if ($AbsoluteUri) { '<absolute uri>' } else { $Path })), $status, $snippet)
    }
}

function Get-GcCursorPages {
    <#  Generic walker for Genesys "after" cursor listings (knowledge documents, feedback, suggestions).
        Those endpoints return nextUri (a relative /api/v2 path including the after= cursor); we follow it. #>
    param([string] $Path, [hashtable] $Query, [switch] $AllowNotFound)
    $all = @()
    $q = @{}; if ($Query) { foreach ($k in $Query.Keys) { $q[$k] = $Query[$k] } }
    $page = Invoke-GcApi -Method GET -Path $Path -Query $q -AllowNotFound:$AllowNotFound
    $guard = 0
    while ($null -ne $page) {
        $ents = Get-Prop $page 'entities'
        if ($ents) { $all += @($ents) }
        $next = Get-Prop $page 'nextUri'
        if (-not $next) { break }
        $guard++; if ($guard -gt 5000) { throw "Pagination guard tripped for $Path" }
        $page = Invoke-GcApi -Method GET -Path $next
    }
    return $all
}

function Get-GcNumberedPages {
    <# Walker for pageNumber/pageSize listings (routing queues, wrap-up codes). #>
    param([string] $Path, [hashtable] $Query, [int] $PageSize = 100)
    $all = @(); $pageNumber = 1
    while ($true) {
        $q = @{ pageSize = $PageSize; pageNumber = $pageNumber }
        if ($Query) { foreach ($k in $Query.Keys) { $q[$k] = $Query[$k] } }
        $page = Invoke-GcApi -Method GET -Path $Path -Query $q
        $ents = @(Get-Prop $page 'entities')
        if ($ents.Count -gt 0) { $all += $ents }
        $pageCount = Get-Prop $page 'pageCount'
        if ($ents.Count -lt $PageSize) { break }
        if ($pageCount -and $pageNumber -ge [int]$pageCount) { break }
        $pageNumber++
        if ($pageNumber -gt 1000) { throw "Pagination guard tripped for $Path" }
    }
    return $all
}
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region PII masking and text utilities
# ---------------------------------------------------------------------------------------------------------------
function Protect-Text {
    <#  Default output masking for customer speech. Digit runs (account numbers, phone numbers, DOB), e-mail
        addresses and long spelled-digit sequences are redacted, then the text is truncated to -MaskLength.
        -IncludeUtteranceText disables masking (opt-in). #>
    param([string] $Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    if ($IncludeUtteranceText) { return $Text }
    $t = $Text -replace '[\w\.\-\+]+@[\w\.\-]+\.\w+', '[email]'
    $t = $t -replace '\d[\d\s\-]{2,}\d', '###'
    $t = $t -replace '\d{3,}', '###'
    if ($t.Length -gt $MaskLength) { $t = $t.Substring(0, $MaskLength) + '...' }
    return $t
}

$script:StopWords = @('the', 'a', 'an', 'and', 'or', 'but', 'if', 'so', 'to', 'of', 'in', 'on', 'at', 'for', 'with', 'about', 'from', 'by',
    'is', 'are', 'was', 'were', 'be', 'been', 'being', 'am', 'i', 'im', 'ive', 'me', 'my', 'we', 'our', 'you', 'your', 'youre', 'he', 'she', 'it', 'its', 'they',
    'them', 'their', 'this', 'that', 'these', 'those', 'there', 'here', 'what', 'which', 'who', 'when', 'where', 'why', 'how', 'can', 'could', 'would',
    'should', 'will', 'just', 'like', 'um', 'uh', 'yeah', 'yes', 'no', 'ok', 'okay', 'hi', 'hello', 'thanks', 'thank', 'please', 'have', 'has', 'had',
    'do', 'does', 'did', 'dont', 'not', 'get', 'got', 'want', 'wanted', 'need', 'know', 'think', 'say', 'said', 'see', 'one', 'also', 'because', 'then',
    'than', 'as', 'up', 'out', 'into', 'over', 'all', 'some', 'any', 'more', 'very', 'really', 'actually', 'basically', 'sort', 'kind', 'bit', 'lot',
    'going', 'go', 'call', 'calling', 'called', 'speak', 'talk', 'mean', 'well', 'right', 'now', 'today', 'thing', 'things', 'something', 'anything',
    'number', 'name', 'hey', 'oh', 'ah', 'sorry', 'good', 'morning', 'afternoon', 'mate', 'let', 'lets', 'still', 'back', 'again', 'through')

function Get-NormalizedText {
    param([string] $Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = $Text.ToLowerInvariant() -replace "[^a-z0-9' ]", ' '
    $t = $t -replace "'", ''
    return (($t -split '\s+') | Where-Object { $_ -ne '' }) -join ' '
}

function Get-SignificantTokens {
    param([string] $NormalizedText)
    if ([string]::IsNullOrEmpty($NormalizedText)) { return @() }
    return @(($NormalizedText -split ' ') | Where-Object { $_.Length -gt 2 -and ($script:StopWords -notcontains $_) -and ($_ -notmatch '^\d+$') })
}

function Get-WordCount { param([string] $Text) if ([string]::IsNullOrWhiteSpace($Text)) { return 0 }; return @(($Text.Trim()) -split '\s+').Count }

function Join-Unique { param($Values, [string] $Separator = '; ') return (@($Values | Where-Object { $null -ne $_ -and "$_" -ne '' } | Select-Object -Unique) -join $Separator) }

function Get-DateChunks {
    param([datetime] $From, [datetime] $To, [int] $Days)
    $chunks = @()
    $cur = $From
    while ($cur -lt $To) {
        $next = $cur.AddDays($Days); if ($next -gt $To) { $next = $To }
        $chunks += , @($cur, $next)
        $cur = $next
    }
    return $chunks
}
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region Data retrieval - conversations (Analytics Conversation Details async job)
# ---------------------------------------------------------------------------------------------------------------
function Get-GcVoiceConversations {
    <#  POST /api/v2/analytics/conversations/details/jobs - asynchronous job for large ranges.
        segmentFilters mediaType=voice (+ optional queueId OR-clause). Results paged with cursor.
        Note: the jobs endpoint returns conversations whose lifetime overlaps the interval; we post-filter on
        conversationStart to keep the window precise. #>
    param([datetime] $From, [datetime] $To, [string[]] $Queues)

    $segmentPredicates = @(@{ type = 'dimension'; dimension = 'mediaType'; operator = 'matches'; value = 'voice' })
    $segmentFilter = @{ type = 'and'; predicates = $segmentPredicates }
    if ($Queues -and $Queues.Count -gt 0) {
        $qPreds = @(); foreach ($q in $Queues) { $qPreds += @{ type = 'dimension'; dimension = 'queueId'; operator = 'matches'; value = $q } }
        $segmentFilter = @{ type = 'and'; predicates = $segmentPredicates; clauses = @(@{ type = 'or'; predicates = $qPreds }) }
    }
    $body = @{
        interval       = ('{0}/{1}' -f (Get-UtcIso $From), (Get-UtcIso $To))
        order          = 'asc'
        orderBy        = 'conversationStart'
        segmentFilters = @($segmentFilter)
        startOfDayIntervalMatching = $false
    }
    Write-Log "Submitting analytics conversation details job for $($body.interval)"
    $job = Invoke-GcApi -Method POST -Path '/api/v2/analytics/conversations/details/jobs' -Body $body
    $jobId = Get-Prop $job 'jobId'
    if (-not $jobId) { throw 'Analytics job did not return a jobId.' }

    # Poll job status until FULFILLED (FAILED / CANCELLED / EXPIRED raise).
    $state = 'QUEUED'; $polls = 0
    while ($state -in @('QUEUED', 'PENDING')) {
        Start-Sleep -Seconds 3
        $status = Invoke-GcApi -Method GET -Path "/api/v2/analytics/conversations/details/jobs/$jobId"
        $state = Get-Prop $status 'state'
        $polls++
        if ($polls % 10 -eq 0) { Write-Log "Job $jobId state=$state (poll $polls)" }
        if ($polls -gt 600) { throw "Analytics job $jobId did not complete in time (state=$state)." }
    }
    if ($state -ne 'FULFILLED') { throw ("Analytics job {0} ended in state {1}: {2}" -f $jobId, $state, (Get-Prop $status 'errorMessage')) }

    $convs = @(); $cursor = $null; $pages = 0
    while ($true) {
        $q = @{ pageSize = 500 }; if ($cursor) { $q['cursor'] = $cursor }
        $page = Invoke-GcApi -Method GET -Path "/api/v2/analytics/conversations/details/jobs/$jobId/results" -Query $q
        $batch = @(Get-Prop $page 'conversations')
        if ($batch.Count -gt 0) { $convs += $batch }
        $pages++
        Write-Progress -Activity 'Downloading conversation details' -Status ("{0} conversations ({1} pages)" -f $convs.Count, $pages)
        $cursor = Get-Prop $page 'cursor'
        if (-not $cursor) { break }
    }
    Write-Progress -Activity 'Downloading conversation details' -Completed
    Write-Log ("Analytics job returned {0} conversations" -f $convs.Count)
    return $convs
}

function ConvertTo-ConversationSummary {
    <#  Flattens an AnalyticsConversation into the fields we join on: agents, queues, wrap-up codes and the
        agent voice communication IDs (session.sessionId) needed for transcripts. #>
    param($Conv)
    $agentIds = @(); $queueIds = @(); $wrapCodes = @(); $commIds = @(); $hasAgentInteract = $false
    foreach ($p in @(Get-Prop $Conv 'participants')) {
        $purpose = Get-Prop $p 'purpose'
        foreach ($s in @(Get-Prop $p 'sessions')) {
            if ((Get-Prop $s 'mediaType') -ne 'voice') { continue }
            foreach ($seg in @(Get-Prop $s 'segments')) {
                $qid = Get-Prop $seg 'queueId'; if ($qid) { $queueIds += $qid }
                $wc = Get-Prop $seg 'wrapUpCode'; if ($wc) { $wrapCodes += $wc }
                if ($purpose -eq 'agent' -and (Get-Prop $seg 'segmentType') -eq 'interact') { $hasAgentInteract = $true }
            }
            if ($purpose -eq 'agent') {
                $uid = Get-Prop $p 'userId'; if ($uid) { $agentIds += $uid }
                $sid = Get-Prop $s 'sessionId'; if ($sid) { $commIds += $sid }
            }
        }
    }
    return New-Object -TypeName PSObject -Property ([ordered]@{
        ConversationId   = Get-Prop $Conv 'conversationId'
        Start            = Get-Prop $Conv 'conversationStart'
        End              = Get-Prop $Conv 'conversationEnd'
        Direction        = Get-Prop $Conv 'originatingDirection'
        AgentIds         = @($agentIds | Select-Object -Unique)
        QueueIds         = @($queueIds | Select-Object -Unique)
        WrapUpCodeIds    = @($wrapCodes | Select-Object -Unique)
        CommunicationIds = @($commIds | Select-Object -Unique)
        KnowledgeBaseIds = @(Get-Prop $Conv 'knowledgeBaseIds')
        HasAgentInteract = $hasAgentInteract
    })
}
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region Data retrieval - Copilot suggestions, knowledge aggregates, feedback, unanswered groups
# ---------------------------------------------------------------------------------------------------------------
$script:KnowledgeSuggestionTypes = @('Faq', 'Article', 'KnowledgeArticle', 'KnowledgeSearch', 'SuggestedKnowledgeAnswer')

function Get-GcConversationSuggestions {
    <#  GET /api/v2/conversations/{conversationId}/suggestions (permission conversation:suggestion:view)
        Returns the persisted Agent Copilot suggestions with state (Suggested/Accepted/Dismissed/Rated),
        triggerType, confidence, the document reference and the query statement Copilot used.
        Retention of this store is not documented - see README. #>
    param([string] $ConversationId)
    $rows = @()
    $ents = Get-GcCursorPages -Path "/api/v2/conversations/$ConversationId/suggestions" -Query @{ pageSize = 200 } -AllowNotFound
    foreach ($s in @($ents)) {
        $type = Get-Prop $s 'type'
        $docId = $null; $title = $null; $conf = $null; $searchId = $null
        $ks = Get-Prop $s 'knowledgeSearch'
        $ka = Get-Prop $s 'knowledgeArticle'
        $ar = Get-Prop $s 'article'
        $fq = Get-Prop $s 'faq'
        $ga = Get-Prop $s 'suggestedKnowledgeAnswer'
        if ($ks) { $docId = Get-Prop (Get-Prop $ks 'document') 'id'; $title = Get-Prop $ks 'title'; $conf = Get-Prop $ks 'confidence'; $searchId = Get-Prop $ks 'searchId' }
        elseif ($ka) { $docId = Get-Prop (Get-Prop $ka 'document') 'id'; $title = Get-Prop $ka 'title' }
        elseif ($ar) { $title = Get-Prop $ar 'title'; $conf = Get-Prop $ar 'confidence'; $docId = Get-Prop $ar 'uri' }
        elseif ($fq) { $title = Get-Prop $fq 'question'; $conf = Get-Prop $fq 'confidence'; $docId = Get-Prop $fq 'sourceUri' }
        elseif ($ga) {
            $searchId = Get-Prop $ga 'searchId'
            $chunks = @(Get-Prop $ga 'suggestedSearchChunks')
            if ($chunks.Count -gt 0) { $docId = Get-Prop (Get-Prop $chunks[0] 'document') 'id'; $title = Get-Prop $chunks[0] 'title'; $conf = Get-Prop $chunks[0] 'confidence' }
        }
        $ctx = Get-Prop $s 'context'
        $rows += New-Object -TypeName PSObject -Property ([ordered]@{
            ConversationId  = $ConversationId
            SuggestionId    = Get-Prop $s 'id'
            Type            = $type
            IsKnowledge     = ($script:KnowledgeSuggestionTypes -contains $type)
            State           = Get-Prop $s 'state'
            TriggerType     = Get-Prop $s 'triggerType'
            DateCreated     = Get-Prop $s 'dateCreated'
            DocumentId      = $docId
            Title           = $title
            Confidence      = $conf
            SearchId        = $searchId
            ParticipantType = Get-Prop $ctx 'participantType'
            UtteranceId     = Get-Prop (Get-Prop $ctx 'utterance') 'id'
            QueryStatement  = Get-Prop $ctx 'queryStatement'
            QueueId         = Get-Prop (Get-Prop $ctx 'queue') 'id'
        })
    }
    return $rows
}

function Invoke-GcKnowledgeAggregate {
    <#  POST /api/v2/analytics/knowledge/aggregates/query (permission analytics:knowledgeAggregate:view)
        Returns a flat list of {Group(hashtable), Metric, Count} across the whole window.
        The interval is split into chunks (default 7 days, falling back to 1 day on HTTP 400) because the
        service limits the interval length / granularity for aggregate queries. #>
    param([datetime] $From, [datetime] $To, [string[]] $GroupBy, [string[]] $Metrics, [hashtable[]] $ExtraPredicates, [int] $ChunkDays = 7)
    $rows = @()
    foreach ($chunk in (Get-DateChunks -From $From -To $To -Days $ChunkDays)) {
        $preds = @(@{ type = 'dimension'; dimension = 'knowledgeBaseId'; operator = 'matches'; value = $KnowledgeBaseId })
        if ($ExtraPredicates) { $preds += $ExtraPredicates }
        $body = @{
            interval = ('{0}/{1}' -f (Get-UtcIso $chunk[0]), (Get-UtcIso $chunk[1]))
            groupBy  = $GroupBy
            metrics  = $Metrics
            filter   = @{ type = 'and'; predicates = $preds }
            flattenMultivaluedDimensions = $true
        }
        $resp = $null
        try { $resp = Invoke-GcApi -Method POST -Path '/api/v2/analytics/knowledge/aggregates/query' -Body $body }
        catch {
            if ($ChunkDays -gt 1 -and "$($_.Exception.Message)" -match 'HTTP 400') {
                Write-Log ("Knowledge aggregate chunk {0} rejected (400) - retrying as 1-day chunks" -f $body.interval) -Level DEBUG
                $rows += Invoke-GcKnowledgeAggregate -From $chunk[0] -To $chunk[1] -GroupBy $GroupBy -Metrics $Metrics -ExtraPredicates $ExtraPredicates -ChunkDays 1
                continue
            }
            throw
        }
        foreach ($r in @(Get-Prop $resp 'results')) {
            $group = @{}
            $g = Get-Prop $r 'group'
            if ($g) { foreach ($pn in $g.PSObject.Properties.Name) { $group[$pn] = $g.$pn } }
            foreach ($d in @(Get-Prop $r 'data')) {
                foreach ($m in @(Get-Prop $d 'metrics')) {
                    $stats = Get-Prop $m 'stats'
                    $count = Get-Prop $stats 'count'; if ($null -eq $count) { $count = Get-Prop $stats 'sum' }
                    $rows += New-Object -TypeName PSObject -Property ([ordered]@{ Group = $group; Metric = (Get-Prop $m 'metric'); Count = [double]$count })
                }
            }
        }
    }
    return $rows
}

function Get-GcKnowledgeFeedback {
    <#  GET /api/v2/knowledge/knowledgebases/{kbId}/documents/feedback (permission knowledge:feedback:view)
        Thumbs up/down given on articles, with queryType/application/conversation context. The listing is not
        date-filterable server-side, so we page newest-first and stop once we are before the window. #>
    param([datetime] $From, [datetime] $To)
    $rows = @()
    $fromUtc = $From.ToUniversalTime(); $toUtc = $To.ToUniversalTime()
    $path = "/api/v2/knowledge/knowledgebases/$KnowledgeBaseId/documents/feedback"
    $page = Invoke-GcApi -Method GET -Path $path -Query @{ pageSize = 200 } -AllowNotFound
    $guard = 0
    while ($null -ne $page) {
        $ents = @(Get-Prop $page 'entities')
        $older = 0; $firstDt = $null; $lastDt = $null
        foreach ($f in $ents) {
            $dc = Get-Prop $f 'dateCreated'
            $dt = $null; if ($dc) { $dt = ([datetime]$dc).ToUniversalTime() }
            if ($dt) { if ($null -eq $firstDt) { $firstDt = $dt }; $lastDt = $dt }
            if ($dt -and $dt -lt $fromUtc) { $older++; continue }
            if ($dt -and $dt -ge $toUtc) { continue }
            $doc = Get-Prop $f 'document'
            $docId = Get-Prop $doc 'id'; if (-not $docId) { $docId = Get-Prop $doc 'documentId' }
            $rows += New-Object -TypeName PSObject -Property ([ordered]@{
                FeedbackId      = Get-Prop $f 'id'
                DocumentId      = $docId
                Rating          = Get-Prop $f 'rating'
                Reason          = Get-Prop $f 'reason'
                QueryType       = Get-Prop $f 'queryType'
                SurfacingMethod = Get-Prop $f 'surfacingMethod'
                AppType         = Get-Prop (Get-Prop $f 'application') 'type'
                ConversationId  = Get-Prop (Get-Prop (Get-Prop $f 'conversationContext') 'conversation') 'id'
                UserId          = Get-Prop (Get-Prop $f 'user') 'id'
                DateCreated     = $dc
                Comment         = Get-Prop $f 'comment'
            })
        }
        $next = Get-Prop $page 'nextUri'
        if (-not $next) { break }
        # Early stop only when the page is demonstrably newest-first AND entirely older than the window.
        $descending = ($firstDt -and $lastDt -and $firstDt -gt $lastDt)
        if ($descending -and $ents.Count -gt 0 -and $older -eq $ents.Count) { break }
        $guard++; if ($guard -gt 5000) { Add-Warning 'Feedback paging guard tripped (5000 pages); feedback list may be incomplete.'; break }
        $page = Invoke-GcApi -Method GET -Path $next
    }
    return $rows
}

function Get-GcUnansweredGroups {
    <#  GET /api/v2/knowledge/knowledgebases/{kbId}/unanswered/groups?app=Assistant&dateStart&dateEnd
        (permission knowledge:groups:view). This is the Knowledge Optimizer "unanswered queries" data. #>
    param([datetime] $From, [datetime] $To)
    $q = @{ app = 'Assistant'; dateStart = $From.ToUniversalTime().ToString('yyyy-MM-dd'); dateEnd = $To.ToUniversalTime().ToString('yyyy-MM-dd') }
    $resp = Invoke-GcApi -Method GET -Path "/api/v2/knowledge/knowledgebases/$KnowledgeBaseId/unanswered/groups" -Query $q -AllowNotFound
    return @(Get-Prop $resp 'entities')
}
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region Data retrieval - knowledge base inventory, transcripts, knowledge search replay
# ---------------------------------------------------------------------------------------------------------------
function Get-GcKnowledgeDocuments {
    <#  GET /api/v2/knowledge/knowledgebases/{kbId}/documents?includeDrafts=true&expand=category,labels
        (permission knowledge:document:view). 'alternatives' (alternative phrases) are part of the response. #>
    $docs = Get-GcCursorPages -Path "/api/v2/knowledge/knowledgebases/$KnowledgeBaseId/documents" -Query @{ pageSize = 200; includeDrafts = 'true'; expand = 'category,labels' }
    $inv = @{}
    foreach ($d in $docs) {
        $alts = @(); foreach ($a in @(Get-Prop $d 'alternatives')) { $p = Get-Prop $a 'phrase'; if ($p) { $alts += $p } }
        $labels = @(); foreach ($l in @(Get-Prop $d 'labels')) { $n = Get-Prop $l 'name'; if ($n) { $labels += $n } else { $labels += (Get-Prop $l 'id') } }
        $id = Get-Prop $d 'id'
        $inv[$id] = New-Object -TypeName PSObject -Property ([ordered]@{
            DocumentId    = $id
            Title         = Get-Prop $d 'title'
            State         = Get-Prop $d 'state'
            Visible       = Get-Prop $d 'visible'
            Category      = Get-Prop (Get-Prop $d 'category') 'name'
            Labels        = ($labels -join '; ')
            Alternatives  = $alts
            DateModified  = Get-Prop $d 'dateModified'
            DatePublished = Get-Prop $d 'datePublished'
        })
    }
    return $inv
}

function Get-GcTranscriptPhrases {
    <#  GET /api/v2/speechandtextanalytics/conversations/{id}/communications/{commId}/transcripturl
        (permissions recording:recording:view + speechAndTextAnalytics:data:view) returns a pre-signed S3 URL
        valid ~10 minutes; the JSON holds transcripts[].phrases[] {text, startTimeMs, participantPurpose}.
        The presigned URL must be fetched WITHOUT the bearer token. Returns @() when no transcript exists. #>
    param([string] $ConversationId, [string] $CommunicationId)
    $tu = Invoke-GcApi -Method GET -Path "/api/v2/speechandtextanalytics/conversations/$ConversationId/communications/$CommunicationId/transcripturl" -AllowNotFound
    $url = Get-Prop $tu 'url'
    if (-not $url) { return @() }
    $json = Invoke-GcApi -Method GET -AbsoluteUri $url -NoAuth
    $phrases = @()
    foreach ($t in @(Get-Prop $json 'transcripts')) {
        foreach ($ph in @(Get-Prop $t 'phrases')) {
            $text = Get-Prop $ph 'text'
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $purpose = "$(Get-Prop $ph 'participantPurpose')".ToLowerInvariant()
            $speaker = if ($purpose -in @('external', 'customer')) { 'Customer' } elseif ($purpose -in @('internal', 'agent')) { 'Agent' } else { 'Unknown' }
            $phrases += New-Object -TypeName PSObject -Property ([ordered]@{ Text = $text; StartMs = [long](Get-Prop $ph 'startTimeMs'); Speaker = $speaker; CommunicationId = $CommunicationId })
        }
    }
    return $phrases
}

function Select-ReplayUtterances {
    <#  Picks the utterances worth replaying: customer (and optionally agent) phrases with >= MinUtteranceWords,
        de-duplicated, the N longest, returned in call order. #>
    param($Phrases)
    $cands = @($Phrases | Where-Object { ($_.Speaker -eq 'Customer') -or ($IncludeAgentUtterances -and $_.Speaker -eq 'Agent') })
    $seen = @{}; $uniq = @()
    foreach ($p in $cands) {
        $norm = Get-NormalizedText $p.Text
        if ((Get-WordCount $norm) -lt $MinUtteranceWords) { continue }
        if ($seen.ContainsKey($norm)) { continue }
        $seen[$norm] = $true
        $uniq += New-Object -TypeName PSObject -Property ([ordered]@{ Text = $p.Text; Norm = $norm; Words = (Get-WordCount $norm); StartMs = $p.StartMs; Speaker = $p.Speaker; CommunicationId = $p.CommunicationId })
    }
    return @($uniq | Sort-Object -Property Words -Descending | Select-Object -First $MaxUtterancesPerConversation | Sort-Object -Property StartMs)
}

function Search-GcKnowledge {
    <#  POST /api/v2/knowledge/knowledgebases/{kbId}/documents/search (permission knowledge:knowledgebase:search)
        Replays an utterance as an AutoSearch (the query type Copilot uses) sorted by confidence, no threshold,
        so we see what WOULD have matched and with what confidence. Draft documents are excluded.
        NOTE: each call is recorded by knowledge analytics as a search. By default we do NOT send
        application/conversationContext so these replays are distinguishable from live Copilot searches. #>
    param([string] $Query, [string] $ConversationId, [string] $QueueIdForContext)
    $q = $Query; if ($q.Length -gt 1000) { $q = $q.Substring(0, 1000) }
    $body = @{ query = $q; pageSize = $ReplayResultCount; pageNumber = 1; queryType = 'AutoSearch'; includeDraftDocuments = $false; sortBy = 'ConfidenceScore'; sortOrder = 'Desc' }
    if ($TagReplaySearchesWithContext) {
        $body['application'] = @{ type = 'Assistant' }
        $ctx = @{ conversationId = $ConversationId; mediaType = 'Voice' }
        if ($QueueIdForContext) { $ctx['queueId'] = $QueueIdForContext }
        $body['conversationContext'] = $ctx
    }
    $resp = Invoke-GcApi -Method POST -Path "/api/v2/knowledge/knowledgebases/$KnowledgeBaseId/documents/search" -Body $body
    $results = @()
    $rank = 0
    foreach ($r in @(Get-Prop $resp 'results')) {
        $rank++
        $doc = Get-Prop $r 'document'
        $results += New-Object -TypeName PSObject -Property ([ordered]@{ Rank = $rank; DocumentId = (Get-Prop $doc 'id'); Title = (Get-Prop $doc 'title'); Confidence = [double](Get-Prop $r 'confidence') })
    }
    return New-Object -TypeName PSObject -Property ([ordered]@{ SearchId = (Get-Prop $resp 'searchId'); Results = $results })
}
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region Name resolution
# ---------------------------------------------------------------------------------------------------------------
function Get-GcNameMap {
    <# Resolves queue, wrap-up code and user IDs to names. All three are read-only list endpoints. #>
    param([string[]] $QueueIds, [string[]] $UserIds)
    $map = @{ Queue = @{}; WrapUp = @{}; User = @{} }
    try {
        # GET /api/v2/routing/queues?id=a,b,c (permission routing:queue:view)
        for ($i = 0; $i -lt $QueueIds.Count; $i += 50) {
            $batch = @($QueueIds[$i..((Get-MinInt ($i + 49) ($QueueIds.Count - 1)))])
            $page = Invoke-GcApi -Method GET -Path '/api/v2/routing/queues' -Query @{ pageSize = 100; id = ($batch -join ',') }
            foreach ($e in @(Get-Prop $page 'entities')) { $map.Queue[(Get-Prop $e 'id')] = Get-Prop $e 'name' }
        }
    } catch { Add-Warning "Queue name resolution failed: $($_.Exception.Message)" }
    try {
        # GET /api/v2/routing/wrapupcodes (permission routing:wrapupCode:view) - full list, pageNumber paging
        foreach ($e in (Get-GcNumberedPages -Path '/api/v2/routing/wrapupcodes' -PageSize 100)) { $map.WrapUp[(Get-Prop $e 'id')] = Get-Prop $e 'name' }
    } catch { Add-Warning "Wrap-up code name resolution failed: $($_.Exception.Message)" }
    try {
        # GET /api/v2/users?id=a,b,c&state=any
        for ($i = 0; $i -lt $UserIds.Count; $i += 50) {
            $batch = @($UserIds[$i..((Get-MinInt ($i + 49) ($UserIds.Count - 1)))])
            $page = Invoke-GcApi -Method GET -Path '/api/v2/users' -Query @{ pageSize = 100; id = ($batch -join ','); state = 'any' }
            foreach ($e in @(Get-Prop $page 'entities')) { $map.User[(Get-Prop $e 'id')] = Get-Prop $e 'name' }
        }
    } catch { Add-Warning "User name resolution failed: $($_.Exception.Message)" }
    return $map
}

function Resolve-Names { param([hashtable] $Map, $Ids) return (Join-Unique @(@($Ids) | ForEach-Object { if ($Map.ContainsKey($_)) { $Map[$_] } else { $_ } })) }
# endregion

# ---------------------------------------------------------------------------------------------------------------
# region CSV export
# ---------------------------------------------------------------------------------------------------------------
function Export-Report {
    param([string] $Name, $Rows)
    $file = Join-Path -Path $OutputPath -ChildPath $Name
    $list = @($Rows | Where-Object { $null -ne $_ })
    if ($list.Count -eq 0) {
        Set-Content -Path $file -Value '' -Encoding UTF8
        Write-Log "Wrote $Name (0 rows)"
        return
    }
    $list | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-Log ("Wrote {0} ({1} rows)" -f $Name, $list.Count)
}
# endregion

# ===============================================================================================================
# MAIN
# ===============================================================================================================
$runStart = Get-Date
if ($EndDate -le $StartDate) { throw '-EndDate must be after -StartDate.' }
if (-not (Test-Path -Path $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:LogFile = Join-Path -Path $OutputPath -ChildPath "KnowledgeAudit-$stamp.log"
$transcriptFile = Join-Path -Path $OutputPath -ChildPath "KnowledgeAudit-$stamp.transcript.txt"
$transcriptStarted = $false
try { Start-Transcript -Path $transcriptFile -ErrorAction Stop | Out-Null; $transcriptStarted = $true } catch { Write-Warning "Start-Transcript unavailable: $($_.Exception.Message)" }

try {
    Write-Log ("Genesys Cloud knowledge audit starting. PS {0}, LanguageMode={1}" -f $PSVersionTable.PSVersion, $ExecutionContext.SessionState.LanguageMode)
    Write-Log ("Window {0} -> {1} (UTC), region {2}, KB {3}, queues: {4}" -f (Get-UtcIso $StartDate), (Get-UtcIso $EndDate), $Region, $KnowledgeBaseId, ($(if ($QueueId) { $QueueId -join ',' } else { 'all' })))
    if (-not $IncludeUtteranceText) { Write-Log "PII masking ON: utterances are redacted and truncated to $MaskLength chars (use -IncludeUtteranceText to opt in)." }
    else { Write-Log 'PII masking OFF: full utterance text WILL be written to disk (-IncludeUtteranceText).' -Level WARN }

    # ---- Auth -------------------------------------------------------------------------------------------------
    # Client ID precedence: -ClientId parameter > GENESYS_CLIENT_ID env var (parameter default) > $EmbeddedClientId > prompt.
    if ([string]::IsNullOrWhiteSpace($ClientId) -and -not [string]::IsNullOrWhiteSpace($EmbeddedClientId)) { $ClientId = $EmbeddedClientId }
    if ([string]::IsNullOrWhiteSpace($ClientId)) { $ClientId = Read-Host -Prompt 'Enter Genesys Cloud OAuth client ID' }
    $script:ClientIdValue = $ClientId
    $script:ClientSecretValue = Resolve-ClientSecret -Provided $ClientSecret
    Get-GcToken -Id $script:ClientIdValue -Secret $script:ClientSecretValue

    # ---- Data availability sanity check -----------------------------------------------------------------------
    try {
        $avail = Invoke-GcApi -Method GET -Path '/api/v2/analytics/conversations/details/jobs/availability'
        $availDate = Get-Prop $avail 'dataAvailabilityDate'
        Write-Log "Analytics data lake available up to $availDate"
        if ($availDate -and ([datetime]$availDate).ToUniversalTime() -lt $EndDate.ToUniversalTime()) { Add-Warning "EndDate is after the analytics availability date ($availDate); the newest conversations may be missing." }
    } catch { Add-Warning "Could not read analytics availability: $($_.Exception.Message)" }

    # ---- 1. Knowledge base + article inventory ---------------------------------------------------------------
    $kb = Invoke-GcApi -Method GET -Path "/api/v2/knowledge/knowledgebases/$KnowledgeBaseId"
    Write-Log ("Knowledge base: '{0}' ({1}), language {2}" -f (Get-Prop $kb 'name'), $KnowledgeBaseId, (Get-Prop $kb 'coreLanguage'))
    $inventory = Get-GcKnowledgeDocuments
    Write-Log ("Article inventory: {0} documents" -f $inventory.Count)

    # ---- 2. Conversations ------------------------------------------------------------------------------------
    $rawConvs = Get-GcVoiceConversations -From $StartDate -To $EndDate -Queues $QueueId
    $fromUtc = $StartDate.ToUniversalTime(); $toUtc = $EndDate.ToUniversalTime()
    $conversations = @()
    $skippedNoAgent = 0; $skippedOutside = 0
    foreach ($c in $rawConvs) {
        $s = ConvertTo-ConversationSummary -Conv $c
        if (-not $s.ConversationId) { continue }
        $st = if ($s.Start) { ([datetime]$s.Start).ToUniversalTime() } else { $null }
        if ($st -and ($st -lt $fromUtc -or $st -ge $toUtc)) { $skippedOutside++; continue }
        if (-not $s.HasAgentInteract) { $skippedNoAgent++; continue }   # abandoned / IVR-only: Copilot cannot act
        $conversations += $s
    }
    Write-Log ("Voice conversations in scope: {0} (skipped {1} without an agent interaction, {2} starting outside the window)" -f $conversations.Count, $skippedNoAgent, $skippedOutside)
    $convIndex = @{}; foreach ($c in $conversations) { $convIndex[$c.ConversationId] = $c }

    # ---- 3. Name resolution ---------------------------------------------------------------------------------
    $allQueueIds = @($conversations | ForEach-Object { $_.QueueIds } | Select-Object -Unique)
    $allUserIds  = @($conversations | ForEach-Object { $_.AgentIds } | Select-Object -Unique)
    $names = Get-GcNameMap -QueueIds $allQueueIds -UserIds $allUserIds

    # ---- 4. Knowledge aggregates (article + conversation level) --------------------------------------------
    $appPred = @(); if ($AggregateAppType) { $appPred = @(@{ type = 'dimension'; dimension = 'appType'; operator = 'matches'; value = $AggregateAppType }) }
    $aggDocMetrics  = @('nKnowledgeDocumentSurfaced', 'nKnowledgeDocumentPresented', 'nKnowledgeDocumentViewed', 'nKnowledgeDocumentCopied', 'nKnowledgeDocumentFeedback')
    $aggConvMetrics = @('nKnowledgeSearch', 'nKnowledgeSearchAnswered', 'nKnowledgeSearchUnanswered', 'nKnowledgeDocumentSurfaced', 'nKnowledgeDocumentViewed', 'nKnowledgeDocumentCopied')
    $aggByDoc = @{}; $aggByConv = @{}; $aggConvDoc = @{}; $aggFeedback = @{}
    try {
        foreach ($r in (Invoke-GcKnowledgeAggregate -From $StartDate -To $EndDate -GroupBy @('knowledgeBaseDocumentId') -Metrics $aggDocMetrics -ExtraPredicates $appPred)) {
            $d = $r.Group['knowledgeBaseDocumentId']; if (-not $d) { continue }
            if (-not $aggByDoc.ContainsKey($d)) { $aggByDoc[$d] = @{} }
            $aggByDoc[$d][$r.Metric] = [double](Get-Prop $aggByDoc[$d] $r.Metric) + $r.Count
        }
        foreach ($r in (Invoke-GcKnowledgeAggregate -From $StartDate -To $EndDate -GroupBy @('conversationId') -Metrics $aggConvMetrics -ExtraPredicates $appPred)) {
            $cid = $r.Group['conversationId']; if (-not $cid) { continue }
            if (-not $aggByConv.ContainsKey($cid)) { $aggByConv[$cid] = @{} }
            $aggByConv[$cid][$r.Metric] = [double](Get-Prop $aggByConv[$cid] $r.Metric) + $r.Count
        }
        foreach ($r in (Invoke-GcKnowledgeAggregate -From $StartDate -To $EndDate -GroupBy @('conversationId', 'knowledgeBaseDocumentId') -Metrics @('nKnowledgeDocumentSurfaced', 'nKnowledgeDocumentViewed', 'nKnowledgeDocumentCopied') -ExtraPredicates $appPred)) {
            $cid = $r.Group['conversationId']; $d = $r.Group['knowledgeBaseDocumentId']
            if (-not $cid -or -not $d -or $r.Count -le 0) { continue }
            if (-not $aggConvDoc.ContainsKey($cid)) { $aggConvDoc[$cid] = @{} }
            if (-not $aggConvDoc[$cid].ContainsKey($d)) { $aggConvDoc[$cid][$d] = @{} }
            $aggConvDoc[$cid][$d][$r.Metric] = [double](Get-Prop $aggConvDoc[$cid][$d] $r.Metric) + $r.Count
        }
        foreach ($r in (Invoke-GcKnowledgeAggregate -From $StartDate -To $EndDate -GroupBy @('knowledgeBaseDocumentId', 'feedbackRating') -Metrics @('nKnowledgeDocumentFeedback') -ExtraPredicates $appPred)) {
            $d = $r.Group['knowledgeBaseDocumentId']; $rating = "$($r.Group['feedbackRating'])"
            if (-not $d) { continue }
            if (-not $aggFeedback.ContainsKey($d)) { $aggFeedback[$d] = @{ Positive = 0; Negative = 0 } }
            if ($rating -match '^(?i)pos') { $aggFeedback[$d].Positive += $r.Count } elseif ($rating -match '^(?i)neg') { $aggFeedback[$d].Negative += $r.Count }
        }
        Write-Log ("Knowledge aggregates: {0} documents with activity, {1} conversations with knowledge events" -f $aggByDoc.Count, $aggByConv.Count)
        if ($aggByDoc.Count -eq 0) { Add-Warning "Knowledge aggregates returned no activity for KB $KnowledgeBaseId with appType='$AggregateAppType'. If Copilot is live, check the KB id and try -AggregateAppType ''." }
    }
    catch { Add-Warning "Knowledge aggregates unavailable: $($_.Exception.Message)" }

    # ---- 5. Copilot suggestions per conversation ------------------------------------------------------------
    $suggestions = @()               # flat rows
    $suggestionsByConv = @{}         # conversationId -> rows
    $suggestionSource = if ($aggByConv.Count -gt 0) { 'KnowledgeAggregates' } else { 'None' }
    if (-not $SkipSuggestionsApi) {
        $lookupTargets = $conversations
        if ($SuggestionLookupMode -eq 'AggregatesFirst') {
            $lookupTargets = @($conversations | Where-Object { $aggByConv.ContainsKey($_.ConversationId) })
            Write-Log ("SuggestionLookupMode=AggregatesFirst: querying suggestions for {0} of {1} conversations with knowledge events" -f $lookupTargets.Count, $conversations.Count)
        }
        $consecutiveForbidden = 0; $i = 0; $apiDisabled = $false; $successfulCalls = 0
        $forbiddenLimit = Get-MinInt 3 $lookupTargets.Count
        foreach ($c in $lookupTargets) {
            $i++
            if ($i % 25 -eq 0 -or $i -eq $lookupTargets.Count) { Write-Progress -Activity 'Reading Copilot suggestions' -Status "$i / $($lookupTargets.Count)" -PercentComplete ([int](100 * $i / (Get-MaxInt 1 $lookupTargets.Count))) }
            try {
                $rows = @(Get-GcConversationSuggestions -ConversationId $c.ConversationId)
                $consecutiveForbidden = 0; $successfulCalls++
                if ($rows.Count -gt 0) { $suggestionsByConv[$c.ConversationId] = $rows; $suggestions += $rows }
            }
            catch {
                $msg = $_.Exception.Message
                if ($msg -match 'HTTP 403') {
                    $consecutiveForbidden++
                    if ($consecutiveForbidden -ge $forbiddenLimit) {
                        Add-Warning 'GET /conversations/{id}/suggestions returned 403 for every conversation tried. The OAuth client lacks conversation:suggestion:view (or the endpoint is not available to client-credentials tokens). Falling back to knowledge aggregates only.'
                        $apiDisabled = $true; break
                    }
                }
                else { Add-Warning "Suggestions for $($c.ConversationId): $msg" }
            }
        }
        Write-Progress -Activity 'Reading Copilot suggestions' -Completed
        if (-not $apiDisabled -and $successfulCalls -gt 0) { $suggestionSource = 'SuggestionsApi' }
        Write-Log ("Suggestions API: {0} suggestions across {1} conversations" -f $suggestions.Count, $suggestionsByConv.Count)
    }

    # ---- 6. Article feedback records -------------------------------------------------------------------------
    $feedback = @()
    try { $feedback = @(Get-GcKnowledgeFeedback -From $StartDate -To $EndDate); Write-Log ("Feedback records in window: {0}" -f $feedback.Count) }
    catch { Add-Warning "Knowledge feedback listing unavailable: $($_.Exception.Message)" }
    $feedbackByConv = @{}; $feedbackByDoc = @{}
    foreach ($f in $feedback) {
        if ($f.ConversationId) { if (-not $feedbackByConv.ContainsKey($f.ConversationId)) { $feedbackByConv[$f.ConversationId] = @() }; $feedbackByConv[$f.ConversationId] += $f }
        if ($f.DocumentId)     { if (-not $feedbackByDoc.ContainsKey($f.DocumentId)) { $feedbackByDoc[$f.DocumentId] = @() }; $feedbackByDoc[$f.DocumentId] += $f }
    }

    # ---- 7. Build "what surfaced live" index: conversationId -> set of documentIds ----------------------------
    $liveDocsByConv = @{}
    foreach ($cid in $suggestionsByConv.Keys) {
        foreach ($s in $suggestionsByConv[$cid]) { if ($s.IsKnowledge -and $s.DocumentId) { if (-not $liveDocsByConv.ContainsKey($cid)) { $liveDocsByConv[$cid] = @{} }; $liveDocsByConv[$cid][$s.DocumentId] = $true } }
    }
    foreach ($cid in $aggConvDoc.Keys) {
        foreach ($d in $aggConvDoc[$cid].Keys) { if (-not $liveDocsByConv.ContainsKey($cid)) { $liveDocsByConv[$cid] = @{} }; $liveDocsByConv[$cid][$d] = $true }
    }

    # ---- 8. Transcript replay --------------------------------------------------------------------------------
    $replayRows = @()              # one row per replayed utterance
    $replayedConvs = @{}           # conversationId -> @{ Transcript = bool; Utterances = n; BestConfidence = x }
    if (-not $SkipTranscriptReplay) {
        # Conversations with no live knowledge suggestion first, then the rest, capped.
        $ordered = @($conversations | Sort-Object -Property @{ Expression = { if ($liveDocsByConv.ContainsKey($_.ConversationId)) { 1 } else { 0 } } }, @{ Expression = { $_.Start } })
        $target = @($ordered | Select-Object -First $MaxConversationsForReplay)
        Write-Log ("Transcript replay: {0} conversations selected (cap {1}; conversations without live suggestions first)" -f $target.Count, $MaxConversationsForReplay)
        $i = 0; $noTranscript = 0; $searches = 0
        foreach ($c in $target) {
            $i++
            Write-Progress -Activity 'Replaying transcripts against knowledge search' -Status ("{0} / {1}  ({2} searches)" -f $i, $target.Count, $searches) -PercentComplete ([int](100 * $i / (Get-MaxInt 1 $target.Count)))
            $phrases = @()
            foreach ($commId in $c.CommunicationIds) {
                try { $phrases += @(Get-GcTranscriptPhrases -ConversationId $c.ConversationId -CommunicationId $commId) }
                catch { Add-Warning "Transcript for $($c.ConversationId)/$commId failed: $($_.Exception.Message)" }
            }
            if ($phrases.Count -eq 0) { $noTranscript++; $replayedConvs[$c.ConversationId] = @{ Transcript = $false; Utterances = 0; BestConfidence = $null }; continue }
            $utts = @(Select-ReplayUtterances -Phrases $phrases)
            $best = $null
            $primaryQueue = if ($c.QueueIds.Count -gt 0) { $c.QueueIds[0] } else { $null }
            foreach ($u in $utts) {
                try {
                    $sr = Search-GcKnowledge -Query $u.Text -ConversationId $c.ConversationId -QueueIdForContext $primaryQueue
                    $searches++
                }
                catch { Add-Warning "Knowledge search failed for $($c.ConversationId): $($_.Exception.Message)"; continue }
                $top = if ($sr.Results.Count -gt 0) { $sr.Results[0] } else { $null }
                $second = if ($sr.Results.Count -gt 1) { $sr.Results[1] } else { $null }
                $topConf = if ($top) { $top.Confidence } else { 0 }
                if ($null -eq $best -or $topConf -gt $best) { $best = $topConf }
                $surfaced = ($top -and $liveDocsByConv.ContainsKey($c.ConversationId) -and $liveDocsByConv[$c.ConversationId].ContainsKey($top.DocumentId))
                $anyLive = $liveDocsByConv.ContainsKey($c.ConversationId)
                $class = if (-not $top) { 'NoMatch' }
                         elseif ($surfaced) { 'SurfacedLive' }
                         elseif ($topConf -ge $MissedOpportunityMinConfidence) { 'Missed' }
                         elseif ($topConf -lt $GapMaxConfidence) { 'LowConfidence' }
                         else { 'Marginal' }
                $replayRows += New-Object -TypeName PSObject -Property ([ordered]@{
                    ConversationId        = $c.ConversationId
                    CommunicationId       = $u.CommunicationId
                    UtteranceStartMs      = $u.StartMs
                    Speaker               = $u.Speaker
                    UtteranceRaw          = $u.Text
                    UtteranceNorm         = $u.Norm
                    BestDocumentId        = $(if ($top) { $top.DocumentId } else { $null })
                    BestTitle             = $(if ($top) { $top.Title } else { $null })
                    BestConfidence        = $topConf
                    SecondDocumentId      = $(if ($second) { $second.DocumentId } else { $null })
                    SecondTitle           = $(if ($second) { $second.Title } else { $null })
                    SecondConfidence      = $(if ($second) { $second.Confidence } else { $null })
                    AllResults            = $sr.Results
                    SurfacedLive          = $surfaced
                    ConversationHadLiveSuggestion = $anyLive
                    Classification        = $class
                    SearchId              = $sr.SearchId
                })
            }
            $replayedConvs[$c.ConversationId] = @{ Transcript = $true; Utterances = $utts.Count; BestConfidence = $best }
        }
        Write-Progress -Activity 'Replaying transcripts against knowledge search' -Completed
        Write-Log ("Replay complete: {0} conversations, {1} without transcript, {2} knowledge searches, {3} utterance rows" -f $target.Count, $noTranscript, $searches, $replayRows.Count)
        if ($noTranscript -eq $target.Count -and $target.Count -gt 0) { Add-Warning 'No transcripts were found for any selected conversation. Check voice transcription is enabled on the queues and that the OAuth client has recording:recording:view + speechAndTextAnalytics:data:view.' }
    }

    # ---- 9. Knowledge Optimizer unanswered groups ------------------------------------------------------------
    $unanswered = @()
    try { $unanswered = @(Get-GcUnansweredGroups -From $StartDate -To $EndDate); Write-Log ("Knowledge Optimizer unanswered groups (Assistant): {0}" -f $unanswered.Count) }
    catch { Add-Warning "Unanswered groups unavailable: $($_.Exception.Message)" }

    # ===========================================================================================================
    # OUTPUT 1 - ConversationKnowledgeActivity.csv
    # ===========================================================================================================
    $activity = foreach ($c in $conversations) {
        $cid = $c.ConversationId
        $sugs = @(); if ($suggestionsByConv.ContainsKey($cid)) { $sugs = @($suggestionsByConv[$cid]) }
        $ksugs = @($sugs | Where-Object { $_.IsKnowledge })
        $agg = if ($aggByConv.ContainsKey($cid)) { $aggByConv[$cid] } else { @{} }
        $fb = @(); if ($feedbackByConv.ContainsKey($cid)) { $fb = @($feedbackByConv[$cid]) }
        $liveTitles = @()
        if ($ksugs.Count -gt 0) { $liveTitles = @($ksugs | ForEach-Object { if ($_.Title) { $_.Title } elseif ($inventory.ContainsKey("$($_.DocumentId)")) { $inventory[$_.DocumentId].Title } else { $_.DocumentId } }) }
        elseif ($aggConvDoc.ContainsKey($cid)) { $liveTitles = @($aggConvDoc[$cid].Keys | ForEach-Object { if ($inventory.ContainsKey($_)) { $inventory[$_].Title } else { $_ } }) }
        $liveIds = @(); if ($liveDocsByConv.ContainsKey($cid)) { $liveIds = @($liveDocsByConv[$cid].Keys) }
        $rp = if ($replayedConvs.ContainsKey($cid)) { $replayedConvs[$cid] } else { $null }
        $missed = @($replayRows | Where-Object { $_.ConversationId -eq $cid -and $_.Classification -eq 'Missed' }).Count
        $queueNames = Resolve-Names -Map $names.Queue -Ids $c.QueueIds
        New-Object -TypeName PSObject -Property ([ordered]@{
            conversationId            = $cid
            conversationStart         = (ConvertTo-IsoText $c.Start)
            conversationEnd           = (ConvertTo-IsoText $c.End)
            direction                 = $c.Direction
            queueIds                  = ($c.QueueIds -join '; ')
            queueNames                = $queueNames
            agentIds                  = ($c.AgentIds -join '; ')
            agentNames                = (Resolve-Names -Map $names.User -Ids $c.AgentIds)
            wrapUpCodeIds             = ($c.WrapUpCodeIds -join '; ')
            wrapUpCodes               = (Resolve-Names -Map $names.WrapUp -Ids $c.WrapUpCodeIds)
            knowledgeBaseIdsOnConversation = ($c.KnowledgeBaseIds -join '; ')
            suggestionSource          = $(if ($sugs.Count -gt 0) { 'SuggestionsApi' } elseif ($aggConvDoc.ContainsKey($cid) -or $agg.Count -gt 0) { 'KnowledgeAggregates' } else { 'None' })
            suggestionsShown          = $sugs.Count
            knowledgeSuggestionsShown = $(if ($ksugs.Count -gt 0) { $ksugs.Count } else { [int][double](Get-Prop $agg 'nKnowledgeDocumentSurfaced') })
            suggestionTriggerTypes    = (Join-Unique ($ksugs | ForEach-Object { $_.TriggerType }))
            articleIdsShown           = ($liveIds -join '; ')
            articlesShown             = (Join-Unique $liveTitles)
            articlesAccepted          = (Join-Unique ($ksugs | Where-Object { $_.State -eq 'Accepted' } | ForEach-Object { $_.Title }))
            articlesDismissed         = (Join-Unique ($ksugs | Where-Object { $_.State -eq 'Dismissed' } | ForEach-Object { $_.Title }))
            articlesRated             = (Join-Unique ($ksugs | Where-Object { $_.State -eq 'Rated' } | ForEach-Object { $_.Title }))
            aggKnowledgeSearches      = [int][double](Get-Prop $agg 'nKnowledgeSearch')
            aggSearchesAnswered       = [int][double](Get-Prop $agg 'nKnowledgeSearchAnswered')
            aggSearchesUnanswered     = [int][double](Get-Prop $agg 'nKnowledgeSearchUnanswered')
            aggDocsSurfaced           = [int][double](Get-Prop $agg 'nKnowledgeDocumentSurfaced')
            aggDocsViewed             = [int][double](Get-Prop $agg 'nKnowledgeDocumentViewed')
            aggDocsCopied             = [int][double](Get-Prop $agg 'nKnowledgeDocumentCopied')
            feedbackPositive          = @($fb | Where-Object { $_.Rating -eq 'Positive' }).Count
            feedbackNegative          = @($fb | Where-Object { $_.Rating -eq 'Negative' }).Count
            transcriptReplayed        = $(if ($rp) { [bool]$rp.Transcript } else { $false })
            utterancesReplayed        = $(if ($rp) { $rp.Utterances } else { 0 })
            replayBestConfidence      = $(if ($rp -and $null -ne $rp.BestConfidence) { ConvertTo-Rounded ([double]$rp.BestConfidence) } else { $null })
            missedOpportunities       = $missed
        })
    }
    Export-Report -Name 'ConversationKnowledgeActivity.csv' -Rows $activity

    # ===========================================================================================================
    # OUTPUT 2 - ArticlePerformance.csv
    # ===========================================================================================================
    $sugByDoc = @{}
    foreach ($s in $suggestions) { if ($s.IsKnowledge -and $s.DocumentId) { if (-not $sugByDoc.ContainsKey($s.DocumentId)) { $sugByDoc[$s.DocumentId] = @() }; $sugByDoc[$s.DocumentId] += $s } }
    $replayBestByDoc = @{}; $replayMissedByDoc = @{}
    foreach ($r in $replayRows) {
        if ($r.BestDocumentId) {
            $replayBestByDoc[$r.BestDocumentId] = [int](Get-Prop $replayBestByDoc $r.BestDocumentId) + 1
            if ($r.Classification -eq 'Missed') { $replayMissedByDoc[$r.BestDocumentId] = [int](Get-Prop $replayMissedByDoc $r.BestDocumentId) + 1 }
        }
    }
    $allDocIds = @($inventory.Keys) + @($aggByDoc.Keys) + @($sugByDoc.Keys) | Select-Object -Unique
    $articlePerf = foreach ($d in $allDocIds) {
        $inv = if ($inventory.ContainsKey($d)) { $inventory[$d] } else { $null }
        $agg = if ($aggByDoc.ContainsKey($d)) { $aggByDoc[$d] } else { @{} }
        $sg = @(); if ($sugByDoc.ContainsKey($d)) { $sg = @($sugByDoc[$d]) }
        $fb = @(); if ($feedbackByDoc.ContainsKey($d)) { $fb = @($feedbackByDoc[$d]) }
        $aggSurf = [int][double](Get-Prop $agg 'nKnowledgeDocumentSurfaced')
        $surfaced = Get-MaxInt $aggSurf $sg.Count
        $viewed   = Get-MaxInt ([int][double](Get-Prop $agg 'nKnowledgeDocumentViewed')) (@($sg | Where-Object { $_.State -eq 'Accepted' }).Count)
        $copied   = [int][double](Get-Prop $agg 'nKnowledgeDocumentCopied')
        $pos = @($fb | Where-Object { $_.Rating -eq 'Positive' }).Count; $neg = @($fb | Where-Object { $_.Rating -eq 'Negative' }).Count
        if ($aggFeedback.ContainsKey($d)) { $pos = Get-MaxInt $pos ([int]$aggFeedback[$d].Positive); $neg = Get-MaxInt $neg ([int]$aggFeedback[$d].Negative) }
        $missedN = [int](Get-Prop $replayMissedByDoc $d); $bestN = [int](Get-Prop $replayBestByDoc $d)
        $state = if ($inv) { $inv.State } else { 'NotInInventory' }
        $visible = if ($inv) { $inv.Visible } else { $null }
        $status = if ($inv -and ($state -ne 'Published' -or $visible -eq $false)) { 'NotSearchable' }
                  elseif ($surfaced -eq 0 -and $missedN -gt 0) { 'Dormant-MatchesCalls' }
                  elseif ($surfaced -eq 0) { 'Dormant' }
                  elseif ($neg -gt $pos -and $neg -ge 2) { 'NeedsReview' }
                  elseif ($surfaced -le $LowSurfaceThreshold -or ($viewed -eq 0 -and $copied -eq 0)) { 'Low' }
                  else { 'Healthy' }
        $avgConf = $null
        $confs = @($sg | Where-Object { $null -ne $_.Confidence } | ForEach-Object { [double]$_.Confidence })
        if ($confs.Count -gt 0) { $avgConf = ConvertTo-Rounded (($confs | Measure-Object -Average).Average) }
        New-Object -TypeName PSObject -Property ([ordered]@{
            articleId                 = $d
            title                     = $(if ($inv) { $inv.Title } elseif ($sg.Count -gt 0) { $sg[0].Title } else { '' })
            category                  = $(if ($inv) { $inv.Category } else { '' })
            labels                    = $(if ($inv) { $inv.Labels } else { '' })
            state                     = $state
            visible                   = $visible
            datePublished             = $(if ($inv) { ConvertTo-IsoText $inv.DatePublished } else { '' })
            dateModified              = $(if ($inv) { ConvertTo-IsoText $inv.DateModified } else { '' })
            timesSuggested            = $surfaced
            timesSuggestedViaApi      = $sg.Count
            timesSurfacedAggregates   = $aggSurf
            timesPresented            = [int][double](Get-Prop $agg 'nKnowledgeDocumentPresented')
            timesViewed               = $viewed
            timesUsed                 = $copied
            timesAccepted             = @($sg | Where-Object { $_.State -eq 'Accepted' }).Count
            timesDismissed            = @($sg | Where-Object { $_.State -eq 'Dismissed' }).Count
            avgSuggestionConfidence   = $avgConf
            positiveFeedback          = $pos
            negativeFeedback          = $neg
            alternativePhraseCount    = $(if ($inv) { $inv.Alternatives.Count } else { 0 })
            replayBestMatchCount      = $bestN
            replayMissedCount         = $missedN
            status                    = $status
        })
    }
    Export-Report -Name 'ArticlePerformance.csv' -Rows ($articlePerf | Sort-Object -Property @{ Expression = 'timesSuggested'; Descending = $true }, title)

    # ===========================================================================================================
    # OUTPUT 3 - MissedOpportunities.csv  (every replayed utterance, classified; filter on classification)
    # ===========================================================================================================
    $missedRows = foreach ($r in ($replayRows | Sort-Object -Property @{ Expression = 'BestConfidence'; Descending = $true })) {
        $c = $convIndex[$r.ConversationId]
        New-Object -TypeName PSObject -Property ([ordered]@{
            conversationId            = $r.ConversationId
            conversationStart         = (ConvertTo-IsoText $c.Start)
            queueNames                = (Resolve-Names -Map $names.Queue -Ids $c.QueueIds)
            wrapUpCodes               = (Resolve-Names -Map $names.WrapUp -Ids $c.WrapUpCodeIds)
            communicationId           = $r.CommunicationId
            utteranceStartMs          = $r.UtteranceStartMs
            speaker                   = $r.Speaker
            utterance                 = (Protect-Text $r.UtteranceRaw)
            bestMatchingArticleId     = $r.BestDocumentId
            bestMatchingArticle       = $r.BestTitle
            matchConfidence           = ConvertTo-Rounded ([double]$r.BestConfidence)
            secondArticle             = $r.SecondTitle
            secondConfidence          = $(if ($null -ne $r.SecondConfidence) { ConvertTo-Rounded ([double]$r.SecondConfidence) } else { $null })
            surfacedLive              = $(if ($r.SurfacedLive) { 'Y' } else { 'N' })
            conversationHadAnySuggestion = $(if ($r.ConversationHadLiveSuggestion) { 'Y' } else { 'N' })
            classification            = $r.Classification
            replaySearchId            = $r.SearchId
        })
    }
    Export-Report -Name 'MissedOpportunities.csv' -Rows $missedRows

    # ===========================================================================================================
    # OUTPUT 4 - PhraseRecommendations.csv
    #   Real utterances that match an under-performing article (best or runner-up at >= GapMaxConfidence) and
    #   were not surfaced live, grouped by normalised text with frequency. Add the high-frequency ones as
    #   alternative phrases on the article.
    # ===========================================================================================================
    $targetStatuses = @('Dormant', 'Dormant-MatchesCalls', 'Low', 'NeedsReview')
    $perfIndex = @{}; foreach ($a in $articlePerf) { $perfIndex[$a.articleId] = $a }
    $phraseAgg = @{}   # key docId|norm -> @{ DocId; Norm; Raw; Count; MaxConf; Convs(hashtable) }
    foreach ($r in $replayRows) {
        if ($r.SurfacedLive) { continue }
        foreach ($res in @($r.AllResults)) {
            if ($res.Confidence -lt $GapMaxConfidence) { continue }
            $perf = if ($perfIndex.ContainsKey($res.DocumentId)) { $perfIndex[$res.DocumentId] } else { $null }
            $underperforming = ($null -ne $perf -and ($targetStatuses -contains $perf.status))
            if (-not $underperforming -and $res.Rank -gt 1) { continue }   # healthy articles: only keep clear misses
            if (-not $underperforming -and $r.Classification -ne 'Missed') { continue }
            $key = "$($res.DocumentId)|$($r.UtteranceNorm)"
            if (-not $phraseAgg.ContainsKey($key)) { $phraseAgg[$key] = @{ DocId = $res.DocumentId; Title = $res.Title; Norm = $r.UtteranceNorm; Raw = $r.UtteranceRaw; Count = 0; MaxConf = 0; Rank = $res.Rank; Convs = @{} } }
            $e = $phraseAgg[$key]
            $e.Count++; $e.Convs[$r.ConversationId] = $true
            if ($res.Confidence -gt $e.MaxConf) { $e.MaxConf = $res.Confidence }
            if ($res.Rank -lt $e.Rank) { $e.Rank = $res.Rank }
        }
    }
    $phraseRows = foreach ($key in $phraseAgg.Keys) {
        $e = $phraseAgg[$key]
        $inv = if ($inventory.ContainsKey($e.DocId)) { $inventory[$e.DocId] } else { $null }
        $existing = @(); if ($inv) { $existing = @($inv.Alternatives) }
        $already = ($existing | ForEach-Object { Get-NormalizedText $_ }) -contains $e.Norm
        New-Object -TypeName PSObject -Property ([ordered]@{
            articleId              = $e.DocId
            title                  = $(if ($inv) { $inv.Title } else { $e.Title })
            articleStatus          = $(if ($perfIndex.ContainsKey($e.DocId)) { $perfIndex[$e.DocId].status } else { '' })
            existingPhraseCount    = $existing.Count
            existingPhrases        = ($existing -join ' | ')
            suggestedPhrase        = (Protect-Text $e.Raw)
            suggestedPhraseWords   = (Get-WordCount $e.Norm)
            frequency              = $e.Count
            distinctConversations  = $e.Convs.Count
            bestRankInReplay       = $e.Rank
            maxConfidence          = ConvertTo-Rounded ([double]$e.MaxConf)
            alreadyAnAlternative   = $already
        })
    }
    Export-Report -Name 'PhraseRecommendations.csv' -Rows ($phraseRows | Sort-Object -Property @{ Expression = 'frequency'; Descending = $true }, @{ Expression = 'maxConfidence'; Descending = $true })

    # ===========================================================================================================
    # OUTPUT 5 - KnowledgeGaps.csv
    #   a) WrapUpCode x Queue clusters of calls where nothing surfaced and (if replayed) nothing matched well
    #   b) utterance keyword clusters of low-confidence utterances
    #   c) Knowledge Optimizer unanswered groups for the Assistant app
    # ===========================================================================================================
    function Get-GapRating {
        <# Volume-led rating; replay confidence (when known) only pushes a cluster up, never down. #>
        param([int] $Volume, $Conf)
        $lowConf = ($null -ne $Conf -and [double]$Conf -lt 0.2)
        if ($Volume -ge 10) { return 'High' }
        if ($Volume -ge 5 -or ($lowConf -and $Volume -ge 2)) { return 'Medium' }
        return 'Low' }

    $gapRows = @()
    # (a) wrap-up x queue
    $noKnowledgeConvs = @($conversations | Where-Object { -not $liveDocsByConv.ContainsKey($_.ConversationId) })
    $groups = @{}
    foreach ($c in $noKnowledgeConvs) {
        $rp = if ($replayedConvs.ContainsKey($c.ConversationId)) { $replayedConvs[$c.ConversationId] } else { $null }
        if ($rp -and $rp.Transcript -and $null -ne $rp.BestConfidence -and [double]$rp.BestConfidence -ge $GapMaxConfidence) { continue }  # replay found a usable article -> missed opportunity, not a gap
        $wc = if ($c.WrapUpCodeIds.Count -gt 0) { Resolve-Names -Map $names.WrapUp -Ids $c.WrapUpCodeIds } else { '(no wrap-up)' }
        $qn = if ($c.QueueIds.Count -gt 0) { Resolve-Names -Map $names.Queue -Ids $c.QueueIds } else { '(no queue)' }
        $key = "$wc|$qn"
        if (-not $groups.ContainsKey($key)) { $groups[$key] = @{ WrapUp = $wc; Queue = $qn; Convs = @(); Replayed = 0; Confs = @(); BestDocs = @{}; BestConf = @{} } }
        $g = $groups[$key]; $g.Convs += $c.ConversationId
        if ($rp -and $rp.Transcript) { $g.Replayed++; if ($null -ne $rp.BestConfidence) { $g.Confs += [double]$rp.BestConfidence } }
        foreach ($r in @($replayRows | Where-Object { $_.ConversationId -eq $c.ConversationId -and $_.BestDocumentId })) {
            $g.BestDocs[$r.BestTitle] = [int](Get-Prop $g.BestDocs $r.BestTitle) + 1
            if ([double]$r.BestConfidence -gt [double](Get-Prop $g.BestConf $r.BestTitle)) { $g.BestConf[$r.BestTitle] = [double]$r.BestConfidence }
        }
    }
    foreach ($key in $groups.Keys) {
        $g = $groups[$key]
        $avg = if ($g.Confs.Count -gt 0) { ConvertTo-Rounded (($g.Confs | Measure-Object -Average).Average) } else { $null }
        $closest = ''; if ($g.BestDocs.Count -gt 0) { $closest = (@($g.BestDocs.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1)[0]).Key }
        $gapRows += New-Object -TypeName PSObject -Property ([ordered]@{
            clusterType            = 'WrapUpCode x Queue'
            topic                  = $g.WrapUp
            queue                  = $g.Queue
            callVolume             = $g.Convs.Count
            conversationsReplayed  = $g.Replayed
            avgBestConfidence      = $avg
            closestExistingArticle = $closest
            closestArticleConfidence = $(if ($closest) { ConvertTo-Rounded ([double](Get-Prop $g.BestConf $closest)) } else { $null })
            samplePhrases          = ''
            sampleConversationIds  = (@($g.Convs | Select-Object -First 5) -join '; ')
            gapRating              = (Get-GapRating -Volume $g.Convs.Count -Conf $avg)
            source                 = 'Analytics + replay'
        })
    }
    # (b) utterance keyword clusters (low-confidence / no-match utterances only)
    $lowRows = @($replayRows | Where-Object { $_.Classification -in @('LowConfidence', 'NoMatch') })
    if ($lowRows.Count -gt 0) {
        $df = @{}
        $tokensByRow = @{}
        $idx = 0
        foreach ($r in $lowRows) { $toks = @(Get-SignificantTokens $r.UtteranceNorm | Select-Object -Unique); $tokensByRow[$idx] = $toks; foreach ($t in $toks) { $df[$t] = [int](Get-Prop $df $t) + 1 }; $idx++ }
        $clusters = @{}
        for ($k = 0; $k -lt $lowRows.Count; $k++) {
            $toks = $tokensByRow[$k]
            if ($toks.Count -eq 0) { continue }
            $top = @($toks | Sort-Object -Property @{ Expression = { $df[$_] }; Descending = $true }, @{ Expression = { $_ } } | Select-Object -First 2)
            if (($df[$top[0]] -lt 2)) { continue }   # singleton topics are noise
            $sig = ($top | Sort-Object) -join ' + '
            if (-not $clusters.ContainsKey($sig)) { $clusters[$sig] = @{ Rows = @(); Convs = @{}; BestDocs = @{}; BestConf = @{} } }
            $cl = $clusters[$sig]; $cl.Rows += $lowRows[$k]; $cl.Convs[$lowRows[$k].ConversationId] = $true
            if ($lowRows[$k].BestTitle) {
                $cl.BestDocs[$lowRows[$k].BestTitle] = [int](Get-Prop $cl.BestDocs $lowRows[$k].BestTitle) + 1
                if ([double]$lowRows[$k].BestConfidence -gt [double](Get-Prop $cl.BestConf $lowRows[$k].BestTitle)) { $cl.BestConf[$lowRows[$k].BestTitle] = [double]$lowRows[$k].BestConfidence }
            }
        }
        foreach ($sig in $clusters.Keys) {
            $cl = $clusters[$sig]
            if ($cl.Convs.Count -lt 2) { continue }
            $avg = ConvertTo-Rounded (($cl.Rows | ForEach-Object { [double]$_.BestConfidence } | Measure-Object -Average).Average)
            $closest = ''; if ($cl.BestDocs.Count -gt 0) { $closest = (@($cl.BestDocs.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1)[0]).Key }
            $samples = @($cl.Rows | Select-Object -First 3 | ForEach-Object { Protect-Text $_.UtteranceRaw })
            $gapRows += New-Object -TypeName PSObject -Property ([ordered]@{
                clusterType            = 'Utterance keywords'
                topic                  = $sig
                queue                  = (Join-Unique ($cl.Rows | ForEach-Object { Resolve-Names -Map $names.Queue -Ids $convIndex[$_.ConversationId].QueueIds }))
                callVolume             = $cl.Convs.Count
                conversationsReplayed  = $cl.Convs.Count
                avgBestConfidence      = $avg
                closestExistingArticle = $closest
                closestArticleConfidence = $(if ($closest) { ConvertTo-Rounded ([double](Get-Prop $cl.BestConf $closest)) } else { $null })
                samplePhrases          = ($samples -join ' || ')
                sampleConversationIds  = (@($cl.Convs.Keys | Select-Object -First 5) -join '; ')
                gapRating              = (Get-GapRating -Volume $cl.Convs.Count -Conf $avg)
                source                 = 'Transcript replay'
            })
        }
    }
    # (c) Knowledge Optimizer unanswered groups
    foreach ($u in $unanswered) {
        $stats = Get-Prop $u 'statistics'
        $hits = [int](Get-Prop $stats 'totalPhraseHitCount'); if (-not $hits) { $hits = [int](Get-Prop $stats 'unlinkedPhraseHitCount') }
        $pgLabels = @(); $samples = @()
        foreach ($pg in @(Get-Prop $u 'phraseGroups')) { $l = Get-Prop $pg 'label'; if ($l) { $pgLabels += $l }; foreach ($ph in @(Get-Prop $pg 'phrases')) { if ($samples.Count -lt 3) { $samples += (Protect-Text (Get-Prop $ph 'text')) } } }
        $sugDocs = @(); foreach ($sd in @(Get-Prop $u 'suggestedDocuments')) { $id = Get-Prop $sd 'id'; if ($inventory.ContainsKey("$id")) { $sugDocs += $inventory[$id].Title } elseif ($id) { $sugDocs += $id } }
        $gapRows += New-Object -TypeName PSObject -Property ([ordered]@{
            clusterType            = 'Knowledge Optimizer unanswered group'
            topic                  = Get-Prop $u 'label'
            queue                  = ''
            callVolume             = $hits
            conversationsReplayed  = 0
            avgBestConfidence      = $null
            closestExistingArticle = ($sugDocs -join '; ')
            closestArticleConfidence = $null
            samplePhrases          = (@($pgLabels | Select-Object -First 3) + $samples -join ' || ')
            sampleConversationIds  = ''
            gapRating              = (Get-GapRating -Volume $hits -Conf $null)
            source                 = 'Knowledge Optimizer API'
        })
    }
    Export-Report -Name 'KnowledgeGaps.csv' -Rows ($gapRows | Sort-Object -Property @{ Expression = 'callVolume'; Descending = $true })

    # ---- Run summary -----------------------------------------------------------------------------------------
    $elapsed = (Get-Date) - $runStart
    $summary = @(
        "Genesys Cloud knowledge audit - run summary",
        "Started: $runStart   Elapsed: $([int]$elapsed.TotalMinutes) min",
        "Window (UTC): $(Get-UtcIso $StartDate) -> $(Get-UtcIso $EndDate)   Region: $Region   KB: $KnowledgeBaseId",
        "Conversations in scope: $($conversations.Count)   Suggestion source: $suggestionSource",
        "Suggestions (API): $($suggestions.Count)   Docs with aggregate activity: $($aggByDoc.Count)   Feedback records: $($feedback.Count)",
        "Replay rows: $($replayRows.Count)   Missed: $(@($replayRows | Where-Object { $_.Classification -eq 'Missed' }).Count)   Low-confidence: $(@($replayRows | Where-Object { $_.Classification -in @('LowConfidence','NoMatch') }).Count)",
        "Articles: $($articlePerf.Count)   Dormant: $(@($articlePerf | Where-Object { $_.status -like 'Dormant*' }).Count)   Low: $(@($articlePerf | Where-Object { $_.status -eq 'Low' }).Count)   Healthy: $(@($articlePerf | Where-Object { $_.status -eq 'Healthy' }).Count)",
        "API calls: $($script:RunStats.ApiCalls)   Retries: $($script:RunStats.Retries)   429s: $($script:RunStats.RateLimited)",
        "PII masking: $(if ($IncludeUtteranceText) { 'OFF (full text written)' } else { 'ON' })",
        "Warnings ($($script:Warnings.Count)):"
    ) + @($script:Warnings | ForEach-Object { "  - $_" })
    Set-Content -Path (Join-Path -Path $OutputPath -ChildPath 'RunSummary.txt') -Value $summary -Encoding UTF8
    foreach ($l in $summary) { Write-Log $l }
    Write-Log "Done. Output folder: $OutputPath"
}
catch {
    Write-Log ("FATAL: {0}" -f $_.Exception.Message) -Level ERROR
    if ($_.ScriptStackTrace) { Write-Log $_.ScriptStackTrace -Level ERROR }
    throw
}
finally {
    if ($transcriptStarted) { try { Stop-Transcript | Out-Null } catch { } }
}
