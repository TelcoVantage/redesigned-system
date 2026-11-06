# Genesys Cloud Agent Copilot – Knowledge Surfacing Audit (voice)

`Invoke-GenesysKnowledgeAudit.ps1` answers, for every voice conversation in a date range:

1. Did Agent Copilot surface knowledge, which articles, and did agents open / use / rate them?
2. Which customer (or agent) utterances **should** have matched an article but nothing (or only a weak match) surfaced?
3. Which articles are dormant or low-performing versus the ones that surface most?
4. Which real call phrases could be added as **alternative phrases** to under-performing articles?
5. Which recurring topics (wrap-up code × queue, utterance clusters, Knowledge Optimizer unanswered groups) have **no matching article**?

It is a single, self-contained script written for **PowerShell Constrained Language Mode** (Windows PowerShell 5.1 or PowerShell 7+), uses only
`Invoke-WebRequest` / `Invoke-RestMethod`, built-in cmdlets, hashtables, PSObjects and `Export-Csv`, and authenticates with an OAuth
client-credentials grant whose secret is read from an environment variable or a prompted `SecureString`.

Every endpoint, field and enum used below was validated on 2026-11-06 against the published Platform API definition
(`https://api.mypurecloud.com/api/v2/docs/swagger`, 2,249 paths). Items that the swagger cannot answer (retention windows, product
behaviour) are flagged **[uncertain]**.

---

## 1. Data sources, endpoints and permissions

| Step | Endpoint | Why | Permission(s) on the OAuth client's role |
|---|---|---|---|
| Auth | `POST https://login.mypurecloud.com.au/oauth/token` (`grant_type=client_credentials`) | Token at runtime | – |
| Availability | `GET /api/v2/analytics/conversations/details/jobs/availability` | Warn if the window is newer than the analytics data lake | `analytics:conversationDetail:view` |
| Conversations | `POST /api/v2/analytics/conversations/details/jobs` → `GET …/jobs/{jobId}` → `GET …/jobs/{jobId}/results?cursor=` | Async job for large ranges; `segmentFilters` `mediaType=voice` (+ optional `queueId` OR-clause); cursor paging; post-filtered on `conversationStart` | `analytics:conversationDetail:view` |
| Copilot suggestions | `GET /api/v2/conversations/{conversationId}/suggestions?pageSize=200` (cursor via `nextUri`) | **The persisted record of what Copilot raised.** Each `Suggestion` carries `type` (KnowledgeArticle, KnowledgeSearch, SuggestedKnowledgeAnswer, Faq, Article, CannedResponse, Script…), `state` (Suggested / Accepted / Dismissed / Failed / Rated), `triggerType` (Intent, ConversationStart, ConversationEnd, ConversationTransfer, Fallback), `knowledgeSearch.confidence`, `knowledgeSearch.document.id`, `knowledgeSearch.searchId`, and `context.queryStatement` (the text Copilot actually searched with), `context.utterance.id`, `context.participantType` | `conversation:suggestion:view` |
| Knowledge usage counters | `POST /api/v2/analytics/knowledge/aggregates/query` | Article- and conversation-level counts: `nKnowledgeSearch`, `nKnowledgeSearchAnswered`, `nKnowledgeSearchUnanswered`, `nKnowledgeDocumentSurfaced`, `nKnowledgeDocumentPresented`, `nKnowledgeDocumentViewed`, `nKnowledgeDocumentCopied`, `nKnowledgeDocumentFeedback`; grouped by `knowledgeBaseDocumentId`, `conversationId`, `feedbackRating`; filtered by `knowledgeBaseId` and `appType=Assistant`. Interval is chunked (7 days, falling back to 1 day on HTTP 400) | `analytics:knowledgeAggregate:view` |
| Article inventory | `GET /api/v2/knowledge/knowledgebases/{kbId}` and `GET …/documents?includeDrafts=true&expand=category,labels&pageSize=200` (cursor) | Title, state, visible, category, labels, `alternatives[].phrase` | `knowledge:knowledgebase:view`, `knowledge:document:view` |
| Article feedback | `GET /api/v2/knowledge/knowledgebases/{kbId}/documents/feedback?pageSize=200` (cursor) | Thumbs up/down with `rating`, `reason`, `queryType`, `surfacingMethod`, `application.type`, `conversationContext.conversation.id`, `user.id` | `knowledge:feedback:view` |
| Unanswered groups | `GET /api/v2/knowledge/knowledgebases/{kbId}/unanswered/groups?app=Assistant&dateStart&dateEnd` | Knowledge Optimizer "unanswered queries" clusters with phrase groups, hit counts and suggested documents | `knowledge:groups:view` |
| Transcripts | `GET /api/v2/speechandtextanalytics/conversations/{id}/communications/{communicationId}/transcripturl` → pre-signed S3 URL (fetched **without** the bearer token) | `transcripts[].phrases[] { text, startTimeMs, participantPurpose: internal \| external }`. `communicationId` = the agent voice `session.sessionId` from analytics | `recording:recording:view`, `speechAndTextAnalytics:data:view` |
| Replay | `POST /api/v2/knowledge/knowledgebases/{kbId}/documents/search` body `{ query, queryType: AutoSearch, sortBy: ConfidenceScore, pageSize }` | Returns `results[] { confidence, document{id,title} }` + `searchId`; **no** `confidenceThreshold` is applied so you see what would have matched and how strongly | `knowledge:knowledgebase:search` |
| Names | `GET /api/v2/routing/queues?id=…`, `GET /api/v2/routing/wrapupcodes`, `GET /api/v2/users?id=…&state=any` | Resolve IDs to names | `routing:queue:view`, `routing:wrapupCode:view`, `directory:user:view` |

The OAuth client also needs **division** access to the conversations/queues it reports on (assign the role in every relevant division).

### What is and is not retrievable

| Question | Answer |
|---|---|
| Did Copilot attempt to surface knowledge in conversation X? | **Yes, via API** – `GET /conversations/{id}/suggestions` returns every suggestion raised, with trigger, state and confidence. Aggregates (`nKnowledgeSearch` by `conversationId`) corroborate it. |
| Which articles were suggested / opened / copied / rated? | **Yes** – suggestion `state` (Accepted = opened, Dismissed, Rated) plus aggregates (`…Viewed`, `…Copied`) plus the feedback listing (rating value, since `state=Rated` does not carry the rating). |
| What query text did Copilot search with? | **Yes** – `context.queryStatement` on the suggestion (customer speech → masked by default). |
| Which utterance triggered a suggestion? | **Partially** – `context.utterance.id` is an opaque ID; the transcript JSON has no utterance IDs, so the script joins at conversation level, not utterance level **[not exposed]**. |
| Searches Copilot ran that returned nothing / below threshold | **Not as per-conversation rows.** Aggregates give counts (`nKnowledgeSearchUnanswered` per conversation); the unanswered-groups endpoint gives Knowledge Optimizer's clustered phrases. Whether a `state=Failed` suggestion is persisted for below-threshold searches is **[uncertain]**; the script keeps them if returned. Hence the transcript-replay fallback. |
| Copilot "knowledge fabric / answer generation" feedback detail (the Knowledge Optimizer "queries with negative feedback" panel) | **UI only** – the help centre states that panel's data is kept **10 days** and no public endpoint exposes it. |
| Retention of `GET /conversations/{id}/suggestions` | **[uncertain]** – not documented. Copilot summaries have a configurable `retentionSeconds`; assume suggestions are similar and run the audit frequently (weekly) rather than relying on month-old data. |
| Transcript retention | Follows the recording retention policy of your QM policies (help centre). Transcript *content* search is limited to 60 days; the per-communication transcript URL works as long as the recording exists. |
| Analytics conversation details | Available for the org's interaction-data retention (minimum 90 days, longer for CX 3). |
| Knowledge aggregates interval limits | The service rejects some long/fine-grained intervals (community reports) – the script chunks by 7 days and retries at 1 day on HTTP 400. |
| Client-credentials access to `GET /conversations/{id}/suggestions` | **[uncertain]** – documented permission is `conversation:suggestion:view`; if your client receives 403 the script warns once and falls back to aggregates automatically. |

### Important caveat on the replay

The replay calls the Knowledge Search API directly. Live Copilot surfacing is a pipeline (NLU intent → Copilot rules → auto-search with the
assistant's configured threshold / answer mode). Replay therefore shows **what the knowledge base would return for that utterance and at what
confidence**, which is exactly what you need to decide on phrases and content, but it is not a byte-for-byte reproduction of Copilot's decision.
Each replay search is recorded by knowledge analytics as a search; by default the script sends **no** `application` / `conversationContext`
so replays are distinguishable from (and do not pollute) live Copilot statistics in Knowledge Optimizer. Use `-TagReplaySearchesWithContext`
only if you deliberately want them attributed to the Assistant app.

---

## 2. Output files (CSV, UTF-8, ISO-8601 UTC dates – ready for Power BI)

| File | Grain | Key columns |
|---|---|---|
| `ConversationKnowledgeActivity.csv` | one row per voice conversation with an agent interaction | conversationId, start/end, queue(s), agent(s), wrap-up code(s), `suggestionSource` (SuggestionsApi / KnowledgeAggregates / None), `suggestionsShown`, `knowledgeSuggestionsShown`, `suggestionTriggerTypes`, `articleIdsShown`, `articlesShown`, `articlesAccepted`, `articlesDismissed`, `articlesRated`, `agg*` counters (searches, answered, unanswered, surfaced, viewed, copied), `feedbackPositive/Negative`, `transcriptReplayed`, `utterancesReplayed`, `replayBestConfidence`, `missedOpportunities` |
| `ArticlePerformance.csv` | one row per article (inventory ∪ anything that surfaced) | articleId, title, category, labels, state, visible, `timesSuggested` (max of API count and aggregates), `timesPresented`, `timesViewed`, `timesUsed` (copied), `timesAccepted`, `timesDismissed`, `avgSuggestionConfidence`, `positiveFeedback`, `negativeFeedback`, `alternativePhraseCount`, `replayBestMatchCount`, `replayMissedCount`, **`status`** |
| `MissedOpportunities.csv` | one row per replayed utterance | conversationId, queue, wrap-up, communicationId, `utteranceStartMs`, speaker, `utterance` (masked unless opted in), best article + `matchConfidence`, second article + confidence, **`surfacedLive` (Y/N)**, `conversationHadAnySuggestion`, **`classification`**, `replaySearchId` |
| `PhraseRecommendations.csv` | one row per (article, candidate phrase) | articleId, title, `articleStatus`, `existingPhraseCount`, `existingPhrases`, `suggestedPhrase`, `frequency`, `distinctConversations`, `bestRankInReplay`, `maxConfidence`, `alreadyAnAlternative` |
| `KnowledgeGaps.csv` | one row per cluster | `clusterType` (WrapUpCode × Queue / Utterance keywords / Knowledge Optimizer unanswered group), topic, queue, `callVolume`, `conversationsReplayed`, `avgBestConfidence`, `closestExistingArticle`, `closestArticleConfidence`, `samplePhrases`, `sampleConversationIds`, **`gapRating`**, source |
| `RunSummary.txt`, `KnowledgeAudit-<stamp>.log`, `KnowledgeAudit-<stamp>.transcript.txt` | run metadata | counts, API call / retry / 429 statistics, every warning raised |

**Article `status`** – `NotSearchable` (draft/archived/invisible) · `Dormant` (0 surfaces in window) · `Dormant-MatchesCalls` (0 surfaces but
replay shows real calls it would answer – highest priority) · `Low` (≤ `-LowSurfaceThreshold` surfaces, or surfaced but never viewed/copied) ·
`NeedsReview` (negative feedback outweighs positive) · `Healthy`.

**Utterance `classification`** – `SurfacedLive` (best article was shown in that conversation) · `Missed` (best match ≥ `-MissedOpportunityMinConfidence`
but never surfaced) · `Marginal` (between the two cut-offs) · `LowConfidence` (< `-GapMaxConfidence`) · `NoMatch`. These are **reporting labels**
you can re-slice in Power BI; they are not Copilot thresholds and the script never changes any Copilot setting.

---

## 3. How to run

### Prerequisites
* Windows PowerShell 5.1 or PowerShell 7.x; Constrained Language Mode is fine (the dry-run harness forces it).
* An OAuth **client credentials** client in Genesys Cloud with a role containing the permissions in section 1, assigned in the right divisions.
* Voice transcription enabled on the audited queues (Speech & Text Analytics) if you want the replay.

### Secrets
```powershell
$env:GENESYS_CLIENT_ID     = '<client id>'
$env:GENESYS_CLIENT_SECRET = '<client secret>'     # or omit and you will be prompted for a SecureString
```
Alternatively set `$EmbeddedClientId` / `$EmbeddedClientSecret` in the block near the top of the script. Precedence is
parameter → environment variable → embedded value → prompt. Embedded values are clear text: keep the file out of source control
and lock down its file permissions. The script logs a warning whenever the embedded secret is used.

### Typical weekly run (Sydney region is the default)
```powershell
.\Invoke-GenesysKnowledgeAudit.ps1 `
    -StartDate '2026-10-27' -EndDate '2026-11-03' `
    -KnowledgeBaseId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' `
    -OutputPath 'C:\Reports\KnowledgeAudit\2026-W44'
```

### Restrict to queues, replay more calls, keep full utterance text (PII – opt in)
```powershell
.\Invoke-GenesysKnowledgeAudit.ps1 -StartDate '2026-10-01' -EndDate '2026-11-01' `
    -KnowledgeBaseId '<kb id>' -QueueId '<queue id 1>','<queue id 2>' `
    -MaxConversationsForReplay 2000 -MaxUtterancesPerConversation 8 -IncludeAgentUtterances `
    -IncludeUtteranceText -OutputPath 'D:\Audit\Oct'
```

### Large org, cheapest run: aggregates first, no transcripts
```powershell
.\Invoke-GenesysKnowledgeAudit.ps1 -StartDate '2026-10-01' -EndDate '2026-11-01' `
    -KnowledgeBaseId '<kb id>' -SuggestionLookupMode AggregatesFirst -SkipTranscriptReplay -OutputPath 'D:\Audit\Oct'
```

### Other region / proxy
```powershell
.\Invoke-GenesysKnowledgeAudit.ps1 ... -Region 'mypurecloud.com'      # or -ApiBaseUri / -LoginBaseUri overrides
```

### Dry run against the bundled mock (no Genesys org needed)
```bash
python3 tests/mock-genesys-api.py 8765
```
```powershell
pwsh -File tests/run-mock.ps1        # forces ConstrainedLanguage, writes tests/out/
```
The mock also returns one HTTP 429 (with `Retry-After`) and one HTTP 502 to exercise the retry path; set `MOCK_SUGGESTIONS_403=1` to
exercise the aggregates fallback.

### Parameters at a glance

| Parameter | Default | Notes |
|---|---|---|
| `-StartDate` / `-EndDate` | required | Local time unless you write `…Z`; converted to UTC |
| `-QueueId` | all | one or more queue IDs |
| `-KnowledgeBaseId` | required | the KB Copilot uses on those queues |
| `-OutputPath` | `.\KnowledgeAudit` | created if missing |
| `-Region` | `mypurecloud.com.au` | |
| `-ClientId` / `-ClientSecret` | env vars / prompt | never hard-code |
| `-IncludeUtteranceText` | off | **opt-in** to write full customer speech |
| `-MaskLength` | 60 | truncation length when masking |
| `-SkipTranscriptReplay` / `-SkipSuggestionsApi` | off | speed switches |
| `-SuggestionLookupMode` | `All` | `AggregatesFirst` = only call the suggestions endpoint for conversations with knowledge events |
| `-MaxConversationsForReplay` | 500 | conversations with **no** live suggestion are replayed first |
| `-MaxUtterancesPerConversation` / `-MinUtteranceWords` | 6 / 5 | longest customer utterances are replayed |
| `-IncludeAgentUtterances` | off | also replay agent speech |
| `-ReplayResultCount` | 3 | ranked results per utterance (max 30) |
| `-MissedOpportunityMinConfidence` / `-GapMaxConfidence` | 0.5 / 0.35 | reporting cut-offs only |
| `-LowSurfaceThreshold` | 3 | `Low` status cut-off |
| `-AggregateAppType` | `Assistant` | `''` to include all knowledge apps |
| `-TagReplaySearchesWithContext` | off | attribute replay searches to the Assistant app |

### Run-time expectations
* One API call per conversation for suggestions (plus one per agent communication for transcripts and one per replayed utterance).
  Genesys rate limits are honoured (HTTP 429 `Retry-After`, exponential backoff 2–64 s, 6 attempts, token refresh on 401).
* 5,000 conversations with full replay ≈ 5,000 + 5,000 + 30,000 calls ≈ 2–3 hours at the default limits. Use `-QueueId`,
  `-SuggestionLookupMode AggregatesFirst` and `-MaxConversationsForReplay` to size the run.
* Nothing fails silently: every warning is counted, logged, repeated in `RunSummary.txt`, and a fatal error stops the run with a stack line.

---

## 4. Actioning the results and review cadence

**Weekly (knowledge analyst, 30 min)**
1. `ArticlePerformance.csv` – filter `status = Dormant-MatchesCalls`: these articles answer real calls but never surface. Open
   `PhraseRecommendations.csv` for each, add the top `frequency` phrases (that are not `alreadyAnAlternative`) as **alternative phrases**,
   then publish. Re-run with `-IncludeUtteranceText` on a short window if the masked phrase is too truncated to use.
2. `MissedOpportunities.csv` – filter `classification = Missed` and `surfacedLive = N`. If the same article appears across queues, check the
   Copilot rules / queue assignment for that assistant; if it is one queue, check the queue's knowledge base mapping.
3. `KnowledgeGaps.csv` – `gapRating = High`: create or extend an article. Use `samplePhrases` as the first alternative phrases of the new article
   and `closestExistingArticle` (+ its confidence) to decide between a new article and a rewrite/merge.

**Fortnightly (content owner)**
4. `status = NeedsReview` (negative feedback) – read the feedback comments (`documents/feedback` `comment`), rewrite the article body, and
   check `avgSuggestionConfidence`: low confidence with high dismissals usually means the **title** is ambiguous rather than the body.
5. `status = Low` with high `alternativePhraseCount` – phrases exist but calls still do not match: the phrases are probably agent/internal
   wording, not customer wording. Replace them with phrases from `PhraseRecommendations.csv`.
6. `NotSearchable` articles that appear as `replayBestMatchCount > 0` – draft/invisible content that calls need: publish it.

**Monthly (ops / Copilot owner)**
7. Trend `knowledgeSuggestionsShown`, `articlesAccepted` and `missedOpportunities` per queue from `ConversationKnowledgeActivity.csv`
   (Power BI). A queue with searches but few accepts points at relevance; a queue with no searches at all points at Copilot not being live on
   that queue (`assistants/{id}/queues`) or no transcription.
8. Archive articles that stay `Dormant` for three consecutive months **and** have `replayBestMatchCount = 0`.
9. Re-baseline: after each round of phrase changes, compare `timesSuggested` / `timesAccepted` of the edited articles between runs
   (keep each run in its own dated folder).

Suggested hygiene when adding phrases: 5–15 phrases per article, customer wording ("my card got declined overseas"), one intent per
article; if a phrase matches two articles at similar confidence, merge or sharpen the titles rather than adding more phrases.

---

## 5. Assumptions made in the design

* Copilot auto-search uses the knowledge base passed in `-KnowledgeBaseId`; the script also reports `knowledgeBaseIds` seen on each
  analytics conversation so you can spot mismatches.
* "Viewed" = suggestion `state = Accepted` or `nKnowledgeDocumentViewed`; "Used" = `nKnowledgeDocumentCopied`; Copilot thumbs up/down land in
  the knowledge feedback listing with `application.type = Assistant` (believed, not verifiable from the swagger).
* Conversations with no agent `interact` segment (abandoned / IVR-only) are skipped – Copilot cannot act on them.
* Customer speech is masked by default (digit runs, e-mails, truncation). Agent names are staff data and are written as-is.
* The script never writes to Genesys Cloud (no feedback, no phrase changes); applying phrases is a deliberate manual step.

## 6. Questions worth confirming (none block running the script)

1. Does your OAuth client get HTTP 200 or 403 on `GET /api/v2/conversations/{id}/suggestions`? (403 → the script falls back to aggregates;
   if you want the richer data, a user-context OAuth client (auth code / implicit) is the alternative.)
2. Is it acceptable that replay searches are counted in knowledge analytics (as un-attributed searches by default)? If not, run with
   `-SkipTranscriptReplay` and rely on `nKnowledgeSearchUnanswered` + unanswered groups.
3. Which PowerShell host will run it (5.1 or 7)? Both are supported; the 7.x path was exercised end-to-end in Constrained Language Mode
   against the bundled mock, the 5.1 `WebException` path is implemented but could not be executed in this environment.
