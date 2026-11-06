<#
  Dry-run harness: starts nothing itself - start the mock first:
      python3 tests/mock-genesys-api.py 8765
  then run this file. It forces Constrained Language Mode so CLM violations show up locally.
#>
param([int] $Port = 8765, [string] $OutputPath = (Join-Path $PSScriptRoot 'out'))
$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'
$env:GENESYS_CLIENT_ID = 'mock-client'
$env:GENESYS_CLIENT_SECRET = 'mock-secret'
& (Join-Path $PSScriptRoot '..\Invoke-GenesysKnowledgeAudit.ps1') `
    -StartDate '2026-10-01' -EndDate '2026-11-01' -KnowledgeBaseId 'kb-0001' `
    -OutputPath $OutputPath -ApiBaseUri "http://127.0.0.1:$Port" -LoginBaseUri "http://127.0.0.1:$Port" `
    -MaxConversationsForReplay 10 -Verbose
