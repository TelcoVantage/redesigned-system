#!/usr/bin/env python3
"""
Minimal mock of the Genesys Cloud endpoints used by Invoke-GenesysKnowledgeAudit.ps1.
Purpose: dry-run the script (including Constrained Language Mode) without a real org.
Run:  python3 tests/mock-genesys-api.py 8765
Then: see tests/run-mock.ps1
It also simulates one HTTP 429 (with Retry-After) and one HTTP 500 to exercise the retry logic.
"""
import json, re, sys, random, os
from http.server import BaseHTTPRequestHandler, HTTPServer

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
KB = "kb-0001"
DOCS = [
    {"id": "doc-1", "title": "How to reset your online banking password", "state": "Published", "visible": True,
     "alternatives": [{"phrase": "forgot password", "autocomplete": True}], "category": {"id": "c1", "name": "Online banking"}, "labels": [{"id": "l1", "name": "Self-service"}],
     "dateModified": "2026-09-01T00:00:00Z", "datePublished": "2026-09-01T00:00:00Z", "words": ["password", "reset", "login", "online", "banking", "locked"]},
    {"id": "doc-2", "title": "Dispute a card transaction", "state": "Published", "visible": True, "alternatives": [],
     "category": {"id": "c2", "name": "Cards"}, "labels": [], "dateModified": "2026-08-01T00:00:00Z", "datePublished": "2026-08-01T00:00:00Z",
     "words": ["dispute", "transaction", "charge", "card", "fraud", "unauthorised", "refund"]},
    {"id": "doc-3", "title": "Update your mailing address", "state": "Published", "visible": True, "alternatives": [{"phrase": "change address", "autocomplete": True}, {"phrase": "moved house", "autocomplete": True}],
     "category": {"id": "c3", "name": "Profile"}, "labels": [], "dateModified": "2026-07-01T00:00:00Z", "datePublished": "2026-07-01T00:00:00Z",
     "words": ["address", "moved", "mailing", "update", "house", "postal"]},
    {"id": "doc-4", "title": "Travel notification for overseas card use", "state": "Published", "visible": True, "alternatives": [],
     "category": {"id": "c2", "name": "Cards"}, "labels": [], "dateModified": "2026-06-01T00:00:00Z", "datePublished": "2026-06-01T00:00:00Z",
     "words": ["travel", "overseas", "abroad", "holiday", "card", "notification"]},
    {"id": "doc-5", "title": "Draft: Interest rate change FAQ", "state": "Draft", "visible": True, "alternatives": [],
     "category": {"id": "c4", "name": "Rates"}, "labels": [], "dateModified": "2026-10-01T00:00:00Z", "datePublished": None,
     "words": ["interest", "rate", "change"]},
]
QUEUES = {"q-1": "Retail Support", "q-2": "Cards"}
USERS = {"u-1": "Agent One", "u-2": "Agent Two"}
WRAPS = {"w-1": "Password reset", "w-2": "Card dispute", "w-3": "General enquiry", "w-4": "Hardship"}

def conv(cid, start, end, queue, agent, wrap, comm, abandoned=False):
    parts = [{"participantId": "p-cust", "purpose": "customer", "sessions": [{"mediaType": "voice", "sessionId": "cust-" + comm, "segments": [{"segmentType": "interact", "queueId": queue}]}]}]
    if not abandoned:
        parts.append({"participantId": "p-agent", "purpose": "agent", "userId": agent, "sessions": [{"mediaType": "voice", "sessionId": comm,
                      "segments": [{"segmentType": "alert", "queueId": queue}, {"segmentType": "interact", "queueId": queue}, {"segmentType": "wrapup", "queueId": queue, "wrapUpCode": wrap}]}]})
    return {"conversationId": cid, "conversationStart": start, "conversationEnd": end, "originatingDirection": "inbound", "participants": parts, "knowledgeBaseIds": [KB] if cid == "conv-1" else []}

CONVS = [
    conv("conv-1", "2026-10-02T01:00:00Z", "2026-10-02T01:10:00Z", "q-1", "u-1", "w-1", "comm-1"),   # suggestions via API, accepted
    conv("conv-2", "2026-10-02T02:00:00Z", "2026-10-02T02:08:00Z", "q-2", "u-2", "w-2", "comm-2"),   # nothing surfaced, transcript matches doc-2 -> Missed
    conv("conv-3", "2026-10-03T03:00:00Z", "2026-10-03T03:05:00Z", "q-1", "u-1", "w-4", "comm-3"),   # nothing surfaced, gap (hardship)
    conv("conv-4", "2026-10-03T04:00:00Z", "2026-10-03T04:05:00Z", "q-2", "u-2", "w-3", "comm-4"),   # no transcript
    conv("conv-5", "2026-10-04T05:00:00Z", "2026-10-04T05:05:00Z", "q-1", "u-1", "w-3", "comm-5"),   # travel question, no suggestion -> missed doc-4
    conv("conv-6", "2026-10-04T06:00:00Z", "2026-10-04T06:01:00Z", "q-1", None, None, "comm-6", abandoned=True),
    conv("conv-7", "2026-11-04T06:00:00Z", "2026-11-04T06:01:00Z", "q-1", "u-1", "w-3", "comm-7"),  # outside window (should be filtered)
]
TRANSCRIPTS = {
    "comm-1": ["Hi I've been locked out of online banking and can't reset my password", "My account number is 12345678 and my email is bob@example.com", "Thanks that worked"],
    "comm-2": ["There's a transaction on my card I don't recognise, I think it's fraud and I want to dispute the charge", "It was for two hundred dollars at a shop I've never been to", "Okay thank you"],
    "comm-3": ["I lost my job last month and I'm struggling to make my loan repayments, is there any hardship assistance", "I just need some breathing room on the repayments for a few months"],
    "comm-5": ["I'm going overseas next week on holiday, do I need to tell you so my card keeps working abroad", "Great, and is there a fee for using the card overseas"],
}
SUGGESTIONS = {
    "conv-1": [
        {"id": "s-1", "type": "KnowledgeSearch", "state": "Accepted", "triggerType": "Intent", "dateCreated": "2026-10-02T01:01:00Z",
         "knowledgeSearch": {"title": DOCS[0]["title"], "confidence": 0.87, "searchId": "srch-1", "document": {"id": "doc-1"}},
         "context": {"participantType": "Customer", "utterance": {"id": "utt-1"}, "queryStatement": "locked out of online banking reset password", "queue": {"id": "q-1"}, "mediaType": "Voice"}},
        {"id": "s-2", "type": "Script", "state": "Suggested", "triggerType": "ConversationStart", "dateCreated": "2026-10-02T01:00:05Z", "script": {"script": {"id": "scr-1"}}},
    ],
    "conv-2": [],
    "conv-3": [],
    "conv-5": [],
}
state = {"rate_limited": False, "server_error": False}

def score(query, doc):
    q = set(re.findall(r"[a-z']+", query.lower()))
    hits = len(q & set(doc["words"]))
    return min(0.95, round(0.18 * hits + random.uniform(0, 0.04), 3)) if hits else round(random.uniform(0.02, 0.12), 3)

class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def send(self, code, body=None, headers=None):
        data = json.dumps(body).encode() if body is not None else b""
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        for k, v in (headers or {}).items(): self.send_header(k, v)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers(); self.wfile.write(data)
    def body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}") if n else {}
    def auth_ok(self):
        return self.headers.get("Authorization") == "Bearer mock-token"

    def do_POST(self):
        p = self.path.split("?")[0]
        if p == "/oauth/token":
            if not self.headers.get("Authorization", "").startswith("Basic "): return self.send(401, {"error": "no basic auth"})
            return self.send(200, {"access_token": "mock-token", "token_type": "bearer", "expires_in": 86399})
        if not self.auth_ok(): return self.send(401, {"message": "unauthorized"})
        b = self.body()
        if p == "/api/v2/analytics/conversations/details/jobs":
            assert b["segmentFilters"][0]["predicates"][0]["value"] == "voice"
            return self.send(202, {"jobId": "job-1"})
        if p == "/api/v2/analytics/knowledge/aggregates/query":
            gb = b.get("groupBy", [])
            assert b["filter"]["predicates"][0]["dimension"] == "knowledgeBaseId"
            def row(group, metrics): return {"group": group, "data": [{"interval": b["interval"], "metrics": [{"metric": m, "stats": {"count": c}} for m, c in metrics.items()]}]}
            if gb == ["knowledgeBaseDocumentId"]:
                return self.send(200, {"results": [row({"knowledgeBaseDocumentId": "doc-1"}, {"nKnowledgeDocumentSurfaced": 12, "nKnowledgeDocumentViewed": 7, "nKnowledgeDocumentCopied": 2}),
                                                   row({"knowledgeBaseDocumentId": "doc-3"}, {"nKnowledgeDocumentSurfaced": 2, "nKnowledgeDocumentViewed": 0})]})
            if gb == ["conversationId"]:
                return self.send(200, {"results": [row({"conversationId": "conv-1"}, {"nKnowledgeSearch": 2, "nKnowledgeSearchAnswered": 1, "nKnowledgeSearchUnanswered": 1, "nKnowledgeDocumentSurfaced": 1, "nKnowledgeDocumentViewed": 1}),
                                                   row({"conversationId": "conv-4"}, {"nKnowledgeSearch": 1, "nKnowledgeSearchAnswered": 1, "nKnowledgeDocumentSurfaced": 1})]})
            if gb == ["conversationId", "knowledgeBaseDocumentId"]:
                return self.send(200, {"results": [row({"conversationId": "conv-1", "knowledgeBaseDocumentId": "doc-1"}, {"nKnowledgeDocumentSurfaced": 1, "nKnowledgeDocumentViewed": 1}),
                                                   row({"conversationId": "conv-4", "knowledgeBaseDocumentId": "doc-3"}, {"nKnowledgeDocumentSurfaced": 1})]})
            if gb == ["knowledgeBaseDocumentId", "feedbackRating"]:
                return self.send(200, {"results": [row({"knowledgeBaseDocumentId": "doc-1", "feedbackRating": "Positive"}, {"nKnowledgeDocumentFeedback": 3}),
                                                   row({"knowledgeBaseDocumentId": "doc-3", "feedbackRating": "Negative"}, {"nKnowledgeDocumentFeedback": 2})]})
            return self.send(400, {"message": "unsupported groupBy in mock: %s" % gb})
        m = re.match(r"/api/v2/knowledge/knowledgebases/([^/]+)/documents/search$", p)
        if m:
            if not state["rate_limited"]:
                state["rate_limited"] = True
                return self.send(429, {"message": "Rate limit exceeded"}, {"Retry-After": "1"})
            if not state["server_error"]:
                state["server_error"] = True
                return self.send(502, {"message": "bad gateway"})
            assert b["queryType"] == "AutoSearch" and b["sortBy"] == "ConfidenceScore"
            ranked = sorted([(score(b["query"], d), d) for d in DOCS if d["state"] == "Published"], key=lambda x: -x[0])[: int(b.get("pageSize", 3))]
            return self.send(200, {"query": b["query"], "searchId": "replay-%d" % random.randint(1000, 9999), "pageSize": b.get("pageSize"), "total": len(ranked),
                                   "results": [{"confidence": s, "document": {"id": d["id"], "title": d["title"], "state": d["state"]}} for s, d in ranked]})
        return self.send(404, {"message": "no mock route for POST " + p})

    def do_GET(self):
        p, _, qs = self.path.partition("?")
        q = dict(kv.split("=", 1) for kv in qs.split("&") if "=" in kv)
        if p.startswith("/s3/"):
            comm = p.split("/")[-1]
            phrases = []
            for i, t in enumerate(TRANSCRIPTS.get(comm, [])):
                phrases.append({"text": t, "startTimeMs": 1000 * (i + 1), "participantPurpose": "external", "type": "transcription"})
                phrases.append({"text": "Sure, let me check that for you right now", "startTimeMs": 1000 * (i + 1) + 500, "participantPurpose": "internal"})
            return self.send(200, {"conversationId": "x", "communicationId": comm, "mediaType": "voice", "transcripts": [{"transcriptId": "t-" + comm, "language": "en-AU", "phrases": phrases}], "participants": []})
        if not self.auth_ok(): return self.send(401, {"message": "unauthorized"})
        if p == "/api/v2/analytics/conversations/details/jobs/availability": return self.send(200, {"dataAvailabilityDate": "2026-11-05T00:00:00Z"})
        if p == "/api/v2/analytics/conversations/details/jobs/job-1": return self.send(200, {"state": "FULFILLED"})
        if p == "/api/v2/analytics/conversations/details/jobs/job-1/results":
            if q.get("cursor") == "c2": return self.send(200, {"conversations": CONVS[4:], "dataAvailabilityDate": "2026-11-05T00:00:00Z"})
            return self.send(200, {"conversations": CONVS[:4], "cursor": "c2"})
        if p == "/api/v2/knowledge/knowledgebases/" + KB: return self.send(200, {"id": KB, "name": "Mock KB", "coreLanguage": "en-AU"})
        if p == "/api/v2/knowledge/knowledgebases/%s/documents" % KB:
            if q.get("after") == "page2": return self.send(200, {"entities": DOCS[3:]})
            assert q.get("includeDrafts") == "true" and "category" in q.get("expand", "")
            return self.send(200, {"entities": DOCS[:3], "nextUri": "/api/v2/knowledge/knowledgebases/%s/documents?pageSize=200&after=page2" % KB})
        if p == "/api/v2/knowledge/knowledgebases/%s/documents/feedback" % KB:
            return self.send(200, {"entities": [
                {"id": "fb-1", "rating": "Positive", "reason": "SearchResults", "dateCreated": "2026-10-02T01:05:00Z", "queryType": "AutoSearch", "surfacingMethod": "Snippet",
                 "document": {"id": "doc-1", "versionId": "v1"}, "application": {"type": "Assistant"}, "conversationContext": {"conversation": {"id": "conv-1"}}, "user": {"id": "u-1"}},
                {"id": "fb-2", "rating": "Negative", "reason": "DocumentContent", "dateCreated": "2026-10-03T04:02:00Z", "queryType": "ManualSearch", "surfacingMethod": "Article",
                 "document": {"id": "doc-3", "versionId": "v1"}, "application": {"type": "Assistant"}, "conversationContext": {"conversation": {"id": "conv-4"}}, "user": {"id": "u-2"}, "comment": "Out of date"},
                {"id": "fb-0", "rating": "Positive", "dateCreated": "2026-09-01T00:00:00Z", "document": {"id": "doc-1"}, "application": {"type": "Assistant"}},
            ]})
        if p == "/api/v2/knowledge/knowledgebases/%s/unanswered/groups" % KB:
            assert q.get("app") == "Assistant"
            return self.send(200, {"entities": [{"id": "ug-1", "label": "financial hardship", "statistics": {"unlinkedPhraseCount": 4, "unlinkedPhraseHitCount": 11, "totalPhraseHitCount": 11},
                                                  "phraseGroups": [{"id": "pg-1", "label": "hardship assistance", "unlinkedPhraseHitCount": 11, "unlinkedPhraseCount": 4,
                                                                    "phrases": [{"id": "ph-1", "text": "i need hardship assistance with my loan", "unlinkedPhraseHitCount": 7}]}],
                                                  "suggestedDocuments": []}]})
        m = re.match(r"/api/v2/conversations/([^/]+)/suggestions$", p)
        if m:
            cid = m.group(1)
            if os.environ.get("MOCK_SUGGESTIONS_403"): return self.send(403, {"message": "missing permission conversation:suggestion:view"})
            if cid not in SUGGESTIONS: return self.send(404, {"message": "not found"})
            return self.send(200, {"entities": SUGGESTIONS[cid]})
        m = re.match(r"/api/v2/speechandtextanalytics/conversations/([^/]+)/communications/([^/]+)/transcripturl$", p)
        if m:
            comm = m.group(2)
            if comm not in TRANSCRIPTS: return self.send(404, {"message": "no transcript"})
            return self.send(200, {"url": "http://127.0.0.1:%d/s3/%s" % (PORT, comm)})
        if p == "/api/v2/routing/queues": return self.send(200, {"entities": [{"id": k, "name": v} for k, v in QUEUES.items() if k in q.get("id", "").split("%2C") or k in q.get("id", "").split(",")]})
        if p == "/api/v2/routing/wrapupcodes": return self.send(200, {"entities": [{"id": k, "name": v} for k, v in WRAPS.items()], "pageCount": 1})
        if p == "/api/v2/users": return self.send(200, {"entities": [{"id": k, "name": v} for k, v in USERS.items()]})
        return self.send(404, {"message": "no mock route for GET " + p})

if __name__ == "__main__":
    random.seed(7)
    print("mock Genesys API on port", PORT, flush=True)
    HTTPServer(("127.0.0.1", PORT), H).serve_forever()
