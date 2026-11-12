# ADR-0001: Fully client-side, no backend

- **Status:** Accepted
- **Date:** 2026-06-06

## Context

Mutatis assists legal drafting and review for lawyers in Singapore and England & Wales. Its users handle highly confidential documents under the UK SRA Code of Conduct and the Singapore PDPA / banking-secrecy expectations. Any architecture that transmits document content off the user's machine — even transiently, even to infrastructure the firm controls — is a confidentiality and data-residency problem. The product must also be trivially adoptable: lawyers and PAs cannot stand up servers.

## Decision

The add-in is **fully client-side**. There is no backend, no server, and no cloud processing of document content. Everything runs inside the Word task pane's WebView. Document content **never leaves the user's machine**. The only permitted network calls are (a) loading the add-in's own static assets, and (b) explicit, opt-in storage adapters (GitHub/SharePoint) that carry only user-authored library data — never the working document — and are off by default.

## Rationale

Confidentiality-first is the product's core positioning. A data-residency guarantee *by construction* (no code path sends document bytes anywhere) is far stronger than a policy promise, and it is the only posture compatible with SRA/PDPA obligations without firm-by-firm legal review. It also gives zero-setup adoption: no Docker, no account, no IT involvement.

## Rejected alternatives

- **Self-hosted Docker backend (firm runs the server).** Rejected: adoption friction (most target users cannot deploy containers) and it still raises data-residency objections — IT/security review every install, defeating zero-setup.
- **Project-hosted public backend (we run the server).** Rejected: an outright confidentiality dealbreaker. Sending privileged legal documents to a third-party server is incompatible with SRA/PDPA duties; no amount of TLS or "we don't store it" mitigates the trust problem.
- **Bundled localhost sidecar (native helper process).** Rejected: install friction (a separate native binary per platform, code-signing, updates) negates the "just sideload the add-in" promise, for capability the OOXML loop already covers client-side.

## Consequences

- Operations the typed Office JS API cannot express must use the client-side OOXML read-modify-write loop (ADR-0003) instead of a server.
- A few capabilities are desktop-only or deferred where no client-side path exists; these degrade gracefully rather than reaching for a server.
- A build-time lint / code-review invariant must flag any `fetch`/`XHR`/`WebSocket` carrying document content outside the sanctioned storage-adapter modules.
