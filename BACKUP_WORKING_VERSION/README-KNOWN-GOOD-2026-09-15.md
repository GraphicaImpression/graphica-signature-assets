# Signature system — known-good snapshot, 2026-09-15

This backup captures the entire Graphica signature system in a fully working, confirmed state, as of the date above. If a future change breaks something, this is the version to restore.

## What was confirmed working at this point

- The core bug (native Outlook signature wiping out the Graphica block with no fallback) is fixed and verified end-to-end on Émiliana's own account: everyday signature inserted, switched to a native personal signature, sent, and the Graphica block appeared underneath in the actual delivered email.
- All four personal Exchange transport rules (Émiliana, Luc Leblanc, J.C. Tirman, Hélène Potvin) are deployed with the corrected, marker-only exception logic — confirmed via `Get-TransportRule` showing `ExceptIfHeaderContainsMessageHeader`/`ExceptIfHeaderContainsWords` blank and only `ExceptIfSubjectOrBodyContainsWords: {GraphicaSigMarkerV1}` present, with a fresh `WhenChanged` timestamp on each.
- The Outlook add-in's manual "Aperçu signature" button and variant-switching work correctly.
- The send-time Smart Alerts safety net (`OnMessageSend`) is deployed as a second-layer backstop (status not independently re-confirmed after the marker fix, but architecturally sound and fails open by design).
- The add-in is installed via Exchange Online PowerShell (`New-App -OrganizationApp`), scoped to exactly four people: Émiliana, Luc, Hélène, J.C. (`ProvidedTo: SpecificUsers`).

## What is included but NOT working (by design, left in place)

- `OnNewMessageCompose` (automatic signature preview in a new email's body, without clicking the button first) is implemented in `launchevent.js` but confirmed, via an isolated smoke test on 2026-09-15, to never fire on Émiliana's Outlook client. This is a known, currently unresolved bug in Microsoft's own Outlook add-in platform (documented on Microsoft's own developer Q&A/forums — the event simply doesn't reach some users/clients, independent of manifest or code correctness). It costs nothing to leave in place: it will start working automatically if Microsoft ever fixes this, and doing nothing is exactly its current behavior.

## Files in this backup

- `Deploy-Signature.ps1` — the PowerShell deploy script. Key fact: it sets ONLY `ExceptIfSubjectOrBodyContainsWords = @($MarkerText)` as the transport rule's exception (section "3c"), and explicitly clears `ExceptIfHeaderContainsMessageHeader`/`ExceptIfHeaderContainsWords` with `$null` when updating an existing rule (section "3d") — both are load-bearing fixes; see the script's own header comments ("IMPORTANT HISTORY" and the section 3d comment) for the full story of why, before ever re-adding a header-based exception.
- `manifest.xml` — the Outlook add-in manifest. AppVersion `1.1.0.0`. Contains the `OnMessageSend` LaunchEvent (send-time safety net) and the `OnNewMessageCompose` LaunchEvent (currently inert — see above).
- `launchevent.js` — background event handlers for both LaunchEvents above. Every async path is wrapped so it can never fail to call `event.completed()` silently (a `safeComplete()` helper) — this was itself a real bug found and fixed this session.
- `taskpane.js`, `taskpane.html`, `commands.html` — the task pane itself (the manual "Aperçu signature" experience, auto-insert on pane open, conversation memory, auto-refresh). Unchanged by this session's fixes.
- `signature-console-KNOWN-GOOD-2026-09-15.html` — full exported HTML of the live Signature Console (the Claude Artifact at `https://claude.ai/code/artifact/8582d04b-a1c1-4762-812a-e2c83451fb0b`). Confirmed to contain the hidden `GraphicaSigMarkerV1` marker at the end of every generated signature.

## How this is deployed (for reference, not something to redo unless rolling back)

- **Outlook add-in files** (`manifest.xml`, `launchevent.js`, `taskpane.js`, `taskpane.html`, `commands.html`): live on Cloudflare Pages, project `graphica-signatures`, deployed from `~/Documents/outlook-addin` with:
  ```
  npx wrangler pages deploy . --project-name=graphica-signatures
  ```
- **The add-in's registration in Exchange** (separate from the files above — this is what tells Outlook the add-in exists at all): installed via PowerShell, NOT the Microsoft 365 admin center's normal upload page. That page has had a known, unresolved server-side bug (500 error on `uploadCustomApp`) since mid-2025, affecting many organizations — a support ticket was opened with Microsoft with no fix as of this date. The workaround, confirmed working, is:
  ```powershell
  # To update the manifest after editing it:
  Remove-App -Identity 4c2f6a8e-9c1a-4e1b-8f2d-6a1e2b3c4d5e -OrganizationApp -Confirm:$false
  $bytes = [System.IO.File]::ReadAllBytes("$HOME/Documents/outlook-addin/manifest.xml")
  New-App -FileData $bytes -OrganizationApp -ProvidedTo SpecificUsers -UserList "emiliana.braga@graphica.qc.ca","luc.leblanc@graphica.qc.ca","helene.potvin@graphica.qc.ca","jc.tirman@graphica.qc.ca" -DefaultStateForUser Enabled
  ```
  Note: only a **manifest** change requires this Remove-App/New-App pair. A JS-only change (like `launchevent.js` alone) just needs the `wrangler pages deploy` step above — Outlook fetches those files live from Cloudflare each time, they aren't baked into the Exchange registration.
- **Transport rules**: deployed per-person via `Deploy-Signature.ps1 -ConfigPath <their config file> -RuleName "Signature - <Name>" -TestUser <their email> -Priority <N> -StopProcessing`, using each person's exported HTML from the Console.

## How to roll back to this exact state

1. Copy `manifest.xml`, `launchevent.js`, `taskpane.js`, `taskpane.html`, `commands.html` from this backup back into `~/Documents/outlook-addin/`, overwriting whatever is there.
2. Run `npx wrangler pages deploy . --project-name=graphica-signatures` from that folder.
3. If `manifest.xml` changed from what's currently live, also run the `Remove-App` / `New-App` pair above.
4. Copy `Deploy-Signature.ps1` from this backup back into `~/Documents/Signature Studio/`, overwriting whatever is there. This alone does NOT touch Exchange — it only takes effect the next time it's actually run per person.
5. For the Console: this can't be "restored" the same way (it's a hosted Artifact, not a file) — if it's ever changed and needs reverting, the content of `signature-console-KNOWN-GOOD-2026-09-15.html` in this backup is the exact known-good reference to rebuild from.
