# Signature Graphica — Add-in Not Compatible on Christelle's Personal Mailbox

## Reference document — save this, paste it back when the add-in topic comes up again

---

## 1. The problem

Christelle Beaudoin uses one Outlook for Mac profile with two mailboxes open side by side:

- **prepresse@graphica.qc.ca** (shared mailbox) — the "Signature Graphica" button works fine.
- **christelle.beaudoin@graphica.qc.ca** (her personal mailbox) — the button is greyed out. Hovering it shows: *"Add-ins are not compatible for christelle.beaudoin@graphica.qc.ca."*

## 2. What "Signature Graphica" actually is

It is **not** a Microsoft/ISV marketplace add-in. It's a custom Office.js Outlook add-in Graphica built and hosts itself:

- **Deployed at:** `https://graphica-signatures.pages.dev` (Cloudflare Pages)
- **AppId:** `4c2f6a8e-9c1a-4e1b-8f2d-6a1e2b3c4d5e`
- **What it does:** shows a live preview of the signature (with banners) while composing, and can insert an alternate signature into the message body on demand via `item.body.setSignatureAsync`. The everyday, automatic signature people actually receive still comes from the Exchange transport rule (`Deploy-Signature.ps1`) — the add-in never writes that one into the body, so there's no conflict there.
- The manifest file is `manifest.xml` (attached separately) — this is the real source of truth for what the add-in is and does.

## 3. Diagnosis: why it's missing from her personal mailbox

Ran, in Exchange Online PowerShell:

```powershell
Get-App -Mailbox christelle.beaudoin@graphica.qc.ca | Format-List DisplayName, Enabled, ProviderName, AppId
Get-App -Mailbox prepresse@graphica.qc.ca | Format-List DisplayName, Enabled, ProviderName, AppId
```

Result: "Signature Graphica" is fully **absent** from Christelle's personal mailbox's app list, while it's present and enabled on prepresse's. The "not compatible" message Outlook shows is just its generic error for "this mailbox has no record of this add-in at all" — nothing to do with Mac compatibility specifically.

## 4. The path that does NOT work — Integrated Apps (Microsoft 365 admin center)

This was tried previously and is confirmed **blocked by a real bug** — do not re-attempt or re-suggest this path. It could not be added inside Integrated Apps.

Also tried, and confirmed **not viable technically**:

```powershell
New-App -OrganizationApp -Identity 4c2f6a8e-9c1a-4e1b-8f2d-6a1e2b3c4d5e -DefaultStateForUser Enabled -Mailbox christelle.beaudoin@graphica.qc.ca
```

This fails with:

```
The arguments 'Mailbox' and 'OrganizationApp' are mutually exclusive. Use one or the other.
```

Reason: `-OrganizationApp -Identity` only supports assigning an app **org-wide** (everyone) or through the admin center's own per-user assignment UI (Integrated Apps) — it has no "assign this one existing org app to just one mailbox" mode. Since Integrated Apps itself is bugged for this tenant, that entire route is a dead end.

## 5. The solution that DOES work — sideload via manifest file (`-FileData`)

Skip the org catalog entirely. Upload the add-in's actual manifest.xml file straight to the one mailbox that needs it:

```powershell
$manifestBytes = [System.IO.File]::ReadAllBytes("C:\path\to\manifest.xml")
New-App -Mailbox christelle.beaudoin@graphica.qc.ca -FileData $manifestBytes -DefaultStateForUser Enabled
```

Why this works when the other command didn't: `-Mailbox` is only mutually exclusive with `-OrganizationApp`. It is fully compatible with `-FileData`. This installs the add-in directly onto that one mailbox by manifest file, bypassing both the org catalog reference and the broken Integrated Apps admin center UI completely.

**Verify it worked:**

```powershell
Get-App -Mailbox christelle.beaudoin@graphica.qc.ca | Format-List DisplayName, Enabled, ProviderName, AppId
```

Should now show "Signature Graphica" listed, matching prepresse's entry.

**Then:** have Christelle fully quit and reopen Outlook for Mac. Sideloaded apps don't reliably appear until the client restarts.

## 6. Where the manifest file lives

- On this machine (Claude's workspace): `/home/claude/outlook-addin/manifest.xml`
- Also sent to you directly as a chat attachment on 2026-09-15.
- Save your own local copy somewhere durable (e.g. a Documents or IT-tools folder) so you can point `-FileData` at it any time this needs to be repeated for someone else.

## 7. Who else might hit this

Anyone who, like Christelle, has **a personal mailbox and a shared mailbox open together in the same Outlook profile** is a candidate for the same "add-in missing from the personal mailbox" gap, even though the shared mailbox works. Worth checking proactively as more of the ~50 planned employees are onboarded, especially anyone who also has access to a shared/team mailbox (prepresse, or similar).

---
*Generated 2026-09-15 for Graphica Impression Inc.'s email signature project.*
