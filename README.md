# IKEA Outlook Attachment Archiver

A small Microsoft Outlook VBA macro I wrote during my time at IKEA Logistics Services (Shanghai) in 2017 to automate one specific manual process the quality team was doing by hand every day.

This is the **sanitized 2025 archive copy** of the original deployed code, published as a portfolio piece. The macro stayed in production after I left.

---

## The problem

The Quality Coordination team received hundreds of carton-photo emails per day from suppliers. Each email's subject contained a shipment identifier in a fixed format (`Info DEVRPT <ID> IKEA <some text>`), and each email had attached photos that needed to be filed into a per-shipment folder on a shared drive.

Before the macro: someone opened each email, identified the shipment ID, made or navigated to the right folder, and saved each attachment into it manually. With the email volume the team was carrying, this was a real chunk of every working day.

## The macro

`SaveAtt.bas` does the following:

1. Iterates over the user-selected emails in the active Outlook window (not the entire Inbox — the user picks what to archive).
2. For each email, extracts the shipment ID from the subject line by string-slicing between the two fixed markers `Info DEVRPT` and `IKEA`.
3. Checks whether a per-shipment folder already exists under the configured archive root. If not, creates it.
4. Saves every attachment from that email into the matching shipment folder.
5. Skips emails that have no attachments (no empty folders, no errors).

## Design choices worth flagging

A few decisions I made deliberately at the time, which I'd defend the same way today:

- **Operate on `Selection`, not `Inbox.Items`.** The user keeps manual control over which emails get archived — the macro is a tool, not an autonomous bot. This also avoids the worst-case "macro processed 800 emails because someone misclicked" failure mode.
- **Idempotent folder creation.** `Dir(path, vbDirectory) = ""` followed by conditional `MkDir` means re-running on the same shipment is a no-op, not an error. Safer than wrapping a naked `MkDir` in `On Error Resume Next`.
- **Skip empty-attachment emails.** Avoids creating a per-shipment folder for an email that has nothing to put in it.
- **No exception handling for the parsing step.** Deliberate — if a subject line doesn't contain both markers, I want the macro to fail loudly so the operator notices the schema has changed, rather than quietly mis-archive.

## Limitations (things I'd improve if I rewrote it now)

- **Filename collisions.** If two emails for the same shipment have identically-named attachments (e.g. `IMG_001.jpg`), the second overwrites the first silently. A timestamp prefix or a counter-suffix would fix this.
- **Hardcoded subject markers.** `Info DEVRPT` and `IKEA` are baked in. A modern rewrite would put these in a config sheet so the macro adapts when the email format changes.
- **No audit log.** The macro shows a "Done!" message at the end but doesn't write a per-email log of what was saved where. With the volume the team was running, an audit log would have been worth the extra ten lines.
- **Single-folder root.** No support for archiving to multiple destinations (e.g. by supplier, by date) without code changes.

## Running it

1. Open Outlook → Developer tab → Visual Basic Editor (`Alt+F11`).
2. `File → Import File…` → select `SaveAtt.bas`.
3. Edit the `folder` line in the macro to point to your own archive root (must end with a backslash).
4. In Outlook, select the emails you want to archive.
5. Run the macro (`Alt+F8` → `SaveAtt` → Run).

## Notes on this archive copy

- The original archive root path has been replaced with a generic placeholder (`\\<server>\<share>\photo catalog\`).
- I added inline comments in 2025 for the portfolio version. The original 2017 code was sparser.
- This is a faithful logical reconstruction of what was deployed, not a fork of an actively-maintained tool.
