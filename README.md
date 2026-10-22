# Quick Sign for New Outlook (Windows) — Free prototype

This is a working source-code **prototype**, not a Microsoft-verified or tenant-tested production add-in. It uses Outlook's supported JavaScript API to import *direct* PDF attachments from the current email, uses PDF.js to preview the selected page, and uses PDF-lib to place a graphic signature on that page. It then opens an Outlook reply prepopulated with the signed PDF. **It never sends email automatically.**

No Adobe subscription, Microsoft Graph application registration, Python installation, or server-side PDF processor is required. Running the add-in needs Microsoft 365/Outlook with a supported Exchange mailbox, your tenant's permission to use a custom web add-in, and free HTTPS static hosting.

## Easiest setup (GitHub Pages)

1. Create a GitHub repository named `outlook-quick-sign` and make it **public** (GitHub Pages is free for public repositories).
2. Upload **the files inside this folder** to the repository root, including the `assets` subfolder (not just the ZIP archive). Commit the changes. Do not upload private PDFs or your signature.
3. On GitHub, open **Settings → Pages → Build and deployment → Deploy from a branch → main → /(root) → Save**. Your site URL should look like `https://YOURUSERNAME.github.io/outlook-quick-sign/` after deployment.
4. Open that URL in a regular browser to confirm that the app appears. Use a disposable/sample PDF to test; the standalone page will permit PDF loading and signed PDF downloading, but **not Outlook replies**.
5. Open `https://YOURUSERNAME.github.io/outlook-quick-sign/setup.html`, paste your site URL in the field, and download the generated `manifest.xml`.
6. Go to <https://aka.ms/olksideload> in a browser signed in with the **same Microsoft 365 account** used by New Outlook. Select **My add-ins → Custom Addins → Add a custom add-in → Add from File**, choose `manifest.xml`, and accept the prompts. These labels may be slightly different depending on Microsoft's UI.
7. In New Outlook, open an email with an actual PDF attachment (not just a OneDrive link). In the message menu **Apps / More apps → Quick Sign / Sign PDF**, open the add-in. Save your signature once. Move the signature onto the PDF signing area, then click **Sign PDF & open Outlook reply**.
8. Confirm the generated attachment and recipient before sending. Test the whole flow on a non-sensitive test email and small PDF first.

Official Microsoft instructions: <https://learn.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing> and <https://learn.microsoft.com/en-us/javascript/api/outlook/office.messageread>.

## Features

- Draw and save your signature once, or upload a transparent PNG (JPG also supported).
- Automatically reads PDF attachments from the current email; automatically opens one if exactly one PDF is present.
- PDF preview: page navigation, movable signature overlay, size adjustment, arrow key positioning (Shift for a bigger movement).
- Starts with the last PDF page selected and remembers the last position and size as a convenience; **always check placement before signing another form**.
- Click once to sign the PDF and open an Outlook reply draft with `_signed.pdf` attached.
- Local PDF file upload and manual download are available as backup / standalone mode.
- Document processing takes place in browser memory. Signature and last position are stored in browser localStorage on that device. No app-operated backend is used.

## Important limitations and security

- The signature is a **visual/handwritten-style image stamped onto a PDF**, **not** a certificate-based cryptographic PDF signature. Check whether this meets the recipient's policies and the document's legal requirements.
- Requires **internet access**, because Outlook web add-ins need online-hosted pages. This first prototype also loads the PDF libraries from third-party CDNs (`jsdelivr.net` and `cdnjs.com`). For production or confidential company documents, a developer should download and audit/pin/self-host those dependencies and adopt appropriate CSP; you can change the script references in `index.html` once hosted dependencies are ready. Third-party JavaScript can theoretically access in-memory documents, so do not use this prototype for highly confidential documents until that is addressed.
- Only direct PDF file attachments are supported, **not OneDrive/SharePoint cloud links**, password-protected PDFs, or PDF pages with nonzero rotation. Extremely large documents may require more memory or exceed New Outlook's **25 MB** attachment limit.
- **Only the page shown in the preview is signed.** Default is the last page. The last selected position may not match a different document; check every result. It does not detect signature lines automatically.
- The uploaded signature is persisted **unencrypted in browser localStorage**, which is suitable only on a trusted company-managed Windows account. Remove it via **Remove saved** if using a shared PC.
- If company policy blocks **Custom Addins**, Microsoft 365 admin must explicitly permit and/or deploy the add-in. An app cannot bypass that setting.
- The Outlook reply flow uses `displayReplyFormAsync`, available to supported Outlook clients using Mailbox requirement set 1.9; check compatibility with your particular New Outlook build. The add-in checks for Mailbox 1.8 at startup for reading files.
- This app hasn't been executed inside the user's tenant. A local code/markup check isn't a substitute for integration testing.

## Folder contents

- `index.html`: app layout and library references
- `app.js`: Outlook attachment import, signature handling, PDF stamping, reply draft
- `styles.css`: visual design
- `assets/`: 16/32/80-pixel app icons
- `setup.html`: one-click Outlook `manifest.xml` generator (run after hosting)
- `test-document.pdf`: safe sample PDF to try in standalone mode

## Technical notes

- pdf-lib `1.17.1` for placing the signature and exporting a new PDF.
- PDF.js `3.11.174` for client-side preview.
- `Office.context.mailbox.item.attachments` and `getAttachmentContentAsync` for read-mode attachments.
- `Office.context.mailbox.item.displayReplyFormAsync` with a Base64 encoded `ReplyFormAttachment` object for the prepopulated reply.
- Emails are not read from an outside server. The app works only in the context of an email the user opened and invoked it on.
