# Perfect Align — Word Online Add-in

A Microsoft Word task pane add-in (Office.js) that re-formats a chaotic, mixed,
5–6 layer nested list — numbers, letters, and bullets in any combination — so
that **every layer's bullet/number string aligns exactly with the text-start
position of the layer directly above it**.

It does this dynamically (no static OOXML list templates): it walks the selected
paragraphs top-to-bottom, tracks the running text indent per level, and applies
a clean hanging indent to each paragraph using the native
`leftIndent` / `firstLineIndent` properties.

## How the alignment works

For each selected paragraph (1 inch = 72 points):

| Quantity | Rule |
|---|---|
| **Alignment** (number/bullet position) | `= textIndent` of the level directly above (`level − 1`). Level 0 → `0`. |
| **Text buffer** (bullet) | `0.25 in` (18 pt), fixed. |
| **Text buffer** (number/letter) | `(0.3 + 0.08 × charLength) in` — grows with the list string so `1.1.1.1.1.` never collides with its text. |
| **Text indent** (text position) | `= alignment + buffer`. |
| Applied as | `leftIndent = textIndent`, `firstLineIndent = −(textIndent − alignment)`. |

Paragraphs that aren't part of a list are detected via `listItemOrNullObject`
and left untouched.

## Project layout

```
word-align-addin/
├── manifest.xml      # Task pane add-in manifest + Home-tab ribbon button
├── package.json      # dev scripts (no build step)
├── server.js         # tiny HTTPS static server on :3000
├── README.md
├── assets/           # ribbon/task-pane icons (16/32/80 px)
└── src/
    ├── taskpane.html # Fluent-styled UI: one "Perfectly Align Selection" button
    └── taskpane.js   # Office init + the dynamic indent algorithm
```

## Run it locally

Prereqs: Node.js (LTS).

```bash
cd word-align-addin
npm install
npm start
```

`npm start` installs a locally-trusted HTTPS certificate (via
`office-addin-dev-certs`) and serves the folder at
**https://localhost:3000/**. Visit
`https://localhost:3000/src/taskpane.html` once in your browser and accept the
certificate if prompted.

Validate the manifest any time with:

```bash
npm run validate
```

## Sideload in Word on the web (SharePoint / Microsoft 365)

1. Make sure the dev server is running (`npm start`).
2. Open any document in **Word on the web**.
3. Go to **Home ▸ Add-ins ▸ More Add-ins** (or **Insert ▸ Add-ins**) ▸
   **Upload My Add-in**.
4. Choose this project's **`manifest.xml`** and confirm.
5. A **Perfect Align** group appears on the **Home** tab. Click **Align Lists**
   to open the task pane.

> **Org-wide / production deployment:** host the `src/` and `assets/` files on a
> real HTTPS endpoint, update every `https://localhost:3000` URL in
> `manifest.xml` to that host, then deploy the manifest through the
> **Microsoft 365 admin center ▸ Integrated apps** or your **SharePoint app
> catalog**.

## Usage

1. In the document, highlight all the list paragraphs you want to align.
2. In the task pane, click **Perfectly Align Selection**.
3. The status line reports how many list paragraphs were aligned and how many
   non-list paragraphs were skipped.

## Notes

- Requires the **WordApi 1.3** requirement set (for `leftIndent`,
  `firstLineIndent`, and `listItem`), satisfied by current Word on the web.
- Uses the XML manifest so it sideloads directly via *Upload My Add-in*.
- The icons in `assets/` are simple placeholders — replace them with your own
  16/32/80 px PNGs for production.
