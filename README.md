# Perfect Align — Word Online Add-in

A Microsoft Word task pane add-in (Office.js) that re-formats a chaotic, mixed,
5–6 layer nested list — numbers, letters, and bullets in any combination — so
that **every layer's bullet/number string aligns exactly with the text-start
position of the layer directly above it**.

It does this dynamically (no static OOXML list templates): it walks the selected
paragraphs top-to-bottom, tracks the running text indent per level, and applies
a clean hanging indent to each paragraph using the native
`leftIndent` / `firstLineIndent` properties.

## Add this add-in to Word Online

The add-in is already hosted (at `https://word-perfect-align.kawinr.com`), so
you only need the manifest file to sideload it — **no install or local server
required**.

1. **Download the manifest.** Open
   [`manifest.xml`](manifest.xml) in this repo → click **Raw** → save the file
   (right-click → *Save As*) as `manifest.xml`.
2. **Open a document** in [Word on the web](https://www.office.com/launch/word)
   (a SharePoint / OneDrive / Microsoft 365 document).
3. On the **Home** tab, click **Add-ins** → **More Add-ins**
   (or **Insert** → **Add-ins**).
4. In the dialog, go to the **My Add-ins** tab → **Upload My Add-in**.
5. Choose the `manifest.xml` you downloaded → **Upload**.
6. A **Perfect Align** group appears on the **Home** tab. Click **Align Lists**
   to open the task pane.

**To use it:** select your nested list (and optionally set a starting indent —
see *Usage* below), then click **Perfectly Align Selection**.

> To remove or re-upload later: **Insert ▸ Add-ins ▸ My Add-ins**, then manage
> it from there. Re-uploading the same `manifest.xml` replaces the old version.
>
> For org-wide rollout, deploy the manifest through the **Microsoft 365 admin
> center ▸ Integrated apps** or your **SharePoint app catalog** (details under
> *Sideload* below).

## How the alignment works

For each selected paragraph (1 inch = 72 points):

| Quantity | Rule |
|---|---|
| **Alignment** (number/bullet position) | `= textIndent` of the level directly above (`level − 1`). Level 0 → the starting indent (default `0`). |
| **Text buffer** (bullet) | `0.11 in` (~8 pt), fixed. |
| **Text buffer** (number/letter) | `(0.04 + 0.07 × charLength) in` — the per-char term tracks the glyph width so the visible gap stays a constant ~0.04 in even for long numbers like `1.1.1.1.1.`. |
| **Text indent** (text position) | `= alignment + buffer`. |
| Applied as | `leftIndent = textIndent`, `firstLineIndent = −(textIndent − alignment)`. |

The outermost layer can be shifted to a **starting indent** (slider, or *Copy
indent from selection* to match a heading); the whole list shifts with it while
keeping every relative alignment. Bullets that Word reports at a shallow level
are re-parented one layer below the item above them.

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

## Tests

The alignment algorithm lives in a pure, side-effect-free function
(`computeLayout` in `src/taskpane.js`) so it can be unit-tested without Office.
Run the regression suite (Node's built-in runner — no dependencies):

```bash
npm test
```

Tests live in `test/taskpane.test.js` and cover bullet/number classification,
the constant-gap buffer, perfect vertical alignment, orphan-bullet
re-parenting, the hanging-indent math, and non-list edge cases. Run them before
committing any change to the algorithm.

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

## License

[MIT](LICENSE) © 2026 Kawin Rattanapun
