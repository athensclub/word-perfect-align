/*
 * taskpane.js
 * -----------
 * Perfect Align — dynamically formats a chaotic, mixed, 5–6 layer nested list
 * (numbers, letters, bullets in any combination) so that the bullet/number
 * string of every layer aligns exactly with the TEXT-START position of the
 * layer directly above it.
 *
 * Strategy (no static list templates):
 *   Walk the selected paragraphs top-to-bottom. Track the running text-indent
 *   assigned to each list level. For each paragraph, compute a hanging indent
 *   where:
 *     - the number/bullet position == the parent level's text indent
 *     - the text position         == that alignment + a dynamic buffer
 *   Apply via the native leftIndent / firstLineIndent paragraph properties.
 *
 * All indent math is done in POINTS (1 inch = 72 points), the unit Office.js
 * uses for leftIndent / firstLineIndent.
 */

(function () {
  "use strict";

  // ---------------------------------------------------------------------------
  // Tunable constants (in inches; converted to points where applied).
  // ---------------------------------------------------------------------------
  var POINTS_PER_INCH = 72;

  // Bullets ("•", "-", "▪", ...) get a small, fixed gap before their text.
  var BULLET_BUFFER_INCHES = 0.13; // ~9 pt — bullet glyph + a hair of breathing room

  // Numbered/lettered items ("1.", "1.1.1.", "a.") get a length-aware gap.
  // NUMBER_PER_CHAR ≈ the rendered width of one number glyph, so the buffer
  // grows just enough to clear the number string — which keeps the VISIBLE
  // gap between the number and its text constant (== NUMBER_BASE) no matter how
  // long the number is. NUMBER_BASE is therefore the actual gap after the number.
  var NUMBER_BASE_INCHES = 0.06; // ~4 pt — tight gap after the number
  var NUMBER_PER_CHAR_INCHES = 0.07; // ~5 pt/char ≈ glyph width at ~11pt font

  // ---------------------------------------------------------------------------
  // Office bootstrap
  // ---------------------------------------------------------------------------
  // Guarded so the module can be `require()`d in a plain Node test runner
  // (where Office.js is absent) without throwing — tests target the pure logic.
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) {
      // Only wire up the UI inside Word.
      if (info.host === Office.HostType.Word) {
        var button = document.getElementById("align-button");
        button.disabled = false;
        button.addEventListener("click", alignSelection);
      }
    });
  }

  // ---------------------------------------------------------------------------
  // Small pure helpers (modular + unit-testable)
  // ---------------------------------------------------------------------------

  /** Convert an inch value to points. */
  function inchesToPoints(inches) {
    return inches * POINTS_PER_INCH;
  }

  /**
   * Decide whether a list string represents a BULLET (symbol) vs a
   * NUMBER/LETTER sequence.
   *
   * Word number/letter markers either contain a digit ("1.", "1.1.1.", "10)")
   * or are a letter/roman run ending in a delimiter ("a.", "iv.", "B)"). Bare
   * symbols with no delimiter ("•", "-", "▪", "o", "*", "◦") are bullets — note
   * "o" is the hollow Word bullet, distinct from the lettered marker "o.".
   * Empty / missing strings fall back to "bullet" (the conservative, small gap).
   */
  function isBulletString(listString) {
    var s = (listString || "").trim();
    if (s.length === 0) {
      return true; // unknown -> treat as bullet
    }
    if (/[0-9]/.test(s)) {
      return false; // any digit -> numbered
    }
    if (/[.)\]]$/.test(s)) {
      return false; // letter/roman marker ending in a delimiter -> lettered
    }
    return true; // bare symbol or single letter without delimiter -> bullet
  }

  /**
   * Compute the dynamic text buffer (in points) that separates the
   * number/bullet position from the text start.
   *   - Bullet: fixed small buffer.
   *   - Number/letter: base + per-character, scaled to the string length.
   */
  function computeBufferPoints(listString) {
    if (isBulletString(listString)) {
      return inchesToPoints(BULLET_BUFFER_INCHES);
    }
    var charLen = (listString || "").trim().length;
    var inches = NUMBER_BASE_INCHES + NUMBER_PER_CHAR_INCHES * charLen;
    return inchesToPoints(inches);
  }

  /**
   * Apply a clean hanging indent to a paragraph.
   *   leftIndent       = textIndent           (where the text sits)
   *   firstLineIndent  = -(textIndent - align) (pulls the number/bullet back to `align`)
   * Net effect: the number/bullet renders at `alignment`, text at `textIndent`.
   */
  function applyHangingIndent(paragraph, alignmentPoints, textIndentPoints) {
    paragraph.leftIndent = textIndentPoints;
    paragraph.firstLineIndent = -(textIndentPoints - alignmentPoints);
  }

  // ---------------------------------------------------------------------------
  // Core algorithm (PURE — no Office.js dependency, fully unit-testable)
  // ---------------------------------------------------------------------------

  /**
   * Compute the indent layout for a list of paragraphs.
   *
   * @param {Array<{isList:boolean, level:number, listString:string}>} paras
   *        One entry per paragraph in document order. `isList:false` marks a
   *        non-list paragraph (left untouched). `level` is Word's 0-indexed
   *        list level; `listString` is the rendered marker ("1.", "a.", "•").
   * @returns {{
   *   results: Array<null|{level:number, isBullet:boolean, alignment:number,
   *                        textIndent:number, leftIndent:number,
   *                        firstLineIndent:number}>,
   *   aligned:number, skipped:number, maxLevel:number
   * }}
   *   `results[i]` is null for skipped (non-list) paragraphs, otherwise the
   *   computed indents in points for paragraph i.
   */
  function computeLayout(paras) {
    var textIndentByLevel = {}; // effective level -> text indent (points)
    var results = [];
    var aligned = 0;
    var skipped = 0;
    var maxLevel = 0;

    // Track the previous list item so we can re-parent "orphan" bullets.
    // In chaotic docs, Word often reports stray bullets at level 0 even
    // though they visually belong under the (deeper) item above them.
    var prevLevel = -1; // effective level of the previous list item
    var prevWasBullet = false;

    for (var i = 0; i < paras.length; i++) {
      var p = paras[i];

      // Edge case: paragraph isn't part of a list -> leave it untouched.
      if (!p || !p.isList) {
        results.push(null);
        skipped++;
        continue;
      }

      var bullet = isBulletString(p.listString);
      var wordLevel = p.level || 0; // 0-indexed depth reported by Word

      // Effective level:
      //  - Numbers/letters: trust Word's level (it's reliable for them).
      //  - Bullets: place ONE layer below the item directly above.
      //      * first bullet after a number  -> prevLevel + 1
      //      * bullet after a bullet         -> same level (siblings)
      //    Fall back to Word's level only when there's no item above.
      var level;
      if (!bullet) {
        level = wordLevel;
      } else if (prevLevel < 0) {
        level = wordLevel; // first item in selection, nothing to nest under
      } else if (prevWasBullet) {
        level = prevLevel; // consecutive bullets are siblings
      } else {
        level = prevLevel + 1; // first bullet under a numbered/lettered item
      }

      if (level > maxLevel) maxLevel = level;
      prevLevel = level;
      prevWasBullet = bullet;

      // Alignment = where this layer's number/bullet sits.
      // It must equal the text indent of the layer directly above it.
      // (Level 0 starts flush at the left margin.)
      var alignment =
        level === 0
          ? 0
          : typeof textIndentByLevel[level - 1] === "number"
          ? textIndentByLevel[level - 1]
          : 0; // no parent seen yet -> flush

      // Dynamic buffer based on bullet vs number, and number length.
      var buffer = computeBufferPoints(p.listString);

      // Text indent = where this layer's text sits.
      var textIndent = alignment + buffer;
      textIndentByLevel[level] = textIndent;

      // Any deeper levels recorded earlier are now stale (we've moved up
      // or to a new branch); clear them so a later child recomputes
      // against the current ancestry rather than an old sibling subtree.
      for (var deeper in textIndentByLevel) {
        if (Number(deeper) > level) {
          delete textIndentByLevel[deeper];
        }
      }

      results.push({
        level: level,
        isBullet: bullet,
        alignment: alignment,
        textIndent: textIndent,
        leftIndent: textIndent,
        firstLineIndent: -(textIndent - alignment),
      });
      aligned++;
    }

    return { results: results, aligned: aligned, skipped: skipped, maxLevel: maxLevel };
  }

  // ---------------------------------------------------------------------------
  // Status helpers
  // ---------------------------------------------------------------------------
  function setStatus(message, kind) {
    var el = document.getElementById("status");
    el.textContent = message;
    el.className = kind || "";
  }

  // ---------------------------------------------------------------------------
  // Main action
  // ---------------------------------------------------------------------------
  function alignSelection() {
    var button = document.getElementById("align-button");
    button.disabled = true;
    setStatus("Aligning…", "");

    Word.run(function (context) {
      // 1) Get the paragraphs in the current selection.
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items");

      return context.sync().then(function () {
        var items = paragraphs.items;

        if (!items || items.length === 0) {
          setStatus("No text is selected. Highlight your list first.", "warn");
          return context.sync(); // nothing to do
        }

        // 2) For every paragraph, load its list membership + level + string.
        //    listItemOrNullObject lets us skip non-list paragraphs gracefully.
        var listItems = items.map(function (p) {
          var li = p.listItemOrNullObject;
          li.load("level,listString");
          return li;
        });

        return context.sync().then(function () {
          // 3) Build a plain-data snapshot, then run the pure layout algorithm.
          var paras = listItems.map(function (li) {
            return li.isNullObject
              ? { isList: false }
              : { isList: true, level: li.level || 0, listString: li.listString };
          });

          var layout = computeLayout(paras);

          // 4) Apply the computed hanging indents back onto the paragraphs.
          for (var i = 0; i < layout.results.length; i++) {
            var r = layout.results[i];
            if (r) {
              applyHangingIndent(items[i], r.alignment, r.textIndent);
            }
          }

          // 5) Single sync to push all indentation changes at once.
          return context.sync().then(function () {
            var msg =
              "Aligned " +
              layout.aligned +
              " list paragraph" +
              (layout.aligned === 1 ? "" : "s") +
              " across " +
              (layout.maxLevel + 1) +
              " level" +
              (layout.maxLevel === 0 ? "" : "s") +
              ".";
            if (layout.skipped > 0) {
              msg +=
                "\nSkipped " +
                layout.skipped +
                " non-list paragraph" +
                (layout.skipped === 1 ? "" : "s") +
                ".";
            }
            setStatus(msg, layout.aligned > 0 ? "ok" : "warn");
          });
        });
      });
    }).catch(function (error) {
      // Surface Office/OfficeExtension errors to the user.
      var detail = error && error.message ? error.message : String(error);
      if (error && error.debugInfo) {
        detail += "\n(" + JSON.stringify(error.debugInfo) + ")";
      }
      setStatus("Something went wrong: " + detail, "err");
      // Also log for the dev console.
      console.error("Perfect Align error:", error);
    }).then(function () {
      button.disabled = false;
    });
  }

  // Expose helpers + the pure algorithm for unit testing, without breaking the
  // IIFE scope or the browser bootstrap above.
  if (typeof module !== "undefined" && module.exports) {
    module.exports = {
      inchesToPoints: inchesToPoints,
      isBulletString: isBulletString,
      computeBufferPoints: computeBufferPoints,
      computeLayout: computeLayout,
      // Constants exposed so tests assert against the source of truth, not
      // hard-coded magic numbers that would silently drift.
      constants: {
        POINTS_PER_INCH: POINTS_PER_INCH,
        BULLET_BUFFER_INCHES: BULLET_BUFFER_INCHES,
        NUMBER_BASE_INCHES: NUMBER_BASE_INCHES,
        NUMBER_PER_CHAR_INCHES: NUMBER_PER_CHAR_INCHES,
      },
    };
  }
})();
