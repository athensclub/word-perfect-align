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
  var BULLET_BUFFER_INCHES = 0.25; // 18 pt

  // Numbered/lettered items ("1.", "1.1.1.", "a.") get a length-aware gap so a
  // long string like "1.1.1.1.1." never crashes into its own text.
  var NUMBER_BASE_INCHES = 0.3; // 21.6 pt baseline
  var NUMBER_PER_CHAR_INCHES = 0.08; // +5.76 pt per character in the list string

  // ---------------------------------------------------------------------------
  // Office bootstrap
  // ---------------------------------------------------------------------------
  Office.onReady(function (info) {
    // Only wire up the UI inside Word.
    if (info.host === Office.HostType.Word) {
      var button = document.getElementById("align-button");
      button.disabled = false;
      button.addEventListener("click", alignSelection);
    }
  });

  // ---------------------------------------------------------------------------
  // Small pure helpers (modular + unit-testable)
  // ---------------------------------------------------------------------------

  /** Convert an inch value to points. */
  function inchesToPoints(inches) {
    return inches * POINTS_PER_INCH;
  }

  /**
   * Decide whether a list string represents a BULLET (symbol) vs a
   * NUMBER/LETTER sequence. If the trimmed string contains any digit or
   * latin letter we treat it as numbered/lettered; otherwise it's a bullet.
   * Empty / missing strings fall back to "bullet" (the conservative, small gap).
   */
  function isBulletString(listString) {
    var s = (listString || "").trim();
    if (s.length === 0) {
      return true; // unknown -> treat as bullet
    }
    return !/[0-9A-Za-z]/.test(s);
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
          // 3) Walk top-to-bottom, accumulating the text indent per level.
          var textIndentByLevel = {}; // level (number) -> text indent in points
          var aligned = 0;
          var skipped = 0;
          var maxLevel = 0;

          for (var i = 0; i < items.length; i++) {
            var li = listItems[i];

            // Edge case: paragraph isn't part of a list -> leave it untouched.
            if (li.isNullObject) {
              skipped++;
              continue;
            }

            var level = li.level || 0; // 0-indexed depth
            if (level > maxLevel) maxLevel = level;

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
            var buffer = computeBufferPoints(li.listString);

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

            // 4) Apply the hanging indent.
            applyHangingIndent(items[i], alignment, textIndent);
            aligned++;
          }

          // 5) Single sync to push all indentation changes at once.
          return context.sync().then(function () {
            var msg =
              "Aligned " +
              aligned +
              " list paragraph" +
              (aligned === 1 ? "" : "s") +
              " across " +
              (maxLevel + 1) +
              " level" +
              (maxLevel === 0 ? "" : "s") +
              ".";
            if (skipped > 0) {
              msg +=
                "\nSkipped " +
                skipped +
                " non-list paragraph" +
                (skipped === 1 ? "" : "s") +
                ".";
            }
            setStatus(msg, aligned > 0 ? "ok" : "warn");
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

  // Expose helpers for potential unit testing without breaking the IIFE scope.
  if (typeof module !== "undefined" && module.exports) {
    module.exports = {
      inchesToPoints: inchesToPoints,
      isBulletString: isBulletString,
      computeBufferPoints: computeBufferPoints,
    };
  }
})();
