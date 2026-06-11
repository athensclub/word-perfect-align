// @ts-check
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

  // The single, uniform gap left after EVERY marker (bullet or number) before
  // its text. When the marker width is measured exactly (see measureMarkerWidthPt),
  // buffer = markerWidth + this gap — so the visible gap is identical for bullets
  // and numbers and constant at every depth, and the buffer always just clears
  // the marker (so a child item aligns exactly under the parent's text).
  var MARKER_GAP_INCHES = 0.07; // ~5 pt

  // ---- Fallback estimate (used only when the marker width can't be measured,
  // e.g. in the unit tests or if Office font info is unavailable) ----
  var BULLET_BUFFER_INCHES = 0.18; // ~13 pt
  var NUMBER_BASE_INCHES = 0.05; // ~3.5 pt
  var NUMBER_PER_CHAR_INCHES = 0.09; // ~6.5 pt/char ≈ a wide digit

  // The configurable starting indent for the outermost layer, in POINTS.
  // Driven by the slider / "copy from selection" controls in the task pane.
  var baseIndentPoints = 0;
  var liveAlignTimer = null; // debounce handle for real-time slider dragging

  // ---------------------------------------------------------------------------
  // Typed DOM accessors — keep `// @ts-check` happy (and document element types)
  // without scattering casts through the code.
  // ---------------------------------------------------------------------------
  /** @param {string} id @returns {HTMLInputElement} */
  function getInput(id) {
    return /** @type {HTMLInputElement} */ (document.getElementById(id));
  }
  /** @param {string} id @returns {HTMLButtonElement} */
  function getButton(id) {
    return /** @type {HTMLButtonElement} */ (document.getElementById(id));
  }

  // ---------------------------------------------------------------------------
  // Marker width measurement — measure the EXACT rendered width of a list
  // marker ("1.1.1.", "•", "a.") in the paragraph's own font, via a hidden
  // canvas. This makes the buffer = real width + a fixed gap, so the gap is
  // tight/uniform and child markers land precisely under the parent's text,
  // regardless of font.
  // ---------------------------------------------------------------------------
  var _measureCtx = null;

  /**
   * @param {string} listString  the marker text (e.g. "2.1.1.", "•")
   * @param {string} [fontName]   the paragraph font family (from Office.js)
   * @param {number} [fontSizePt] the paragraph font size in points
   * @returns {number|null} the marker width in POINTS, or null if it can't be measured
   */
  function measureMarkerWidthPt(listString, fontName, fontSizePt) {
    if (!listString) return null;
    if (typeof document === "undefined" || !document.createElement) return null;
    if (!_measureCtx) {
      var canvas = document.createElement("canvas");
      _measureCtx = canvas.getContext("2d");
    }
    if (!_measureCtx) return null;
    var size = fontSizePt && fontSizePt > 0 ? fontSizePt : 11;
    var family = fontName || "Calibri";
    // Canvas accepts pt units; measureText returns CSS px (96 dpi).
    _measureCtx.font = size + 'pt "' + family + '"';
    var widthPx = _measureCtx.measureText(listString).width;
    return widthPx * 0.75; // CSS px (96 dpi) -> points (72 dpi)
  }

  /**
   * Ensure each CSS font spec ("11pt \"Aptos Display\"") is loaded before we
   * measure with the canvas — otherwise the first measureText() falls back to a
   * default font (wrong widths) and alignment would need a second click.
   * @param {string[]} specs
   * @returns {Promise<*>}
   */
  function ensureFontsLoaded(specs) {
    if (typeof document === "undefined" || !document.fonts || !document.fonts.load) {
      return Promise.resolve(); // older host: skip, measurement still approximates
    }
    try {
      return Promise.all(
        specs.map(function (spec) {
          return document.fonts.load(spec).catch(function () {
            return null; // a font that can't load just falls back gracefully
          });
        })
      );
    } catch (e) {
      return Promise.resolve();
    }
  }

  // ---------------------------------------------------------------------------
  // Office bootstrap
  // ---------------------------------------------------------------------------
  // Guarded so the module can be `require()`d in a plain Node test runner
  // (where Office.js is absent) without throwing — tests target the pure logic.
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) {
      // Only wire up the UI inside Word.
      if (info.host === Office.HostType.Word) {
        var button = getButton("align-button");
        button.disabled = false;
        button.addEventListener("click", function () {
          alignSelection({ silent: false });
        });

        // Starting-indent slider: update the label + live-preview on drag,
        // and commit a (debounced) re-align so the change is visible in real time.
        var slider = getInput("indent-slider");
        if (slider) {
          slider.addEventListener("input", function () {
            baseIndentPoints = Number(slider.value);
            updateIndentLabel();
            scheduleLiveAlign();
          });
        }

        // "Copy from selection": read the indent of the currently-selected
        // paragraph (e.g. a heading) and adopt it as the starting indent.
        var copyBtn = document.getElementById("copy-indent-button");
        if (copyBtn) {
          copyBtn.addEventListener("click", copyIndentFromSelection);
        }

        updateIndentLabel();
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
   * Count the numbering segments in a marker: "2.1.1." -> 3, "1." -> 1, "a." -> 1.
   * Multi-segment (>= 2) dotted numbers encode their own depth and Word reports
   * them reliably; single-segment markers ("1.", "a.") are ambiguous (a nested
   * restart list looks identical to a top-level list — both report level 0).
   * @param {string} listString
   * @returns {number}
   */
  function listSegmentCount(listString) {
    var groups = (listString || "").trim().match(/[0-9A-Za-z]+/g);
    return groups ? groups.length : 0;
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
   * @param {Array<{isList:boolean, level?:number, listString?:string, markerWidth?:number}>} paras
   *        One entry per paragraph in document order. `isList:false` marks a
   *        non-list paragraph (left untouched; other fields omitted).
   *        `level` is Word's 0-indexed list level; `listString` is the rendered
   *        marker ("1.", "a.", "•"). `markerWidth` (points) is the measured
   *        width of that marker — when present, the buffer is markerWidth + a
   *        fixed gap (exact); when absent, a per-character estimate is used.
   * @param {number} [baseIndent=0]
   *        Starting indent (points) for the outermost layer (level 0). The
   *        whole list shifts right by this amount, preserving relative
   *        alignment — e.g. set it to a heading's leftIndent so the list
   *        starts flush under the heading.
   * @returns {{
   *   results: Array<null|{level:number, isBullet:boolean, alignment:number,
   *                        textIndent:number, leftIndent:number,
   *                        firstLineIndent:number}>,
   *   aligned:number, skipped:number, maxLevel:number
   * }}
   *   `results[i]` is null for skipped (non-list) paragraphs, otherwise the
   *   computed indents in points for paragraph i.
   */
  function computeLayout(paras, baseIndent) {
    var base = baseIndent || 0; // starting indent (points) for level 0
    var textIndentByLevel = {}; // effective level -> text indent (points)
    var results = [];
    var aligned = 0;
    var skipped = 0;
    var maxLevel = 0;

    // Track the previous list item so we can re-parent "orphan" markers. In
    // chaotic docs Word reports stray bullets AND restart-numbered lists at
    // level 0 even though they visually belong under the (deeper) item above.
    var prevLevel = -1; // effective level of the previous list item
    var prevKind = null; // "hard" | "bullet" | "ordinal" (see classification below)
    var bulletRunAnchor = 0; // effective level the current bullet run started at
    var prevBulletWordLevel = 0; // Word's level for the previous bullet

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

      // Classify the marker:
      //  - "hard"    : multi-segment dotted number ("2.1.1.") — encodes its own
      //                depth and Word reports it reliably -> trust wordLevel.
      //  - "bullet"  : a bullet symbol.
      //  - "ordinal" : single-segment number/letter ("1.", "a.") — Word can't
      //                tell a nested restart list from a top-level one (both 0).
      var kind = bullet
        ? "bullet"
        : listSegmentCount(p.listString) >= 2
        ? "hard"
        : "ordinal";

      // Effective level:
      var level;
      if (prevLevel < 0) {
        // First list item in the selection: trust Word's level.
        level = wordLevel;
        if (bullet) {
          bulletRunAnchor = level;
          prevBulletWordLevel = wordLevel;
        }
      } else if (kind === "hard") {
        level = wordLevel; // dotted numbers carry a reliable absolute level
      } else if (kind === "bullet") {
        // Word's ABSOLUTE bullet level is unreliable, but the DELTA between
        // consecutive bullets is reliable:
        //   * first bullet of a run  -> one layer below the item above (anchor)
        //   * later bullets in a run -> follow the delta (sub-bullets nest,
        //     de-indented bullets pop out) but never shallower than the anchor.
        if (prevKind !== "bullet") {
          level = prevLevel + 1;
          bulletRunAnchor = level;
          prevBulletWordLevel = wordLevel;
        } else {
          level = prevLevel + (wordLevel - prevBulletWordLevel);
          if (level < bulletRunAnchor) level = bulletRunAnchor;
          prevBulletWordLevel = wordLevel;
        }
      } else {
        // ordinal: single-segment number/letter ("1.", "2.", "a.").
        if (prevKind === "ordinal") {
          level = prevLevel; // consecutive 1, 2, 3 are siblings
        } else if (prevKind === "bullet") {
          level = prevLevel + 1; // a numbered list nested under a bullet
        } else {
          level = wordLevel; // after a dotted number / at the top -> trust Word
        }
      }

      if (level > maxLevel) maxLevel = level;
      prevLevel = level;
      prevKind = kind;

      // Alignment = where this layer's number/bullet sits.
      // It must equal the text indent of the layer directly above it.
      // Level 0 starts at the configured base indent (default 0 = left margin).
      var alignment =
        level === 0
          ? base
          : typeof textIndentByLevel[level - 1] === "number"
          ? textIndentByLevel[level - 1]
          : base; // no parent seen yet -> fall back to the base indent

      // Buffer = where the text sits relative to the marker. When we have the
      // marker's measured width, the buffer exactly clears it plus one fixed
      // gap (uniform + precise); otherwise fall back to the per-char estimate.
      var buffer =
        typeof p.markerWidth === "number" && p.markerWidth >= 0
          ? p.markerWidth + inchesToPoints(MARKER_GAP_INCHES)
          : computeBufferPoints(p.listString);

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
  // Starting-indent controls
  // ---------------------------------------------------------------------------

  /** Reflect the current base indent (points) in the slider + inches label. */
  function updateIndentLabel() {
    var slider = getInput("indent-slider");
    var label = document.getElementById("indent-value");
    if (slider) slider.value = String(baseIndentPoints);
    if (label) label.textContent = (baseIndentPoints / POINTS_PER_INCH).toFixed(2) + '"';
  }

  /**
   * Debounced re-align while dragging the slider, so the document updates in
   * (near) real time without firing a Word.run on every pixel of movement.
   */
  function scheduleLiveAlign() {
    if (liveAlignTimer) clearTimeout(liveAlignTimer);
    liveAlignTimer = setTimeout(function () {
      liveAlignTimer = null;
      alignSelection({ silent: true });
    }, 120);
  }

  /**
   * Read the indent of the first paragraph in the current selection (e.g. a
   * heading or intro line) and adopt it as the starting indent for the list.
   *
   * The visible start of a paragraph's first line is leftIndent + firstLineIndent:
   *   - a left-indented block:      leftIndent > 0, firstLineIndent = 0
   *   - a first-line-indented line: leftIndent = 0, firstLineIndent > 0  (this case
   *                                 is why reading leftIndent alone returned 0)
   *   - a hanging indent (a list):  leftIndent > 0, firstLineIndent < 0  -> marker pos
   * Summing them gives the true horizontal start to match against.
   */
  function copyIndentFromSelection() {
    setStatus("Reading indent from selection…", "");
    Word.run(function (context) {
      var para = context.document.getSelection().paragraphs.getFirstOrNullObject();
      para.load("leftIndent,firstLineIndent,isNullObject");
      // Also peek at list membership to give a helpful message if the indent
      // turns out to come from list formatting (not paragraph indents).
      var li = para.listItemOrNullObject;
      li.load("isNullObject");
      return context.sync().then(function () {
        if (para.isNullObject) {
          setStatus("Select a paragraph (e.g. the heading) first.", "warn");
          return;
        }

        var left = para.leftIndent || 0;
        var first = para.firstLineIndent || 0;
        var start = left + first; // where the first line visually begins

        // Clamp into the slider's range so the thumb stays in view.
        var slider = getInput("indent-slider");
        var max = slider ? Number(slider.max) : 216;
        var pts = Math.max(0, Math.min(max, start));
        baseIndentPoints = pts;
        updateIndentLabel();

        if (start <= 0 && !li.isNullObject) {
          // Indent comes from the list definition, not paragraph indents.
          setStatus(
            "That paragraph's indent comes from its list formatting, which " +
              "can't be read directly. Set the starting indent with the slider " +
              "instead.",
            "warn"
          );
          return;
        }

        setStatus(
          "Starting indent set to " +
            (pts / POINTS_PER_INCH).toFixed(2) +
            '" (left ' +
            (left / POINTS_PER_INCH).toFixed(2) +
            '" + first-line ' +
            (first / POINTS_PER_INCH).toFixed(2) +
            '"). Now select your list and click Align.',
          "ok"
        );
      });
    }).catch(reportError);
  }

  // ---------------------------------------------------------------------------
  // Main action
  // ---------------------------------------------------------------------------

  /**
   * Align the current selection's list.
   * @param {{silent?:boolean}} [opts] When silent, suppress the chatty status
   *        text (used for live slider re-aligns) and don't toggle the button.
   */
  function alignSelection(opts) {
    var silent = opts && opts.silent;
    var button = getButton("align-button");
    if (!silent) {
      button.disabled = true;
      setStatus("Aligning…", "");
    }

    return Word.run(function (context) {
      // 1) Get the paragraphs in the current selection.
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items");

      return context.sync().then(function () {
        var items = paragraphs.items;

        if (!items || items.length === 0) {
          if (!silent) {
            setStatus("No text is selected. Highlight your list first.", "warn");
          }
          return context.sync(); // nothing to do
        }

        // 2) For every paragraph, load its list membership + level + string,
        //    plus the font (name/size) so we can measure the marker width.
        //    listItemOrNullObject lets us skip non-list paragraphs gracefully.
        var listItems = items.map(function (p) {
          var li = p.listItemOrNullObject;
          li.load("level,listString");
          p.font.load("name,size");
          return li;
        });

        return context.sync().then(function () {
          // 3) Make sure every paragraph font is actually loaded in this pane
          //    BEFORE measuring — otherwise the first measureText() uses a
          //    fallback font (wrong widths) and alignment needs a second click.
          var fontSpecs = {};
          listItems.forEach(function (li, idx) {
            if (li.isNullObject) return;
            var f = items[idx].font;
            var size = f.size && f.size > 0 ? f.size : 11;
            fontSpecs[size + 'pt "' + (f.name || "Calibri") + '"'] = true;
          });

          return ensureFontsLoaded(Object.keys(fontSpecs)).then(function () {
            // Build a plain-data snapshot (measuring each marker's real width in
            // its own font), then run the pure layout algorithm from the base.
            var paras = listItems.map(function (li, idx) {
              if (li.isNullObject) return { isList: false };
              var font = items[idx].font;
              return {
                isList: true,
                level: li.level || 0,
                listString: li.listString,
                markerWidth: measureMarkerWidthPt(li.listString, font.name, font.size),
              };
            });

            var layout = computeLayout(paras, baseIndentPoints);

            // 4) Apply the computed hanging indents back onto the paragraphs.
            for (var i = 0; i < layout.results.length; i++) {
              var r = layout.results[i];
              if (r) {
                applyHangingIndent(items[i], r.alignment, r.textIndent);
              }
            }

            // 5) Single sync to push all indentation changes at once.
            return context.sync().then(function () {
              if (silent) return;
              var msg =
                "Aligned " +
                layout.aligned +
                " list paragraph" +
                (layout.aligned === 1 ? "" : "s") +
                " across " +
                (layout.maxLevel + 1) +
                " level" +
                (layout.maxLevel === 0 ? "" : "s") +
                " (start " +
                (baseIndentPoints / POINTS_PER_INCH).toFixed(2) +
                '").';
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
      });
    }).catch(reportError).then(function () {
      if (!silent) button.disabled = false;
    });
  }

  /** Shared error reporter for Office/OfficeExtension failures. */
  function reportError(error) {
    var detail = error && error.message ? error.message : String(error);
    if (error && error.debugInfo) {
      detail += "\n(" + JSON.stringify(error.debugInfo) + ")";
    }
    setStatus("Something went wrong: " + detail, "err");
    console.error("Perfect Align error:", error);
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
        MARKER_GAP_INCHES: MARKER_GAP_INCHES,
        BULLET_BUFFER_INCHES: BULLET_BUFFER_INCHES,
        NUMBER_BASE_INCHES: NUMBER_BASE_INCHES,
        NUMBER_PER_CHAR_INCHES: NUMBER_PER_CHAR_INCHES,
      },
    };
  }
})();
