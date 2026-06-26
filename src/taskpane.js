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
  var lastShiftPoints = 0; // slider value already applied as a live translate
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
   * Measure the ink inset (left side bearing) of a marker's first glyph, in
   * points. Bullet glyphs like "•" carry noticeable blank space before their
   * visible dot, so a marker box placed exactly at the parent's text start
   * LOOKS shifted right; subtracting this inset aligns the visible ink instead
   * of the invisible glyph box. Returns 0 when it can't be measured.
   * @param {string} listString
   * @param {string} [fontName]
   * @param {number} [fontSizePt]
   * @returns {number}
   */
  function measureMarkerInkLeftPt(listString, fontName, fontSizePt) {
    if (!listString) return 0;
    if (typeof document === "undefined" || !document.createElement) return 0;
    if (!_measureCtx) {
      var canvas = document.createElement("canvas");
      _measureCtx = canvas.getContext("2d");
    }
    if (!_measureCtx) return 0;
    var size = fontSizePt && fontSizePt > 0 ? fontSizePt : 11;
    var family = fontName || "Calibri";
    _measureCtx.font = size + 'pt "' + family + '"';
    var m = _measureCtx.measureText(listString);
    if (typeof m.actualBoundingBoxLeft !== "number") return 0;
    // With left-aligned text the ink's left edge sits at -actualBoundingBoxLeft
    // px right of the origin (the value is negative for a normal side bearing).
    var insetPx = -m.actualBoundingBoxLeft;
    if (!(insetPx > 0)) return 0;
    var insetPt = insetPx * 0.75; // CSS px (96 dpi) -> points
    return Math.min(insetPt, 6); // safety clamp for odd font metrics
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

        // Starting-indent slider: live-translate the selection on drag (shifts
        // only the starting indent; relative/hanging indents are untouched).
        var slider = getInput("indent-slider");
        if (slider) {
          slider.addEventListener("input", function () {
            baseIndentPoints = Number(slider.value);
            updateIndentLabel();
            scheduleLiveShift();
          });
        }

        // "Copy from selection": read the indent of the currently-selected
        // paragraph (e.g. a heading) and adopt it as the starting indent.
        var copyBtn = document.getElementById("copy-indent-button");
        if (copyBtn) {
          copyBtn.addEventListener("click", copyIndentFromSelection);
        }

        // Paragraph-style dropdown: apply a built-in style while preserving the
        // selection's indent and list (unlike Word's built-in apply).
        var styleSelect = /** @type {HTMLSelectElement} */ (
          document.getElementById("style-select")
        );
        if (styleSelect) {
          styleSelect.addEventListener("change", function () {
            if (styleSelect.value) {
              applyStylePreservingFormat(styleSelect.value);
              styleSelect.value = ""; // reset to the placeholder
            }
          });
          populateStyleDropdown(); // fill with the document's actual styles
        }

        // "Highlight all (TBC)": scan the whole document body and yellow-
        // highlight every "(TBC)" occurrence.
        var highlightBtn = document.getElementById("highlight-tbc-button");
        if (highlightBtn) {
          highlightBtn.addEventListener("click", highlightTbc);
        }

        // "Renumber error IDs": find every error-handling table and reset its
        // Error ID to {prefix}-{n} in document order.
        var renumberBtn = document.getElementById("renumber-errors-button");
        if (renumberBtn) {
          renumberBtn.addEventListener("click", function () {
            var input = /** @type {HTMLInputElement} */ (
              document.getElementById("error-prefix")
            );
            renumberErrorIds(input ? input.value : "");
          });
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
   * The numeric value of a single-segment ordinal marker: "3." -> 3, "a." -> 1,
   * "10)" -> 10. Used to detect when a numbered run breaks (e.g. 5 then 3),
   * which means the new item belongs to an OUTER list, not a continuation.
   * Returns null when it can't be determined.
   * @param {string} listString
   * @returns {number|null}
   */
  function ordinalValue(listString) {
    var t = (listString || "").trim();
    var digits = t.match(/\d+/);
    if (digits) return parseInt(digits[0], 10);
    var letter = t.match(/[A-Za-z]/);
    if (letter) return letter[0].toLowerCase().charCodeAt(0) - 96; // a=1, b=2, ...
    return null;
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
   * Apply a computed layout result to a paragraph as a clean hanging indent.
   *   leftIndent      = textIndent (where the text sits)
   *   firstLineIndent = negative hang pulling the marker's visible ink back to
   *                     the alignment position (includes the ink-inset shift).
   * @param {Word.Paragraph} paragraph
   * @param {{leftIndent:number, firstLineIndent:number}} r
   */
  function applyHangingIndent(paragraph, r) {
    paragraph.leftIndent = r.leftIndent;
    paragraph.firstLineIndent = r.firstLineIndent;
  }

  // ---------------------------------------------------------------------------
  // Core algorithm (PURE — no Office.js dependency, fully unit-testable)
  // ---------------------------------------------------------------------------

  /**
   * Compute the indent layout for a list of paragraphs.
   *
   * @param {Array<{isList:boolean, level?:number, listString?:string, markerWidth?:number, markerInk?:number}>} paras
   *        One entry per paragraph in document order. `isList:false` marks a
   *        non-list paragraph (left untouched; other fields omitted).
   *        `level` is Word's 0-indexed list level; `listString` is the rendered
   *        marker ("1.", "a.", "•"). `markerWidth` (points) is the measured
   *        width of that marker — when present, the buffer is markerWidth + a
   *        fixed gap (exact); when absent, a per-character estimate is used.
   *        `markerInk` (points) is the marker glyph's ink inset (left side
   *        bearing) — the marker is pulled back by this so its visible ink,
   *        not its glyph box, sits at the alignment position.
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
    // A bullet run is anchored one level below the item it follows. Within the
    // run we trust the DIRECTION of Word's ilvl but COMPRESS its magnitude via a
    // stack: a deeper ilvl is one level in, a shallower ilvl one level out — so
    // a tab that jumps ilvl 1 -> 8 still means just "one deeper", and an "o"
    // list that restarts at ilvl 0 after a "•" at ilvl 1 moves one level out.
    var bulletStack = []; // [{ ilvl, eff }] from outer (small eff) to inner
    var bulletRunParentLevel = 0; // floor: don't outdent past the run's parent
    // Open numbered runs: effective level -> {value, letter} of the last ordinal
    // seen there. Lets "4." return to the level of an open "…3." run even when
    // a bullet run sits in between (continuation beats nesting).
    var ordinalRunByLevel = {};

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
      var ordVal = kind === "ordinal" ? ordinalValue(p.listString) : null;
      var ordIsLetter =
        kind === "ordinal" && !/\d/.test((p.listString || "").trim());

      // Effective level:
      var level;
      if (prevLevel < 0) {
        // First list item in the selection: trust Word's level.
        level = wordLevel;
        if (bullet) {
          bulletRunParentLevel = 0;
          bulletStack = [{ ilvl: wordLevel, eff: level }];
        }
      } else if (kind === "hard") {
        level = wordLevel; // dotted numbers carry a reliable absolute level
      } else if (kind === "bullet") {
        if (prevKind !== "bullet") {
          // Start a bullet run one level below the item above (the anchor).
          level = prevLevel + 1;
          bulletRunParentLevel = prevLevel;
          bulletStack = [{ ilvl: wordLevel, eff: level }];
        } else {
          // Pop entries deeper than this ilvl, then place relative to the top.
          var lastPoppedEff = -1;
          while (
            bulletStack.length &&
            bulletStack[bulletStack.length - 1].ilvl > wordLevel
          ) {
            lastPoppedEff = bulletStack.pop().eff;
          }
          var top = bulletStack[bulletStack.length - 1];
          if (top && top.ilvl === wordLevel) {
            level = top.eff; // sibling at the same depth
          } else if (top && top.ilvl < wordLevel) {
            level = top.eff + 1; // one level deeper
            bulletStack.push({ ilvl: wordLevel, eff: level });
          } else {
            // shallower than everything in the run -> one level outward
            level = (lastPoppedEff >= 0 ? lastPoppedEff : prevLevel) - 1;
            if (level < bulletRunParentLevel) level = bulletRunParentLevel;
            if (level < 0) level = 0;
            bulletStack.push({ ilvl: wordLevel, eff: level });
          }
        }
      } else {
        // ordinal: single-segment number/letter ("1.", "2.", "a.").
        // Continuation beats nesting: if this value extends an OPEN numbered
        // run at some level (e.g. "4." while "…3." is still open, with a
        // bullet run in between), return to that run's level — that's how
        // 3./4./5. stay siblings even when each has bullets under it. Only a
        // fresh restart (typically "1.") starts a new list nested one layer
        // under whatever came before.
        var contLevel = -1;
        if (ordVal !== null) {
          for (var lv in ordinalRunByLevel) {
            var run = ordinalRunByLevel[lv];
            if (
              run.letter === ordIsLetter &&
              ordVal > run.value &&
              Number(lv) > contLevel
            ) {
              contLevel = Number(lv); // prefer the deepest continuing run
            }
          }
        }
        if (contLevel >= 0) {
          level = contLevel;
        } else if (prevKind === "bullet" || prevKind === "ordinal") {
          level = prevLevel + 1; // restarted list -> nest under the item above
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
      for (var deeperRun in ordinalRunByLevel) {
        if (Number(deeperRun) > level) {
          delete ordinalRunByLevel[deeperRun]; // runs deeper than us are closed
        }
      }
      if (kind === "ordinal" && ordVal !== null) {
        ordinalRunByLevel[level] = { value: ordVal, letter: ordIsLetter };
      }

      var markerInk =
        typeof p.markerInk === "number" && p.markerInk > 0 ? p.markerInk : 0;
      results.push({
        level: level,
        isBullet: bullet,
        alignment: alignment,
        textIndent: textIndent,
        leftIndent: textIndent,
        // Negative hang pulls the marker back so its visible INK (glyph box
        // minus the ink inset) lands exactly at `alignment`.
        firstLineIndent: alignment - markerInk - textIndent,
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
   * Debounced live TRANSLATE while dragging the slider: shift the whole
   * selection by the change in the slider value, without re-running the
   * alignment. This only moves each paragraph's leftIndent (its starting
   * position) — relative indents / hanging indents are left untouched.
   */
  function scheduleLiveShift() {
    if (liveAlignTimer) clearTimeout(liveAlignTimer);
    liveAlignTimer = setTimeout(function () {
      liveAlignTimer = null;
      var delta = baseIndentPoints - lastShiftPoints;
      if (delta === 0) return;
      shiftSelectionIndent(delta);
      lastShiftPoints = baseIndentPoints;
    }, 120);
  }

  /**
   * Shift every selected paragraph's left indent by `deltaPoints`, leaving
   * firstLineIndent (and thus the marker hang / relative structure) untouched —
   * the whole selection translates horizontally.
   * @param {number} deltaPoints
   */
  function shiftSelectionIndent(deltaPoints) {
    Word.run(function (context) {
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        var items = paragraphs.items;
        if (!items || items.length === 0) return context.sync();
        items.forEach(function (p) {
          p.load("leftIndent");
        });
        return context.sync().then(function () {
          items.forEach(function (p) {
            var next = (p.leftIndent || 0) + deltaPoints;
            p.leftIndent = next < 0 ? 0 : next;
          });
          return context.sync();
        });
      });
    }).catch(reportError);
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
        lastShiftPoints = pts; // copying sets the value; it doesn't translate
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
                markerInk: measureMarkerInkLeftPt(li.listString, font.name, font.size),
              };
            });

            var layout = computeLayout(paras, baseIndentPoints);

            // 4) Apply the computed hanging indents back onto the paragraphs.
            for (var i = 0; i < layout.results.length; i++) {
              var r = layout.results[i];
              if (r) {
                applyHangingIndent(items[i], r);
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
      // After a full align the selection's level 0 sits at baseIndentPoints, so
      // sync the slider's translate baseline (a later drag shifts from here).
      lastShiftPoints = baseIndentPoints;
      if (!silent) button.disabled = false;
    });
  }

  /**
   * Populate the style dropdown with EVERY paragraph style defined in the
   * document (via Document.getStyles), used styles first. Falls back to the
   * static built-in options already in the HTML if getStyles is unavailable.
   */
  function populateStyleDropdown() {
    var select = /** @type {HTMLSelectElement} */ (
      document.getElementById("style-select")
    );
    if (!select || typeof Word === "undefined" || !Word.run) return;
    // getStyles needs WordApi 1.5; if absent, keep the static fallback list.
    try {
      if (
        Office.context &&
        Office.context.requirements &&
        !Office.context.requirements.isSetSupported("WordApi", "1.5")
      ) {
        return;
      }
    } catch (e) {
      /* ignore and try anyway */
    }

    Word.run(function (context) {
      var styles = context.document.getStyles();
      styles.load("nameLocal,type,inUse");
      return context.sync().then(function () {
        var paras = styles.items.filter(function (s) {
          return s.type === Word.StyleType.paragraph;
        });
        // Sort: headings first (by their number), then Normal, then the rest
        // (in-use first within the rest, then alphabetical).
        var rankOf = function (name) {
          var n = (name || "").toLowerCase();
          if (/heading/.test(n)) return 0;
          if (n === "normal") return 1;
          return 2;
        };
        var headingNum = function (name) {
          var m = (name || "").match(/\d+/);
          return m ? parseInt(m[0], 10) : Infinity; // un-numbered headings last
        };
        paras.sort(function (a, b) {
          var ra = rankOf(a.nameLocal);
          var rb = rankOf(b.nameLocal);
          if (ra !== rb) return ra - rb;
          if (ra === 0) {
            var na = headingNum(a.nameLocal);
            var nb = headingNum(b.nameLocal);
            if (na !== nb) return na - nb;
          } else if (ra === 2 && a.inUse !== b.inUse) {
            return a.inUse ? -1 : 1;
          }
          return a.nameLocal.localeCompare(b.nameLocal);
        });
        // Rebuild options, keeping the first (placeholder) one.
        while (select.options.length > 1) select.remove(1);
        paras.forEach(function (s) {
          var opt = document.createElement("option");
          opt.value = s.nameLocal;
          opt.textContent = s.nameLocal + (s.inUse ? "" : " (unused)");
          select.appendChild(opt);
        });
      });
    }).catch(function (e) {
      console.warn("getStyles unavailable; using built-in style list", e);
    });
  }

  /**
   * Apply a paragraph style (by its display name) to the selection while
   * PRESERVING each paragraph's number/bullet and indent. Changing a paragraph's
   * style resets indentation and — for style/list-linked numbering like a "4.1."
   * heading — drops the number. We snapshot each paragraph's list (id + level)
   * and indents, apply the style, then: if the style dropped the list, re-attach
   * it (which restores the number AND its indent); for non-list paragraphs we
   * restore the saved direct indent.
   * @param {string} styleName the style's display name (Style.nameLocal)
   */
  function applyStylePreservingFormat(styleName) {
    setStatus("Applying " + styleName + "…", "");

    // Plain-data snapshot shared between the two passes (proxies don't survive
    // across Word.run calls, but plain JS does).
    var saved = null;

    // Pass 1: snapshot indents + list (id/level), then apply the style.
    Word.run(function (context) {
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        var items = paragraphs.items;
        if (!items || items.length === 0) return context.sync();
        items.forEach(function (p) {
          p.load("leftIndent,firstLineIndent");
          p.listOrNullObject.load("id,isNullObject");
          p.listItemOrNullObject.load("level,isNullObject");
        });
        return context.sync().then(function () {
          saved = items.map(function (p) {
            var inList = !p.listOrNullObject.isNullObject;
            return {
              left: p.leftIndent || 0,
              first: p.firstLineIndent || 0,
              inList: inList,
              listId: inList ? p.listOrNullObject.id : null,
              level: p.listItemOrNullObject.isNullObject
                ? 0
                : p.listItemOrNullObject.level || 0,
            };
          });
          items.forEach(function (p) {
            p.style = styleName;
          });
          return context.sync();
        });
      });
    })
      .then(function () {
        if (!saved) {
          setStatus("No text is selected. Highlight some paragraphs first.", "warn");
          return;
        }
        // Pass 2: in a FRESH context (so list proxies aren't the ones the style
        // change invalidated), restore numbering/indent.
        return Word.run(function (context) {
          var paragraphs = context.document.getSelection().paragraphs;
          paragraphs.load("items");
          return context.sync().then(function () {
            var items = paragraphs.items;
            var n = Math.min(items.length, saved.length);
            for (var k = 0; k < n; k++) {
              items[k].listItemOrNullObject.load("isNullObject");
            }
            return context.sync().then(function () {
              var reattached = 0;
              for (var i = 0; i < n; i++) {
                var s = saved[i];
                var p = items[i];
                var stillInList = !p.listItemOrNullObject.isNullObject;
                if (s.inList && !stillInList && s.listId != null) {
                  // Re-join the original list -> restores the number/bullet.
                  p.attachToList(s.listId, s.level);
                  reattached++;
                }
                // Restore the indent for EVERY paragraph: the new style sets its
                // own indent, so re-apply the captured (effective) one — this is
                // what a numbered heading needs, since its indent came from the
                // old style, not from the list.
                p.leftIndent = s.left;
                p.firstLineIndent = s.first;
              }
              return context.sync().then(function () {
                setStatus(
                  "Applied " +
                    styleName +
                    " to " +
                    n +
                    " paragraph" +
                    (n === 1 ? "" : "s") +
                    ", keeping indent" +
                    (reattached ? " and restoring the list numbering" : "") +
                    ".",
                  "ok"
                );
              });
            });
          });
        });
      })
      .catch(reportError);
  }

  /**
   * Scan the whole document body and apply a yellow highlight to every
   * "(TBC)" occurrence.
   */
  function highlightTbc() {
    setStatus("Scanning for “(TBC)”…", "");
    Word.run(function (context) {
      // Literal (non-wildcard) search; parentheses match as-is.
      var results = context.document.body.search("(TBC)", {
        matchCase: false,
        matchWildcards: false,
      });
      results.load("items");
      return context.sync().then(function () {
        var count = results.items.length;
        results.items.forEach(function (r) {
          r.font.highlightColor = "#FFFF00"; // yellow
        });
        return context.sync().then(function () {
          setStatus(
            count > 0
              ? "Highlighted " +
                  count +
                  " “(TBC)” occurrence" +
                  (count === 1 ? "" : "s") +
                  "."
              : "No “(TBC)” found in the document.",
            count > 0 ? "ok" : "warn"
          );
        });
      });
    }).catch(reportError);
  }

  /**
   * Find every error-handling table (one whose first cell reads "Error ID")
   * and reset its Error ID value (the cell to the right) to {prefix}-{n},
   * numbered from 1 in document order.
   * @param {string} prefixRaw
   */
  function renumberErrorIds(prefixRaw) {
    var prefix = (prefixRaw || "").trim();
    if (!prefix) {
      setStatus("Enter an error ID prefix first.", "warn");
      return;
    }
    setStatus("Renumbering error IDs…", "");
    Word.run(function (context) {
      var tables = context.document.body.tables;
      tables.load("items");
      return context.sync().then(function () {
        var items = tables.items;
        items.forEach(function (t) {
          t.load("values");
        });
        return context.sync().then(function () {
          var n = 0;
          items.forEach(function (t) {
            var v = t.values;
            var isErrorTable =
              v &&
              v[0] &&
              v[0].length >= 2 &&
              typeof v[0][0] === "string" &&
              v[0][0].trim().toLowerCase() === "error id";
            if (isErrorTable) {
              n++;
              t.getCell(0, 1).value = prefix + "-" + n;
            }
          });
          return context.sync().then(function () {
            setStatus(
              n > 0
                ? "Renumbered " +
                    n +
                    " error table" +
                    (n === 1 ? "" : "s") +
                    " as " +
                    prefix +
                    "-1 … " +
                    prefix +
                    "-" +
                    n +
                    "."
                : "No error-handling tables (a cell reading “Error ID”) were found.",
              n > 0 ? "ok" : "warn"
            );
          });
        });
      });
    }).catch(reportError);
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
