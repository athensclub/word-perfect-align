/*
 * Regression tests for the Perfect Align layout algorithm.
 *
 * Run with:  npm test   (uses Node's built-in test runner — no dependencies)
 *
 * These lock in the behaviour we hand-verified during QA:
 *   - bullet vs number/letter classification
 *   - the dynamic buffer (constant gap after the marker, regardless of length)
 *   - perfect vertical alignment (a level's marker == parent level's text)
 *   - orphan-bullet re-parenting (Word reports stray bullets at level 0)
 *   - hanging-indent math (leftIndent / firstLineIndent)
 *   - graceful handling of non-list paragraphs
 */

"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");

const {
  inchesToPoints,
  isBulletString,
  computeBufferPoints,
  computeLayout,
  constants,
} = require("../src/taskpane.js");

// Small helpers to build inputs the way the add-in feeds computeLayout.
const num = (listString, level) => ({ isList: true, level, listString });
const bul = (listString, level) => ({ isList: true, level, listString });
const nonList = () => ({ isList: false });

// Floating-point tolerant compare (points).
const approx = (a, b, eps = 1e-6) => Math.abs(a - b) <= eps;

// --------------------------------------------------------------------------
// inchesToPoints
// --------------------------------------------------------------------------
test("inchesToPoints converts using 72pt/inch", () => {
  assert.equal(inchesToPoints(1), 72);
  assert.equal(inchesToPoints(0.5), 36);
  assert.equal(inchesToPoints(0), 0);
});

// --------------------------------------------------------------------------
// isBulletString
// --------------------------------------------------------------------------
test("isBulletString: bullets are symbols, numbers/letters are not", () => {
  // Bullets / symbols
  for (const s of ["•", "-", "▪", "o", "*", "‣", "◦"]) {
    assert.equal(isBulletString(s), true, `"${s}" should be a bullet`);
  }
  // Numbered / lettered
  for (const s of ["1.", "1.1.1.", "a.", "A.", "iv.", "10)"]) {
    assert.equal(isBulletString(s), false, `"${s}" should be numbered/lettered`);
  }
});

test("isBulletString: empty/missing falls back to bullet", () => {
  assert.equal(isBulletString(""), true);
  assert.equal(isBulletString("   "), true);
  assert.equal(isBulletString(null), true);
  assert.equal(isBulletString(undefined), true);
});

// --------------------------------------------------------------------------
// computeBufferPoints
// --------------------------------------------------------------------------
test("computeBufferPoints: bullet uses the fixed bullet buffer", () => {
  const expected = inchesToPoints(constants.BULLET_BUFFER_INCHES);
  assert.ok(approx(computeBufferPoints("•"), expected));
  assert.ok(approx(computeBufferPoints("-"), expected));
});

test("computeBufferPoints: number buffer = base + perChar * length", () => {
  const { NUMBER_BASE_INCHES, NUMBER_PER_CHAR_INCHES } = constants;
  for (const s of ["1.", "2.1.", "2.2.2.1.", "1.1.1.1.1."]) {
    const expected = inchesToPoints(
      NUMBER_BASE_INCHES + NUMBER_PER_CHAR_INCHES * s.length
    );
    assert.ok(
      approx(computeBufferPoints(s), expected),
      `buffer for "${s}" should match base+perChar*len`
    );
  }
});

test("computeBufferPoints: longer numbers get a larger buffer (no crash into text)", () => {
  assert.ok(computeBufferPoints("1.1.1.1.1.") > computeBufferPoints("1."));
});

test("computeBufferPoints: visible gap after the number is ~constant across lengths", () => {
  // The gap-after-number ≈ buffer − (numberWidth ≈ perChar * len) ≈ NUMBER_BASE.
  // i.e. removing the per-char glyph allowance leaves a constant base gap.
  const perCharPt = inchesToPoints(constants.NUMBER_PER_CHAR_INCHES);
  const basePt = inchesToPoints(constants.NUMBER_BASE_INCHES);
  for (const s of ["1.", "2.1.", "2.2.2.1.", "1.1.1.1.1."]) {
    const gap = computeBufferPoints(s) - perCharPt * s.length;
    assert.ok(approx(gap, basePt), `constant base gap for "${s}"`);
  }
});

// --------------------------------------------------------------------------
// computeLayout — core alignment
// --------------------------------------------------------------------------
test("computeLayout: level 0 marker sits at the left margin (alignment 0)", () => {
  const { results } = computeLayout([num("1.", 0)]);
  assert.equal(results[0].alignment, 0);
  assert.equal(results[0].level, 0);
});

test("computeLayout: each level's marker aligns with the parent's text start", () => {
  // 1. -> 1.1. -> 1.1.1.  (clean nested numbers)
  const { results } = computeLayout([
    num("1.", 0),
    num("1.1.", 1),
    num("1.1.1.", 2),
  ]);
  // The defining invariant: alignment(level L) === textIndent(level L-1).
  assert.ok(approx(results[1].alignment, results[0].textIndent));
  assert.ok(approx(results[2].alignment, results[1].textIndent));
});

test("computeLayout: hanging-indent math (leftIndent/firstLineIndent)", () => {
  const { results } = computeLayout([num("1.", 0), num("1.1.", 1)]);
  for (const r of results) {
    // leftIndent is the text position.
    assert.equal(r.leftIndent, r.textIndent);
    // firstLineIndent pulls the marker back to `alignment` (negative hang).
    assert.ok(approx(r.firstLineIndent, -(r.textIndent - r.alignment)));
    // Marker actually renders at the alignment position.
    assert.ok(approx(r.leftIndent + r.firstLineIndent, r.alignment));
  }
});

// --------------------------------------------------------------------------
// computeLayout — orphan bullet re-parenting (the QA regression)
// --------------------------------------------------------------------------
test("computeLayout: orphan bullets (Word level 0) nest under the item above", () => {
  // Word reports the bullets at level 0, but they follow "2.1.1." (level 2)
  // and should render one layer below it (effective level 3).
  const { results } = computeLayout([
    num("2.1.1.", 2),
    bul("•", 0),
    bul("•", 0),
    bul("•", 0),
  ]);
  const parent = results[0];
  for (let i = 1; i <= 3; i++) {
    assert.equal(results[i].level, 3, `bullet ${i} should be effective level 3`);
    // Its marker aligns exactly with the numbered parent's text start.
    assert.ok(
      approx(results[i].alignment, parent.textIndent),
      `bullet ${i} marker should sit at parent's text start`
    );
  }
});

test("computeLayout: consecutive bullets at the same Word level stay siblings", () => {
  const { results } = computeLayout([
    num("2.", 0),
    bul("•", 0),
    bul("•", 0),
  ]);
  assert.equal(results[1].level, results[2].level);
  assert.ok(approx(results[1].alignment, results[2].alignment));
  assert.ok(approx(results[1].textIndent, results[2].textIndent));
});

test("computeLayout: a bullet nested deeper in Word nests one layer deeper", () => {
  // A is a bullet under the number; B is indented one more in Word (level 1).
  const { results } = computeLayout([
    num("2.", 0), // eff 0
    bul("•", 0), // first bullet -> eff 1 (anchor), word baseline 0
    bul("•", 1), // delta +1 -> eff 2
  ]);
  assert.equal(results[1].level, 1);
  assert.equal(results[2].level, 2);
  // The deeper bullet's marker aligns under the shallower bullet's text.
  assert.ok(approx(results[2].alignment, results[1].textIndent));
});

test("computeLayout: a de-indented bullet pops back out (delta down)", () => {
  const { results } = computeLayout([
    num("2.", 0), // eff 0
    bul("•", 0), // eff 1 (anchor)
    bul("•", 1), // eff 2 (sub-bullet)
    bul("•", 1), // eff 2 (sibling sub-bullet)
    bul("•", 0), // delta -1 -> eff 1 (back out to the anchor level)
  ]);
  assert.deepEqual(
    results.map((r) => r.level),
    [0, 1, 2, 2, 1]
  );
});

test("computeLayout: bullets never pop out shallower than their run anchor", () => {
  // Even if Word reports a bullet shallower than the run's first bullet, it is
  // floored at the anchor so it stays within the numbered parent's subtree.
  const { results } = computeLayout([
    num("2.2.1.", 2), // eff 2
    bul("•", 1), // first bullet -> eff 3 (anchor), word baseline 1
    bul("•", 0), // delta -1 -> 2, but floored at anchor 3
  ]);
  assert.equal(results[1].level, 3);
  assert.equal(results[2].level, 3);
});

test("computeLayout: a leading bullet with no item above falls back to its Word level", () => {
  const { results } = computeLayout([bul("•", 0)]);
  assert.equal(results[0].level, 0);
  assert.equal(results[0].alignment, 0);
});

test("computeLayout: a restart-numbered list nested under a bullet nests under it", () => {
  // The reported case: 2. -> • (bullet) -> 1./2./3. (restart numbers at Word 0)
  // The inner numbers must sit one layer below the bullet, not jump to the top.
  const { results } = computeLayout([
    num("2.", 0), // ordinal, top level -> 0
    bul("•", 0), // bullet under it -> 1
    num("1.", 0), // ordinal after a bullet -> nests under it -> 2
    num("2.", 0), // consecutive ordinal -> sibling -> 2
    num("3.", 0), // sibling -> 2
  ]);
  assert.deepEqual(
    results.map((r) => r.level),
    [0, 1, 2, 2, 2]
  );
});

test("computeLayout: an ordinal that continues an outer run returns to that run's level", () => {
  // 2. -> • -> 1..5 (nested) -> 3. ("3" continues the outer "…2." run, not the
  // inner "…5." one, so it pops all the way back out to level 0)
  const { results } = computeLayout([
    num("2.", 0), // 0  (outer run: …2)
    bul("•", 0), // 1
    num("1.", 0), // 2  (inner restart)
    num("2.", 0), // 2
    num("5.", 0), // 2  (inner run: …5)
    num("3.", 0), // 3 > outer 2 -> continues the outer run -> 0
  ]);
  assert.deepEqual(
    results.map((r) => r.level),
    [0, 1, 2, 2, 2, 0]
  );
});

test("computeLayout: a numbered run continues at its own level across bullet runs", () => {
  // The reported case: 3./4./5. each have bullets under them; 4 and 5 continue
  // the run instead of nesting ever deeper under the preceding bullets.
  const { results } = computeLayout([
    num("3.", 0), // 0
    bul("•", 0), // 1
    bul("•", 0), // 1
    num("4.", 0), // continues the …3 run -> 0
    bul("•", 0), // 1
    num("5.", 0), // continues the …4 run -> 0
  ]);
  assert.deepEqual(
    results.map((r) => r.level),
    [0, 1, 1, 0, 1, 0]
  );
});

test("computeLayout: markerInk pulls the marker back without moving text or children", () => {
  const plain = computeLayout([num("1.", 0)]).results[0];
  const inked = computeLayout([
    { isList: true, level: 0, listString: "1.", markerInk: 2 },
  ]).results[0];
  // Text position unchanged; only the marker hang grows by the ink inset.
  assert.equal(inked.leftIndent, plain.leftIndent);
  assert.ok(approx(inked.firstLineIndent, plain.firstLineIndent - 2));
  assert.ok(approx(inked.textIndent, plain.textIndent));
});

test("computeLayout: ascending ordinals (even with a skip) stay siblings", () => {
  const { results } = computeLayout([
    num("2.", 0),
    bul("•", 0),
    num("1.", 0), // 2
    num("3.", 0), // skip ahead 1->3, still ascending -> sibling (2)
  ]);
  assert.deepEqual(
    results.map((r) => r.level),
    [0, 1, 2, 2]
  );
});

test("computeLayout: top-level numbered siblings stay at the same level", () => {
  const { results } = computeLayout([num("1.", 0), num("2.", 0), num("3.", 0)]);
  assert.deepEqual(
    results.map((r) => r.level),
    [0, 0, 0]
  );
});

test("computeLayout: multi-segment dotted numbers keep their reliable Word level", () => {
  // A dotted number after a bullet is NOT re-parented — it trusts its own depth.
  const { results } = computeLayout([
    num("2.", 0),
    bul("•", 0), // -> 1
    num("2.1.", 1), // dotted (hard) -> trusts Word level 1, not nested under bullet
  ]);
  assert.deepEqual(
    results.map((r) => r.level),
    [0, 1, 1]
  );
});

test("computeLayout: number after bullets returns to its own Word level", () => {
  const { results } = computeLayout([
    num("2.2.1.", 2),
    bul("•", 0), // -> level 3
    num("2.2.2.", 2), // trusts Word again -> level 2
  ]);
  assert.equal(results[1].level, 3);
  assert.equal(results[2].level, 2);
});

// --------------------------------------------------------------------------
// computeLayout — measured marker width (exact buffer)
// --------------------------------------------------------------------------
const mnum = (listString, level, markerWidth) => ({
  isList: true,
  level,
  listString,
  markerWidth,
});

test("computeLayout: with markerWidth, buffer = markerWidth + uniform gap", () => {
  const gap = inchesToPoints(constants.MARKER_GAP_INCHES);
  const { results } = computeLayout([mnum("1.", 0, 10)]);
  // textIndent (level 0, base 0) = alignment(0) + buffer(markerWidth + gap)
  assert.ok(approx(results[0].textIndent, 10 + gap));
});

test("computeLayout: the gap is uniform for a number and a bullet of any width", () => {
  const gap = inchesToPoints(constants.MARKER_GAP_INCHES);
  // A wide number and a narrow bullet both get exactly the same trailing gap.
  const number = computeLayout([mnum("2.1.1.1.", 0, 52)]).results[0];
  const bullet = computeLayout([mnum("•", 0, 5)]).results[0];
  assert.ok(approx(number.textIndent - number.alignment - 52, gap));
  assert.ok(approx(bullet.textIndent - bullet.alignment - 5, gap));
});

test("computeLayout: a measured child marker aligns exactly under the parent text", () => {
  // Parent number measured at 52pt; its child bullet should sit at the parent's
  // text start regardless of the parent marker's width.
  const { results } = computeLayout([
    mnum("2.1.1.1.", 3, 52), // some deep number
    mnum("•", 0, 5), // orphan bullet -> child of the number
  ]);
  assert.ok(approx(results[1].alignment, results[0].textIndent));
});

test("computeLayout: markerWidth absent -> falls back to the per-char estimate", () => {
  const withWidth = computeLayout([num("1.", 0)]).results[0]; // no markerWidth
  assert.ok(approx(withWidth.textIndent, computeBufferPoints("1.")));
});

// --------------------------------------------------------------------------
// computeLayout — starting (base) indent
// --------------------------------------------------------------------------
test("computeLayout: baseIndent shifts level 0 to the given start", () => {
  const base = 36; // 0.5"
  const { results } = computeLayout([num("1.", 0)], base);
  assert.equal(results[0].alignment, base);
  // text indent = base + buffer
  assert.ok(approx(results[0].textIndent, base + computeBufferPoints("1.")));
});

test("computeLayout: baseIndent shifts the whole tree but preserves relative alignment", () => {
  const paras = [num("1.", 0), num("1.1.", 1), bul("•", 0)];
  const a = computeLayout(paras, 0).results;
  const b = computeLayout(paras, 50).results;
  for (let i = 0; i < paras.length; i++) {
    // Every position is shifted right by exactly the base offset...
    assert.ok(approx(b[i].alignment, a[i].alignment + 50));
    assert.ok(approx(b[i].textIndent, a[i].textIndent + 50));
    // ...and the hanging-indent (the gap) is unchanged.
    assert.ok(approx(b[i].firstLineIndent, a[i].firstLineIndent));
  }
});

test("computeLayout: the vertical-alignment invariant still holds with a base offset", () => {
  const { results } = computeLayout(
    [num("1.", 0), num("1.1.", 1), num("1.1.1.", 2)],
    72
  );
  assert.equal(results[0].alignment, 72);
  assert.ok(approx(results[1].alignment, results[0].textIndent));
  assert.ok(approx(results[2].alignment, results[1].textIndent));
});

test("computeLayout: omitted baseIndent defaults to 0 (back-compat)", () => {
  const { results } = computeLayout([num("1.", 0)]);
  assert.equal(results[0].alignment, 0);
});

// --------------------------------------------------------------------------
// computeLayout — bookkeeping & edge cases
// --------------------------------------------------------------------------
test("computeLayout: non-list paragraphs are skipped (null result, untouched)", () => {
  const { results, aligned, skipped } = computeLayout([
    num("1.", 0),
    nonList(),
    num("2.", 0),
  ]);
  assert.equal(results[1], null);
  assert.equal(aligned, 2);
  assert.equal(skipped, 1);
});

test("computeLayout: empty input yields zero counts", () => {
  const { results, aligned, skipped, maxLevel } = computeLayout([]);
  assert.deepEqual(results, []);
  assert.equal(aligned, 0);
  assert.equal(skipped, 0);
  assert.equal(maxLevel, 0);
});

test("computeLayout: maxLevel reflects the deepest effective level reached", () => {
  const { maxLevel } = computeLayout([
    num("1.", 0),
    num("1.1.", 1),
    num("1.1.1.", 2),
    bul("•", 0), // orphan -> effective level 3
  ]);
  assert.equal(maxLevel, 3);
});

// --------------------------------------------------------------------------
// computeLayout — full chaotic document (the real QA sample)
// --------------------------------------------------------------------------
test("computeLayout: full chaotic sample keeps every marker under its parent's text", () => {
  const paras = [
    num("2.", 0),
    num("2.1.", 1),
    num("2.1.1.", 2),
    bul("•", 0),
    bul("•", 0),
    bul("•", 0),
    num("2.1.2.", 2),
    num("2.2.", 1),
    num("2.2.1.", 2),
    bul("•", 0),
    bul("•", 0),
    bul("•", 0),
    num("2.2.2.", 2),
    num("2.2.2.1.", 3),
    bul("•", 0),
    bul("•", 0),
  ];
  const { results, aligned } = computeLayout(paras);
  assert.equal(aligned, paras.length);

  // For every aligned paragraph at level > 0, its marker must coincide with
  // SOME ancestor's text start that we recorded — specifically the most recent
  // paragraph at (level - 1). We verify the core invariant holds end-to-end by
  // replaying the recorded text indents per level.
  const textAtLevel = {};
  for (const r of results) {
    if (r.level > 0) {
      assert.ok(
        approx(r.alignment, textAtLevel[r.level - 1]),
        `level ${r.level} marker must align to level ${r.level - 1} text`
      );
    } else {
      assert.equal(r.alignment, 0);
    }
    textAtLevel[r.level] = r.textIndent;
  }
});
