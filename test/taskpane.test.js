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

test("computeLayout: consecutive bullets are siblings (same level)", () => {
  const { results } = computeLayout([
    num("2.", 0),
    bul("•", 0),
    bul("•", 0),
  ]);
  assert.equal(results[1].level, results[2].level);
  assert.ok(approx(results[1].alignment, results[2].alignment));
  assert.ok(approx(results[1].textIndent, results[2].textIndent));
});

test("computeLayout: a leading bullet with no item above falls back to its Word level", () => {
  const { results } = computeLayout([bul("•", 0)]);
  assert.equal(results[0].level, 0);
  assert.equal(results[0].alignment, 0);
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
