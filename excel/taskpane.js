// @ts-check
/*
 * excel/taskpane.js
 * -----------------
 * Excel companion add-in: one button that walks every worksheet tab and sets
 * ALL cells' font to "Aptos Display" at size 11.
 */

(function () {
  "use strict";

  var FONT_NAME = "Aptos Display";
  var FONT_SIZE = 11;

  // Guarded so the file is harmless if ever loaded outside Excel.
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) {
      if (info.host === Office.HostType.Excel) {
        var button = getButton("apply-font-button");
        if (button) {
          button.disabled = false;
          button.addEventListener("click", applyFontToAllSheets);
        }
        var regenBtn = getButton("regen-index-button");
        if (regenBtn) {
          regenBtn.disabled = false;
          regenBtn.addEventListener("click", regenerateIndex);
        }
      }
    });
  }

  /** @param {string} id @returns {HTMLButtonElement} */
  function getButton(id) {
    return /** @type {HTMLButtonElement} */ (document.getElementById(id));
  }

  function setStatus(message, kind) {
    var el = document.getElementById("status");
    if (!el) return;
    el.textContent = message;
    el.className = kind || "";
  }

  // Sheets never listed in the Index (case-insensitive).
  var INDEX_NAME = "Index";
  var EXCLUDE_FROM_INDEX = { index: true, status: true };

  /**
   * Compute a sheet's Index status from its used-range values (2D array).
   * Finds the "Status" header cell, scans the column below it:
   *   - "Done"    -> no "Pending" AND at least one "Confirm"
   *   - "Pending" -> otherwise (any Pending, or only N/A/blank, or no column)
   * @param {any[][]|null} values
   * @returns {"Done"|"Pending"}
   */
  function computeSheetStatus(values) {
    if (!values || !values.length) return "Pending";
    var hr = -1;
    var hc = -1;
    for (var r = 0; r < values.length && hr < 0; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var v = values[r][c];
        if (typeof v === "string" && v.trim().toLowerCase() === "status") {
          hr = r;
          hc = c;
          break;
        }
      }
    }
    if (hr < 0) return "Pending"; // no Status column found
    var hasPending = false;
    var hasConfirm = false;
    for (var rr = hr + 1; rr < values.length; rr++) {
      var cell = values[rr][hc];
      if (typeof cell !== "string") continue;
      var t = cell.trim().toLowerCase();
      if (t === "pending") hasPending = true;
      else if (t === "confirm" || t === "confirmed") hasConfirm = true;
    }
    return !hasPending && hasConfirm ? "Done" : "Pending";
  }

  /** Build an Excel cell reference to a sheet's A1, quoting the sheet name. */
  function sheetA1Ref(name) {
    return "'" + String(name).replace(/'/g, "''") + "'!A1";
  }

  /**
   * Rebuild the "Index" sheet: every worksheet in tab order (excluding Index
   * and Status) with No, Sheet name, a "Link" hyperlink to that sheet, and a
   * Done/Pending status computed from the sheet's Status column.
   */
  function regenerateIndex() {
    var regenBtn = getButton("regen-index-button");
    if (regenBtn) regenBtn.disabled = true;
    setStatus("Re-generating the Index…", "");

    var canHyperlink = true;
    try {
      canHyperlink =
        !Office.context ||
        !Office.context.requirements ||
        Office.context.requirements.isSetSupported("ExcelApi", "1.7");
    } catch (e) {
      canHyperlink = false;
    }

    Excel.run(function (context) {
      var wbSheets = context.workbook.worksheets;
      wbSheets.load("items/name,items/position");
      return context.sync().then(function () {
        var ordered = wbSheets.items.slice().sort(function (a, b) {
          return a.position - b.position;
        });
        var targets = ordered.filter(function (s) {
          return !EXCLUDE_FROM_INDEX[(s.name || "").trim().toLowerCase()];
        });

        // Load each target's used-range values to compute its status.
        var usedRanges = targets.map(function (s) {
          var r = s.getUsedRangeOrNullObject(true); // valuesOnly
          r.load("values,isNullObject");
          return r;
        });
        var indexNull = wbSheets.getItemOrNullObject(INDEX_NAME);
        indexNull.load("isNullObject");

        return context.sync().then(function () {
          var rows = targets.map(function (s, i) {
            var ur = usedRanges[i];
            return {
              name: s.name,
              status: computeSheetStatus(ur.isNullObject ? null : ur.values),
            };
          });

          // Get or create the Index sheet, and clear its old contents.
          var indexSheet;
          if (indexNull.isNullObject) {
            indexSheet = wbSheets.add(INDEX_NAME);
            indexSheet.position = 0;
          } else {
            indexSheet = wbSheets.getItem(INDEX_NAME);
            indexSheet.getUsedRangeOrNullObject().clear(Excel.ClearApplyTo.contents);
          }

          return context.sync().then(function () {
            // Header + rows as a single block.
            var data = /** @type {any[][]} */ ([["No", "Sheet", "Link", "Status"]]);
            rows.forEach(function (row, i) {
              data.push([i + 1, row.name, "Link", row.status]);
            });
            var block = indexSheet.getRangeByIndexes(0, 0, data.length, 4);
            block.values = data;
            indexSheet.getRange("A1:D1").format.font.bold = true;

            // Per-cell internal hyperlinks in the Link column.
            if (canHyperlink) {
              rows.forEach(function (row, i) {
                indexSheet.getRange("C" + (i + 2)).hyperlink = {
                  textToDisplay: "Link",
                  documentReference: sheetA1Ref(row.name),
                };
              });
            }
            indexSheet.getUsedRange().format.autofitColumns();

            return context.sync().then(function () {
              setStatus(
                "Index rebuilt with " +
                  rows.length +
                  " sheet" +
                  (rows.length === 1 ? "" : "s") +
                  ".",
                "ok"
              );
            });
          });
        });
      });
    })
      .catch(function (error) {
        var detail = error && error.message ? error.message : String(error);
        if (error && error.debugInfo) {
          detail += "\n(" + JSON.stringify(error.debugInfo) + ")";
        }
        setStatus("Something went wrong: " + detail, "err");
        console.error("Regenerate Index error:", error);
      })
      .then(function () {
        if (regenBtn) regenBtn.disabled = false;
      });
  }

  /** Set every cell on every worksheet to FONT_NAME / FONT_SIZE. */
  function applyFontToAllSheets() {
    var button = getButton("apply-font-button");
    if (button) button.disabled = true;
    setStatus("Applying " + FONT_NAME + " " + FONT_SIZE + " to all sheets…", "");

    Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.load("items/name");
      return context.sync().then(function () {
        sheets.items.forEach(function (sheet) {
          // getRange() with no address = the entire worksheet (all cells).
          var range = sheet.getRange();
          range.format.font.name = FONT_NAME;
          range.format.font.size = FONT_SIZE;
        });
        return context.sync().then(function () {
          var n = sheets.items.length;
          setStatus(
            "Set " +
              n +
              " sheet" +
              (n === 1 ? "" : "s") +
              " to " +
              FONT_NAME +
              ", size " +
              FONT_SIZE +
              ".",
            "ok"
          );
        });
      });
    })
      .catch(function (error) {
        var detail = error && error.message ? error.message : String(error);
        if (error && error.debugInfo) {
          detail += "\n(" + JSON.stringify(error.debugInfo) + ")";
        }
        setStatus("Something went wrong: " + detail, "err");
        console.error("Excel font tool error:", error);
      })
      .then(function () {
        if (button) button.disabled = false;
      });
  }
})();
