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
