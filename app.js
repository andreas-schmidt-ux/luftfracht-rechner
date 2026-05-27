/* ============================================================
   Luftfracht Rechner – app.js
   Pure vanilla JS, no build step required.
   ============================================================ */

'use strict';

// ── Office.onReady / graceful degradation ──────────────────────
if (typeof Office !== 'undefined') {
  Office.onReady(function () {
    initApp();
  });
} else {
  // Running in plain browser without Office.js (development / testing)
  document.addEventListener('DOMContentLoaded', initApp);
}

// ── State ──────────────────────────────────────────────────────
var items = []; // Array of { qty, l, b, h, weight }

// ── DOM refs (populated after DOMContentLoaded) ────────────────
var elPhaseInput, elPhaseResults, elEmailText;
var elItemsTbody, elResPieces, elResWeight, elResVolume;
var elResVolWeight, elResChargeable, elWeightWarning, elWarningText;
var elCopyFeedback;

// ── Init ───────────────────────────────────────────────────────
function initApp() {
  elPhaseInput    = document.getElementById('phase-input');
  elPhaseResults  = document.getElementById('phase-results');
  elEmailText     = document.getElementById('email-text');
  elItemsTbody    = document.getElementById('items-tbody');
  elResPieces     = document.getElementById('res-pieces');
  elResWeight     = document.getElementById('res-weight');
  elResVolume     = document.getElementById('res-volume');
  elResVolWeight  = document.getElementById('res-vol-weight');
  elResChargeable = document.getElementById('res-chargeable');
  elWeightWarning = document.getElementById('weight-warning');
  elWarningText   = document.getElementById('warning-text');
  elCopyFeedback  = document.getElementById('copy-feedback');

  document.getElementById('btn-parse').addEventListener('click', onParse);
  document.getElementById('btn-clear').addEventListener('click', onClear);
  document.getElementById('btn-add-row').addEventListener('click', onAddRow);
  document.getElementById('btn-copy').addEventListener('click', onCopy);
  document.getElementById('btn-restart').addEventListener('click', onRestart);
}

// ── Parser ─────────────────────────────────────────────────────

/**
 * Parse raw email text into an array of items.
 * Each item: { qty: number, l: number, b: number, h: number, weight: number|null }
 *
 * Handles formats:
 *   1x 145*90*100 cm
 *   2x 85*38*40 cm
 *   120 x 80 x 156cm / 181 kg
 *   80 x 60 x 55cm / 70 kg
 *   1x 120x82x70cm
 *   Gw 285 kg           -> assigns weight to previous item (if it has none)
 *   summary lines       -> skipped
 */
function parseEmailText(text) {
  var lines = text.split(/\r?\n/);
  var result = [];

  // Dimension pattern: three numbers separated by x / * / x
  var reDim = /(\d+(?:\.\d+)?)\s*[x*x]\s*(\d+(?:\.\d+)?)\s*[x*x]\s*(\d+(?:\.\d+)?)/i;

  // Weight on same line (after "/" preferred, then bare)
  var reWeightAfterSlash = /\/\s*(\d+(?:\.\d+)?)\s*kg\b/i;
  var reWeightBare       = /\b(\d+(?:\.\d+)?)\s*kg\b/i;

  // Weight-only line (Gw, GW, Brutto, Bruttogewicht, etc.)
  var reWeightOnlyLine = /^\s*(?:gw|gew\.?|brutto(?:gewicht)?|total\s+brutto)[\s:]*(\d+(?:\.\d+)?)\s*kg\b/i;

  // Summary line keywords - if present AND no dim pattern -> skip
  var reSummaryKeywords = /\b(?:total|gesamt|pieces|pcs|cbm|cll|brutto)\b/i;

  for (var i = 0; i < lines.length; i++) {
    var raw = lines[i];
    var line = raw.trim();
    if (!line) continue;

    // ── Weight-only line (e.g. "Gw 285 kg") ──────────────────
    var mWeightOnly = line.match(reWeightOnlyLine);
    if (mWeightOnly && !reDim.test(line)) {
      var gw = parseFloat(mWeightOnly[1]);
      // Assign to previous item that has no weight
      for (var j = result.length - 1; j >= 0; j--) {
        if (result[j].weight === null) {
          result[j].weight = gw;
          break;
        }
      }
      continue;
    }

    // ── Summary / garbage line -> skip ────────────────────────
    if (reSummaryKeywords.test(line) && !reDim.test(line)) {
      continue;
    }

    // ── Step 1: try to strip explicit qty prefix like "2x " or "2 X " ─
    var qty = 1;
    var dimLine = line;
    var mQtyPrefix = line.match(/^(\d+)\s*[xX]\s+/); // requires whitespace after the x
    if (mQtyPrefix) {
      qty = parseInt(mQtyPrefix[1], 10);
      dimLine = line.substring(mQtyPrefix[0].length);
    }

    // ── Step 2: tight format NxAxBxC (no space) only if no prefix found ─
    var mTight = null;
    if (!mQtyPrefix) {
      mTight = line.match(/^(\d+)[xX](\d+(?:\.\d+)?)[xX](\d+(?:\.\d+)?)[xX](\d+(?:\.\d+)?)/i);
      if (mTight) {
        qty = parseInt(mTight[1], 10);
      }
    }

    // ── Step 3: find dimensions in dimLine (or use tight groups) ─────────
    var dimMatch = mTight
      ? [null, mTight[2], mTight[3], mTight[4]]
      : dimLine.match(reDim);
    if (!dimMatch) continue; // No dimensions -> skip

    var l = parseFloat(dimMatch[1]);
    var b = parseFloat(dimMatch[2]);
    var h = parseFloat(dimMatch[3]);

    // ── Weight on same line ───────────────────────────────────
    var weight = null;
    var mWSlash = line.match(reWeightAfterSlash);
    if (mWSlash) {
      weight = parseFloat(mWSlash[1]);
    } else {
      // bare kg - but make sure it's not one of the dimension numbers
      // by checking it appears after the dimension block in the original line
      var dimStr = mTight ? mTight[0] : (mQtyPrefix ? mQtyPrefix[0] : '') + dimMatch[0];
      var dimIndex = line.indexOf(dimStr);
      var afterDim = line.substring(dimIndex + dimStr.length);
      var mWBare = afterDim.match(reWeightBare);
      if (mWBare) {
        weight = parseFloat(mWBare[1]);
      }
    }

    result.push({ qty: qty, l: l, b: b, h: h, weight: weight });
  }

  return result;
}

// ── Calculations ───────────────────────────────────────────────

function calcTotals(rows) {
  var totalPieces    = 0;
  var totalVol_cm3   = 0;
  var totalWeight    = 0;
  var missingWeight  = 0;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var qty    = Number(r.qty)    || 0;
    var l      = Number(r.l)      || 0;
    var b      = Number(r.b)      || 0;
    var h      = Number(r.h)      || 0;
    var weight = (r.weight !== null && r.weight !== '' && !isNaN(Number(r.weight)))
                 ? Number(r.weight) : null;

    totalPieces  += qty;
    totalVol_cm3 += l * b * h * qty;

    if (weight !== null) {
      totalWeight += weight * qty;
    } else {
      missingWeight += qty;
    }
  }

  var totalVolume_m3   = totalVol_cm3 / 1000000;
  var volumeWeight_kg  = totalVol_cm3 / 6000;
  var chargeableWeight = Math.max(
    missingWeight === rows.length ? 0 : totalWeight,
    volumeWeight_kg
  );

  return {
    totalPieces:     totalPieces,
    totalWeight:     totalWeight,
    missingWeight:   missingWeight,
    totalVolume_m3:  totalVolume_m3,
    volumeWeight_kg: volumeWeight_kg,
    chargeableWeight: chargeableWeight,
    hasAllWeights:   missingWeight === 0,
    hasAnyWeight:    missingWeight < rows.length,
  };
}

// ── Render table ───────────────────────────────────────────────

function renderTable() {
  // Clear existing rows safely
  while (elItemsTbody.firstChild) {
    elItemsTbody.removeChild(elItemsTbody.firstChild);
  }

  for (var i = 0; i < items.length; i++) {
    (function (idx) {
      var item = items[idx];
      var tr = document.createElement('tr');
      tr.dataset.idx = idx;

      // Row number cell
      var tdNum = document.createElement('td');
      tdNum.textContent = String(idx + 1);
      tr.appendChild(tdNum);

      // Input cells: qty, l, b, h, weight
      var fields = ['qty', 'l', 'b', 'h', 'weight'];
      fields.forEach(function (field) {
        var td = document.createElement('td');
        var inp = document.createElement('input');
        inp.type = 'number';
        inp.min = '0';
        inp.step = field === 'weight' ? '0.01' : '1';
        inp.placeholder = field === 'weight' ? '' : '0';
        if (item[field] !== null && item[field] !== undefined && item[field] !== '') {
          inp.value = item[field];
        }
        inp.addEventListener('input', function () {
          var val = inp.value === '' ? null : parseFloat(inp.value);
          items[idx][field] = val;
          updateResults();
        });
        td.appendChild(inp);
        tr.appendChild(td);
      });

      // Delete button
      var tdDel = document.createElement('td');
      var btnDel = document.createElement('button');
      btnDel.className = 'btn-del-row';
      btnDel.title = 'Zeile löschen';
      btnDel.textContent = '×'; // x character
      btnDel.addEventListener('click', function () {
        items.splice(idx, 1);
        renderTable();
        updateResults();
      });
      tdDel.appendChild(btnDel);
      tr.appendChild(tdDel);

      elItemsTbody.appendChild(tr);
    })(i);
  }
}

// ── Update results display ─────────────────────────────────────

function updateResults() {
  var t = calcTotals(items);

  elResPieces.textContent = t.totalPieces + ' Stk.';

  if (t.hasAnyWeight) {
    elResWeight.textContent = fmtDE(t.totalWeight, 2) + ' kg';
  } else {
    elResWeight.textContent = '- kg';
  }

  if (!t.hasAllWeights && items.length > 0) {
    var missing = t.missingWeight;
    elWarningText.textContent = '⚠ ' + missing + (missing === 1 ? ' Packstück' : ' Packstücke') + ' ohne Gewicht';
    elWeightWarning.classList.remove('hidden');
  } else {
    elWeightWarning.classList.add('hidden');
  }

  elResVolume.textContent    = fmtDE(t.totalVolume_m3, 3) + ' m³';
  elResVolWeight.textContent = fmtDE(t.volumeWeight_kg, 2) + ' kg';

  if (t.hasAnyWeight || t.volumeWeight_kg > 0) {
    var prefix = (!t.hasAllWeights && t.hasAnyWeight) ? '≥ ' : '';
    elResChargeable.textContent = prefix + fmtDE(t.chargeableWeight, 2) + ' kg';
  } else {
    elResChargeable.textContent = '-';
  }
}

// ── Number formatting helpers ──────────────────────────────────

function fmtDE(num, decimals) {
  return num.toLocaleString('de-DE', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function fmtEN(num, decimals) {
  return num.toFixed(decimals);
}

function padLeft(str, width) {
  str = String(str);
  while (str.length < width) str = ' ' + str;
  return str;
}

// ── Event handlers ─────────────────────────────────────────────

function onParse() {
  var text = elEmailText.value;
  if (!text.trim()) return;

  items = parseEmailText(text);

  if (items.length === 0) {
    elEmailText.style.borderColor = '#d83b01';
    elEmailText.placeholder = 'Keine Dimensionen gefunden. Bitte Text prüfen.';
    return;
  }

  elEmailText.style.borderColor = '';
  showPhase('results');
  renderTable();
  updateResults();
}

function onClear() {
  elEmailText.value = '';
  elEmailText.style.borderColor = '';
  elEmailText.placeholder = 'Dimensionen im Format: 1x 120*80*100 oder 120 x 80 x 100 / 50 kg';
  elEmailText.focus();
}

function onAddRow() {
  items.push({ qty: 1, l: null, b: null, h: null, weight: null });
  renderTable();
  updateResults();
  // Focus first input of new row
  var rows = elItemsTbody.querySelectorAll('tr');
  if (rows.length > 0) {
    var lastRow = rows[rows.length - 1];
    var firstInput = lastRow.querySelector('input');
    if (firstInput) firstInput.focus();
  }
}

function onRestart() {
  items = [];
  elEmailText.value = '';
  elEmailText.style.borderColor = '';
  elEmailText.placeholder = 'Dimensionen im Format: 1x 120*80*100 oder 120 x 80 x 100 / 50 kg';
  showPhase('input');
  elEmailText.focus();
}

function onCopy() {
  var text = buildClipboardText();
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(function () {
      showCopyFeedback();
    }).catch(function () {
      fallbackCopy(text);
    });
  } else {
    fallbackCopy(text);
  }
}

function fallbackCopy(text) {
  var ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand('copy');
    showCopyFeedback();
  } catch (e) {
    alert('Kopieren fehlgeschlagen. Bitte manuell kopieren:\n\n' + text);
  }
  document.body.removeChild(ta);
}

function showCopyFeedback() {
  elCopyFeedback.classList.remove('hidden');
  setTimeout(function () {
    elCopyFeedback.classList.add('hidden');
  }, 2500);
}

// ── Clipboard text builder ─────────────────────────────────────

function buildClipboardText() {
  var t = calcTotals(items);
  var today = new Date();
  var dd = String(today.getDate()).padStart(2, '0');
  var mm = String(today.getMonth() + 1).padStart(2, '0');
  var yyyy = today.getFullYear();
  var dateStr = dd + '.' + mm + '.' + yyyy;

  var sep = '─'.repeat(50);
  var lines = [];

  lines.push('LUFTFRACHT SENDUNG – ' + dateStr);
  lines.push(sep);
  lines.push('Pos. | Anz. | L(cm) | B(cm) | H(cm) | Gew.(kg)');

  for (var i = 0; i < items.length; i++) {
    var r = items[i];
    var qty    = Number(r.qty)  || 0;
    var l      = Number(r.l)    || 0;
    var b      = Number(r.b)    || 0;
    var h      = Number(r.h)    || 0;
    var wStr   = (r.weight !== null && r.weight !== '' && !isNaN(Number(r.weight)))
                 ? fmtEN(Number(r.weight), 2)
                 : '-';

    var pos  = padLeft(i + 1, 4);
    var qtyS = padLeft(qty, 4);
    var lS   = padLeft(l, 6);
    var bS   = padLeft(b, 6);
    var hS   = padLeft(h, 6);
    var wS   = padLeft(wStr, 9);

    lines.push(pos + ' |' + qtyS + ' |' + lS + ' |' + bS + ' |' + hS + ' |' + wS);
  }

  lines.push(sep);
  lines.push(padLeft('Packstücke gesamt:', 28)      + padLeft(t.totalPieces + ' Stk.', 12));
  lines.push(padLeft('Gesamtgewicht:', 28)           + padLeft(fmtEN(t.totalWeight, 2) + ' kg', 12));
  lines.push(padLeft('Gesamtvolumen:', 28)           + padLeft(fmtEN(t.totalVolume_m3, 3) + ' m³', 12));
  lines.push(padLeft('Volumengewicht (÷6000):', 28)  + padLeft(fmtEN(t.volumeWeight_kg, 2) + ' kg', 12));
  lines.push(padLeft('CHARGEABLE WEIGHT:', 28)       + padLeft(fmtEN(t.chargeableWeight, 2) + ' kg', 12));

  return lines.join('\n');
}

// ── Phase switching ────────────────────────────────────────────

function showPhase(phase) {
  if (phase === 'input') {
    elPhaseInput.classList.remove('hidden');
    elPhaseResults.classList.add('hidden');
  } else {
    elPhaseInput.classList.add('hidden');
    elPhaseResults.classList.remove('hidden');
  }
}
