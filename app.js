// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  STICKER LABEL GENERATOR  â€“  app.js
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

// â”€â”€ STATE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
let stickerList   = [];
let selectedIndex = -1;
let zoom          = 2.0;

// â”€â”€ DOM â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const styleInput    = document.getElementById('styleNameInput');
const colourInput   = document.getElementById('colourInput');
const sizeInput     = document.getElementById('sizeInput');
const prevStyle     = document.getElementById('prev-style');
const prevColour    = document.getElementById('prev-colour');
const prevSize      = document.getElementById('prev-size');
const liveSticker   = document.getElementById('liveSticker');
const stickerListEl = document.getElementById('stickerList');
const countBadge    = document.getElementById('stickerCountBadge');
const zoomLabel     = document.getElementById('zoomLabel');
const sheetSection  = document.getElementById('sheetSection');
const printSheetEl  = document.getElementById('printSheetPreview');
const sheetInfo     = document.getElementById('sheetInfo');
const printArea     = document.getElementById('printArea');

applyZoom();

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  EXCEL PARSING  â€“  handles Childish London format
//  Multiple sections in same sheet, each with its own header row
//  Header cols: S.NO | STYLE NO | COLOR | FABRIC | PLACEMENT PRINT or AOP PRINT | 0-3m | 3-6m | 6-12m ...
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

/**
 * Clean colour: strip Pantone/TCX codes
 * e.g. "TRAVERTINE - Pantone 15-1114 TCX"  -->  "TRAVERTINE"
 *      "OATMEAL - Pantone13-0401 TCX"       -->  "OATMEAL"
 *      "WHITE"                              -->  "WHITE"
 */
function cleanColour(raw) {
  if (!raw) return '';
  return String(raw)
    .replace(/-?\s*Pantone\s*[\d\-]+\s*TCX\b/gi, '')
    .replace(/\bTCX\b/gi, '')
    .replace(/Pantone\s*[\d\-]+/gi, '')
    .replace(/\s*-\s*$/, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

/**
 * Clean size label: normalise case, e.g. "0-3m" -> "0-3M", "3y" -> "3Y"
 */
function cleanSize(hdr) {
  return String(hdr).trim()
    .replace(/\bm\b/g, 'M')   // lowercase m â†’ M
    .replace(/\by\b/g, 'Y');  // lowercase y â†’ Y
}

/**
 * Returns true if a value in a size cell means "has quantity"
 */
function hasQty(val) {
  if (val === null || val === undefined || val === '') return false;
  const n = Number(val);
  return !isNaN(n) && n > 0;
}

/**
 * Is this row a data header row? (contains PLACEMENT PRINT or AOP PRINT)
 */
function isHeaderRow(row) {
  return row.some(c => {
    const s = String(c || '').toUpperCase().trim();
    return s === 'PLACEMENT PRINT' || s === 'AOP PRINT';
  });
}

/**
 * Is this a data row? First cell should be a number (S.NO)
 */
function isDataRow(row) {
  const first = String(row[0] || '').trim();
  return /^\d+$/.test(first);
}

/**
 * Parse the Excel ArrayBuffer. Returns [{style, colour, size}]
 */
function parseExcel(arrayBuffer) {
  const wb   = XLSX.read(arrayBuffer, { type: 'array' });
  const ws   = wb.Sheets[wb.SheetNames[0]];   // Sheet1
  // Get as 2D array (no header processing â€” raw)
  const raw  = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const results = [];
  let currentHeader = null;   // { styleCol, colourCol, printCol, sizeCols:[{idx,label}] }

  for (let r = 0; r < raw.length; r++) {
    const row = raw[r];

    // â”€â”€ Detect header row â”€â”€
    if (isHeaderRow(row)) {
      currentHeader = buildHeaderMap(row);
      continue;
    }

    // â”€â”€ Process data row â”€â”€
    if (currentHeader && isDataRow(row)) {
      const styleName = String(row[currentHeader.printCol] || '').trim();
      const colour    = cleanColour(row[currentHeader.colourCol]);
      if (!styleName) continue;

      currentHeader.sizeCols.forEach(sc => {
        let qty = Number(row[sc.idx]) || 0;
        if (qty > 0) {
          // Odd qty â†’ round up to even (21â†’22, 35â†’36)
          if (qty % 2 !== 0) qty += 1;
          for (let q = 0; q < qty; q++) {
            results.push({
              style:  styleName.toUpperCase(),
              colour: colour || '-',
              size:   sc.label
            });
          }
        }
      });
    }
  }

  return results;
}

/**
 * Build column index map from a header row array
 */
function buildHeaderMap(row) {
  let colourCol = -1;
  let printCol  = -1;
  const sizeCols = [];

  // Size keywords to detect size columns
  const SIZE_RE = /^(\d+[-\/]\d+[mMyY]?|\d+[mMyY])$/i;

  row.forEach((cell, idx) => {
    const s = String(cell || '').trim().toUpperCase();
    if (s === 'COLOR' || s === 'COLOUR')             colourCol = idx;
    if (s === 'PLACEMENT PRINT' || s === 'AOP PRINT') printCol  = idx;

    // Size column: MUST end with M (months) or Y (years)
    // e.g. 0-3m âœ“  3-6m âœ“  6-12m âœ“  3Y âœ“  0-6 âœ— (total column, skip)
    const raw = String(cell || '').trim();
    if (raw && /^[\d]/.test(raw)) {
      if (/^\d+-\d+[mM]$/.test(raw) || /^\d+[mM]$/.test(raw) ||
          /^\d+-\d+[yY]$/.test(raw) || /^\d+[yY]$/.test(raw)) {
        sizeCols.push({ idx, label: cleanSize(raw) });
      }
    }
  });

  return { colourCol, printCol, sizeCols };
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  FILE IMPORT UI
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const excelInput    = document.getElementById('excelFileInput');
const importDrop    = document.getElementById('importDrop');
const dropClickArea = document.getElementById('dropClickArea');
const importResult  = document.getElementById('importResult');
const irFilename    = document.getElementById('irFilename');
const irCount       = document.getElementById('irCount');

dropClickArea.addEventListener('click', () => excelInput.click());
importDrop.addEventListener('click', e => { if (e.target !== dropClickArea) excelInput.click(); });
importDrop.addEventListener('dragover', e => { e.preventDefault(); importDrop.classList.add('drag-over'); });
importDrop.addEventListener('dragleave', ()  => importDrop.classList.remove('drag-over'));
importDrop.addEventListener('drop', e => {
  e.preventDefault(); importDrop.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});
excelInput.addEventListener('change', () => { if (excelInput.files[0]) handleFile(excelInput.files[0]); });

document.getElementById('clearImportBtn').addEventListener('click', () => {
  importResult.style.display = 'none';
  importDrop.style.display   = '';
  excelInput.value = '';
  stickerList = []; selectedIndex = -1;
  renderList(); updateSheetPreview();
});

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const stickers = parseExcel(e.target.result);
      if (!stickers.length) {
        alert('No sticker data found in this file.\n\nExpected columns:\nâ€¢ PLACEMENT PRINT or AOP PRINT\nâ€¢ COLOR\nâ€¢ Size columns (0-3m, 3-6m, 6-12m ...)');
        return;
      }
      stickerList   = stickers;
      selectedIndex = 0;
      loadIntoEditor(stickerList[0]);
      renderList();
      updateSheetPreview();
      importDrop.style.display    = 'none';
      importResult.style.display  = 'flex';
      irFilename.textContent      = file.name;
      irCount.textContent         = stickers.length + ' sticker' + (stickers.length !== 1 ? 's' : '') + ' generated (1 per size)';
      flashBadge();
    } catch(err) {
      console.error(err);
      alert('Error reading Excel: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  LIVE PREVIEW
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function updateLivePreview() {
  prevStyle.textContent  = (styleInput.value  || 'Style Name').toUpperCase();
  prevColour.textContent =  colourInput.value  || 'Colour';
  prevSize.textContent   =  sizeInput.value    || 'Size';
}
[styleInput, colourInput, sizeInput].forEach(el => el.addEventListener('input', updateLivePreview));
document.getElementById('updateStickerBtn').addEventListener('click', updateLivePreview);

document.querySelectorAll('.chip').forEach(chip => {
  chip.addEventListener('click', () => {
    document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
    chip.classList.add('active');
    sizeInput.value = chip.dataset.val;
    updateLivePreview();
  });
});

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  QUEUE
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
document.getElementById('addToListBtn').addEventListener('click', () => {
  stickerList.push({
    style:  (styleInput.value.trim()  || 'Style Name').toUpperCase(),
    colour:  colourInput.value.trim() || 'Colour',
    size:    sizeInput.value.trim()   || 'Size'
  });
  renderList(); updateSheetPreview(); flashBadge();
});

document.getElementById('addStickerBtn').addEventListener('click', () => {
  styleInput.focus();
  styleInput.scrollIntoView({ behavior: 'smooth' });
});

function renderList() {
  countBadge.textContent = stickerList.length + ' sticker' + (stickerList.length !== 1 ? 's' : '');
  if (!stickerList.length) {
    stickerListEl.innerHTML = '<div class="empty-list">No stickers yet.<br/>Import Excel or add manually.</div>';
    return;
  }
  stickerListEl.innerHTML = '';
  stickerList.forEach((s, i) => {
    const item = document.createElement('div');
    item.className = 'queue-item' + (i === selectedIndex ? ' selected' : '');
    item.innerHTML = `
      <div class="queue-item-info">
        <div class="qi-style">${s.style}</div>
        <div class="qi-meta">${s.colour} &nbsp;Â·&nbsp; ${s.size}</div>
      </div>
      <button class="qi-delete" title="Delete">âœ•</button>`;
    item.addEventListener('click', e => {
      if (e.target.classList.contains('qi-delete')) return;
      selectedIndex = i; loadIntoEditor(s); renderList();
    });
    item.querySelector('.qi-delete').addEventListener('click', () => {
      stickerList.splice(i, 1);
      if (selectedIndex >= stickerList.length) selectedIndex = stickerList.length - 1;
      renderList(); updateSheetPreview();
    });
    stickerListEl.appendChild(item);
  });
}

function loadIntoEditor(s) {
  styleInput.value  = s.style;
  colourInput.value = s.colour;
  sizeInput.value   = s.size;
  updateLivePreview();
}

document.getElementById('clearAllBtn').addEventListener('click', () => {
  if (!stickerList.length || !confirm('Clear all stickers?')) return;
  stickerList = []; selectedIndex = -1;
  renderList(); updateSheetPreview();
});

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  SHEET PREVIEW
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function updateSheetPreview() {
  if (!stickerList.length) { sheetSection.style.display = 'none'; return; }
  sheetSection.style.display = '';
  sheetInfo.textContent = stickerList.length + ' label' + (stickerList.length > 1 ? 's' : '');
  printSheetEl.innerHTML = '';
  stickerList.forEach(s => printSheetEl.appendChild(buildStickerEl(s, true)));
}

function buildStickerEl(s, small) {
  const div = document.createElement('div');
  div.className = 'sticker';
  if (small) {
    div.style.transform       = 'scale(0.65)';
    div.style.transformOrigin = 'top left';
    div.style.marginRight     = '-72px';
    div.style.marginBottom    = '-40px';
  }
  div.innerHTML = `
    <div class="sticker-field style-name">${s.style}</div>
    <div class="sticker-field">${s.colour}</div>
    <div class="sticker-field">${s.size}</div>`;
  return div;
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  ZOOM
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function applyZoom() {
  liveSticker.style.transform = `scale(${zoom})`;
  liveSticker.style.margin    = `${(zoom-1)*56}px ${(zoom-1)*104}px`;
  zoomLabel.textContent       = Math.round(zoom * 100) + '%';
}
document.getElementById('zoomInBtn').addEventListener('click',  () => { zoom = Math.min(zoom+.25, 4);   applyZoom(); });
document.getElementById('zoomOutBtn').addEventListener('click', () => { zoom = Math.max(zoom-.25, .5); applyZoom(); });

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  PRINT
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
document.getElementById('printAllBtn').addEventListener('click', () => {
  const list = stickerList.length ? stickerList : currentPreview();
  printArea.innerHTML = '';
  list.forEach(s => {
    const el = document.createElement('div');
    el.className = 'sticker';
    el.innerHTML = `
      <div class="sticker-field style-name">${s.style}</div>
      <div class="sticker-field">${s.colour}</div>
      <div class="sticker-field">${s.size}</div>`;
    printArea.appendChild(el);
  });
  window.print();
});

function currentPreview() {
  return [{
    style:  (styleInput.value.trim() || 'Style Name').toUpperCase(),
    colour:  colourInput.value.trim() || 'Colour',
    size:    sizeInput.value.trim()   || 'Size'
  }];
}

// ══════════════════════════════════════════════════════
//  PDF DOWNLOAD
//  Page 1  = Measurement ref (50mm x 30mm)
//  Page 1  = Measurement ref (100mm x 30mm landscape)
//  Page 2+ = LEFT sticker | RIGHT sticker (100mm x 30mm landscape)
// ══════════════════════════════════════════════════════
document.getElementById('downloadPdfBtn').addEventListener('click', () => {
  const toExport = stickerList.length ? stickerList : currentPreview();
  const btn = document.getElementById('downloadPdfBtn');
  btn.disabled = true; btn.textContent = 'Generating...';
  try {
    const { jsPDF } = window.jspdf;
    const SW = 50, SH = 30;   // single sticker size
    const PW = 100, PH = 30;  // page = 2 stickers side by side (landscape)

    // Draw one sticker at xOffset (0=left, 50=right)
    function drawSticker(pdf, s, xOff) {
      // Border: 48mm x 28mm (1mm inset from 50x30 sticker edge)
      pdf.setDrawColor(0); pdf.setLineWidth(0.2);
      pdf.rect(xOff + 1, 1, 48, 28);
      // Text
      pdf.setFont('helvetica', 'bold'); pdf.setFontSize(10); pdf.setTextColor(0);
      pdf.text(s.style.toUpperCase(), xOff + 3, 10);
      pdf.setFont('helvetica', 'normal'); pdf.setFontSize(10);
      pdf.text(s.colour, xOff + 3, 18);
      pdf.text(s.size,   xOff + 3, 26);
    }

    // PAGE 1: Measurement Reference (landscape 100x30)
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: [PH, PW] });
    pdf.setFont('helvetica', 'bold'); pdf.setFontSize(5); pdf.setTextColor(60);
    pdf.text('STICKER LABEL - SIZE REFERENCE  |  50mm x 30mm', PW / 2, 3, { align: 'center' });
    pdf.setFontSize(3.5); pdf.setFont('helvetica','normal');
    pdf.text('Bio-degradable  |  Arial 16pt  |  Border: 48x28mm', PW/2, 6, {align:'center'});

    // Ref sticker box (left half)
    const bx=8, by=8, bw=34, bh=18;
    pdf.setDrawColor(0); pdf.setLineWidth(0.35); pdf.rect(bx,by,bw,bh);
    pdf.setFont('helvetica','bold'); pdf.setFontSize(4.5); pdf.setTextColor(0);
    pdf.text('STYLE NAME', bx+2, by+5);
    pdf.setFont('helvetica','normal'); pdf.setFontSize(4);
    pdf.text('Colour', bx+2, by+9); pdf.text('Size', bx+2, by+14);

    // Width arrow
    const aY=by-2.5; pdf.setLineWidth(0.2);
    pdf.line(bx,aY,bx+bw/2-4,aY); pdf.line(bx+bw/2+4,aY,bx+bw,aY);
    pdf.line(bx,aY,bx+1,aY-0.7); pdf.line(bx,aY,bx+1,aY+0.7);
    pdf.line(bx+bw,aY,bx+bw-1,aY-0.7); pdf.line(bx+bw,aY,bx+bw-1,aY+0.7);
    pdf.setFontSize(3.5); pdf.setFont('helvetica','bold');
    pdf.text('50 mm', bx+bw/2, aY+0.7, {align:'center'});

    // Height arrow
    const aX=bx-3;
    pdf.line(aX,by,aX,by+bh/2-2); pdf.line(aX,by+bh/2+2,aX,by+bh);
    pdf.line(aX,by,aX-0.7,by+1); pdf.line(aX,by,aX+0.7,by+1);
    pdf.line(aX,by+bh,aX-0.7,by+bh-1); pdf.line(aX,by+bh,aX+0.7,by+bh-1);
    pdf.setFontSize(3.5); pdf.text('30 mm', aX, by+bh/2+0.7, {angle:90, align:'center'});

    // Specs (right half)
    const tx = SW + 4;
    pdf.setFont('helvetica','bold'); pdf.setFontSize(4.5); pdf.setTextColor(0);
    pdf.text('Specifications', tx, 9);
    pdf.setFont('helvetica','normal'); pdf.setFontSize(3.8); pdf.setTextColor(50);
    ['Width   : 50 mm','Height  : 30 mm','Border  : 48 x 28 mm',
     'Font    : Arial 16pt','Material: Bio-degradable']
      .forEach((l,i) => pdf.text(l, tx, 13+i*3.5));

    // PAGES 2+: LEFT + RIGHT consecutive stickers (landscape)
    for (let i = 0; i < toExport.length; i += 2) {
      pdf.addPage([PH, PW], 'landscape');
      drawSticker(pdf, toExport[i], 0);                            // LEFT
      if (toExport[i + 1]) drawSticker(pdf, toExport[i + 1], SW); // RIGHT
      // Light divider between left and right
      pdf.setDrawColor(210); pdf.setLineWidth(0.1);
      pdf.line(SW, 1, SW, SH - 1);
    }

    const name = toExport[0].style.replace(/\s+/g,'_').substring(0,20);
    const cnt  = toExport.length > 1 ? '_x'+toExport.length : '';
    pdf.save('sticker_'+name+cnt+'.pdf');
    btn.textContent = 'Downloaded!';
    setTimeout(() => { btn.disabled=false; btn.innerHTML='Download PDF'; }, 2000);
  } catch(err) {
    console.error(err); alert('PDF error: '+err.message);
    btn.disabled=false; btn.innerHTML='Download PDF';
  }
});


// â”€â”€ BADGE FLASH â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function flashBadge() {
  countBadge.style.background = 'rgba(34,208,122,.35)';
  countBadge.style.color = '#22d07a';
  setTimeout(() => { countBadge.style.background=''; countBadge.style.color=''; }, 700);
}