import {
  extractOverviewNumbers,
  getOverviewColor,
} from './overviewNumbers.js';

const borderStyles = {
  thin: { width: 1, style: 'solid' },
  medium: { width: 2, style: 'solid' },
  thick: { width: 3, style: 'solid' },
  hair: { width: 1, style: 'solid' },
  dashed: { width: 1, style: 'dashed' },
  mediumDashed: { width: 2, style: 'dashed' },
  dotted: { width: 1, style: 'dotted' },
  dashDot: { width: 1, style: 'dashed' },
  mediumDashDot: { width: 2, style: 'dashed' },
  dashDotDot: { width: 1, style: 'dashed' },
  mediumDashDotDot: { width: 2, style: 'dashed' },
  slantDashDot: { width: 1, style: 'dashed' },
  double: { width: 3, style: 'double' },
};

function readBorder(cell, side) {
  const border = cell?.s?.border?.[side];
  if (!border?.style || border.style === 'none') return null;

  if (!Object.prototype.hasOwnProperty.call(borderStyles, border.style)) {
    return null;
  }
  const appearance = borderStyles[border.style];
  const result = { ...appearance };
  const rgb = border.color?.rgb;
  if (typeof rgb === 'string' && /^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(rgb)) {
    result.color = `#${rgb.slice(-6)}`;
  } else if (!border.color || border.color.auto || border.color.indexed === 64) {
    result.color = '#000000';
  }

  return result;
}

function getCellBorders(sheet, XLSX, row, col, merge) {
  if (!merge) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
    return Object.fromEntries(
      ['top', 'right', 'bottom', 'left']
        .map((side) => [side, readBorder(cell, side)])
        .filter(([, border]) => border),
    );
  }

  const { startRow, endRow, startCol, endCol } = merge;
  const borders = {};
  const edges = {
    top: Array.from(
      { length: endCol - startCol + 1 },
      (_, i) => [startRow, startCol + i, 'top'],
    ),
    right: Array.from(
      { length: endRow - startRow + 1 },
      (_, i) => [startRow + i, endCol, 'right'],
    ),
    bottom: Array.from(
      { length: endCol - startCol + 1 },
      (_, i) => [endRow, startCol + i, 'bottom'],
    ),
    left: Array.from(
      { length: endRow - startRow + 1 },
      (_, i) => [startRow + i, startCol, 'left'],
    ),
  };

  for (const [side, cells] of Object.entries(edges)) {
    for (const [cellRow, cellCol, cellSide] of cells) {
      const cell = sheet[XLSX.utils.encode_cell({ r: cellRow, c: cellCol })];
      const border = readBorder(cell, cellSide);
      if (border) {
        borders[side] = border;
        break;
      }
    }
  }

  return borders;
}

function applyCellBorders(cellElement, borders) {
  for (const [side, border] of Object.entries(borders)) {
    cellElement.style.setProperty(
      `border-${side}-width`,
      `calc(${border.width}px * var(--grid-scale))`,
    );
    cellElement.style.setProperty(`border-${side}-style`, border.style);
    if (border.color) {
      cellElement.style.setProperty(`border-${side}-color`, border.color);
    }
  }
}

/** Build a read-only HTML table, preserving blank positions and merged ranges. */
export function renderGrid(container, sheet) {
  container.replaceChildren();

  if (!sheet || !sheet['!ref']) {
    container.textContent = '表示するセルがありません。';
    return;
  }

  const XLSX = globalThis.XLSX;
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const merges = sheet['!merges'] || [];

  // Guard against very large / malformed sheets freezing the browser.
  const rows = range.e.r - range.s.r + 1;
  const cols = range.e.c - range.s.c + 1;
  if (rows * cols > 100000) {
    container.textContent = 'シートが大きすぎます（最大10万セル）。';
    return;
  }

  const mergeStart = new Map();
  const covered = new Set();

  for (const merge of merges) {
    const startRow = Math.max(merge.s.r, range.s.r);
    const endRow = Math.min(merge.e.r, range.e.r);
    const startCol = Math.max(merge.s.c, range.s.c);
    const endCol = Math.min(merge.e.c, range.e.c);

    if (startRow > endRow || startCol > endCol) continue;

    mergeStart.set(`${startRow}:${startCol}`, {
      rowspan: endRow - startRow + 1,
      colspan: endCol - startCol + 1,
      startRow,
      endRow,
      startCol,
      endCol,
    });

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        if (row !== startRow || col !== startCol) {
          covered.add(`${row}:${col}`);
        }
      }
    }
  }

  const table = document.createElement('table');
  table.className = 'excel-grid';

  const head = document.createElement('thead');
  const headRow = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'corner';
  headRow.append(corner);

  for (let col = range.s.c; col <= range.e.c; col++) {
    const header = document.createElement('th');
    header.textContent = XLSX.utils.encode_col(col);
    headRow.append(header);
  }

  head.append(headRow);
  table.append(head);

  const body = document.createElement('tbody');
  for (let row = range.s.r; row <= range.e.r; row++) {
    const tableRow = document.createElement('tr');
    const number = document.createElement('th');
    number.className = 'row-number';
    number.textContent = row + 1;
    tableRow.append(number);

    for (let col = range.s.c; col <= range.e.c; col++) {
      const key = `${row}:${col}`;
      if (covered.has(key)) continue;

      const cellElement = document.createElement('td');
      const merge = mergeStart.get(key);

      if (merge) {
        cellElement.rowSpan = merge.rowspan;
        cellElement.colSpan = merge.colspan;
      }

      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      cellElement.textContent = cell
        ? String(cell.w ?? cell.v ?? '')
        : '';
      applyCellBorders(
        cellElement,
        getCellBorders(sheet, XLSX, row, col, merge),
      );

      const overviewNumbers = extractOverviewNumbers(cellElement.textContent);
      if (overviewNumbers.length > 0) {
        cellElement.dataset.overviewNumbers = overviewNumbers.join(' ');
        cellElement.classList.add('overview-reference');
        cellElement.style.setProperty(
          '--overview-color',
          getOverviewColor(overviewNumbers[0]),
        );
      }

      tableRow.append(cellElement);
    }

    body.append(tableRow);
  }

  table.append(body);

  const zoomWrapper = document.createElement('div');
  zoomWrapper.className = 'grid-zoom-wrapper';
  zoomWrapper.append(table);
  container.append(zoomWrapper);
}
