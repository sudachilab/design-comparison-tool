import { readExcel, getSheet } from './excelReader.js';
import { renderGrid } from './gridRenderer.js';
import { highlightMatches } from './comparison.js';

const $ = (id) => document.getElementById(id);
const documents = [
  {
    id: 'basic',
    label: '基本設計書',
    fileId: 'basicFile',
    sheetId: 'basicSheet',
    gridId: 'basicGrid',
  },
  {
    id: 'external',
    label: '外部設計書',
    fileId: 'externalFile',
    sheetId: 'externalSheet',
    gridId: 'externalGrid',
  },
  {
    id: 'detailed',
    label: '詳細設計書',
    fileId: 'detailedFile',
    sheetId: 'detailedSheet',
    gridId: 'detailedGrid',
  },
  {
    id: 'internal',
    label: '内部設計書',
    fileId: 'internalFile',
    sheetId: 'internalSheet',
    gridId: 'internalGrid',
  },
  {
    id: 'requirements',
    label: '要件定義書',
    fileId: 'requirementsFile',
    sheetId: 'requirementsSheet',
    gridId: 'requirementsGrid',
  },
];

const state = {
  workbooks: Object.fromEntries(documents.map(({ id }) => [id, null])),
  activeOverviewNumber: null,
  zoom: {
    basic: 100,
    internal: 100,
  },
};

const status = (message) => {
  $('status').textContent = message;
};

async function upload(documentConfig) {
  const file = $(documentConfig.fileId).files[0];
  if (!file) return;

  status(`${file.name} を読み込み中…`);

  try {
    const workbook = await readExcel(file);
    state.workbooks[documentConfig.id] = workbook;

    const select = $(documentConfig.sheetId);
    select.replaceChildren();

    for (const name of workbook.SheetNames) {
      const option = window.document.createElement('option');
      option.value = name;
      option.textContent = name;
      select.append(option);
    }

    showSheet(documentConfig);
    status(`${file.name} を読み込みました。`);
  } catch (err) {
    state.workbooks[documentConfig.id] = null;
    $(documentConfig.sheetId).replaceChildren();
    $(documentConfig.gridId).replaceChildren();
    status(`読み込みエラー: ${err.message}`);
    console.error(`「${documentConfig.label}」の読み込みに失敗しました。`, err);
  }
}

function showSheet(document) {
  const book = state.workbooks[document.id];
  if (!book) return;

  const grid = $(document.gridId);
  renderGrid(grid, getSheet(book, $(document.sheetId).value));
  applyZoom(document.id);

  const query = $('search').value.trim();
  if (query) highlightMatches([grid], query);
  if (state.activeOverviewNumber) {
    highlightOverviewNumber(state.activeOverviewNumber, false);
  }
}

function applyZoom(documentId) {
  const zoom = state.zoom[documentId];
  if (zoom === undefined) return;

  const input = $(`${documentId}Zoom`);
  const wrapper = $(`${documentId}Grid`).querySelector('.grid-zoom-wrapper');
  const table = wrapper?.querySelector('.excel-grid');
  const scale = zoom / 100;

  input.value = zoom;
  $(`${documentId}ZoomValue`).value = `${zoom}%`;
  $(`${documentId}ZoomValue`).textContent = `${zoom}%`;

  if (!wrapper || !table) return;

  table.style.setProperty('--grid-scale', scale);
}

for (const documentId of ['basic', 'internal']) {
  const controls = document.querySelector(
    `[data-zoom-controls="${documentId}"]`,
  );
  const input = $(`${documentId}Zoom`);

  input.addEventListener('input', () => {
    state.zoom[documentId] = Number(input.value);
    applyZoom(documentId);
  });

  controls.addEventListener('click', (event) => {
    const stepButton = event.target.closest('[data-zoom-step]');
    if (stepButton) {
      state.zoom[documentId] = Math.min(
        200,
        Math.max(50, state.zoom[documentId] + Number(stepButton.dataset.zoomStep)),
      );
      applyZoom(documentId);
      return;
    }

    if (event.target.closest('[data-zoom-reset]')) {
      state.zoom[documentId] = 100;
      applyZoom(documentId);
    }
  });
}

function readOverviewNumbers(cell) {
  return (cell.dataset.overviewNumbers || '').split(' ').filter(Boolean);
}

function highlightOverviewNumber(overviewNumber, scrollToMatches = true) {
  state.activeOverviewNumber = overviewNumber;

  for (const document of documents) {
    const grid = $(document.gridId);
    const isRequirements = document.id === 'requirements';
    let firstMatch = null;

    for (const cell of grid.querySelectorAll('.overview-reference')) {
      const matches =
        !isRequirements &&
        readOverviewNumbers(cell).includes(overviewNumber);

      cell.classList.toggle('linked-match', matches);
      if (matches && !firstMatch) firstMatch = cell;
    }

    if (scrollToMatches && firstMatch && document.id !== 'basic') {
      firstMatch.scrollIntoView({
        behavior: 'smooth',
        block: 'center',
        inline: 'center',
      });
    }
  }

  status(
    `概要番号「${overviewNumber}」を選択中です。要件定義書は連動対象外です。`,
  );
}

for (const document of documents) {
  $(document.fileId).addEventListener('change', () => upload(document));
  $(document.sheetId).addEventListener('change', () => showSheet(document));
}

$('basicGrid').addEventListener('click', (event) => {
  const cell = event.target.closest('td.overview-reference');
  if (!cell) return;

  const [overviewNumber] = readOverviewNumbers(cell);
  if (overviewNumber) highlightOverviewNumber(overviewNumber);
});

$('searchBtn').addEventListener('click', () => {
  const term = $('search').value.trim();
  const count = highlightMatches(
    documents.map(({ gridId }) => $(gridId)),
    term,
  );

  status(
    term
      ? `「${term}」を含むセル: ${count}件`
      : '検索語を入力してください。',
  );
});

$('search').addEventListener('keydown', (event) => {
  if (event.key === 'Enter') $('searchBtn').click();
});

$('resetBtn').addEventListener('click', () => {
  for (const document of documents) {
    $(document.fileId).value = '';
    $(document.sheetId).replaceChildren();
    $(document.gridId).innerHTML =
      '<p class="empty">ファイル未選択</p>';
    state.workbooks[document.id] = null;
  }

  state.activeOverviewNumber = null;
  $('search').value = '';
  status('クリアしました。');
});
