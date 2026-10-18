/** XLSX is loaded from the SheetJS browser bundle. */
const spreadsheetNamespace =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const documentRelationshipsNamespace =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationshipsNamespace =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const drawingNamespace =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

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

function normalizePackagePath(path) {
  const parts = [];
  for (const part of path.replace(/^\/+/, '').split('/')) {
    if (part === '..') parts.pop();
    else if (part && part !== '.') parts.push(part);
  }
  return parts.join('/');
}

function readXmlPart(files, path) {
  const content = files[path]?.content;
  if (!content) return null;

  const xml = typeof content === 'string'
    ? content
    : new TextDecoder().decode(content);
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  if (document.documentElement.localName === 'parsererror') {
    console.warn(`罫線情報を読み込めませんでした: ${path} のXMLを解析できません。`);
    return null;
  }

  return document;
}

function getChildren(element, namespace, localName) {
  return Array.from(element?.childNodes || []).filter(
    (node) =>
      node.nodeType === 1 &&
      node.namespaceURI === namespace &&
      node.localName === localName,
  );
}

function getFirstChild(element, namespace, localName) {
  return getChildren(element, namespace, localName)[0] || null;
}

function resolveColor(color, themeColors) {
  let rgb = color?.getAttribute('rgb');

  if (!rgb && color?.hasAttribute('theme')) {
    rgb = themeColors[Number(color.getAttribute('theme'))];
  }

  if (!rgb && color?.hasAttribute('indexed')) {
    const indexed = Number(color.getAttribute('indexed'));
    if (indexed === 64) rgb = '000000';
    else if (indexed === 65) rgb = 'FFFFFF';
  }

  if (!rgb && color?.getAttribute('auto') === '1') rgb = '000000';
  if (!rgb || !/^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(rgb)) {
    return '#000000';
  }

  const channels = rgb.slice(-6).match(/.{2}/g).map((value) => parseInt(value, 16));
  const tint = Number(color?.getAttribute('tint') || 0);
  const tinted = channels.map((channel) => Math.round(
    tint < 0
      ? channel * (1 + tint)
      : channel * (1 - tint) + 255 * tint,
  ));
  return `#${tinted.map((channel) => channel.toString(16).padStart(2, '0')).join('')}`;
}

function readThemeColors(files) {
  const document = readXmlPart(files, 'xl/theme/theme1.xml');
  const colorScheme = document
    ? document.getElementsByTagNameNS(drawingNamespace, 'clrScheme')[0]
    : null;
  const names = [
    'lt1', 'dk1', 'lt2', 'dk2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink',
  ];
  const colors = [];

  for (const name of names) {
    const schemeColor = getFirstChild(colorScheme, drawingNamespace, name);
    const color = schemeColor?.firstElementChild;
    colors.push(
      color?.getAttribute('lastClr') ||
      color?.getAttribute('val') ||
      '000000',
    );
  }

  return colors;
}

function readBorderDefinitions(stylesDocument, themeColors) {
  const bordersElement = stylesDocument.getElementsByTagNameNS(
    spreadsheetNamespace,
    'borders',
  )[0];
  return getChildren(bordersElement, spreadsheetNamespace, 'border').map(
    (borderElement) => {
      const sides = {};
      for (const side of ['top', 'right', 'bottom', 'left']) {
        const sideElement = getFirstChild(
          borderElement,
          spreadsheetNamespace,
          side,
        );
        const style = sideElement?.getAttribute('style');
        if (
          !style ||
          style === 'none' ||
          !Object.prototype.hasOwnProperty.call(borderStyles, style)
        ) {
          continue;
        }

        const color = getFirstChild(
          sideElement,
          spreadsheetNamespace,
          'color',
        );
        sides[side] = {
          ...borderStyles[style],
          color: resolveColor(color, themeColors),
        };
      }
      return sides;
    },
  );
}

function getWorksheetPaths(workbook, files) {
  const workbookDocument = readXmlPart(files, 'xl/workbook.xml');
  const relationshipsDocument = readXmlPart(
    files,
    'xl/_rels/workbook.xml.rels',
  );
  if (!workbookDocument || !relationshipsDocument) return null;

  const relationships = new Map(
    Array.from(
      relationshipsDocument.getElementsByTagNameNS(
        packageRelationshipsNamespace,
        'Relationship',
      ),
      (relationship) => [
        relationship.getAttribute('Id'),
        relationship.getAttribute('Target'),
      ],
    ),
  );
  const paths = new Map();

  for (const sheetElement of workbookDocument.getElementsByTagNameNS(
    spreadsheetNamespace,
    'sheet',
  )) {
    const name = sheetElement.getAttribute('name');
    const relationshipId = sheetElement.getAttributeNS(
      documentRelationshipsNamespace,
      'id',
    );
    const target = relationships.get(relationshipId);
    if (target) {
      const packagePath = target.startsWith('/') ? target : `xl/${target}`;
      paths.set(name, normalizePackagePath(packagePath));
    }
  }

  return workbook.SheetNames.every((name) => paths.has(name))
    ? paths
    : null;
}

function addXlsxCellBorders(workbook) {
  const files = workbook.files;
  const stylesDocument = readXmlPart(files, 'xl/styles.xml');
  const worksheetPaths = getWorksheetPaths(workbook, files);
  if (!stylesDocument || !worksheetPaths) {
    console.warn('XLSXの罫線情報を特定できませんでした。');
    return;
  }

  const stylesElement = stylesDocument.getElementsByTagNameNS(
    spreadsheetNamespace,
    'cellXfs',
  )[0];
  const borderDefinitions = readBorderDefinitions(
    stylesDocument,
    readThemeColors(files),
  );
  const styleBorders = getChildren(
    stylesElement,
    spreadsheetNamespace,
    'xf',
  ).map((style) => {
    const borderId = Number(style.getAttribute('borderId') || 0);
    return borderDefinitions[borderId] || {};
  });

  for (const [sheetName, path] of worksheetPaths) {
    const sheetDocument = readXmlPart(files, path);
    const sheet = workbook.Sheets[sheetName];
    if (!sheetDocument || !sheet) {
      console.warn(`シート「${sheetName}」の罫線情報を取得できませんでした。`);
      continue;
    }

    const cellBorders = new Map();
    for (const cell of sheetDocument.getElementsByTagNameNS(
      spreadsheetNamespace,
      'c',
    )) {
      const styleId = Number(cell.getAttribute('s') || 0);
      const borders = styleBorders[styleId];
      const address = cell.getAttribute('r');
      if (address && borders && Object.keys(borders).length > 0) {
        cellBorders.set(address, borders);
      }
    }
    sheet['!cellBorders'] = cellBorders;
  }
}

export async function readExcel(file) {
  if (!globalThis.XLSX) {
    throw new Error(
      'Excel解析ライブラリを読み込めません。インターネット接続を確認してください。',
    );
  }

  const buffer = await file.arrayBuffer();
  const isXlsx = String(file.name || '').toLowerCase().endsWith('.xlsx');
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellText: true,
    cellDates: true,
    cellStyles: true,
    bookFiles: isXlsx,
  });

  if (isXlsx) addXlsxCellBorders(workbook);
  return workbook;
}

export function getSheet(workbook, name) {
  return workbook.Sheets[name];
}
