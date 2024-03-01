// Instelling of je de labels wil samenvoegen. true -> normale werking, false -> enkel de rijen opsplitsen in aparte werkbladen
// Bij true, wordt het afplakbandjesprogramma ook uitgevoerd
const MERGE_LABELS = true;

// Enkele instellingen die je kan aanpassen voor het mergeprogramma
const MERGE_SETTINGS = {
  // Het character dat wordt geplaatst tussen labels. BV: Boven/Onder
  CONCAT_CHAR: '/',
  // Het character dat wordt geplaatst tussen labels die als gelijk worden behandeld. BV: Zij L/R
  LABEL_MERGE_CONCAT_CHAR: '/',
  // Characters waarbij labels als gelijk worden behandeld. BV: Zij L & Zij R -> Zij L/R
  // De volgorde hoe ze hier gedefineerd staan zal ook gereflecteerd worden in de uiteindelijke labels. BV L zal altijd voor R staan (nooit R|L)
  LABEL_MERGE_STRINGS: [
    ['L', 'R'],
    ['B', 'O'],
    ['V', 'A'],
    ['links', 'rechts'],
  ],
  // Vervangend materiaallabel indien er geen materiaal gedefineerd is
  NO_MATERIAL_LABEL: 'Onbekend Materiaal',
  // Het character dat wordt geplaatst tussen het originele label en de unieke ID (indien labels niet gemerged worden)
  CHAR_BEFORE_UNIQUE_ID: ' ',
  // Maximale breedte van alle kolommen samen (kolom F wordt automatisch aangepast met resterende breedte)
  MAX_COLUMNS_WIDTH: 450,
  // Nieuwe naam voor originale data werkblad
  DATA_WORKSHEET_NAME: 'Utilized Sheets',
};

// Enkele instellingen die je kan aanpassen voor het afplakbandjesprogramma
const AFPLAKBANDJES_SETTINGS = {
  OVERMAAT: 50, // overmaat in mm (aan 1 kant, dus wordt 2x toegevoegd in berekening)
  VERLIES: 0.3, // verlies (0.3 = 30%)
  // Welke afplakband materialen gefilterd worden. 'equals' betekend volledig overeenkomstig, 'contains' checkt of materiaal het woord bevat
  IGNORE: [
    {
      equals: 'Niet afplakken',
    },
    {
      contains: 'Verstek',
    },
    {
      contains: 'Overmaat',
    },
  ],
  HEADERS: ['Materiaal', 'NETTO Lengte', 'Bruto lengte incl. verlies'],
};

// Types
type ExcelCell = string | number | boolean;
type MergedCell = ExcelCell | string[];
type Length = { netto: number; bruto: number };
type ColumnKey = keyof typeof COLUMNS;

const COLUMNS = {
  length: { idx: 0, unique: true },
  width: { idx: 1, unique: true },
  amount: { idx: 2, unique: false },
  material: { idx: 3, unique: true },
  rotation: { idx: 4, unique: true },
  label: { idx: 5, unique: false },
  side_a: { idx: 6, unique: false },
  side_b: { idx: 7, unique: false },
  side_c: { idx: 8, unique: false },
  side_d: { idx: 9, unique: false },
  id: { idx: 17, unique: false },
};

const FORMATTING_AUTOFIT_COLUMNS = ['D', 'G', 'H', 'I', 'J', 'K'];
const FORMATTING_CENTER_COLUMNS = ['C', 'E'];

const SIDE_TO_INCREASE_COLUMN: [ColumnKey, ColumnKey][] = [
  ['side_a', 'length'],
  ['side_b', 'width'],
  ['side_c', 'length'],
  ['side_d', 'width'],
];

//@ts-ignore
function main(workbook: ExcelScript.Workbook) {
  const dataWorksheet = workbook.getActiveWorksheet();
  const dataWorksheetId = dataWorksheet.getId();
  dataWorksheet.setName(MERGE_SETTINGS.DATA_WORKSHEET_NAME);

  // Delete all worksheets apart from active
  const sheets = workbook.getWorksheets();
  for (const sheet of sheets) {
    if (sheet.getId() === dataWorksheetId) continue;
    sheet.delete();
  }

  const usedRange = dataWorksheet.getUsedRange();
  const values: ExcelCell[][] = [
    ...usedRange.getValues().map((r: ExcelCell[]) => [...r.map(i => i)]),
  ];
  const columnCount = usedRange.getColumnCount();

  // Always calculate afplakbandjes data, even if we dont use
  const afplakbandjesData = getAfplakbandjesData(values);

  const uniqueColumns = Object.values(COLUMNS)
    .filter(c => c.unique)
    .map(c => c.idx);

  // Merge data and populate new rows
  const mergedRows: MergedCell[][] = [];
  for (const row of values) {
    // find idx of row in already mergedRows that has same unique fields
    const existingRowIdx = mergedRows.findIndex(r =>
      uniqueColumns.every(j => row[j] === r[j])
    );

    // if no row found that matches the unique columns then just add row
    if (existingRowIdx === -1 || !MERGE_LABELS) {
      const newRow: MergedCell[] = [...row];
      newRow[COLUMNS.label.idx] = [String(newRow[COLUMNS.label.idx])];
      mergedRows.push(newRow);
      continue;
    }

    const existingRow = mergedRows[existingRowIdx];

    // increase amount
    existingRow[COLUMNS.amount.idx] =
      Number(existingRow[COLUMNS.amount.idx]) + 1;

    // create new label
    const existingLabels = existingRow[COLUMNS.label.idx];
    if (Array.isArray(existingLabels)) {
      existingLabels.push(String(row[COLUMNS.label.idx]));
    } else {
      throw new Error('Merged labels was not array');
    }
  }

  // loop through newrows to create labels
  const finalRows: ExcelCell[][] = [];
  for (const row of mergedRows) {
    let labels = row[COLUMNS.label.idx];
    if (!Array.isArray(labels)) {
      throw new Error('Merged labels are not array');
    }

    labels = concatLabels(labels);

    if (!MERGE_LABELS) {
      labels = `${labels}${MERGE_SETTINGS.CHAR_BEFORE_UNIQUE_ID}${
        row[COLUMNS.id.idx]
        }`;
    }

    const finalRow = [...row];
    finalRow[COLUMNS.label.idx] = labels;

    if (typeof finalRow[COLUMNS.label.idx] !== 'string') {
      throw new Error(`Label ${labels} was not a string after concat`);
    }

    // if statement above catches errors so we can cast
    finalRows.push(finalRow as ExcelCell[]);
  }

  const groupedByMaterial = groupByMaterial(finalRows);

  // Excel doesnt allow iterating of mapss
  for (const [material, rows] of Array.from(groupedByMaterial)) {
    const materialWorksheet = workbook.addWorksheet();
    const worksheetName =
      material.slice(0, 30) || MERGE_SETTINGS.NO_MATERIAL_LABEL;

    try {
      materialWorksheet.setName(worksheetName);
    } catch (e) {
      console.log(`Failed to set worksheetname: ${worksheetName}`);
    }

    const finalColumnCount = MERGE_LABELS ? 6 : columnCount;
    const sortedRows = sortRowsByLabel(rows);
    const range = materialWorksheet.getRangeByIndexes(
      0,
      0,
      sortedRows.length,
      finalColumnCount
    );
    range.setValues(sortedRows.map(r => r.slice(0, finalColumnCount)));

    formatWorksheet(materialWorksheet);

    if (MERGE_LABELS) {
      insertAfplakbandjesData(materialWorksheet, afplakbandjesData, material, sortedRows.length)
    }
  }
}

//@ts-ignore
const insertAfplakbandjesData = (worksheet: ExcelScript.Worksheet, afplakbandjesData: ReturnType<typeof getAfplakbandjesData>, material: string, startIndex: number) => {
  const lenghtsPerMaterial = afplakbandjesData.get(material);
  if (!lenghtsPerMaterial) return;

  // Fill cells
  const afplakbandjesCells: ExcelCell[][] = [AFPLAKBANDJES_SETTINGS.HEADERS];
  for (const [sideMaterial, lenghts] of Array.from(lenghtsPerMaterial)) {
    const brutoWithLoss = lenghts.bruto * (1 + AFPLAKBANDJES_SETTINGS.VERLIES);
    afplakbandjesCells.push([
      sideMaterial,
      `${millimeterToMeter(lenghts.netto)}m`,
      `${millimeterToMeter(brutoWithLoss)}m`
    ])
  }

  // Transform cells so we put them in specific columns
  const transformedCells: ExcelCell[][] = [];
  const TRANSFORM_COLUMNS: Record<number, number> = {
    0: 0,
    3: 1,
    5: 2
  }
  const highestTransformColumnId = Math.max(...Object.keys(TRANSFORM_COLUMNS).map(x => Number(x)))
  for (const row of afplakbandjesCells) {
    const transformedRow: ExcelCell[] = [];
    for (let i = 0; i <= highestTransformColumnId; i++) {
      const cellId = TRANSFORM_COLUMNS[i];
      if (cellId === undefined) {
        transformedRow.push('')
      } else {
        transformedRow.push(row[cellId])
      }
    }
    transformedCells.push(transformedRow)
  }

  // place cells in worksheet
  const columnCount = Math.max(...transformedCells.map(r => r.length));
  const range = worksheet.getRangeByIndexes(
    startIndex + 1,
    0,
    transformedCells.length,
    columnCount
  );
  range.setValues(transformedCells);

  // Place borders around cells
  const BORDERS_TO_COLOR = [
    //@ts-ignore
    ExcelScript.BorderIndex.edgeTop,
    //@ts-ignore
    ExcelScript.BorderIndex.edgeBottom,
    //@ts-ignore
    ExcelScript.BorderIndex.edgeLeft,
    //@ts-ignore
    ExcelScript.BorderIndex.edgeRight,
  ]

  const rangeFormat = range.getFormat()
  const rangeBorders = rangeFormat.getBorders();
  for (const rangeBorder of rangeBorders) {
    if (!BORDERS_TO_COLOR.includes(rangeBorder.getSideIndex())) continue
      //@ts-ignore
    rangeBorder.setStyle(ExcelScript.BorderLineStyle.continuous);
  }

  // align to right
  const alignRightRange = worksheet.getRangeByIndexes(
    startIndex + 1,
    1,
    transformedCells.length,
    columnCount - 1
  );
  const alignRightRangeFormat = alignRightRange.getFormat();
  //@ts-ignore
  alignRightRangeFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
}

const getAfplakbandjesData = (values: ExcelCell[][]) => {
  // k: plaatmateriaal, k: (k: afplakmateriaal, v: lengte)
  const lengthsForMaterialPerMaterial = new Map<string, Map<string, Length>>();

  // calculate total length for each material and side
  for (const rowValues of values) {
    for (const [side, increaseColumn] of SIDE_TO_INCREASE_COLUMN) {
      const sideMaterial = String(rowValues[COLUMNS[side].idx]);
      if (shouldIgnoreMaterialForSide(sideMaterial)) continue;

      const increase =
        Number(rowValues[COLUMNS[increaseColumn].idx]) *
        Number(rowValues[COLUMNS.amount.idx]);
      if (Number.isNaN(increase)) {
        throw new Error(`${increaseColumn} is not a number`);
      }

      const mainMaterial = String(rowValues[COLUMNS.material.idx]);
      let lengthsForMaterial = lengthsForMaterialPerMaterial.get(mainMaterial);
      if (!lengthsForMaterial) {
        lengthsForMaterial = new Map();
        lengthsForMaterialPerMaterial.set(mainMaterial, lengthsForMaterial);
      }

      const existingLengths = lengthsForMaterial.get(sideMaterial);
      lengthsForMaterial.set(sideMaterial, {
        netto: (existingLengths?.netto ?? 0) + increase,
        bruto:
          (existingLengths?.bruto ?? 0) +
          increase +
          2 * AFPLAKBANDJES_SETTINGS.OVERMAAT,
      });
    }
  }

  return lengthsForMaterialPerMaterial;
};

//@ts-ignore
function formatWorksheet(worksheet: ExcelScript.Worksheet) {
  const fullRange = worksheet.getUsedRange();
  const fullRangeFormat = fullRange.getFormat();
  //@ts-ignore
  fullRangeFormat.setVerticalAlignment(ExcelScript.VerticalAlignment.top);

  if (MERGE_LABELS) {
    const fullRangeBorder = fullRangeFormat.getRangeBorder(
      //@ts-ignore
      ExcelScript.BorderIndex.insideHorizontal
    );
    //@ts-ignore
    fullRangeBorder.setStyle(ExcelScript.BorderLineStyle.continuous);
  }

  for (const column of FORMATTING_AUTOFIT_COLUMNS) {
    const range = worksheet.getRange(`${column}:${column}`);
    const rangeFormat = range.getFormat();
    rangeFormat.autofitColumns();
  }

  // let column F to remainder of max width
  let widthRemaining = MERGE_SETTINGS.MAX_COLUMNS_WIDTH;
  for (const column of ['A', 'B', 'C', 'D', 'E']) {
    const range = worksheet.getRange(`${column}:${column}`);
    const columnWidth = range.getFormat().getColumnWidth();
    widthRemaining -= columnWidth;
  }
  const fRange = worksheet.getRange('f:f');
  const fRangeFormat = fRange.getFormat();
  fRangeFormat.setColumnWidth(widthRemaining);
  fRangeFormat.setWrapText(true);

  for (const column of FORMATTING_CENTER_COLUMNS) {
    const range = worksheet.getRange(`${column}:${column}`);
    const rangeFormat = range.getFormat();
    //@ts-ignore
    rangeFormat.setHorizontalAlignment(
      //@ts-ignore
      ExcelScript.HorizontalAlignment.center
    );
  }
}

function concatLabels(labels: string[]) {
  const prefixes = new Set<string>();
  const names = new Set<string>();

  for (const label of labels) {
    const splitLabel = label.split('.');
    const splitLabelLength = splitLabel.length;

    let prefix = '';
    let name: string;

    if (splitLabelLength === 1) {
      name = splitLabel[0];
    } else {
      prefix = splitLabel[0];
      name = splitLabel[splitLabelLength - 1];
    }

    if (splitLabelLength > 2) {
      const extras = new Set(splitLabel.slice(1, splitLabelLength - 1));

      for (const extra of Array.from(extras)) {
        const ignoreExtra = /^[0-9]{4}$/g.test(extra);
        if (ignoreExtra) {
          extras.delete(extra);
        }
      }

      if (extras.size > 0) {
        name = `${name} ${Array.from(extras).join(' ')}`;
      }
    }

    // remove unique id from name
    name = name.replace(/[0-9]*_/g, '');

    prefixes.add(prefix);
    names.add(name);
  }

  const newPrefix = mergeByLabelEnding(prefixes);
  const newName = mergeByLabelEnding(names);

  return `${newPrefix}_${newName}`;
}

function mergeByLabelEnding(stringsSet: Set<string>) {
  const stringsArr = Array.from(stringsSet);
  for (let i = 0; i < stringsArr.length; i++) {
    const str = stringsArr[i];

    const splitStr = str.split(' ');
    const merger = splitStr.pop();
    if (!merger) continue;

    const identifier = splitStr.join(' ');

    const mergeChars = MERGE_SETTINGS.LABEL_MERGE_STRINGS.find(c =>
      c.includes(merger)
    );
    if (!mergeChars) continue;

    const companion = mergeChars[mergeChars[0] === merger ? 1 : 0];

    let foundCompanionIdx: number | undefined;
    // only search next names
    for (let j = i + 1; j < stringsArr.length; j++) {
      const searchStr = stringsArr[j];
      const searchSplitStr = searchStr.split(' ');
      const searchMerger = searchSplitStr.pop();
      const searchIdentifier = searchSplitStr.join(' ');

      if (searchIdentifier === identifier && searchMerger === companion) {
        foundCompanionIdx = j;
        break;
      }
    }
    if (!foundCompanionIdx) continue;

    const sortedMergeChars = [merger, companion].sort(
      (a, b) => mergeChars.indexOf(a) - mergeChars.indexOf(b)
    );

    stringsArr.splice(foundCompanionIdx, 1);
    stringsArr[i] = `${identifier} ${sortedMergeChars.join(
      MERGE_SETTINGS.LABEL_MERGE_CONCAT_CHAR
    )}`;
  }

  stringsArr.sort((a, b) => {
    const aNum = +a;
    const bNum = +b;
    if (Number.isNaN(aNum)) return 1;
    if (Number.isNaN(bNum)) return -1;
    return aNum - bNum;
  });

  return stringsArr.join(MERGE_SETTINGS.CONCAT_CHAR);
}

function groupByMaterial(rows: ExcelCell[][]) {
  // loop through new rows and sort by
  const groupedByMaterial: Map<string, typeof rows> = new Map();
  for (const row of rows) {
    const material = String(row[COLUMNS.material.idx]);
    const materialRows = groupedByMaterial.get(material) ?? [];
    groupedByMaterial.set(material, [...materialRows, row]);
  }
  return groupedByMaterial;
}

function sortRowsByLabel(rows: ExcelCell[][]) {
  return rows.sort((aRow, bRow) => {
    const aNum = parseInt(aRow[COLUMNS.label.idx].toString());
    const bNum = parseInt(bRow[COLUMNS.label.idx].toString());

    if (Number.isNaN(aNum)) return 1;
    if (Number.isNaN(bNum)) return -1;

    return aNum - bNum;
  });
}

function shouldIgnoreMaterialForSide(material: string) {
  const lowercaseMaterial = material.toLocaleLowerCase();
  for (const req of AFPLAKBANDJES_SETTINGS.IGNORE) {
    if (req.equals) {
      if (lowercaseMaterial === req.equals.toLocaleLowerCase()) return true;
    } else if (req.contains) {
      if (lowercaseMaterial.includes(req.contains.toLocaleLowerCase()))
        return true;
    } else {
      throw new Error(
        `Unknown ignore requirement: ${Object.keys(req).join(', ')}`
      );
    }
  }
  return false;
}

function millimeterToMeter(millimeters: number) {
  const meters = millimeters / 1000;
  return Math.round(meters * 10) / 10;
}
