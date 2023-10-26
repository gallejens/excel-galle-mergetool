// Enkele instellingen die je kan aanpassen
const SETTINGS = {
  WORKSHEET_NAME: 'Afplakbandjes TEST', // naam die het nieuwe werkblad krijgt
  OVERMAAT: 50, // overmaat in mm (aan 1 kant, dus wordt 2x toegevoegd in berekening)
  VERLIES: 0.3, // verlies (0.3 = 30%)
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
};

const COLUMNS = {
  length: 0,
  width: 1,
  material: 3,
  side_a: 6,
  side_b: 7,
  side_c: 8,
  side_d: 9,
};

type ColumnKey = keyof typeof COLUMNS;

const SIDE_TO_INCREASE_COLUMN: [ColumnKey, ColumnKey][] = [
  ['side_a', 'length'],
  ['side_b', 'width'],
  ['side_c', 'length'],
  ['side_d', 'width'],
];

const TEMPLATE_DATA: ExcelCell[][] = [
  [
    'Materiaal',
    'Afplakmateriaal',
    'NETTO Lengte',
    'Bruto lengte incl. verlies',
  ],
];

// Types
type ExcelCell = string | number | boolean;

//@ts-ignore
function main(workbook: ExcelScript.Workbook) {
  const dataWorksheet = workbook.getActiveWorksheet();
  const dataRange = dataWorksheet.getUsedRange();
  const values: ExcelCell[][] = [
    ...dataRange.getValues().map((r: ExcelCell[]) => [...r.map(i => i)]),
  ];

  // Remove existing worksheet with name
  const existingWorksheet = workbook.getWorksheet(SETTINGS.WORKSHEET_NAME);
  if (existingWorksheet) {
    existingWorksheet.delete();
  }

  // k: plaatmateriaal, v: (k: afplakmateriaal, v: lengte)
  const grouppedValues = new Map<string, Map<string, number>>();

  // calculate total length for each material and side
  for (const rowValues of values) {
    const material = String(rowValues[COLUMNS.material]);
    const lengthsForMaterial =
      grouppedValues.get(material) ?? new Map<string, number>();

    for (const [side, increaseColumn] of SIDE_TO_INCREASE_COLUMN) {
      const sideMaterial = String(rowValues[COLUMNS[side]]);
      if (shouldIgnoreMaterial(sideMaterial)) continue;

      const increase = Number(rowValues[COLUMNS[increaseColumn]]);
      if (Number.isNaN(increase)) {
        throw new Error(`${increaseColumn} is not a number`);
      }

      const existingLength = lengthsForMaterial.get(sideMaterial) ?? 0;
      lengthsForMaterial.set(sideMaterial, existingLength + increase);
    }

    grouppedValues.set(material, lengthsForMaterial);
  }

  // Copy template data
  const data = [...TEMPLATE_DATA.map(r => [...r])];

  for (const [material, lengths] of Array.from(grouppedValues)) {
    for (const [sideMaterial, length] of Array.from(lengths)) {
      data.push([material, sideMaterial, length, 0]);
    }
  }

  const worksheet = workbook.addWorksheet();
  try {
    worksheet.setName(SETTINGS.WORKSHEET_NAME);
  } catch (e) {
    console.log(`Failed to set worksheet for afplakbandjes`);
  }

  const columnCount = Math.max(...data.map(r => r.length));
  const range = worksheet.getRangeByIndexes(0, 0, data.length, columnCount);
  range.setValues(data);

  range.getFormat().autofitColumns();
}

function shouldIgnoreMaterial(material: string) {
  const lowercaseMaterial = material.toLocaleLowerCase();
  for (const req of SETTINGS.IGNORE) {
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
