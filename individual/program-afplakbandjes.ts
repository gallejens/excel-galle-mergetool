// Enkele instellingen die je kan aanpassen
const SETTINGS = {
  WORKSHEET_NAME: 'Afplakbandjes', // naam die het nieuwe werkblad krijgt
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
  // Of final values in meter moeten staan (zoniet staan ze in mm)
  IN_METERS: true,
};

const COLUMNS = {
  length: 0,
  width: 1,
  amount: 2,
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

// Types
type ExcelCell = string | number | boolean;
type Length = { netto: number; bruto: number };

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

  // k: afplakmateriaal, v: lengte
  const lengthsForMaterial = new Map<string, Length>();

  // calculate total length for each material and side
  for (const rowValues of values) {
    for (const [side, increaseColumn] of SIDE_TO_INCREASE_COLUMN) {
      const sideMaterial = String(rowValues[COLUMNS[side]]);
      if (shouldIgnoreMaterial(sideMaterial)) continue;

      const increase =
        Number(rowValues[COLUMNS[increaseColumn]]) *
        Number(rowValues[COLUMNS.amount]);
      if (Number.isNaN(increase)) {
        throw new Error(`${increaseColumn} is not a number`);
      }

      const existingLengths = lengthsForMaterial.get(sideMaterial);
      lengthsForMaterial.set(sideMaterial, {
        netto: (existingLengths?.netto ?? 0) + increase,
        bruto: (existingLengths?.bruto ?? 0) + increase + 2 * SETTINGS.OVERMAAT,
      });
    }
  }

  const data: ExcelCell[][] = [SETTINGS.HEADERS];
  for (const [sideMaterial, lenghts] of Array.from(lengthsForMaterial)) {
    let netto = lenghts.netto;
    let bruto = lenghts.bruto * (1 + SETTINGS.VERLIES);
    if (SETTINGS.IN_METERS) {
      netto = millimeterToMeter(netto);
      bruto = millimeterToMeter(bruto);
    }
    data.push([sideMaterial, netto, bruto]);
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

function millimeterToMeter(millimeters: number) {
  const meters = millimeters / 1000;
  return Math.round(meters * 10) / 10;
}
