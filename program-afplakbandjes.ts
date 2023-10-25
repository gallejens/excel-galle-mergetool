// Enkele instellingen die je kan aanpassen
const SETTINGS = {
  WORKSHEET_NAME: 'Afplakbandjes TEST', // naam die het nieuwe werkblad krijgt
  OVERMAAT: 50, // overmaat in mm (aan 1 kant, dus wordt 2x toegevoegd in berekening)
  VERLIES: 0.3, // verlies (0.3 = 30%)
};

const COLUMNS = {
  length: { idx: 0, unique: true },
  width: { idx: 1, unique: true },
  amount: { idx: 2, unique: false },
  material: { idx: 3, unique: true },
  rotation: { idx: 4, unique: true },
  label: { idx: 5, unique: false },
  id: { idx: 17, unique: false },
};

const TEMPLATE_DATA: ExcelCell[][] = [
  ['Materiaal', 'NETTO Lengte', 'Bruto lengte incl. verlies'],
  ['Test materiaal', 60, 100],
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

  // Copy template data
  const data = [...TEMPLATE_DATA.map(r => [...r])];

  const worksheet = workbook.addWorksheet();
  try {
    worksheet.setName(SETTINGS.WORKSHEET_NAME);
  } catch (e) {
    console.log(`Failed to set worksheet for afplakbandjes, duplicate name?`);
  }

  const columnCount = Math.max(...data.map(r => r.length));
  const range = worksheet.getRangeByIndexes(0, 0, data.length, columnCount);
  range.setValues(data);

  range.getFormat().autofitColumns();
}
