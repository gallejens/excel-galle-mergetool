// Enkele instellingen die je kan aanpassen
const SETTINGS = {
	// Het character dat wordt geplaatst tussen labels. BV: Boven/Onder
	CONCAT_CHAR: '/',
	// Het character dat wordt geplaatst tussen labels die als gelijk worden behandeld. BV: Zij L|R
	LABEL_MERGE_CONCAT_CHAR: '|',
	// Characters waarbij labels als gelijk worden behandeld. BV: Zij L & Zij R -> Zij L|R
	// De volgorde hoe ze hier gedefineerd staan zal ook gereflecteerd worden in de uiteindelijke labels. BV L zal altijd voor R staan (nooit R|L)
	LABEL_MERGE_STRINGS: [
		['L', 'R'],
		['O', 'B'],
		['V', 'A'],
		['links', 'rechts'],
	],
	// Vervangend materiaallabel indien er geen materiaal gedefineerd is
	NO_MATERIAL_LABEL: 'Onbekend Materiaal',
	// Instelling of je de labels wil samenvoegen. true -> normale werking, false -> enkel de rijen opsplitsen in aparte werkbladen
	MERGE_LABELS: true,
	// Het character dat wordt geplaatst tussen het originele label en de unieke ID
	CHAR_BEFORE_UNIQUE_ID: ' ',
};

const COLUMNS = {
	length: {idx: 0, unique: true},
	width: {idx: 1, unique: true},
	amount: {idx: 2, unique: false},
	material: {idx: 3, unique: true},
	rotation: {idx: 4, unique: true},
	label: {idx: 5, unique: false},
	id: {idx: 17, unique: false},
};

const AUTOFIT_COLUMNS = ['D', 'F', 'G', 'H', 'I', 'J', 'K'];
const CENTER_COLUMNS = ['C', 'E'];

const uniqueColumns = Object.values(COLUMNS)
	.filter((c) => c.unique)
	.map((c) => c.idx);

// Types
type ExcelCell = string | number | boolean;
type MergedCell = ExcelCell | string[];

//@ts-ignore
function main(workbook: ExcelScript.Workbook) {
	const dataWorksheet = workbook.getActiveWorksheet();
	const dataWorksheetId = dataWorksheet.getId();

	// Delete all worksheets apart from active
	const sheets = workbook.getWorksheets();
	for (const sheet of sheets) {
		if (sheet.getId() === dataWorksheetId) continue;
		sheet.delete();
	}

	const usedRange = dataWorksheet.getUsedRange();
	const values: ExcelCell[][] = [...usedRange.getValues().map((r) => [...r.map((i) => i)])];
	const columnCount = usedRange.getColumnCount();

	// formatWorksheet(dataWorksheet)

	const mergedRows: MergedCell[][] = [];

	// Merge data and populate new rows
	for (let y = 0; y < values.length; y++) {
		const row = values[y];

		// find idx of row in newrows array that has same unique fields
		const existingRowIdx = mergedRows.findIndex((r) => uniqueColumns.every((j) => row[j] === r[j]));

		// if no row found what matches the unique columns then just add row
		if (existingRowIdx === -1 || !SETTINGS.MERGE_LABELS) {
			mergedRows.push([...row]);
			continue;
		}

		const existingRow = mergedRows[existingRowIdx];

		// increase amount
		existingRow[COLUMNS.amount.idx] = Number(existingRow[COLUMNS.amount.idx]) + 1;

		// create new label
		const existingLabels = existingRow[COLUMNS.label.idx];
		if (Array.isArray(existingLabels)) {
			existingLabels.push(String(row[COLUMNS.label.idx]));
		} else {
			existingRow[COLUMNS.label.idx] = [String(existingLabels), String(row[COLUMNS.label.idx])];
		}
	}

	const finalRows: ExcelCell[][] = [];

	// loop through newrows to create labels
	for (const row of mergedRows) {
		let labels = row[COLUMNS.label.idx];
		if (Array.isArray(labels)) {
			labels = concatLabels(labels);
		} else {
			if (!SETTINGS.MERGE_LABELS) {
				labels = `${labels}${SETTINGS.CHAR_BEFORE_UNIQUE_ID}${row[COLUMNS.id.idx]}`;
			}
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

	// Excel doesnt allow iterating of maps
	for (const [material, rows] of Array.from(groupedByMaterial)) {
		const materialWorksheet = workbook.addWorksheet();
		const worksheetName = material.slice(0, 30) || SETTINGS.NO_MATERIAL_LABEL;

		try {
			materialWorksheet.setName(worksheetName);
		} catch (e) {
			console.log(`Failed to set worksheetname: ${worksheetName}`);
		}

		const sortedRows = sortRowsByLabel(rows);
		const range = materialWorksheet.getRangeByIndexes(0, 0, sortedRows.length, columnCount);
		range.setValues(sortedRows);

		formatWorksheet(materialWorksheet);
	}
}

function concatLabels(labels: string[]) {
	const prefixes = new Set<string>();
	const names = new Set<string>();

	for (const label of labels) {
		const splitLabel = label.split('_');
		if (splitLabel.length !== 2) {
			throw new Error(`Label ${label} had none or more than one underscore`);
		}

		const [prefix, name] = splitLabel;
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

		const mergeChars = SETTINGS.LABEL_MERGE_STRINGS.find((c) => c.includes(merger));
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

		const sortedMergeChars = [merger, companion].sort((a, b) => mergeChars.indexOf(a) - mergeChars.indexOf(b));

		stringsArr.splice(foundCompanionIdx, 1);
		stringsArr[i] = `${identifier} ${sortedMergeChars.join(SETTINGS.LABEL_MERGE_CONCAT_CHAR)}`;
	}

	stringsArr.sort((a, b) => {
		const aNum = +a;
		const bNum = +b;
		if (Number.isNaN(aNum)) return 1;
		if (Number.isNaN(bNum)) return -1;
		return aNum - bNum;
	});

	return stringsArr.join(SETTINGS.CONCAT_CHAR);
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

//@ts-ignore
function formatWorksheet(worksheet: ExcelScript.Worksheet) {
	for (const column of AUTOFIT_COLUMNS) {
		const range = worksheet.getRange(`${column}:${column}`);
		const rangeFormat = range.getFormat();
		rangeFormat.autofitColumns();
	}

	for (const column of CENTER_COLUMNS) {
		const range = worksheet.getRange(`${column}:${column}`);
		const rangeFormat = range.getFormat();
		//@ts-ignore
		rangeFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	}
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
