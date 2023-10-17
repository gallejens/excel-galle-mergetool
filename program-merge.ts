// Enkele instellingen die je kan aanpassen
const SETTINGS = {
	// Het character dat wordt geplaatst tussen labels. BV: Boven/Onder
	CONCAT_CHAR: '/',
	// Het character dat wordt geplaatst tussen labels die als gelijk worden behandeld. BV: Zij L|R
	LABEL_MERGE_CONCAT_CHAR: '|',
	// Characters waarbij labels als gelijk worden behandeld. BV: Zij L & Zij R -> Zij L|R
	LABEL_MERGE_STRINGS: [
		['L', 'R'],
		['O', 'B'],
		['V', 'A'],
		['links', 'rechts'],
	],
	// Vervangend materiaallabel indien er geen materiaal gedefineerd is
	NO_MATERIAL_LABEL: 'Onbekend Materiaal',
	MERGE_LABELS: true,
};

// Types
type ExcelCell = string | number | boolean;
type MergedCell = ExcelCell | string[];

// Do not touch
const COLUMNS = {
	length: {idx: 0, unique: true, center: false},
	width: {idx: 1, unique: true, center: false},
	amount: {idx: 2, unique: false, center: true},
	material: {idx: 3, unique: true, center: false},
	rotation: {idx: 4, unique: true, center: true},
	label: {idx: 5, unique: false, center: false},
};

const uniqueColumns = Object.values(COLUMNS)
	.filter((c) => c.unique)
	.map((c) => c.idx);

//@ts-ignore
function main(workbook: ExcelScript.Workbook) {
	const dataWorksheet = workbook.getActiveWorksheet();
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
		materialWorksheet.setName(worksheetName);

		const range = materialWorksheet.getRangeByIndexes(0, 0, rows.length, columnCount);
		range.setValues(rows as ExcelCell[][]);

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

		const companionChar = mergeChars[mergeChars[0] === merger ? 1 : 0];

		let foundCompanionIdx: number | undefined;
		// only search next names
		for (let j = i + 1; j < stringsArr.length; j++) {
			const searchStr = stringsArr[j];
			const searchSplitStr = searchStr.split(' ');
			const searchMerger = searchSplitStr.pop();
			const searchIdentifier = searchSplitStr.join(' ');

			if (searchIdentifier === identifier && searchMerger === companionChar) {
				foundCompanionIdx = j;
				break;
			}
		}
		if (!foundCompanionIdx) continue;

		stringsArr.splice(foundCompanionIdx, 1);
		stringsArr[i] = `${str}${SETTINGS.LABEL_MERGE_CONCAT_CHAR}${companionChar}`;
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
	const fullRange = worksheet.getUsedRange();
	const fullRangeFormat = fullRange.getFormat();
	fullRangeFormat.autofitColumns();

	// We increase with for every columns a bit
	const rowCount = fullRange.getRowCount();
	const startColumn = fullRange.getColumnIndex();
	const endColumn = fullRange.getColumnCount() + startColumn;
	for (let i = startColumn; i < endColumn; i++) {
		const columnRange = worksheet.getRangeByIndexes(0, i, rowCount, 1);
		const columnRangeFormat = columnRange.getFormat();
		const columnWidth = columnRange.getWidth();
		columnRangeFormat.setColumnWidth(columnWidth + 5);

		if (Object.values(COLUMNS).find((c) => c.idx === i)?.center) {
			//@ts-ignore
			columnRangeFormat.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
		}
	}
}
