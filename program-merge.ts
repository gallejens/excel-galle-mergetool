// Enkele instellingen die je kan aanpassen
const SETTINGS = {
	// Het character dat wordt geplaatst tussen labels. BV: Boven/Onder
	CONCAT_CHAR: '/',
	// Het character dat wordt geplaatst tussen labels die als gelijk worden behandeld. BV: Zij L|R
	LABEL_MERGE_CONCAT_CHAR: '|',
	// Characters waarbij labels als gelijk worden behandeld. BV: Zij L & Zij R -> Zij L|R
	LABEL_MERGE_CHARACTERS: [
		['L', 'R'],
		['O', 'B'],
		['V', 'A'],
	],
	// Vervangend materiaallabel indien er geen materiaal gedefineerd is
	NO_MATERIAL_LABEL: 'Onbekend Materiaal',
};

// Types
type ExcelRow = string | number | boolean;
type MergedRow = ExcelRow | string[];

// Do not touch
const COLUMNS = {
	length: {idx: 0, unique: true},
	width: {idx: 1, unique: true},
	amount: {idx: 2, unique: false},
	material: {idx: 3, unique: true},
	rotation: {idx: 4, unique: true},
	label: {idx: 5, unique: false},
} satisfies Record<string, {idx: number; unique: boolean}>;

const uniqueColumns = Object.values(COLUMNS)
	.filter((c) => c.unique)
	.map((c) => c.idx);

//@ts-ignore
function main(workbook: ExcelScript.Workbook) {
	const dataSheet = workbook.getActiveWorksheet();
	const usedRange = dataSheet.getUsedRange();
	const values = [...usedRange.getValues().map((r) => [...r.map((i) => i)])];
	const columnCount = usedRange.getColumnCount();

	const mergedRows: MergedRow[][] = [];

	// Merge data and populate new rows
	for (let y = 0; y < values.length; y++) {
		const row = values[y];

		// find idx of row in newrows array that has same unique fields
		const existingRowIdx = mergedRows.findIndex((r) => uniqueColumns.every((j) => row[j] === r[j]));

		// if no row found what matches the unique columns then just add row
		if (existingRowIdx === -1) {
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

	const finalRows: ExcelRow[][] = [];

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
		finalRows.push(finalRow as ExcelRow[]);
	}

	const groupedByMaterial = groupByMaterial(finalRows);

	for (const [material, rows] of Array.from(groupedByMaterial)) {
		const materialWorksheet = workbook.addWorksheet();
		const worksheetName = material.slice(0, 30) || SETTINGS.NO_MATERIAL_LABEL;
		materialWorksheet.setName(worksheetName);
		const range = materialWorksheet.getRangeByIndexes(0, 0, rows.length, columnCount);
		range.setValues(rows as ExcelRow[][]);
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

	const newPrefix = mergeByLastCharacter(prefixes);
	const newName = mergeByLastCharacter(names);

	return `${newPrefix}_${newName}`;
}

function mergeByLastCharacter(stringsSet: Set<string>) {
	const stringsArr = Array.from(stringsSet);
	for (let i = 0; i < stringsArr.length; i++) {
		const str = stringsArr[i];

		const splitStr = str.split(' ');
		const merger = splitStr.pop();
		const identifier = splitStr.join(' ');
		if (!merger) continue;

		const mergeChars = SETTINGS.LABEL_MERGE_CHARACTERS.find((c) => c.includes(merger));
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

function groupByMaterial(rows: ExcelRow[][]) {
	// loop through new rows and sort by
	const groupedByMaterial: Map<string, typeof rows> = new Map();
	for (const row of rows) {
		const material = String(row[COLUMNS.material.idx]);
		const materialRows = groupedByMaterial.get(material) ?? [];
		groupedByMaterial.set(material, [...materialRows, row]);
	}
	return groupedByMaterial;
}
