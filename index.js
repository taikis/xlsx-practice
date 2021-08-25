const XlsxPopulate = require("xlsx-populate");
const Path = require("path");
const { endianness } = require("os");
const { count } = require("console");

workingDirectry = Path.dirname(process.cwd());

filePath = process.argv[2];
if (!filePath) {
	console.error("パスを入力してください");
}

if (!(process.argv[3] && process.argv[4])) {
	console.error("列番号を入力してください");
}
itemNumberColumnId = process.argv[3];
statusColumnId = process.argv[4];

startSheet = 0;
if (process.argv[5]) {
	startSheet = process.argv[5];
	if (process.argv[6]) {
		endSheet = process.argv[6];
	}
}

var countItems = [];

const stringsDone = ["完了", "◯", "OK", "ok", "Ok"];

// Load an existing workbook
XlsxPopulate.fromFileAsync(filePath).then((workbook) => {
	const sheets = workbook.sheets();
	for (let sheetNum = startSheet; sheetNum < sheets.length; sheetNum++) {
		const sheet = sheets[sheetNum];
		const itemNumberColumn = sheet.column(itemNumberColumnId);
		const statusColumn = sheet.column(statusColumnId);
		var countItemNum = 0;
		var countDoneNum = 0;
		var countNull = 0;
		const COUNTEND = 10;
		var i = 1;
		while (countNull < COUNTEND) {
			if (itemNumberColumn.cell(i++).value()) {
				countItemNum++;
			} else {
				countNull++;
			}
		}
		countNull = 0;
		i = 1;
		while (countNull < COUNTEND) {
			if (statusColumn.cell(i).value()) {
				if (0 <= stringsDone.indexOf(statusColumn.cell(i).value())) {
					countDoneNum++;
				}
			} else {
				countNull++;
			}
            i++;
		}

		countItemObj = {};
		countItemObj.sheetName = sheet.name();
		countItemObj.countItemNum = countItemNum;
		countItemObj.countDoneNum = countDoneNum;
		countItems.push = countItemObj;
	}
	console.log(countItems);
});
