const fs = require('fs').promises;

const xml2js = require('xml2js');
const ExcelJS = require('exceljs');

async function main(input, output) {
	var parser = new xml2js.Parser();
	const str = await fs.readFile(input);
	const res = await parser.parseStringPromise(str);

	// find the root key
	let root = undefined
	for (const [key, value] of Object.entries(res)) {
		const regex = /tracklist\d*/g;
		if (key.match(regex)) {
			root = value
			break
		}
	}

	if (!root) {
		console.error("root not found")
	}

	let arr = []

	// find the events in "event list"
	root.list.forEach(l => {
		l.obj.forEach(o => {
			if (o.$.class === "MAudioTrackEvent") {
				o.obj.forEach(obj => {
					if (obj.$.class === "MListNode") {
						obj.list.forEach(list => {
							if (list.$.name === "Events") {
								list.obj.forEach(evt => {
									// console.log(evt.$.class)
									let name = ""
									let start = 0
									let length = 0
									let end = 0
									if (evt.$.class === "MAudioEvent") {
										name = evt.string[0].$.value
										// console.log("audio evt", evt)
									} else if (evt.$.class === "MAudioPartEvent") {
										// console.log("audio part evt", evt)
										name = evt.obj[0].string[0].$.value
									}
									evt.float.forEach(f => {
										if (f.$.name === "Start") {
											start = Math.round(parseFloat(f.$.value) / 960)
										}
										if (f.$.name === "Length") {
											length = Math.round(parseFloat(f.$.value) / 960)
										}
									})

									end = start + length
									arr.push({ name, start, length: length, end })
									// console.log(name, start, length, end)
								})
							}
						})
					}
				})
			}
		})
	})

	// console.log(arr)
	arr.sort(compare)
	const workbook = new ExcelJS.Workbook();
	const ws = workbook.addWorksheet('1st Sheet', { properties: { tabColor: { argb: 'FFC0000' } } });
	ws.columns = [{ header: "Name", key: "name" }, { header: "Length", key: "length" }, { header: "Start", key: "start" }, { header: "End", key: "end" }]
	arr.forEach(item => {
		ws.addRow([item.name, format(item.length), format(item.start), format(item.end)])
	})
	// const row = workbook.worksheets[0].getRow(1);
	await workbook.xlsx.writeFile(output);
	// console.log(row)
}

function compare(a, b) {
	if (a.start < b.start) {
		return -1;
	}
	if (a.start > b.start) {
		return 1;
	}
	return 0;
}

function format(seconds) {
	const date = new Date(null);
	date.setSeconds(seconds); // specify value for SECONDS here
	return date.toISOString().slice(12, 19);
}

// main();
//
var { argv } = require("yargs")
	.scriptName("area")
	.usage("Usage: $0 -i input_file -o output_file")
	.example(
		"$0 -i input.xml -o output.xlsx",
		"transform xml into xlsx"
	)
	.option("i", {
		alias: "input",
		describe: "The input file name.",
		demandOption: "The input file name is required.",
		type: "string",
		nargs: 1,
	})
	.option("o", {
		alias: "output",
		describe: "The onput file name.",
		demandOption: "The onput file name is required.",
		type: "string",
		nargs: 1,
	})
	.describe("help", "Show help.") // Override --help usage message.
	.describe("version", "Show version number.") // Override --version usage message.

const { input, output } = argv;
main(input, output)

// console.log(`The area is ${input} ${output}`);
