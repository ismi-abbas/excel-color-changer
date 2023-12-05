const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

// Function to process an XLSX file
async function processXLSXFile(filePath) {
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.readFile(filePath);
	// Loop through each worksheet in the XLSX file
	workbook.eachSheet(worksheet => {
		// Loop through all rows and columns in the worksheet
		worksheet.eachRow(row => {
			row.eachCell(cell => {
				if (
					cell.value === 'i-PIMS (PETRONAS Integrated Pipeline Integrity Assurance Solutions)' ||
					cell.value === 'IPIMS (PETRONAS Integrated Pipeline Integrity Assurance Solutions)'
				) {
					cell.value = 'i-PIMS (Integrated Pipeline Integrity Assurance Solutions)';
				}
				if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor && cell.fill.fgColor.argb === 'FF00B3BC') {
					// Change the background color to #20419A
					cell.fill = {
						type: 'pattern',
						pattern: 'solid',
						fgColor: { argb: 'FF20419A' } // Set the new color here
					};
				}
			});
		});
	});

	// Save the modified workbook
	await workbook.xlsx.writeFile(filePath);
}

// Get a list of all XLSX files in the current directory
const currentDir = __dirname + '/files';
let successList = [];
fs.readdirSync(currentDir).forEach(file => {
	const filePath = path.join(currentDir, file);

	// Check if the file is an XLSX file
	if (path.extname(filePath) === '.xlsx') {
		processXLSXFile(filePath)
			.then(() => {
				console.log(`Processed and saved ${filePath}`);
				const file = filePath.split('\\').pop().split('.')[0];
				successList.push(file);
				console.log(successList);
			})
			.catch(error => {
				console.error(`Error processing ${filePath}: ${error.message}`);
			});
	}
});
